"""
PPTX 병합 검증기 (v1.0)

Gamma PPTX 병합 시 발생하는 3대 구조적 결함을 사전·사후 검증합니다.

┌─────────────────────────────────────────────────────────────────────────┐
│  Issue #1: NotesSlide 미복사/미연결                                     │
│    - Donor의 notesSlide가 복사되지 않거나 번호 재지정이 안 된 경우       │
│    - Slide rels → notesSlide 참조가 잘못된 파일을 가리키는 경우          │
│                                                                         │
│  Issue #2: Donor Layout 미등록 (Orphaned Layouts)                       │
│    - slideLayout 파일이 존재하지만 slideMaster.xml.rels에 등록 안 됨     │
│    - slideMaster.xml의 <p:sldLayoutIdLst>에 누락                       │
│                                                                         │
│  Issue #3: NotesSlide → Slide 역참조 불일치                              │
│    - notesSlide .rels가 원본 donor의 slide 번호를 참조하는 경우          │
│    - 병합 후 renumber된 slide를 정확히 가리키는지 검증                   │
└─────────────────────────────────────────────────────────────────────────┘

사용법:
    from src.validators import PptxMergeValidator

    # 사후 검증 (병합 결과물)
    validator = PptxMergeValidator("merged.pptx")
    result = validator.validate()
    result.print_report()

    # 사전 검증 (병합 전 두 파일 분석)
    report = PptxMergeValidator.pre_merge_check("part1.pptx", "part2.pptx")
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

try:
    from ..utils.logger import get_logger
    logger = get_logger("pptx_merge_validator")
except Exception:
    # Standalone 실행 시 fallback
    import logging
    logger = logging.getLogger("pptx_merge_validator")


# ─── Data Classes ───────────────────────────────────────────────────

class Severity(Enum):
    """검증 결과 심각도"""
    PASS = "PASS"
    WARN = "WARN"
    FAIL = "FAIL"


@dataclass
class ValidationIssue:
    """단일 검증 항목 결과"""
    category: str          # "notes_slide" | "layout_orphan" | "back_reference" | ...
    severity: Severity
    message: str
    detail: str = ""       # 상세 정보 (영향받는 파일 목록 등)

    def __str__(self):
        icon = {"PASS": "[PASS]", "WARN": "[WARN]", "FAIL": "[FAIL]"}[self.severity.value]
        s = f"  {icon} {self.message}"
        if self.detail:
            s += f"\n         {self.detail}"
        return s


@dataclass
class ValidationResult:
    """전체 검증 결과"""
    pptx_path: str
    issues: List[ValidationIssue] = field(default_factory=list)
    stats: Dict[str, int] = field(default_factory=dict)

    @property
    def pass_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.PASS)

    @property
    def warn_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.WARN)

    @property
    def fail_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.FAIL)

    @property
    def grade(self) -> str:
        if self.fail_count > 0:
            return "FAIL"
        if self.warn_count > 2:
            return "WARN"
        return "OK"

    @property
    def is_valid(self) -> bool:
        return self.fail_count == 0

    def print_report(self):
        """콘솔 리포트 출력"""
        sep = "=" * 64
        print(f"\n{sep}")
        print(f"  PPTX MERGE VALIDATION: {self.pptx_path}")
        print(sep)

        if self.stats:
            print(f"  Slides: {self.stats.get('slides', '?')}"
                  f"  Layouts: {self.stats.get('layouts', '?')}"
                  f"  Notes: {self.stats.get('notes', '?')}"
                  f"  Media: {self.stats.get('media', '?')}")
            print()

        for issue in self.issues:
            print(str(issue))

        print(f"\n{sep}")
        print(f"  PASS: {self.pass_count}  WARN: {self.warn_count}  FAIL: {self.fail_count}")
        print(f"  Grade: {self.grade}")
        print(sep)

    def to_dict(self) -> dict:
        """JSON 직렬화용"""
        return {
            "pptx_path": self.pptx_path,
            "grade": self.grade,
            "is_valid": self.is_valid,
            "stats": self.stats,
            "summary": {
                "pass": self.pass_count,
                "warn": self.warn_count,
                "fail": self.fail_count,
            },
            "issues": [
                {
                    "category": i.category,
                    "severity": i.severity.value,
                    "message": i.message,
                    "detail": i.detail,
                }
                for i in self.issues
            ],
        }


# ─── Helper: PPTX 내부 파일 목록 추출 ───────────────────────────────

def _part_numbers(names: set, pattern: str) -> List[int]:
    """ZIP 내부 파일명에서 번호 추출 (예: slide1.xml → 1)"""
    nums = []
    for n in names:
        m = re.match(pattern, n)
        if m:
            nums.append(int(m.group(1)))
    return sorted(nums)


def _extract_targets(rels_xml: str) -> List[Tuple[str, str, str]]:
    """rels XML에서 (Id, Type-short, Target) 추출"""
    results = []
    for m in re.finditer(
        r'<Relationship\s+[^>]*Id="([^"]+)"[^>]*Type="[^"]*?/([^/"]+)"[^>]*Target="([^"]+)"',
        rels_xml,
    ):
        results.append((m.group(1), m.group(2), m.group(3)))
    # Alternative attribute order: Target before Type
    for m in re.finditer(
        r'<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*Type="[^"]*?/([^/"]+)"',
        rels_xml,
    ):
        rid = m.group(1)
        if not any(r[0] == rid for r in results):
            results.append((rid, m.group(3), m.group(2)))
    return results


# ─── Main Validator ──────────────────────────────────────────────────

class PptxMergeValidator:
    """
    Gamma PPTX 병합 검증기

    3대 핵심 검증 + 7개 추가 검증 = 총 10개 검증 항목:

    [Critical — 병합 오류 3종]
    1. notes_slide_integrity    : NotesSlide 복사 및 연결 정합성
    2. layout_registration      : Donor Layout → slideMaster 등록 여부
    3. notes_back_reference     : NotesSlide → Slide 역참조 일치

    [Standard — 일반 OOXML 무결성]
    4. zip_integrity            : ZIP 파일 손상 여부
    5. content_types            : [Content_Types].xml 완전성
    6. presentation_refs        : presentation.xml sldId 수 일치
    7. presentation_rels        : presentation.xml.rels 슬라이드 수 일치
    8. slide_rels_exist         : 모든 slide .rels 파일 존재
    9. media_references         : 미디어 파일 참조 유효성
    10. id_uniqueness           : rId / sldId 중복 없음
    """

    def __init__(self, pptx_path: str | Path):
        self.pptx_path = Path(pptx_path)
        self._zf: Optional[zipfile.ZipFile] = None
        self._names: Set[str] = set()

    # ── Public API ─────────────────────────────────────────────

    def validate(self) -> ValidationResult:
        """전체 검증 실행 → ValidationResult 반환"""
        result = ValidationResult(pptx_path=str(self.pptx_path))

        if not self.pptx_path.exists():
            result.issues.append(ValidationIssue(
                "file", Severity.FAIL,
                f"파일을 찾을 수 없습니다: {self.pptx_path}",
            ))
            return result

        try:
            self._zf = zipfile.ZipFile(self.pptx_path, "r")
            self._names = set(self._zf.namelist())
        except zipfile.BadZipFile as e:
            result.issues.append(ValidationIssue(
                "zip_integrity", Severity.FAIL,
                f"ZIP 파일 손상: {e}",
            ))
            return result

        # 통계 수집
        result.stats = self._collect_stats()

        # 10개 검증 항목 순차 실행
        result.issues.extend(self._check_zip_integrity())
        result.issues.extend(self._check_content_types())
        result.issues.extend(self._check_presentation_refs())
        result.issues.extend(self._check_presentation_rels())
        result.issues.extend(self._check_slide_rels_exist())
        result.issues.extend(self._check_notes_slide_integrity())       # Critical #1
        result.issues.extend(self._check_layout_registration())         # Critical #2
        result.issues.extend(self._check_notes_back_reference())        # Critical #3
        result.issues.extend(self._check_media_references())
        result.issues.extend(self._check_id_uniqueness())

        self._zf.close()
        return result

    @staticmethod
    def pre_merge_check(
        part1_path: str | Path, part2_path: str | Path
    ) -> Dict[str, any]:
        """
        병합 전 사전 진단 — 두 PPTX의 구조를 비교하여 잠재 충돌 예측.

        Returns:
            {
                "compatible": bool,
                "warnings": List[str],
                "part1": { slides, layouts, notes, media, ... },
                "part2": { slides, layouts, notes, media, ... },
            }
        """
        warnings = []

        def _analyze(path: Path) -> dict:
            zf = zipfile.ZipFile(path, "r")
            names = set(zf.namelist())
            info = {
                "slides": len(_part_numbers(names, r"ppt/slides/slide(\d+)\.xml$")),
                "layouts": len(_part_numbers(names, r"ppt/slideLayouts/slideLayout(\d+)\.xml$")),
                "notes": len(_part_numbers(names, r"ppt/notesSlides/notesSlide(\d+)\.xml$")),
                "media": len([n for n in names if "/media/" in n]),
                "masters": len([n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]),
                "themes": len([n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n)]),
            }

            # slideMaster1 layout 등록 수
            master_rels_name = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
            if master_rels_name in names:
                rels_xml = zf.read(master_rels_name).decode("utf-8")
                info["master_layout_refs"] = len(re.findall(r"slideLayout", rels_xml)) // 3
            else:
                info["master_layout_refs"] = 0

            # sldLayoutIdLst 항목 수
            master_xml_name = "ppt/slideMasters/slideMaster1.xml"
            if master_xml_name in names:
                master_xml = zf.read(master_xml_name).decode("utf-8")
                info["sld_layout_id_count"] = len(re.findall(r"sldLayoutId\s+id=", master_xml))
            else:
                info["sld_layout_id_count"] = 0

            # slide→notesSlide 매핑 확인
            slides_with_notes = 0
            for sn in _part_numbers(names, r"ppt/slides/slide(\d+)\.xml$"):
                rels_name = f"ppt/slides/_rels/slide{sn}.xml.rels"
                if rels_name in names:
                    rels_xml = zf.read(rels_name).decode("utf-8")
                    if "notesSlide" in rels_xml:
                        slides_with_notes += 1
            info["slides_with_notes"] = slides_with_notes

            zf.close()
            return info

        p1 = _analyze(Path(part1_path))
        p2 = _analyze(Path(part2_path))

        # 잠재 충돌 분석
        if p2["notes"] > 0:
            warnings.append(
                f"Donor에 notesSlide {p2['notes']}개 존재 → 복사 + 번호 재지정 필요"
            )
        if p2["layouts"] > 0:
            warnings.append(
                f"Donor에 slideLayout {p2['layouts']}개 존재 → "
                f"slideMaster rels + sldLayoutIdLst 등록 필요"
            )
        if p2["slides_with_notes"] > 0 and p2["notes"] > 0:
            warnings.append(
                f"Donor slide {p2['slides_with_notes']}개가 notesSlide를 참조 → "
                f"rels의 notesSlide 번호 재지정 필요"
            )
        if p2["masters"] > 1 or p1["masters"] > 1:
            warnings.append(
                f"다중 slideMaster 감지 (Part1={p1['masters']}, Part2={p2['masters']}) → "
                f"테마 충돌 주의"
            )

        return {
            "compatible": len(warnings) == 0 or all("필요" in w for w in warnings),
            "warnings": warnings,
            "part1": p1,
            "part2": p2,
        }

    # ── Internal Checks ────────────────────────────────────────

    def _read(self, name: str) -> str:
        """ZIP 내부 파일 읽기"""
        return self._zf.read(name).decode("utf-8")

    def _collect_stats(self) -> dict:
        return {
            "slides": len(_part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$")),
            "layouts": len(_part_numbers(self._names, r"ppt/slideLayouts/slideLayout(\d+)\.xml$")),
            "notes": len(_part_numbers(self._names, r"ppt/notesSlides/notesSlide(\d+)\.xml$")),
            "media": len([n for n in self._names if "/media/" in n]),
            "total_files": len(self._names),
        }

    # ── [Check 1] ZIP Integrity ────────────────────────────────

    def _check_zip_integrity(self) -> List[ValidationIssue]:
        bad = self._zf.testzip()
        if bad:
            return [ValidationIssue(
                "zip_integrity", Severity.FAIL,
                f"ZIP 내부 파일 손상: {bad}",
            )]
        return [ValidationIssue(
            "zip_integrity", Severity.PASS,
            "ZIP 무결성 정상",
        )]

    # ── [Check 2] Content_Types ────────────────────────────────

    def _check_content_types(self) -> List[ValidationIssue]:
        issues = []
        ct_xml = self._read("[Content_Types].xml")

        missing = []
        # Slides
        for num in _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"):
            pn = f"/ppt/slides/slide{num}.xml"
            if pn not in ct_xml:
                missing.append(pn)
        # Layouts
        for num in _part_numbers(self._names, r"ppt/slideLayouts/slideLayout(\d+)\.xml$"):
            pn = f"/ppt/slideLayouts/slideLayout{num}.xml"
            if pn not in ct_xml:
                missing.append(pn)
        # NotesSlides
        for num in _part_numbers(self._names, r"ppt/notesSlides/notesSlide(\d+)\.xml$"):
            pn = f"/ppt/notesSlides/notesSlide{num}.xml"
            if pn not in ct_xml:
                missing.append(pn)

        if missing:
            issues.append(ValidationIssue(
                "content_types", Severity.FAIL,
                f"[Content_Types].xml 누락 {len(missing)}건",
                detail=", ".join(missing[:10]) + ("..." if len(missing) > 10 else ""),
            ))
        else:
            issues.append(ValidationIssue(
                "content_types", Severity.PASS,
                "[Content_Types].xml 완전 - 모든 파트 등록됨",
            ))

        # 중복 체크
        overrides = re.findall(r'PartName="([^"]+)"', ct_xml)
        seen = set()
        dups = []
        for o in overrides:
            if o in seen:
                dups.append(o)
            seen.add(o)
        if dups:
            issues.append(ValidationIssue(
                "content_types", Severity.WARN,
                f"[Content_Types].xml 중복 항목 {len(dups)}건",
                detail=", ".join(dups[:5]),
            ))

        return issues

    # ── [Check 3] presentation.xml sldId Count ─────────────────

    def _check_presentation_refs(self) -> List[ValidationIssue]:
        pres_xml = self._read("ppt/presentation.xml")
        sld_count = len(re.findall(r"sldId\s+id=", pres_xml))
        slide_files = len(_part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"))

        if sld_count == slide_files:
            return [ValidationIssue(
                "presentation_refs", Severity.PASS,
                f"presentation.xml sldId {sld_count}개 = 슬라이드 파일 수 일치",
            )]
        return [ValidationIssue(
            "presentation_refs", Severity.FAIL,
            f"presentation.xml sldId {sld_count}개 ≠ 슬라이드 파일 {slide_files}개",
        )]

    # ── [Check 4] presentation.xml.rels Slide Count ────────────

    def _check_presentation_rels(self) -> List[ValidationIssue]:
        pres_rels = self._read("ppt/_rels/presentation.xml.rels")
        slide_rels = len(re.findall(r'relationships/slide"', pres_rels))
        slide_files = len(_part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"))

        if slide_rels == slide_files:
            return [ValidationIssue(
                "presentation_rels", Severity.PASS,
                f"presentation.xml.rels 슬라이드 관계 {slide_rels}개 일치",
            )]
        return [ValidationIssue(
            "presentation_rels", Severity.FAIL,
            f"presentation.xml.rels {slide_rels}개 ≠ 슬라이드 파일 {slide_files}개",
        )]

    # ── [Check 5] Slide .rels Existence ────────────────────────

    def _check_slide_rels_exist(self) -> List[ValidationIssue]:
        missing = []
        for num in _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"):
            rels_name = f"ppt/slides/_rels/slide{num}.xml.rels"
            if rels_name not in self._names:
                missing.append(f"slide{num}")

        if missing:
            return [ValidationIssue(
                "slide_rels_exist", Severity.FAIL,
                f"슬라이드 .rels 파일 누락 {len(missing)}건",
                detail=", ".join(missing[:10]),
            )]
        total = len(_part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"))
        return [ValidationIssue(
            "slide_rels_exist", Severity.PASS,
            f"슬라이드 .rels 파일 {total}개 모두 존재",
        )]

    # ── [Critical #1] NotesSlide Integrity ─────────────────────

    def _check_notes_slide_integrity(self) -> List[ValidationIssue]:
        """
        Issue #1: NotesSlide 복사 및 연결 검증

        검증 항목:
        a) slide .rels에서 참조하는 notesSlide 파일이 실제 존재하는가?
        b) 서로 다른 slide가 동일한 notesSlide를 참조하지 않는가?
           (병합 시 번호 재지정 누락 → 여러 slide가 같은 note를 공유)
        c) 모든 notesSlide가 최소 1개 slide에 의해 참조되는가?
           (고아 notesSlide 감지)
        """
        issues = []
        slide_nums = _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$")

        # a) 참조 유효성 + b) 중복 참조
        broken_refs = []
        notes_ref_map: Dict[str, List[int]] = {}  # notesSlide파일 → [참조하는 slide 번호들]

        for sn in slide_nums:
            rels_name = f"ppt/slides/_rels/slide{sn}.xml.rels"
            if rels_name not in self._names:
                continue
            rels_xml = self._read(rels_name)

            for m in re.finditer(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"', rels_xml):
                notes_file = m.group(1)
                full_path = f"ppt/notesSlides/{notes_file}"

                # a) 존재 확인
                if full_path not in self._names:
                    broken_refs.append(f"slide{sn} → {notes_file}")

                # b) 중복 참조 추적
                notes_ref_map.setdefault(notes_file, []).append(sn)

        if broken_refs:
            issues.append(ValidationIssue(
                "notes_slide_integrity", Severity.FAIL,
                f"[Issue #1a] 존재하지 않는 notesSlide 참조 {len(broken_refs)}건",
                detail="; ".join(broken_refs[:5]) + ("..." if len(broken_refs) > 5 else ""),
            ))
        else:
            issues.append(ValidationIssue(
                "notes_slide_integrity", Severity.PASS,
                "[Issue #1a] 모든 slide → notesSlide 참조 유효",
            ))

        # b) 중복 참조 (번호 재지정 누락 징후)
        dup_refs = {f: slides for f, slides in notes_ref_map.items() if len(slides) > 1}
        if dup_refs:
            detail_parts = [f"{f}: slide{','.join(str(s) for s in sl)}" for f, sl in list(dup_refs.items())[:5]]
            issues.append(ValidationIssue(
                "notes_slide_integrity", Severity.FAIL,
                f"[Issue #1b] notesSlide 중복 참조 {len(dup_refs)}건 "
                f"(병합 시 번호 재지정 누락 가능성)",
                detail="; ".join(detail_parts),
            ))
        else:
            issues.append(ValidationIssue(
                "notes_slide_integrity", Severity.PASS,
                "[Issue #1b] notesSlide 중복 참조 없음 (1:1 매핑 정상)",
            ))

        # c) 고아 notesSlide
        all_notes = set(
            f"notesSlide{n}.xml"
            for n in _part_numbers(self._names, r"ppt/notesSlides/notesSlide(\d+)\.xml$")
        )
        referenced_notes = set(notes_ref_map.keys())
        orphans = all_notes - referenced_notes
        if orphans:
            issues.append(ValidationIssue(
                "notes_slide_integrity", Severity.WARN,
                f"[Issue #1c] 참조되지 않는 고아 notesSlide {len(orphans)}개",
                detail=", ".join(sorted(orphans)[:5]),
            ))

        return issues

    # ── [Critical #2] Layout Registration ──────────────────────

    def _check_layout_registration(self) -> List[ValidationIssue]:
        """
        Issue #2: Donor Layout → slideMaster 등록 검증

        검증 항목:
        a) 모든 slideLayout 파일이 slideMaster1.xml.rels에 등록되어 있는가?
        b) slideMaster1.xml의 <p:sldLayoutIdLst>에 모든 layout이 포함되어 있는가?
        c) slide가 참조하는 layout이 모두 실제 존재하는가?
        """
        issues = []

        # 존재하는 layout 파일 목록
        all_layouts = set(
            f"slideLayout{n}.xml"
            for n in _part_numbers(self._names, r"ppt/slideLayouts/slideLayout(\d+)\.xml$")
        )

        # a) slideMaster1.xml.rels 등록 확인
        master_rels_name = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
        if master_rels_name in self._names:
            master_rels = self._read(master_rels_name)
            registered_in_rels = set(
                re.findall(r'Target="\.\./slideLayouts/(slideLayout\d+\.xml)"', master_rels)
            )
            orphaned_rels = all_layouts - registered_in_rels
            if orphaned_rels:
                issues.append(ValidationIssue(
                    "layout_registration", Severity.FAIL,
                    f"[Issue #2a] slideMaster1.xml.rels 미등록 layout {len(orphaned_rels)}개",
                    detail=", ".join(sorted(orphaned_rels)[:10]),
                ))
            else:
                issues.append(ValidationIssue(
                    "layout_registration", Severity.PASS,
                    f"[Issue #2a] 모든 {len(all_layouts)} layout이 slideMaster1.xml.rels에 등록됨",
                ))
        else:
            issues.append(ValidationIssue(
                "layout_registration", Severity.FAIL,
                "[Issue #2a] slideMaster1.xml.rels 파일을 찾을 수 없음",
            ))

        # b) <p:sldLayoutIdLst> 확인
        master_xml_name = "ppt/slideMasters/slideMaster1.xml"
        if master_xml_name in self._names:
            master_xml = self._read(master_xml_name)
            id_list_count = len(re.findall(r"sldLayoutId\s+id=", master_xml))
            if id_list_count == len(all_layouts):
                issues.append(ValidationIssue(
                    "layout_registration", Severity.PASS,
                    f"[Issue #2b] <p:sldLayoutIdLst> 항목 {id_list_count}개 = layout 파일 수 일치",
                ))
            elif id_list_count < len(all_layouts):
                diff = len(all_layouts) - id_list_count
                issues.append(ValidationIssue(
                    "layout_registration", Severity.FAIL,
                    f"[Issue #2b] <p:sldLayoutIdLst> {id_list_count}개 < layout 파일 {len(all_layouts)}개"
                    f" (누락 {diff}개)",
                ))
            else:
                issues.append(ValidationIssue(
                    "layout_registration", Severity.WARN,
                    f"[Issue #2b] <p:sldLayoutIdLst> {id_list_count}개 > layout 파일 {len(all_layouts)}개"
                    f" (여분 {id_list_count - len(all_layouts)}개)",
                ))

        # c) slide → layout 참조 유효성
        broken_layouts = []
        for sn in _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"):
            rels_name = f"ppt/slides/_rels/slide{sn}.xml.rels"
            if rels_name not in self._names:
                continue
            rels_xml = self._read(rels_name)
            for m in re.finditer(r'Target="\.\./slideLayouts/(slideLayout\d+\.xml)"', rels_xml):
                layout_file = m.group(1)
                full_path = f"ppt/slideLayouts/{layout_file}"
                if full_path not in self._names:
                    broken_layouts.append(f"slide{sn} → {layout_file}")

        if broken_layouts:
            issues.append(ValidationIssue(
                "layout_registration", Severity.FAIL,
                f"[Issue #2c] 존재하지 않는 layout 참조 {len(broken_layouts)}건",
                detail="; ".join(broken_layouts[:5]),
            ))
        else:
            issues.append(ValidationIssue(
                "layout_registration", Severity.PASS,
                "[Issue #2c] 모든 slide → layout 참조 유효",
            ))

        return issues

    # ── [Critical #3] NotesSlide → Slide Back-Reference ────────

    def _check_notes_back_reference(self) -> List[ValidationIssue]:
        """
        Issue #3: notesSlide .rels → slide 역참조 검증

        검증 항목:
        a) 각 notesSlide의 .rels가 존재하는 slide를 참조하는가?
        b) slide N의 notesSlide가 역으로 slide N을 참조하는가? (일관성)
        """
        issues = []
        slide_nums = _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$")
        notes_nums = _part_numbers(self._names, r"ppt/notesSlides/notesSlide(\d+)\.xml$")

        # slide → notes 정방향 매핑 구축
        slide_to_notes: Dict[int, str] = {}
        for sn in slide_nums:
            rels_name = f"ppt/slides/_rels/slide{sn}.xml.rels"
            if rels_name not in self._names:
                continue
            rels_xml = self._read(rels_name)
            m = re.search(r'Target="\.\./notesSlides/(notesSlide(\d+)\.xml)"', rels_xml)
            if m:
                slide_to_notes[sn] = m.group(1)

        # a) notesSlide .rels → slide 존재 확인
        broken_back = []
        notes_to_slide: Dict[int, int] = {}  # notes번호 → 참조하는 slide번호
        for nn in notes_nums:
            rels_name = f"ppt/notesSlides/_rels/notesSlide{nn}.xml.rels"
            if rels_name not in self._names:
                continue
            rels_xml = self._read(rels_name)
            m = re.search(r'Target="\.\./slides/slide(\d+)\.xml"', rels_xml)
            if m:
                ref_slide = int(m.group(1))
                notes_to_slide[nn] = ref_slide
                full_path = f"ppt/slides/slide{ref_slide}.xml"
                if full_path not in self._names:
                    broken_back.append(f"notesSlide{nn} → slide{ref_slide}")

        if broken_back:
            issues.append(ValidationIssue(
                "notes_back_reference", Severity.FAIL,
                f"[Issue #3a] notesSlide → 존재하지 않는 slide 참조 {len(broken_back)}건",
                detail="; ".join(broken_back[:5]) + ("..." if len(broken_back) > 5 else ""),
            ))
        else:
            issues.append(ValidationIssue(
                "notes_back_reference", Severity.PASS,
                f"[Issue #3a] 모든 notesSlide → slide 역참조 유효 ({len(notes_to_slide)}건)",
            ))

        # b) 양방향 일관성: slide N → notesSlide M이면, notesSlide M → slide N인가?
        mismatches = []
        for sn, notes_file in slide_to_notes.items():
            nn = int(re.search(r"(\d+)", notes_file).group(1))
            if nn in notes_to_slide:
                back_slide = notes_to_slide[nn]
                if back_slide != sn:
                    mismatches.append(
                        f"slide{sn} → notesSlide{nn}, but notesSlide{nn} → slide{back_slide}"
                    )

        if mismatches:
            issues.append(ValidationIssue(
                "notes_back_reference", Severity.FAIL,
                f"[Issue #3b] 양방향 참조 불일치 {len(mismatches)}건 "
                f"(병합 시 역참조 번호 미갱신 가능성)",
                detail="; ".join(mismatches[:5]),
            ))
        else:
            issues.append(ValidationIssue(
                "notes_back_reference", Severity.PASS,
                "[Issue #3b] 양방향 참조 일관성 정상 (slide↔notesSlide 1:1 매칭)",
            ))

        return issues

    # ── [Check 9] Media References ─────────────────────────────

    def _check_media_references(self) -> List[ValidationIssue]:
        broken = []
        for sn in _part_numbers(self._names, r"ppt/slides/slide(\d+)\.xml$"):
            rels_name = f"ppt/slides/_rels/slide{sn}.xml.rels"
            if rels_name not in self._names:
                continue
            rels_xml = self._read(rels_name)
            for m in re.finditer(r'Target="\.\./media/([^"]+)"', rels_xml):
                media_file = m.group(1)
                full_path = f"ppt/media/{media_file}"
                if full_path not in self._names:
                    broken.append(f"slide{sn}: {media_file}")

        if broken:
            return [ValidationIssue(
                "media_references", Severity.FAIL,
                f"존재하지 않는 미디어 참조 {len(broken)}건",
                detail="; ".join(broken[:5]) + ("..." if len(broken) > 5 else ""),
            )]
        return [ValidationIssue(
            "media_references", Severity.PASS,
            "모든 슬라이드 미디어 참조 유효",
        )]

    # ── [Check 10] ID Uniqueness ───────────────────────────────

    def _check_id_uniqueness(self) -> List[ValidationIssue]:
        issues = []

        # presentation.xml.rels rId 중복
        pres_rels = self._read("ppt/_rels/presentation.xml.rels")
        rids = re.findall(r'Id="(rId\d+)"', pres_rels)
        dup_rids = len(rids) - len(set(rids))
        if dup_rids > 0:
            issues.append(ValidationIssue(
                "id_uniqueness", Severity.FAIL,
                f"presentation.xml.rels rId 중복 {dup_rids}건",
            ))
        else:
            issues.append(ValidationIssue(
                "id_uniqueness", Severity.PASS,
                f"presentation.xml.rels rId {len(rids)}개 고유",
            ))

        # presentation.xml sldId 중복
        pres_xml = self._read("ppt/presentation.xml")
        sld_ids = re.findall(r'sldId\s+id="(\d+)"', pres_xml)
        dup_slds = len(sld_ids) - len(set(sld_ids))
        if dup_slds > 0:
            issues.append(ValidationIssue(
                "id_uniqueness", Severity.FAIL,
                f"presentation.xml sldId 중복 {dup_slds}건",
            ))
        else:
            issues.append(ValidationIssue(
                "id_uniqueness", Severity.PASS,
                f"presentation.xml sldId {len(sld_ids)}개 고유",
            ))

        return issues


# ─── Standalone CLI ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import io

    # Windows cp949 인코딩 오류 방지
    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    if len(sys.argv) < 2:
        print("Usage: python pptx_merge_validator.py <merged.pptx> [part1.pptx part2.pptx]")
        sys.exit(1)

    target = sys.argv[1]

    # 사전 진단 모드
    if len(sys.argv) >= 4:
        print("\n[PRE-MERGE CHECK]")
        report = PptxMergeValidator.pre_merge_check(sys.argv[2], sys.argv[3])
        print(f"  Part 1: {report['part1']}")
        print(f"  Part 2: {report['part2']}")
        if report["warnings"]:
            for w in report["warnings"]:
                print(f"  [WARN] {w}")
        else:
            print("  [OK] 잠재 충돌 없음")
        print()

    # 사후 검증
    validator = PptxMergeValidator(target)
    result = validator.validate()
    result.print_report()

    sys.exit(0 if result.is_valid else 1)
