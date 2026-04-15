"""PPTX 내부 폰트명 일괄 치환기.

- OOXML 내 typeface 속성만 교체 (좌표/도형/이미지 불변)
- 슬라이드/레이아웃/마스터/테마 모두 커버
"""
from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path


TARGET_XML_PREFIXES = (
    "ppt/slides/",
    "ppt/slideLayouts/",
    "ppt/slideMasters/",
    "ppt/theme/",
    "ppt/notesSlides/",
    "ppt/notesMasters/",
    "ppt/handoutMasters/",
    "ppt/charts/",
    "ppt/diagrams/",
)


@dataclass
class ReplaceStats:
    files_scanned: int = 0
    files_modified: int = 0
    total_replacements: int = 0
    per_file: dict[str, int] = field(default_factory=dict)


def replace_fonts(
    src: Path,
    dst: Path,
    mapping: dict[str, str],
) -> ReplaceStats:
    """src의 PPTX 내부 XML에서 폰트명을 mapping에 따라 치환 → dst."""
    stats = ReplaceStats()

    # typeface="XXX" 또는 latin/ea/cs typeface 패턴 모두 커버
    patterns: list[tuple[re.Pattern, str]] = []
    for old, new in mapping.items():
        old_esc = re.escape(old)
        # typeface="Pretendard" (속성 전체 정확 매칭)
        patterns.append(
            (re.compile(rf'typeface="{old_esc}"'), f'typeface="{new}"')
        )
        # 폰트 세트에 등장하는 경우 (<latin typeface="...") — 이미 위에서 커버됨

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9
    ) as zout:
        for entry in zin.infolist():
            name = entry.filename
            data = zin.read(name)

            if name.endswith(".xml") and any(
                name.startswith(p) for p in TARGET_XML_PREFIXES
            ):
                stats.files_scanned += 1
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    zout.writestr(entry, data)
                    continue

                hits = 0
                for pat, repl in patterns:
                    text, n = pat.subn(repl, text)
                    hits += n

                if hits > 0:
                    stats.files_modified += 1
                    stats.total_replacements += hits
                    stats.per_file[name] = hits
                    data = text.encode("utf-8")

            zout.writestr(entry, data, compress_type=zipfile.ZIP_DEFLATED)

    return stats


def print_stats(stats: ReplaceStats) -> None:
    print(f"\n{'='*60}")
    print("폰트 치환 완료")
    print(f"{'='*60}")
    print(f"스캔한 XML 파일 : {stats.files_scanned}")
    print(f"수정한 파일     : {stats.files_modified}")
    print(f"총 치환 횟수    : {stats.total_replacements}")
    if stats.per_file:
        top = sorted(stats.per_file.items(), key=lambda x: -x[1])[:10]
        print("상위 10개 파일:")
        for name, n in top:
            print(f"  {n:>4}회  {name}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    import sys

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    mapping = {"Pretendard": "Noto Sans KR"}
    stats = replace_fonts(src, dst, mapping)
    print_stats(stats)
    print(f"출력: {dst}")
