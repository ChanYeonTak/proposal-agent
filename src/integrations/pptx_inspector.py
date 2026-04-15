"""PPTX 내부 분석기 — Google Slides 업로드 전 최적화 전략 수립용."""
from __future__ import annotations

import hashlib
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from PIL import Image


@dataclass
class ImageInfo:
    name: str
    size: int
    format: str
    width: int = 0
    height: int = 0
    sha1: str = ""


@dataclass
class InspectionReport:
    path: Path
    total_mb: float
    slide_count: int
    images: list[ImageInfo] = field(default_factory=list)
    fonts: list[tuple[str, int]] = field(default_factory=list)
    embeddings: list[tuple[str, int]] = field(default_factory=list)
    xml_bytes: int = 0
    media_bytes: int = 0
    duplicate_groups: dict[str, list[str]] = field(default_factory=dict)
    emf_wmf_count: int = 0

    def format_counter(self) -> Counter:
        return Counter(i.format for i in self.images)

    def oversized_images(self, max_dim: int = 2400) -> list[ImageInfo]:
        return [i for i in self.images if max(i.width, i.height) > max_dim]

    def duplicate_bytes(self) -> int:
        total = 0
        for sha, names in self.duplicate_groups.items():
            if len(names) > 1:
                sizes = [img.size for img in self.images if img.sha1 == sha]
                if sizes:
                    # 중복분(1개만 남기고 나머지)의 크기 합
                    total += sum(sizes) - max(sizes)
        return total


def inspect(pptx_path: Path) -> InspectionReport:
    rep = InspectionReport(
        path=pptx_path,
        total_mb=pptx_path.stat().st_size / 1e6,
        slide_count=0,
    )

    hash_map: dict[str, list[str]] = defaultdict(list)

    with zipfile.ZipFile(pptx_path, "r") as z:
        for info in z.infolist():
            name = info.filename
            lname = name.lower()

            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                rep.slide_count += 1

            if name.endswith(".xml") or name.endswith(".rels"):
                rep.xml_bytes += info.file_size

            if name.startswith("ppt/media/"):
                rep.media_bytes += info.file_size
                data = z.read(name)
                sha = hashlib.sha1(data).hexdigest()
                hash_map[sha].append(name)

                ext = lname.rsplit(".", 1)[-1] if "." in lname else "?"
                img_info = ImageInfo(
                    name=name, size=len(data), format=ext, sha1=sha
                )
                if ext in {"emf", "wmf"}:
                    rep.emf_wmf_count += 1
                else:
                    try:
                        with Image.open(BytesIO(data)) as im:
                            img_info.width, img_info.height = im.size
                    except Exception:
                        pass
                rep.images.append(img_info)

            if name.startswith("ppt/fonts/") or name.startswith(
                "ppt/embeddings/fonts/"
            ):
                rep.fonts.append((name, info.file_size))

            if name.startswith("ppt/embeddings/") and not name.startswith(
                "ppt/embeddings/fonts/"
            ):
                rep.embeddings.append((name, info.file_size))

    rep.duplicate_groups = {k: v for k, v in hash_map.items() if len(v) > 1}
    return rep


def print_report(rep: InspectionReport) -> None:
    print(f"\n{'='*70}")
    print(f"PPTX 분석 리포트: {rep.path.name}")
    print(f"{'='*70}")
    print(f"전체 크기     : {rep.total_mb:>8.1f} MB")
    print(f"슬라이드 수   : {rep.slide_count}")
    print(f"XML 용량      : {rep.xml_bytes/1e6:>8.1f} MB ({rep.xml_bytes/(rep.total_mb*1e6)*100:.1f}%)")
    print(f"미디어 용량   : {rep.media_bytes/1e6:>8.1f} MB ({rep.media_bytes/(rep.total_mb*1e6)*100:.1f}%)")
    print(f"이미지 개수   : {len(rep.images)}")

    fmt = rep.format_counter()
    print(f"포맷 분포     : {dict(fmt)}")

    oversized = rep.oversized_images(2400)
    print(f"2400px 초과   : {len(oversized)}개")
    if oversized[:5]:
        print("  상위 5개:")
        for i in sorted(oversized, key=lambda x: -x.size)[:5]:
            print(f"    {i.name} {i.width}x{i.height} {i.size/1e6:.1f}MB")

    print(f"EMF/WMF 벡터  : {rep.emf_wmf_count}개  (Google Slides에서 깨질 위험)")

    font_total = sum(s for _, s in rep.fonts)
    print(f"임베드 폰트   : {len(rep.fonts)}개, {font_total/1e6:.1f} MB")

    embed_total = sum(s for _, s in rep.embeddings)
    print(f"임베드 객체   : {len(rep.embeddings)}개, {embed_total/1e6:.1f} MB")

    dup_bytes = rep.duplicate_bytes()
    print(f"중복 이미지   : {len(rep.duplicate_groups)}그룹, 절감 가능 {dup_bytes/1e6:.1f} MB")

    print(f"\n최적화 시뮬레이션 (예상치):")
    saved_oversize = sum(
        max(0, i.size - i.size * (2400 / max(i.width, i.height)) ** 2)
        for i in oversized if i.width and i.height
    )
    projected = rep.total_mb - (font_total + dup_bytes + saved_oversize) / 1e6
    print(f"  폰트 strip      : -{font_total/1e6:.1f} MB")
    print(f"  dedupe          : -{dup_bytes/1e6:.1f} MB")
    print(f"  이미지 다운샘플 : -{saved_oversize/1e6:.1f} MB (추정)")
    print(f"  예상 결과       : {projected:.1f} MB")
    print(f"  100MB 목표 달성 : {'OK' if projected < 100 else 'FAIL - 추가 최적화 필요'}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    import sys
    path = Path(sys.argv[1])
    print_report(inspect(path))
