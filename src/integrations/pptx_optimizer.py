"""PPTX 시각적 무손실 최적화기 — Google Slides 업로드 대응.

- 이미지: 1920~2400px 다운샘플 + 고품질 재인코딩
- 임베드 폰트/메타데이터 제거
- 슬라이드 XML/도형 좌표 절대 불변
"""
from __future__ import annotations

import shutil
import subprocess
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from PIL import Image


@dataclass
class OptimizeStats:
    original_mb: float = 0.0
    optimized_mb: float = 0.0
    images_processed: int = 0
    images_downsampled: int = 0
    png_recompressed: int = 0
    jpeg_recompressed: int = 0
    fonts_stripped: int = 0
    bytes_saved: int = 0

    @property
    def ratio(self) -> float:
        return self.optimized_mb / self.original_mb if self.original_mb else 0.0

    @property
    def reduction_pct(self) -> float:
        return (1 - self.ratio) * 100


class LosslessPPTXOptimizer:
    """레이아웃/도형/텍스트 불변. 미디어와 폰트만 최적화."""

    def __init__(
        self,
        max_image_dim: int = 1920,
        jpeg_quality: int = 90,
        use_oxipng: bool = True,
        strip_fonts: bool = True,
        zip_level: int = 9,
    ):
        self.max_image_dim = max_image_dim
        self.jpeg_quality = jpeg_quality
        self.use_oxipng = use_oxipng and shutil.which("oxipng") is not None
        self.strip_fonts = strip_fonts
        self.zip_level = zip_level

    def optimize(self, src: Path, dst: Path) -> OptimizeStats:
        stats = OptimizeStats(original_mb=src.stat().st_size / 1e6)

        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
            dst, "w", zipfile.ZIP_DEFLATED, compresslevel=self.zip_level
        ) as zout:
            for entry in zin.infolist():
                name = entry.filename
                data = zin.read(name)
                original = len(data)

                # 임베드 폰트 제거 (Google Slides는 무시함)
                if self.strip_fonts and (
                    name.startswith("ppt/fonts/")
                    or name.startswith("ppt/embeddings/fonts/")
                ):
                    stats.fonts_stripped += 1
                    stats.bytes_saved += original
                    continue

                # 이미지 처리
                if name.startswith("ppt/media/"):
                    new_data = self._process_image(data, name, stats)
                    if new_data is not None and len(new_data) < original:
                        stats.bytes_saved += original - len(new_data)
                        data = new_data
                    stats.images_processed += 1

                zout.writestr(
                    entry,
                    data,
                    compress_type=zipfile.ZIP_DEFLATED,
                    compresslevel=self.zip_level,
                )

        stats.optimized_mb = dst.stat().st_size / 1e6
        return stats

    def _process_image(
        self, data: bytes, name: str, stats: OptimizeStats
    ) -> bytes | None:
        lower = name.lower()
        if lower.endswith((".emf", ".wmf")):
            return None

        try:
            img = Image.open(BytesIO(data))
            img.load()
        except Exception:
            return None

        w, h = img.size
        resized = False
        if max(w, h) > self.max_image_dim:
            ratio = self.max_image_dim / max(w, h)
            img = img.resize(
                (max(1, int(w * ratio)), max(1, int(h * ratio))),
                Image.LANCZOS,
            )
            resized = True
            stats.images_downsampled += 1

        buf = BytesIO()
        if lower.endswith(".png"):
            # 알파 채널 유지
            save_img = img
            if save_img.mode not in ("RGBA", "LA", "P"):
                save_img = save_img.convert("RGBA" if "A" in img.mode else "RGB")
            save_img.save(buf, format="PNG", optimize=True, compress_level=9)
            out = buf.getvalue()
            if self.use_oxipng:
                out = self._oxipng(out) or out
            stats.png_recompressed += 1
            return out

        if lower.endswith((".jpg", ".jpeg")):
            if not resized:
                # 리사이즈 안 했으면 건드리지 않음 (재인코딩 손실 방지)
                return None
            rgb = img.convert("RGB")
            rgb.save(
                buf,
                format="JPEG",
                quality=self.jpeg_quality,
                optimize=True,
                progressive=True,
                subsampling=0,
            )
            stats.jpeg_recompressed += 1
            return buf.getvalue()

        return None

    def _oxipng(self, png_bytes: bytes) -> bytes | None:
        try:
            r = subprocess.run(
                ["oxipng", "--opt", "max", "--strip", "safe", "-"],
                input=png_bytes,
                capture_output=True,
                timeout=60,
            )
            if r.returncode == 0 and len(r.stdout) < len(png_bytes):
                return r.stdout
        except Exception:
            pass
        return None


def print_stats(stats: OptimizeStats) -> None:
    print(f"\n{'='*60}")
    print("최적화 완료")
    print(f"{'='*60}")
    print(f"원본        : {stats.original_mb:>8.1f} MB")
    print(f"결과        : {stats.optimized_mb:>8.1f} MB")
    print(f"감소율      : {stats.reduction_pct:>8.1f} %")
    print(f"절감        : {stats.bytes_saved/1e6:>8.1f} MB")
    print(f"이미지 처리 : {stats.images_processed}개")
    print(f"  다운샘플  : {stats.images_downsampled}개")
    print(f"  PNG 재압축: {stats.png_recompressed}개")
    print(f"  JPEG 재압축: {stats.jpeg_recompressed}개")
    print(f"폰트 제거   : {stats.fonts_stripped}개")
    print(f"100MB 목표  : {'OK' if stats.optimized_mb < 100 else 'FAIL'}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    import sys

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(
        src.stem + "_optimized.pptx"
    )
    opt = LosslessPPTXOptimizer(max_image_dim=1920, jpeg_quality=90)
    stats = opt.optimize(src, dst)
    print_stats(stats)
    print(f"출력: {dst}")
