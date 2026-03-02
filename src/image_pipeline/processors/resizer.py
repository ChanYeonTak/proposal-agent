"""
이미지 리사이저

슬라이드 크기에 맞게 이미지를 조정합니다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from ...utils.logger import get_logger

logger = get_logger("image_resizer")


class ImageResizer:
    """이미지 크기 조정"""

    @staticmethod
    def resize(
        image_path: Path,
        target_size: Tuple[int, int] = (1920, 1080),
        output_path: Optional[Path] = None,
    ) -> Path:
        """
        이미지 크기 조정 (비율 유지)

        Args:
            image_path: 원본 이미지 경로
            target_size: 목표 크기 (width, height)
            output_path: 출력 경로 (None이면 원본 덮어쓰기)

        Returns:
            출력 파일 경로
        """
        try:
            from PIL import Image

            img = Image.open(str(image_path))

            # 비율 유지하면서 리사이즈
            img.thumbnail(target_size, Image.Resampling.LANCZOS)

            out_path = output_path or image_path
            img.save(str(out_path), quality=90)

            logger.info(f"이미지 리사이즈: {img.size} -> {out_path}")
            return out_path

        except ImportError:
            logger.warning("Pillow 미설치 - 리사이즈 건너뜀")
            return image_path
        except Exception as e:
            logger.error(f"이미지 리사이즈 실패: {e}")
            return image_path
