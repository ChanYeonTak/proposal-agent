"""
다이어그램 렌더링

평면도, 동선도, 조직도 등을 자동 생성합니다.
matplotlib + PIL 기반.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from ...utils.logger import get_logger

logger = get_logger("diagram_renderer")


class DiagramRenderer:
    """
    다이어그램 렌더링

    matplotlib을 사용하여 간단한 다이어그램을 생성합니다.
    """

    def __init__(self, cache_dir: Optional[Path] = None):
        self.cache_dir = cache_dir or Path("output/.image_cache/diagram")
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    async def search(self, request) -> "ImageResult":
        """
        다이어그램 생성

        Args:
            request: ImageRequest

        Returns:
            ImageResult
        """
        from ..manager import ImageResult

        if request.category != "diagram":
            return ImageResult(
                placeholder_id=request.placeholder_id,
                success=False,
                error="다이어그램 카테고리가 아님",
            )

        try:
            file_path = await self._render_placeholder(request)
            if file_path:
                return ImageResult(
                    placeholder_id=request.placeholder_id,
                    file_path=file_path,
                    source="diagram_renderer",
                    success=True,
                )
        except Exception as e:
            logger.warning(f"다이어그램 렌더링 실패: {e}")

        return ImageResult(
            placeholder_id=request.placeholder_id,
            success=False,
            error="다이어그램 렌더링 실패",
        )

    async def _render_placeholder(self, request) -> Optional[Path]:
        """플레이스홀더 다이어그램 생성"""
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            import matplotlib.patches as patches

            fig, ax = plt.subplots(1, 1, figsize=(16, 9), dpi=120)
            ax.set_xlim(0, 16)
            ax.set_ylim(0, 9)
            ax.set_aspect('equal')
            ax.axis('off')

            # 배경
            bg = patches.Rectangle((0, 0), 16, 9, fill=True, color='#F5F5F5')
            ax.add_patch(bg)

            # 중앙 텍스트
            desc = request.description[:50] if request.description else "다이어그램"
            ax.text(
                8, 5, desc,
                ha='center', va='center',
                fontsize=14, color='#666666',
                fontweight='bold',
            )
            ax.text(
                8, 3.5, "[이미지 삽입 영역]",
                ha='center', va='center',
                fontsize=10, color='#AAAAAA',
            )

            # 테두리
            border = patches.Rectangle(
                (0.2, 0.2), 15.6, 8.6,
                fill=False, edgecolor='#CCCCCC', linewidth=2, linestyle='--',
            )
            ax.add_patch(border)

            file_path = self.cache_dir / f"{request.placeholder_id}.png"
            plt.savefig(str(file_path), bbox_inches='tight', pad_inches=0)
            plt.close(fig)

            return file_path

        except ImportError:
            logger.warning("matplotlib 미설치")
            return None
