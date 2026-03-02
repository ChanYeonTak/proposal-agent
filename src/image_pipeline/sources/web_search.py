"""
웹 이미지 검색 (Unsplash/Pexels)

무료 이미지 API를 통해 슬라이드에 사용할 이미지를 검색합니다.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from ...utils.logger import get_logger

logger = get_logger("web_image_search")


class WebImageSearch:
    """
    웹 이미지 검색

    Unsplash API 또는 Pexels API를 사용합니다.
    API 키가 없으면 비활성화됩니다.
    """

    def __init__(self, cache_dir: Optional[Path] = None):
        self.cache_dir = cache_dir or Path("output/.image_cache/web")
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        self.unsplash_key = os.getenv("UNSPLASH_ACCESS_KEY", "")
        self.pexels_key = os.getenv("PEXELS_API_KEY", "")

        if not self.unsplash_key and not self.pexels_key:
            logger.warning(
                "이미지 검색 API 키 미설정 "
                "(UNSPLASH_ACCESS_KEY 또는 PEXELS_API_KEY)"
            )

    async def search(self, request) -> "ImageResult":
        """
        이미지 검색

        Args:
            request: ImageRequest

        Returns:
            ImageResult
        """
        from ..manager import ImageResult

        keywords = " ".join(request.keywords) if request.keywords else request.description

        if not keywords:
            return ImageResult(
                placeholder_id=request.placeholder_id,
                success=False,
                error="검색 키워드 없음",
            )

        # Unsplash 시도
        if self.unsplash_key:
            result = await self._search_unsplash(request, keywords)
            if result and result.success:
                return result

        # Pexels 시도
        if self.pexels_key:
            result = await self._search_pexels(request, keywords)
            if result and result.success:
                return result

        return ImageResult(
            placeholder_id=request.placeholder_id,
            success=False,
            error="이미지 검색 API 키 미설정 또는 검색 결과 없음",
        )

    async def _search_unsplash(self, request, keywords: str):
        """Unsplash API 검색"""
        from ..manager import ImageResult
        import urllib.request
        import json

        try:
            url = (
                f"https://api.unsplash.com/search/photos"
                f"?query={urllib.parse.quote(keywords)}"
                f"&per_page=1"
                f"&orientation=landscape"
            )

            req = urllib.request.Request(url)
            req.add_header("Authorization", f"Client-ID {self.unsplash_key}")

            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read())

            if data.get("results"):
                photo = data["results"][0]
                image_url = photo["urls"]["regular"]

                # 다운로드
                file_path = self.cache_dir / f"{request.placeholder_id}.jpg"
                urllib.request.urlretrieve(image_url, str(file_path))

                return ImageResult(
                    placeholder_id=request.placeholder_id,
                    file_path=file_path,
                    source="unsplash",
                    attribution=f"Photo by {photo['user']['name']} on Unsplash",
                    success=True,
                )

        except Exception as e:
            logger.warning(f"Unsplash 검색 실패: {e}")

        return None

    async def _search_pexels(self, request, keywords: str):
        """Pexels API 검색"""
        from ..manager import ImageResult
        import urllib.request
        import json

        try:
            url = (
                f"https://api.pexels.com/v1/search"
                f"?query={urllib.parse.quote(keywords)}"
                f"&per_page=1"
                f"&orientation=landscape"
            )

            req = urllib.request.Request(url)
            req.add_header("Authorization", self.pexels_key)

            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read())

            if data.get("photos"):
                photo = data["photos"][0]
                image_url = photo["src"]["large"]

                file_path = self.cache_dir / f"{request.placeholder_id}.jpg"
                urllib.request.urlretrieve(image_url, str(file_path))

                return ImageResult(
                    placeholder_id=request.placeholder_id,
                    file_path=file_path,
                    source="pexels",
                    attribution=f"Photo by {photo['photographer']} on Pexels",
                    success=True,
                )

        except Exception as e:
            logger.warning(f"Pexels 검색 실패: {e}")

        return None
