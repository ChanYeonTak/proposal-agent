"""
이미지 파이프라인 매니저 (v5.0)

IMG_PH 플레이스홀더를 분석하고, 적절한 이미지 소스에서
이미지를 수급하여 PPTX에 삽입합니다.

v5.0: MergedDesignBrief 통합 — 디자인 에이전트가 결정한
이미지 스타일/키워드를 반영하여 소스 자동 선택
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from ..utils.logger import get_logger

logger = get_logger("image_pipeline")


class ImageRequest(object):
    """이미지 요청"""
    def __init__(
        self,
        placeholder_id: str = "",
        description: str = "",
        category: str = "",       # photo, diagram, chart, illustration, logo
        keywords: List[str] = None,
        width: int = 1920,
        height: int = 1080,
        source_preference: str = "web",  # web, ai, diagram
    ):
        self.placeholder_id = placeholder_id
        self.description = description
        self.category = category
        self.keywords = keywords or []
        self.width = width
        self.height = height
        self.source_preference = source_preference


class ImageResult(object):
    """이미지 결과"""
    def __init__(
        self,
        placeholder_id: str = "",
        file_path: Optional[Path] = None,
        source: str = "",          # unsplash, pexels, dalle, stable_diffusion, diagram
        attribution: str = "",
        success: bool = False,
        error: str = "",
    ):
        self.placeholder_id = placeholder_id
        self.file_path = file_path
        self.source = source
        self.attribution = attribution
        self.success = success
        self.error = error


class ImagePipelineManager:
    """
    이미지 파이프라인 매니저

    1. PPTX 또는 콘텐츠에서 IMG_PH 플레이스홀더 추출
    2. 각 플레이스홀더에 맞는 이미지 소스 선택
    3. 이미지 다운로드/생성
    4. 크기 조정 후 PPTX에 삽입

    MergedDesignBrief 통합:
        brief가 제공되면 image_style.source_preference에 따라
        카테고리별 소스 우선순위가 자동 결정됩니다.
    """

    def __init__(
        self,
        cache_dir: Optional[Path] = None,
        web_search_enabled: bool = True,
        ai_generation_enabled: bool = False,
        diagram_rendering_enabled: bool = True,
        design_brief: Optional[Any] = None,
    ):
        self.cache_dir = cache_dir or Path("output/.image_cache")
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        self.web_search_enabled = web_search_enabled
        self.ai_generation_enabled = ai_generation_enabled
        self.diagram_rendering_enabled = diagram_rendering_enabled

        # MergedDesignBrief에서 이미지 설정 추출
        self._brief_keywords: List[str] = []
        self._brief_source_map: Dict[str, str] = {}
        if design_brief:
            img_style = getattr(design_brief, "image_style", None)
            if img_style:
                self._brief_keywords = getattr(img_style, "keywords", [])
                self._brief_source_map = getattr(
                    img_style, "source_preference", {}
                )

        self._sources: Dict[str, Any] = {}
        self._init_sources()

    def _init_sources(self):
        """이미지 소스 초기화"""
        if self.web_search_enabled:
            try:
                from .sources.web_search import WebImageSearch
                self._sources["web"] = WebImageSearch(cache_dir=self.cache_dir)
                logger.info("웹 이미지 검색 활성화")
            except ImportError:
                logger.warning("웹 이미지 검색 모듈 없음")

        if self.ai_generation_enabled:
            try:
                from .sources.ai_generator import AIImageGenerator
                self._sources["ai"] = AIImageGenerator(cache_dir=self.cache_dir)
                logger.info("AI 이미지 생성 활성화")
            except ImportError:
                logger.warning("AI 이미지 생성 모듈 없음")

        if self.diagram_rendering_enabled:
            try:
                from .sources.diagram_renderer import DiagramRenderer
                self._sources["diagram"] = DiagramRenderer(cache_dir=self.cache_dir)
                logger.info("다이어그램 렌더링 활성화")
            except ImportError:
                logger.warning("다이어그램 렌더링 모듈 없음")

    async def process_requests(
        self,
        requests: List[ImageRequest],
    ) -> Dict[str, ImageResult]:
        """
        이미지 요청 일괄 처리

        Args:
            requests: 이미지 요청 목록

        Returns:
            Dict[placeholder_id, ImageResult]
        """
        results = {}

        for req in requests:
            try:
                result = await self._process_single(req)
                results[req.placeholder_id] = result
            except Exception as e:
                logger.error(f"이미지 처리 실패 [{req.placeholder_id}]: {e}")
                results[req.placeholder_id] = ImageResult(
                    placeholder_id=req.placeholder_id,
                    success=False,
                    error=str(e),
                )

        success_count = sum(1 for r in results.values() if r.success)
        logger.info(f"이미지 처리 완료: {success_count}/{len(requests)} 성공")

        return results

    async def _process_single(self, req: ImageRequest) -> ImageResult:
        """단일 이미지 요청 처리"""
        # 소스 우선순위 결정
        source_order = self._get_source_order(req)

        for source_name in source_order:
            source = self._sources.get(source_name)
            if source is None:
                continue

            try:
                result = await source.search(req)
                if result and result.success:
                    return result
            except Exception as e:
                logger.warning(f"[{source_name}] 실패: {e}")
                continue

        return ImageResult(
            placeholder_id=req.placeholder_id,
            success=False,
            error="모든 이미지 소스 실패",
        )

    def _get_source_order(self, req: ImageRequest) -> List[str]:
        """
        요청에 맞는 소스 우선순위.

        MergedDesignBrief의 source_preference가 있으면 반영:
            {"photo": "pexels", "illustration": "gamma_ai", "diagram": "renderer"}
        """
        # MergedDesignBrief 기반 소스 매핑
        if self._brief_source_map and req.category in self._brief_source_map:
            pref = self._brief_source_map[req.category]
            if pref in ("pexels", "unsplash", "web"):
                return ["web", "ai", "diagram"]
            elif pref in ("gamma_ai", "dalle", "ai"):
                return ["ai", "web", "diagram"]
            elif pref in ("renderer", "diagram"):
                return ["diagram", "ai", "web"]

        # 기본 카테고리별 규칙
        if req.category == "diagram":
            return ["diagram", "ai", "web"]
        elif req.category == "illustration":
            return ["ai", "web"]
        elif req.source_preference == "ai":
            return ["ai", "web"]
        else:
            return ["web", "ai", "diagram"]

    def extract_placeholders_from_content(self, content) -> List[ImageRequest]:
        """ProposalContent에서 IMG_PH 플레이스홀더 추출"""
        requests = []

        # teaser slides
        if content.teaser:
            for slide in content.teaser.slides:
                requests.extend(self._extract_from_slide(slide, "teaser"))

        # phase slides
        for phase in content.phases:
            for slide in phase.slides:
                requests.extend(self._extract_from_slide(slide, f"phase{phase.phase_number}"))

        logger.info(f"플레이스홀더 추출: {len(requests)}개")
        return requests

    def _extract_from_slide(self, slide, context: str) -> List[ImageRequest]:
        """슬라이드에서 이미지 플레이스홀더 추출"""
        requests = []

        # visual_elements 필드에서 IMG_PH 참조 검색
        if hasattr(slide, 'visual_elements') and slide.visual_elements:
            for i, elem in enumerate(slide.visual_elements):
                if "IMG_PH" in str(elem) or "이미지" in str(elem) or "사진" in str(elem):
                    req = ImageRequest(
                        placeholder_id=f"{context}_slide{i}_{len(requests)}",
                        description=str(elem),
                        category=self._guess_category(str(elem)),
                        keywords=self._extract_keywords(str(elem)),
                    )
                    requests.append(req)

        return requests

    def _guess_category(self, description: str) -> str:
        """설명에서 이미지 카테고리 추정"""
        desc_lower = description.lower()
        if any(kw in desc_lower for kw in ["평면도", "동선도", "아키텍처", "다이어그램", "플로우"]):
            return "diagram"
        elif any(kw in desc_lower for kw in ["일러스트", "캐릭터", "컨셉아트"]):
            return "illustration"
        elif any(kw in desc_lower for kw in ["로고", "아이콘"]):
            return "logo"
        elif any(kw in desc_lower for kw in ["차트", "그래프"]):
            return "chart"
        return "photo"

    def _extract_keywords(self, description: str) -> List[str]:
        """설명에서 검색 키워드 추출 (brief 키워드 보강)"""
        keywords = []
        import re
        # 괄호 내 텍스트 추출
        bracket_content = re.findall(r'[（(]([^)）]+)[)）]', description)
        keywords.extend(bracket_content)

        # 쉼표/공백으로 분리된 주요 단어
        for word in description.replace(",", " ").split():
            if len(word) > 1 and word not in ["IMG_PH", "이미지", "사진", "자료"]:
                keywords.append(word)

        # MergedDesignBrief 키워드 보강 (검색 정확도 향상)
        if self._brief_keywords:
            for bk in self._brief_keywords[:2]:
                if bk not in keywords:
                    keywords.append(bk)

        return keywords[:7]
