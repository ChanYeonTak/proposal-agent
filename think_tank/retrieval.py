"""
Think Tank 검색 엔진 (v5.0)

유사 레퍼런스, 디자인 패턴, 콘텐츠 패턴, 콘텐츠 톤을 검색합니다.
v5.0: get_content_tone_patterns() 추가 — 산업/유형별 콘텐츠 톤 규칙 검색
"""

from __future__ import annotations

from typing import List, Optional

from .db import ThinkTankDB
from .models import (
    ContentPattern,
    ContentToneProfile,
    DesignProfile,
    ProgramTemplate,
    ReferenceDocument,
    SearchResult,
)
from src.utils.logger import get_logger

logger = get_logger("think_tank_retrieval")


class ThinkTankRetrieval:
    """
    Think Tank 검색 엔진

    저장된 레퍼런스에서 유사 사례, 디자인 패턴, 콘텐츠 패턴을 검색합니다.
    """

    def __init__(self, db: Optional[ThinkTankDB] = None):
        self.db = db or ThinkTankDB()

    def search_similar(
        self,
        project_type: Optional[str] = None,
        industry: Optional[str] = None,
        won_bid_only: bool = True,
        top_k: int = 3,
    ) -> List[SearchResult]:
        """
        유사 레퍼런스 검색

        Args:
            project_type: 프로젝트 유형 (event, marketing_pr 등)
            industry: 산업 분류
            won_bid_only: 수주 성공 사례만
            top_k: 최대 결과 수

        Returns:
            List[SearchResult]: 유사 레퍼런스 목록 (관련도 순)
        """
        # 1차: 동일 유형 + 산업 + 수주 성공
        results = self.db.search_by_type(
            doc_type="proposal",
            industry=industry,
            project_type=project_type,
            won_bid_only=won_bid_only,
            limit=top_k,
        )

        # 부족하면 유형만으로 검색
        if len(results) < top_k and project_type:
            more = self.db.search_by_type(
                doc_type="proposal",
                project_type=project_type,
                won_bid_only=won_bid_only,
                limit=top_k - len(results),
            )
            existing_ids = {r.id for r in results}
            results.extend([r for r in more if r.id not in existing_ids])

        # 그래도 부족하면 산업만으로
        if len(results) < top_k and industry:
            more = self.db.search_by_type(
                doc_type="proposal",
                industry=industry,
                won_bid_only=False,
                limit=top_k - len(results),
            )
            existing_ids = {r.id for r in results}
            results.extend([r for r in more if r.id not in existing_ids])

        # SearchResult로 변환 (간단한 점수 부여)
        search_results = []
        for i, doc in enumerate(results[:top_k]):
            score = 1.0 - (i * 0.1)

            reasons = []
            if doc.project_type == project_type:
                reasons.append(f"동일 유형({project_type})")
            if doc.industry and doc.industry.value == industry:
                reasons.append(f"동일 산업({industry})")
            if doc.won_bid:
                reasons.append("수주 성공")

            search_results.append(SearchResult(
                document=doc,
                relevance_score=score,
                match_reason=", ".join(reasons) if reasons else "전체 검색",
            ))

        logger.info(f"유사 레퍼런스 검색: {len(search_results)}건 (type={project_type}, industry={industry})")
        return search_results

    def get_design_patterns(
        self,
        project_type: Optional[str] = None,
        industry: Optional[str] = None,
    ) -> List[DesignProfile]:
        """
        해당 유형의 디자인 패턴 반환

        Args:
            project_type: 프로젝트 유형
            industry: 산업 분류

        Returns:
            List[DesignProfile]: 디자인 프로파일 목록
        """
        docs = self.db.search_by_type(
            doc_type="proposal",
            industry=industry,
            project_type=project_type,
            limit=5,
        )

        profiles = [doc.design_profile for doc in docs if doc.design_profile]
        logger.info(f"디자인 패턴 검색: {len(profiles)}건")
        return profiles

    def get_content_patterns(
        self,
        section: Optional[str] = None,
        project_type: Optional[str] = None,
    ) -> List[ContentPattern]:
        """
        섹션별 콘텐츠 패턴 반환

        Args:
            section: 섹션 이름 (예: "ACTION PLAN", "CONCEPT")
            project_type: 프로젝트 유형

        Returns:
            List[ContentPattern]: 콘텐츠 패턴 목록
        """
        docs = self.db.search_by_type(
            doc_type="proposal",
            project_type=project_type,
            won_bid_only=True,
            limit=10,
        )

        patterns = []
        for doc in docs:
            for pattern in doc.content_patterns:
                if section is None or section.lower() in pattern.section_context.lower():
                    patterns.append(pattern)

        logger.info(f"콘텐츠 패턴 검색: {len(patterns)}건 (section={section})")
        return patterns

    def get_program_templates(
        self,
        industry: Optional[str] = None,
        category: Optional[str] = None,
    ) -> List[ProgramTemplate]:
        """
        프로그램/이벤트 템플릿 반환

        Args:
            industry: 산업 분류
            category: 프로그램 카테고리 (booth_design, event_pack 등)

        Returns:
            List[ProgramTemplate]: 프로그램 템플릿 목록
        """
        docs = self.db.search_by_type(
            doc_type="proposal",
            industry=industry,
            won_bid_only=True,
            limit=10,
        )

        templates = []
        for doc in docs:
            for tmpl in doc.program_templates:
                if category is None or tmpl.category == category:
                    templates.append(tmpl)

        logger.info(f"프로그램 템플릿 검색: {len(templates)}건 (industry={industry})")
        return templates

    def get_section_structure_stats(
        self,
        project_type: Optional[str] = None,
    ) -> dict:
        """
        섹션별 평균 슬라이드 수 통계

        수주 성공 레퍼런스들의 섹션 구조를 분석하여
        각 섹션에 평균적으로 할당되는 슬라이드 수를 반환합니다.
        """
        docs = self.db.search_by_type(
            doc_type="proposal",
            project_type=project_type,
            won_bid_only=True,
            limit=20,
        )

        section_stats: dict = {}  # section_name → [slide_counts]

        for doc in docs:
            for section in doc.sections:
                name = section.name.upper().strip()
                if name not in section_stats:
                    section_stats[name] = []
                section_stats[name].append(section.slide_count)

        # 평균 계산
        result = {}
        for name, counts in section_stats.items():
            result[name] = {
                "avg_slides": round(sum(counts) / len(counts), 1),
                "min_slides": min(counts),
                "max_slides": max(counts),
                "sample_count": len(counts),
            }

        return result

    def get_content_tone_patterns(
        self,
        industry: Optional[str] = None,
        project_type: Optional[str] = None,
        won_bid_only: bool = True,
    ) -> List[ContentToneProfile]:
        """
        산업/유형별 콘텐츠 톤 프로파일 반환

        수주 성공 레퍼런스의 ContentToneProfile을 검색하여
        해당 산업에서 효과적인 콘텐츠 톤 패턴을 제공합니다.

        Args:
            industry: 산업 분류 (game_event, marketing_pr 등)
            project_type: 프로젝트 유형
            won_bid_only: 수주 성공 사례만

        Returns:
            List[ContentToneProfile]: 콘텐츠 톤 프로파일 목록
        """
        docs = self.db.search_by_type(
            doc_type="proposal",
            industry=industry,
            project_type=project_type,
            won_bid_only=won_bid_only,
            limit=10,
        )

        profiles = []
        for doc in docs:
            if doc.content_tone and doc.content_tone.emotional_tone_level > 0:
                profiles.append(doc.content_tone)

        logger.info(
            f"콘텐츠 톤 패턴 검색: {len(profiles)}건 "
            f"(industry={industry}, type={project_type})"
        )
        return profiles

    def get_tone_rules_for_industry(
        self,
        industry: str,
    ) -> List[str]:
        """
        특정 산업에 대한 콘텐츠 톤 규칙만 반환 (편의 메서드)

        Args:
            industry: 산업 분류

        Returns:
            List[str]: 톤 규칙 목록
        """
        profiles = self.get_content_tone_patterns(
            industry=industry, won_bid_only=True,
        )

        rules = []
        for p in profiles:
            rules.extend(p.tone_rules)

        # 중복 제거하면서 순서 유지
        seen = set()
        unique_rules = []
        for r in rules:
            if r not in seen:
                seen.add(r)
                unique_rules.append(r)

        return unique_rules
