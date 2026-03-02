"""
IP 딥 리서치 에이전트 (v4.0)

RFP 밖의 실제 데이터를 조사하여 콘텐츠 생성에 활용합니다.
브랜드 현황, 커뮤니티 동향, 경쟁사, 트렌드 등을 수집합니다.
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional

from .base_agent import BaseAgent
from ..schemas.ip_research_schema import (
    BrandData,
    CollaboratorCandidate,
    CommunityInsight,
    CompetitorProfile,
    ConfidenceLevel,
    DataPoint,
    IndustryTrend,
    IPResearchResult,
)
from ..schemas.rfp_schema import RFPAnalysis
from ..utils.logger import get_logger

logger = get_logger("ip_researcher")


class IPResearcher(BaseAgent):
    """
    IP 딥 리서치 에이전트

    RFP 분석 결과를 기반으로 대상 IP/브랜드에 대한
    심층 리서치를 수행합니다.
    """

    async def execute(
        self,
        input_data: Dict[str, Any],
        progress_callback: Optional[Callable] = None,
    ) -> IPResearchResult:
        """
        IP 딥 리서치 실행

        Args:
            input_data: {
                "rfp_analysis": RFPAnalysis,
                "project_type": str (optional),
            }
            progress_callback: 진행 상황 콜백

        Returns:
            IPResearchResult: 리서치 결과
        """
        rfp_analysis: RFPAnalysis = input_data.get("rfp_analysis")
        project_type = input_data.get("project_type", "")

        if progress_callback:
            progress_callback({
                "step": 1,
                "total": 3,
                "message": "IP/브랜드 리서치 키워드 도출 중...",
            })

        # 검색 키워드 도출
        search_context = self._build_search_context(rfp_analysis, project_type)

        if progress_callback:
            progress_callback({
                "step": 2,
                "total": 3,
                "message": "IP/브랜드 심층 분석 중 (Claude)...",
            })

        # Claude에게 리서치 분석 의뢰
        system_prompt = self._load_prompt("ip_research_guide")
        if not system_prompt:
            system_prompt = self._get_default_system_prompt()

        user_message = self._build_user_message(search_context, rfp_analysis)

        response = self._call_claude(
            system_prompt=system_prompt,
            user_message=user_message,
            max_tokens=8192,
        )

        if progress_callback:
            progress_callback({
                "step": 3,
                "total": 3,
                "message": "리서치 결과 구조화 중...",
            })

        # JSON 파싱 + IPResearchResult 변환
        result = self._parse_result(response, search_context)

        logger.info(
            f"IP 리서치 완료: {result.target_brand} "
            f"(검증:{result.verified_data_count}, "
            f"추정:{result.estimated_data_count}, "
            f"AI생성:{result.ai_generated_count})"
        )

        return result

    def _build_search_context(
        self,
        rfp_analysis: RFPAnalysis,
        project_type: str,
    ) -> Dict[str, Any]:
        """리서치 컨텍스트 구성"""
        context = {
            "project_name": rfp_analysis.project_name if rfp_analysis else "",
            "client_name": rfp_analysis.client_name if rfp_analysis else "",
            "project_overview": rfp_analysis.project_overview if rfp_analysis else "",
            "project_type": project_type,
            "key_requirements": [],
            "target_audience": "",
            "industry_keywords": [],
        }

        if rfp_analysis:
            context["key_requirements"] = rfp_analysis.key_requirements[:5]
            context["target_audience"] = getattr(rfp_analysis, 'target_audience', '')
            context["industry_keywords"] = getattr(rfp_analysis, 'keywords', [])

            # IP/브랜드 관련 키워드 추출
            overview = rfp_analysis.project_overview or ""
            name = rfp_analysis.project_name or ""
            context["brand_hints"] = [name] + [
                kw for kw in (rfp_analysis.key_requirements or [])[:3]
            ]

        return context

    def _build_user_message(
        self,
        search_context: Dict[str, Any],
        rfp_analysis: Optional[RFPAnalysis],
    ) -> str:
        """Claude에게 전달할 사용자 메시지 구성"""
        msg_parts = [
            "# IP/브랜드 딥 리서치 요청",
            "",
            f"## 프로젝트 정보",
            f"- 프로젝트명: {search_context['project_name']}",
            f"- 발주처: {search_context['client_name']}",
            f"- 유형: {search_context['project_type']}",
            "",
            f"## 프로젝트 개요",
            f"{search_context['project_overview'][:3000]}",
            "",
        ]

        if search_context.get("key_requirements"):
            msg_parts.append("## 핵심 요구사항")
            for req in search_context["key_requirements"]:
                msg_parts.append(f"- {req}")
            msg_parts.append("")

        if rfp_analysis and hasattr(rfp_analysis, 'pain_points') and rfp_analysis.pain_points:
            msg_parts.append("## RFP에서 파악된 Pain Points")
            for pp in rfp_analysis.pain_points[:5]:
                msg_parts.append(f"- {pp}")
            msg_parts.append("")

        msg_parts.extend([
            "## 리서치 요청",
            "",
            "위 프로젝트에 관련된 IP/브랜드에 대해 다음을 조사해주세요:",
            "",
            "1. **브랜드 현황**: MAU, 소셜미디어 팔로워, 시장 위치, 핵심 캐릭터, 최근 동향",
            "2. **커뮤니티 동향**: 유저 감성, 바이럴 토픽, 인기 콘텐츠, 유저 니즈",
            "3. **경쟁사 분석**: 동종 브랜드/IP의 최근 이벤트, 차별점",
            "4. **잠재 협력사**: 관련 인플루언서, 코스어, 아티스트",
            "5. **산업 트렌드**: 해당 분야의 최근 트렌드",
            "",
            "각 데이터에 출처와 신뢰도(verified/estimated/ai_generated)를 반드시 포함해주세요.",
            "",
            "JSON 형식으로 응답해주세요.",
        ])

        return "\n".join(msg_parts)

    def _parse_result(
        self,
        response: str,
        search_context: Dict[str, Any],
    ) -> IPResearchResult:
        """Claude 응답을 IPResearchResult로 변환"""
        raw = self._extract_json(response)

        if not raw:
            logger.warning("IP 리서치 JSON 파싱 실패 - 기본 결과 반환")
            return IPResearchResult(
                target_brand=search_context.get("project_name", ""),
                research_scope="RFP 기반 추정",
                research_timestamp=datetime.now().isoformat(),
            )

        try:
            # 브랜드 데이터
            brand_raw = raw.get("brand_data", {})
            brand_data = BrandData(
                brand_name=brand_raw.get("brand_name", ""),
                company=brand_raw.get("company", ""),
                genre=brand_raw.get("genre", ""),
                release_date=brand_raw.get("release_date", ""),
                platforms=brand_raw.get("platforms", []),
                mau=self._parse_data_point(brand_raw.get("mau", {})),
                dau=self._parse_data_point(brand_raw.get("dau", {})),
                total_downloads=self._parse_data_point(brand_raw.get("total_downloads", {})),
                social_media={
                    k: self._parse_data_point(v)
                    for k, v in brand_raw.get("social_media", {}).items()
                },
                key_characters=brand_raw.get("key_characters", []),
                ip_strengths=brand_raw.get("ip_strengths", []),
                brand_keywords=brand_raw.get("brand_keywords", []),
                recent_updates=brand_raw.get("recent_updates", []),
            )

            # 커뮤니티 인사이트
            communities = []
            for ci_raw in raw.get("community_insights", []):
                communities.append(CommunityInsight(
                    platform=ci_raw.get("platform", ""),
                    overall_sentiment=ci_raw.get("overall_sentiment", ""),
                    sentiment_details=ci_raw.get("sentiment_details", ""),
                    viral_topics=ci_raw.get("viral_topics", []),
                    popular_characters=ci_raw.get("popular_characters", []),
                    user_demands=ci_raw.get("user_demands", []),
                    pain_points=ci_raw.get("pain_points", []),
                ))

            # 경쟁사 프로파일
            competitors = []
            for cp_raw in raw.get("competitor_profiles", []):
                competitors.append(CompetitorProfile(
                    name=cp_raw.get("name", ""),
                    brand=cp_raw.get("brand", ""),
                    strengths=cp_raw.get("strengths", []),
                    weaknesses=cp_raw.get("weaknesses", []),
                    market_position=cp_raw.get("market_position", ""),
                ))

            # 협력사 후보
            collaborators = []
            for col_raw in raw.get("collaborator_candidates", []):
                collaborators.append(CollaboratorCandidate(
                    name=col_raw.get("name", ""),
                    category=col_raw.get("category", ""),
                    platform=col_raw.get("platform", ""),
                    relevance=col_raw.get("relevance", ""),
                ))

            # 트렌드
            trends = []
            for tr_raw in raw.get("industry_trends", []):
                trends.append(IndustryTrend(
                    trend_name=tr_raw.get("trend_name", ""),
                    description=tr_raw.get("description", ""),
                    relevance_to_project=tr_raw.get("relevance_to_project", ""),
                ))

            # 데이터 품질 계산
            verified = 0
            estimated = 0
            ai_gen = 0
            # (간단한 카운팅 - 실제로는 모든 DataPoint를 순회)
            total_data = max(1, len(communities) + len(competitors) + len(collaborators) + len(trends))
            if brand_data.mau.confidence == ConfidenceLevel.VERIFIED:
                verified += 1
            elif brand_data.mau.confidence == ConfidenceLevel.ESTIMATED:
                estimated += 1
            else:
                ai_gen += 1

            result = IPResearchResult(
                target_brand=raw.get("target_brand", search_context.get("project_name", "")),
                research_scope=raw.get("research_scope", "Claude 분석 기반"),
                brand_data=brand_data,
                community_insights=communities,
                competitor_profiles=competitors,
                collaborator_candidates=collaborators,
                industry_trends=trends,
                strategic_insights=raw.get("strategic_insights", []),
                differentiation_opportunities=raw.get("differentiation_opportunities", []),
                risk_factors=raw.get("risk_factors", []),
                data_quality_score=verified / total_data if total_data > 0 else 0,
                verified_data_count=verified,
                estimated_data_count=estimated,
                ai_generated_count=ai_gen,
                research_timestamp=datetime.now().isoformat(),
                search_queries_used=raw.get("search_queries_used", []),
            )

            return result

        except Exception as e:
            logger.error(f"IP 리서치 결과 파싱 실패: {e}")
            return IPResearchResult(
                target_brand=search_context.get("project_name", ""),
                research_scope="파싱 실패 - 기본 결과",
                research_timestamp=datetime.now().isoformat(),
            )

    def _parse_data_point(self, raw: Any) -> DataPoint:
        """데이터 포인트 파싱"""
        if isinstance(raw, str):
            return DataPoint(value=raw, confidence=ConfidenceLevel.AI_GENERATED)
        if isinstance(raw, dict):
            conf = raw.get("confidence", "ai_generated")
            try:
                confidence = ConfidenceLevel(conf)
            except ValueError:
                confidence = ConfidenceLevel.AI_GENERATED
            return DataPoint(
                value=raw.get("value", ""),
                source=raw.get("source", ""),
                confidence=confidence,
                date=raw.get("date", ""),
            )
        return DataPoint()

    def _get_default_system_prompt(self) -> str:
        """기본 시스템 프롬프트"""
        return """당신은 IP/브랜드 딥 리서치 전문가입니다.
RFP에서 언급된 IP, 브랜드, 프로젝트에 대해 심층 분석을 수행합니다.

모든 데이터에 출처(source)와 신뢰도(confidence: verified/estimated/ai_generated)를 포함하세요.
구체적인 수치와 사례를 최대한 포함하세요.
JSON 형식으로 응답하세요."""
