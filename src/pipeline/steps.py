"""
파이프라인 단계 정의 (v4.0)

각 단계는 PipelineStep 추상 클래스를 상속하고,
execute(context, step_config) → context 인터페이스를 구현합니다.

기존 v3.6 기능을 Step으로 래핑하여 하위 호환성을 유지합니다.
"""

from __future__ import annotations

import json
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, Optional

from .config import StepConfig
from .engine import PipelineContext
from ..utils.logger import get_logger

logger = get_logger("pipeline_steps")


# ═══════════════════════════════════════════════════════════════
# 추상 클래스
# ═══════════════════════════════════════════════════════════════

class PipelineStep(ABC):
    """
    파이프라인 단계 추상 클래스

    Attributes:
        name: 단계 이름 (pipeline_config.yaml 키와 매칭)
        description: 단계 설명 (진행 상황 표시용)
        is_critical: True면 실패 시 파이프라인 중단
    """

    name: str = "unnamed_step"
    description: str = "처리 중"
    is_critical: bool = False

    @abstractmethod
    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        """
        단계 실행

        Args:
            context: 파이프라인 컨텍스트 (이전 단계의 산출물 포함)
            step_config: 이 단계의 설정 (options 포함)

        Returns:
            PipelineContext: 산출물이 추가된 컨텍스트
        """
        ...


# ═══════════════════════════════════════════════════════════════
# Step 1: 문서 파싱
# ═══════════════════════════════════════════════════════════════

class DocumentParsingStep(PipelineStep):
    """RFP 문서 파싱 (PDF/DOCX → Dict)"""

    name = "document_parsing"
    description = "RFP 문서 파싱 중"
    is_critical = True

    def __init__(self):
        from ..parsers.pdf_parser import PDFParser
        from ..parsers.docx_parser import DOCXParser
        self.pdf_parser = PDFParser()
        self.docx_parser = DOCXParser()

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        if context.rfp_path is None:
            raise ValueError("rfp_path가 설정되지 않았습니다")

        suffix = context.rfp_path.suffix.lower()

        if suffix == ".pdf":
            parsed = self.pdf_parser.parse(context.rfp_path)
        elif suffix in [".docx", ".doc"]:
            parsed = self.docx_parser.parse(context.rfp_path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {suffix}")

        context.parsed_rfp = parsed
        logger.info(f"RFP 파싱 완료: {len(parsed.get('raw_text', ''))} 문자")

        # 회사 데이터 로드
        if context.company_data_path and context.company_data_path.exists():
            try:
                context.company_data = json.loads(
                    context.company_data_path.read_text(encoding="utf-8")
                )
                logger.info("회사 데이터 로드 완료")
            except Exception as e:
                logger.warning(f"회사 데이터 로드 실패: {e}")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 2: RFP 분석
# ═══════════════════════════════════════════════════════════════

class RFPAnalysisStep(PipelineStep):
    """RFP 분석 (Claude API)"""

    name = "rfp_analysis"
    description = "RFP 분석 중 (Claude)"
    is_critical = True

    def __init__(self, api_key: str = ""):
        self._api_key = api_key

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        from ..agents.rfp_analyzer import RFPAnalyzer

        api_key = self._api_key or context.api_key
        analyzer = RFPAnalyzer(api_key=api_key)

        rfp_analysis = await analyzer.execute(
            input_data=context.parsed_rfp or {},
            progress_callback=lambda p: context.progress_callback({
                "phase": "rfp_analysis",
                "sub_step": p.get("step"),
                "sub_total": p.get("total"),
                "message": p.get("message", "분석 중..."),
            }) if context.progress_callback else None,
        )

        context.rfp_analysis = rfp_analysis

        # 프로젝트명/발주처명 결정
        if not context.project_name:
            context.project_name = rfp_analysis.project_name
        if not context.client_name:
            context.client_name = rfp_analysis.client_name

        logger.info(f"RFP 분석 완료: {context.project_name} ({context.client_name})")
        return context


# ═══════════════════════════════════════════════════════════════
# Step 3: IP 딥 리서치
# ═══════════════════════════════════════════════════════════════

class IPResearchStep(PipelineStep):
    """IP 딥 리서치 - RFP 밖의 실제 데이터 조사"""

    name = "ip_research"
    description = "IP/브랜드 딥 리서치 중"
    is_critical = False

    def __init__(self, api_key: str = ""):
        self._api_key = api_key

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        try:
            from ..agents.ip_researcher import IPResearcher

            api_key = self._api_key or context.api_key
            researcher = IPResearcher(api_key=api_key)

            result = await researcher.execute(
                input_data={
                    "rfp_analysis": context.rfp_analysis,
                    "project_type": context.proposal_type,
                },
                progress_callback=lambda p: context.progress_callback({
                    "phase": "ip_research",
                    "sub_step": p.get("step"),
                    "sub_total": p.get("total"),
                    "message": p.get("message", "리서치 중..."),
                }) if context.progress_callback else None,
            )

            context.ip_research = result
            logger.info(
                f"IP 리서치 완료: {result.target_brand} "
                f"(커뮤니티 {len(result.community_insights)}건, "
                f"경쟁사 {len(result.competitor_profiles)}건)"
            )
        except ImportError:
            logger.warning("IP 리서치 모듈 미설치 - 건너뜀")
        except Exception as e:
            logger.warning(f"IP 리서치 실패: {e} - 건너뜀")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 4: Think Tank 검색
# ═══════════════════════════════════════════════════════════════

class ThinkTankRetrievalStep(PipelineStep):
    """Think Tank 레퍼런스 검색 + 디자인 브리프 생성

    싱크탱크 역할 (★ 명확한 책임 정의):
        1. 수주 성공 레퍼런스 기반 디자인 규칙 제공
        2. 레이아웃 배분 (complex_diagram 42%, image_focused 20% 등)
        3. Phase별 장표 수 배분 (ACTION PLAN 40%, INSIGHT 10% 등)
        4. 배경 스케줄 (슬라이드별 white/dark/light/gradient 패턴)
        5. 시각 밀도 목표 (이미지 25%, 도식 35%, 텍스트 20%)
        6. 컴포넌트 사용 빈도 목표 (IMG_PH 18개, COLS 8개 등)

    slide_kit과의 관계 (충돌 없음):
        - 싱크탱크 = 규칙 (WHAT: 어떤 레이아웃을 얼마나 쓸지)
        - slide_kit = 도구 (HOW: 레이아웃을 어떻게 그릴지)
        - DesignAgent가 싱크탱크 규칙 → slide_kit 테마로 변환하는 브릿지
    """

    name = "think_tank_retrieval"
    description = "Think Tank 레퍼런스 검색 + 디자인 규칙 생성 중"
    is_critical = False

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        try:
            from think_tank.retrieval import ThinkTankRetrieval
            from think_tank.db import ThinkTankDB

            db = ThinkTankDB()
            retrieval = ThinkTankRetrieval(db=db)

            top_k = step_config.get_option("top_k", 3)

            # 유사 레퍼런스 검색
            industry = None
            if context.rfp_analysis and hasattr(context.rfp_analysis, 'industry'):
                industry = context.rfp_analysis.industry

            search_results = retrieval.search_similar(
                project_type=context.proposal_type,
                industry=industry,
                won_bid_only=True,
                top_k=top_k,
            )

            context.similar_references = [sr.document for sr in search_results]

            # 디자인 패턴
            context.design_patterns = retrieval.get_design_patterns(
                project_type=context.proposal_type,
                industry=industry,
            )

            # 콘텐츠 패턴
            context.content_patterns = retrieval.get_content_patterns(
                project_type=context.proposal_type,
            )

            # ── 디자인 브리프 생성 (★ 싱크탱크의 핵심 역할) ──
            # 수주 성공 레퍼런스의 디자인 규칙을 구조화된 DesignBrief로 변환
            generate_brief = step_config.get_option("generate_design_brief", True)
            if generate_brief:
                try:
                    from think_tank.design_brief import DesignBriefBuilder

                    # 목표 슬라이드 수: context.extras에서 가져오거나 기본값 70
                    target_slides = context.extras.get("target_slides", 70)

                    builder = DesignBriefBuilder(db=db)
                    design_brief = builder.build(
                        project_type=context.proposal_type,
                        industry=industry,
                        target_slides=target_slides,
                    )

                    # 디자인 브리프를 context.extras에 저장
                    # → DesignAgentStep 또는 ContentGenerationStep에서 활용
                    context.extras["think_tank_design_brief"] = design_brief.to_dict()

                    logger.info(
                        f"디자인 브리프 생성 완료:\n{design_brief.summary()}"
                    )

                    # 섹션별 슬라이드 통계 (콘텐츠 생성 시 장표 배분 참조용)
                    section_stats = retrieval.get_section_structure_stats(
                        project_type=context.proposal_type,
                    )
                    if section_stats:
                        context.extras["section_structure_stats"] = section_stats

                except Exception as e:
                    logger.warning(f"디자인 브리프 생성 실패: {e} — 레퍼런스 검색만 사용")

            logger.info(
                f"Think Tank 검색 완료: "
                f"유사 레퍼런스 {len(context.similar_references)}건, "
                f"디자인 패턴 {len(context.design_patterns)}건, "
                f"콘텐츠 패턴 {len(context.content_patterns)}건"
                f"{', 디자인 브리프 생성됨' if 'think_tank_design_brief' in context.extras else ''}"
            )
        except ImportError:
            logger.warning("Think Tank 모듈 미설치 — 건너뜀")
        except Exception as e:
            logger.warning(f"Think Tank 검색 실패: {e} — 건너뜀")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 5: 콘텐츠 생성
# ═══════════════════════════════════════════════════════════════

class ContentGenerationStep(PipelineStep):
    """제안서 콘텐츠 생성 (Claude API - Impact-8 Framework)"""

    name = "content_generation"
    description = "제안서 콘텐츠 생성 중 (Impact-8)"
    is_critical = True

    def __init__(self, api_key: str = ""):
        self._api_key = api_key

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        from ..agents.content_generator import ContentGenerator

        api_key = self._api_key or context.api_key
        generator = ContentGenerator(api_key=api_key)

        # 입력 데이터 구성
        input_data = {
            "rfp_analysis": context.rfp_analysis,
            "company_data": context.company_data,
            "project_name": context.project_name,
            "client_name": context.client_name,
            "submission_date": context.submission_date,
            "proposal_type": context.proposal_type,
        }

        # Phase 3/2에서 추가된 데이터가 있으면 포함
        if context.ip_research is not None:
            input_data["ip_research"] = context.ip_research
        if context.similar_references:
            input_data["similar_references"] = context.similar_references
        if context.design_patterns:
            input_data["design_patterns"] = context.design_patterns
        if context.content_patterns:
            input_data["content_patterns"] = context.content_patterns

        # ★ 싱크탱크 디자인 브리프 (장표 배분/레이아웃 규칙)
        tt_brief = context.extras.get("think_tank_design_brief")
        if tt_brief:
            input_data["design_brief"] = tt_brief

        # 섹션별 슬라이드 통계 (레퍼런스 기반 장표 수 가이드)
        section_stats = context.extras.get("section_structure_stats")
        if section_stats:
            input_data["section_structure_stats"] = section_stats

        proposal_content = await generator.execute(
            input_data=input_data,
            progress_callback=lambda p: context.progress_callback({
                "phase": "content_generation",
                "sub_step": p.get("step"),
                "sub_total": p.get("total"),
                "message": p.get("message", "생성 중..."),
            }) if context.progress_callback else None,
        )

        context.proposal_content = proposal_content
        logger.info("제안서 콘텐츠 생성 완료")
        return context


# ═══════════════════════════════════════════════════════════════
# Step 6: 이미지 파이프라인
# ═══════════════════════════════════════════════════════════════

class ImagePipelineStep(PipelineStep):
    """이미지 파이프라인 - IMG_PH를 실제 이미지로 교체"""

    name = "image_pipeline"
    description = "이미지 수급 중"
    is_critical = False

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        try:
            from ..image_pipeline.manager import ImagePipelineManager

            web_search = step_config.get_option("web_search", True)
            ai_generation = step_config.get_option("ai_generation", False)
            diagram_rendering = step_config.get_option("diagram_rendering", True)

            manager = ImagePipelineManager(
                web_search_enabled=web_search,
                ai_generation_enabled=ai_generation,
                diagram_rendering_enabled=diagram_rendering,
            )

            if context.proposal_content:
                requests = manager.extract_placeholders_from_content(context.proposal_content)
                if requests:
                    results = await manager.process_requests(requests)
                    context.image_map = {
                        r.placeholder_id: r.file_path
                        for r in results.values()
                        if r.success and r.file_path
                    }
                    logger.info(f"이미지 파이프라인: {len(context.image_map)}개 이미지 수급")
                else:
                    logger.info("이미지 플레이스홀더 없음")

        except ImportError:
            logger.warning("이미지 파이프라인 모듈 미설치 - 건너뜀")
        except Exception as e:
            logger.warning(f"이미지 파이프라인 실패: {e} - 건너뜀")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 7: 검증
# ═══════════════════════════════════════════════════════════════

class ValidationStep(PipelineStep):
    """생성된 콘텐츠 검증"""

    name = "validation"
    description = "콘텐츠 검증 중"
    is_critical = False

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        if context.proposal_content is None:
            logger.warning("검증할 콘텐츠가 없습니다")
            return context

        content = context.proposal_content
        issues = []

        # 기본 검증
        teaser_slides = len(content.teaser.slides) if content.teaser else 0
        phase_slides = sum(len(p.slides) for p in content.phases)
        total_slides = teaser_slides + phase_slides

        if total_slides < 20:
            issues.append(f"슬라이드 수 부족: {total_slides}장 (최소 20장 권장)")

        if not content.project_name:
            issues.append("프로젝트명이 비어있습니다")

        if not content.slogan:
            issues.append("슬로건이 비어있습니다")

        # Win Theme 검증
        if not content.win_themes:
            issues.append("Win Theme이 설정되지 않았습니다")

        context.validation_result = {
            "total_slides": total_slides,
            "teaser_slides": teaser_slides,
            "phase_slides": phase_slides,
            "issues": issues,
            "is_valid": len(issues) == 0,
        }

        if issues:
            for issue in issues:
                logger.warning(f"[검증] {issue}")
        else:
            logger.info(f"콘텐츠 검증 통과: {total_slides}장")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 8: 디자인 에이전트 (싱크탱크 + Gamma 병합 → MergedDesignBrief)
# ═══════════════════════════════════════════════════════════════

class DesignAgentStep(PipelineStep):
    """디자인 에이전트 — 싱크탱크 디자인 브리프 + Gamma 테마 병합

    ThinkTankRetrievalStep에서 생성한 design_brief를 가져오고,
    (선택적) Gamma 테마 추천과 병합하여 MergedDesignBrief를 생성합니다.

    이 MergedDesignBrief는:
        - slide_kit의 동적 테마 등록에 사용 (register_theme + apply_theme)
        - 생성 스크립트가 참조하는 디자인 규칙의 최종 버전
    """

    name = "design_agent"
    description = "디자인 에이전트 (싱크탱크+Gamma 병합) 실행 중"
    is_critical = False

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        try:
            from ..agents.design_agent import DesignAgent

            agent = DesignAgent()

            # ── 1. 싱크탱크 디자인 브리프 (ThinkTankRetrievalStep에서 이미 생성됨) ──
            tt_brief = context.extras.get("think_tank_design_brief")

            if not tt_brief:
                # ThinkTankRetrievalStep을 건너뛴 경우, 직접 조회
                industry = None
                if context.rfp_analysis and hasattr(context.rfp_analysis, 'industry'):
                    industry = context.rfp_analysis.industry

                target_slides = context.extras.get("target_slides", 70)
                tt_brief = agent.get_think_tank_brief(
                    project_type=context.proposal_type or "event",
                    industry=industry or "",
                    target_slides=target_slides,
                )

            # ── 2. Gamma 테마 (extras에 있으면 활용) ──
            gamma_themes_data = context.extras.get("gamma_themes_data")
            project_keywords = context.extras.get("project_keywords", [])

            gamma_recs = []
            if gamma_themes_data:
                gamma_recs = agent.interpret_gamma_themes(
                    gamma_themes_data, project_keywords
                )

            # ── 3. 사용자 커스텀 컬러 (디자인 오버라이드) ──
            custom_colors = context.extras.get("custom_colors")

            # ── 4. 병합 → MergedDesignBrief ──
            industry = ""
            if context.rfp_analysis and hasattr(context.rfp_analysis, 'industry'):
                industry = context.rfp_analysis.industry or ""

            merged_brief = agent.merge(
                tt_brief=tt_brief,
                gamma_recs=gamma_recs,
                project_name=context.project_name,
                project_type=context.proposal_type or "event",
                industry=industry,
                custom_colors=custom_colors,
            )

            # ── 5. 결과 저장 ──
            context.extras["merged_design_brief"] = merged_brief.model_dump()
            context.design_bridge_result = merged_brief

            logger.info(
                f"DesignAgent 완료: "
                f"theme={merged_brief.theme_name}, "
                f"confidence={merged_brief.confidence:.2f}, "
                f"colors={list(merged_brief.colors.keys())}"
            )

        except ImportError as e:
            logger.warning(f"DesignAgent 모듈 미설치 - 건너뜀: {e}")
        except Exception as e:
            logger.warning(f"DesignAgent 실패: {e} - 기본값 사용")

        return context


# ═══════════════════════════════════════════════════════════════
# Step 9: 디자인 브릿지 (Gamma MCP 연동 — 선택적)
# ═══════════════════════════════════════════════════════════════

class DesignBridgeStep(PipelineStep):
    """디자인 브릿지 — Gamma MCP를 통한 디자인 워싱 (선택적)

    PPTX를 텍스트로 추출 → Gamma generate() 파라미터를 준비하여
    context.extras에 저장합니다. 실제 Gamma API 호출은 Claude Code가 수행.

    이 Step은 PPTX가 이미 생성된 후에만 의미가 있으므로,
    자동 파이프라인에서는 보통 비활성화 상태입니다.
    """

    name = "design_bridge"
    description = "Gamma 디자인 워싱 준비 중"
    is_critical = False

    async def execute(
        self,
        context: PipelineContext,
        step_config: StepConfig,
    ) -> PipelineContext:
        try:
            from ..integrations.design_bridge import GammaMCPBridge

            # PPTX 경로 확인 (extras에서 가져오거나 output_path 구성)
            pptx_path = context.extras.get("output_pptx_path")
            if not pptx_path:
                logger.info("디자인 브릿지: PPTX 경로 미설정 - 건너뜀")
                return context

            pptx_path = Path(pptx_path) if isinstance(pptx_path, str) else pptx_path

            if not pptx_path.exists():
                logger.warning(f"디자인 브릿지: PPTX 파일 없음 - {pptx_path}")
                return context

            # PPTX에서 텍스트 추출 → Gamma inputText 변환
            input_text = GammaMCPBridge.prepare_content_for_gamma(pptx_path)
            if not input_text:
                logger.warning("디자인 브릿지: PPTX 텍스트 추출 실패")
                return context

            # MergedDesignBrief가 있으면 Gamma 파라미터에 반영
            merged_brief = context.extras.get("merged_design_brief")
            brief_obj = None
            if merged_brief:
                from ..schemas.design_schema import MergedDesignBrief
                brief_obj = MergedDesignBrief(**merged_brief)

            # Gamma generate() 파라미터 준비
            num_cards = context.extras.get("target_slides")
            export_as = step_config.get_option("export_as", "pptx")

            gamma_params = GammaMCPBridge.build_gamma_params(
                input_text=input_text,
                brief=brief_obj,
                num_cards=num_cards,
                export_as=export_as,
            )

            # 결과 저장 — Claude Code가 이 파라미터로 Gamma MCP 호출
            context.extras["gamma_params"] = gamma_params
            context.extras["gamma_input_text"] = input_text

            logger.info(
                f"디자인 브릿지 완료: "
                f"inputText {len(input_text)}자, "
                f"numCards={gamma_params.get('numCards', 'auto')}, "
                f"themeId={gamma_params.get('themeId', 'default')}"
            )

        except ImportError:
            logger.warning("디자인 브릿지 모듈 미설치 - 건너뜀")
        except Exception as e:
            logger.warning(f"디자인 브릿지 실패: {e} - 건너뜀")

        return context


# ═══════════════════════════════════════════════════════════════
# 팩토리: 기본 파이프라인 구성
# ═══════════════════════════════════════════════════════════════

def build_default_pipeline(api_key: str = "") -> "PipelineEngine":
    """
    기본 파이프라인 구성 (v5.0)

    실행 순서:
        1. document_parsing  — RFP 문서 파싱 (필수)
        2. rfp_analysis      — Claude RFP 분석 (필수)
        3. ip_research       — 웹 검색 기반 IP/시장 리서치
        4. think_tank        — 싱크탱크 레퍼런스 검색 + 디자인 브리프
        5. content_generation — Impact-8 콘텐츠 생성 (필수)
        6. image_pipeline    — IMG_PH → 실제 이미지 교체
        7. validation        — 콘텐츠 검증
        8. design_agent      — 싱크탱크+Gamma 병합 → MergedDesignBrief
        9. design_bridge     — Gamma 디자인 워싱 준비 (선택적)

    pipeline_config.yaml에서 disabled된 단계는 엔진이 자동으로 건너뜁니다.

    Args:
        api_key: Anthropic API 키

    Returns:
        PipelineEngine: 구성된 파이프라인 엔진
    """
    from .engine import PipelineEngine

    engine = PipelineEngine()

    engine.register_step("document_parsing", DocumentParsingStep())
    engine.register_step("rfp_analysis", RFPAnalysisStep(api_key=api_key))
    engine.register_step("ip_research", IPResearchStep(api_key=api_key))
    engine.register_step("think_tank_retrieval", ThinkTankRetrievalStep())
    engine.register_step("content_generation", ContentGenerationStep(api_key=api_key))
    engine.register_step("image_pipeline", ImagePipelineStep())
    engine.register_step("validation", ValidationStep())
    engine.register_step("design_agent", DesignAgentStep())
    engine.register_step("design_bridge", DesignBridgeStep())

    return engine
