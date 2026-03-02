"""
제안서 생성 오케스트레이터 (v4.0 - Pipeline Engine + Impact-8 Framework)

전체 워크플로우 조율: 파이프라인 기반 실행
v4.0: PipelineEngine 도입으로 모든 단계 ON/OFF 가능

하위 호환: pipeline_config.yaml 없으면 기존 v3.6 동작 유지
"""

import json
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from ..parsers.pdf_parser import PDFParser
from ..parsers.docx_parser import DOCXParser
from ..agents.rfp_analyzer import RFPAnalyzer
from ..agents.content_generator import ContentGenerator
from ..schemas.proposal_schema import ProposalContent, ProposalType
from ..schemas.rfp_schema import RFPAnalysis
from ..pipeline.engine import PipelineContext
from ..pipeline.steps import build_default_pipeline
from ..utils.logger import get_logger
from config.settings import get_settings

logger = get_logger("proposal_orchestrator")


class ProposalOrchestrator:
    """
    제안서 콘텐츠 생성 오케스트레이터 (v4.0 - Pipeline Engine)

    v4.0 변경사항:
    - PipelineEngine 기반 실행 (ON/OFF 제어)
    - 하위 호환: pipeline_config.yaml 없으면 기존 동작 유지
    - PipelineContext로 단계 간 데이터 전달
    """

    def __init__(self, api_key: Optional[str] = None):
        settings = get_settings()
        self.api_key = api_key or settings.anthropic_api_key

        # v3.6 레거시 호환용 (직접 호출 시 사용)
        self.pdf_parser = PDFParser()
        self.docx_parser = DOCXParser()
        self.rfp_analyzer = RFPAnalyzer(api_key=self.api_key)
        self.content_generator = ContentGenerator(api_key=self.api_key)

    async def execute(
        self,
        rfp_path: Path,
        company_data_path: Optional[Path] = None,
        project_name: str = "",
        client_name: str = "",
        submission_date: str = "",
        proposal_type: Optional[str] = None,
        progress_callback: Optional[Callable] = None,
    ) -> ProposalContent:
        """
        전체 제안서 콘텐츠 생성 워크플로우 실행 (v4.0 Pipeline)

        Args:
            rfp_path: RFP 문서 경로
            company_data_path: 회사 정보 JSON 경로
            project_name: 프로젝트명 (미입력시 RFP에서 추출)
            client_name: 발주처명 (미입력시 RFP에서 추출)
            submission_date: 제출일
            proposal_type: 제안서 유형
            progress_callback: 진행 상황 콜백

        Returns:
            ProposalContent: 생성된 제안서 콘텐츠
        """
        try:
            return await self._execute_pipeline(
                rfp_path=rfp_path,
                company_data_path=company_data_path,
                project_name=project_name,
                client_name=client_name,
                submission_date=submission_date,
                proposal_type=proposal_type,
                progress_callback=progress_callback,
            )
        except Exception as e:
            logger.warning(f"파이프라인 실행 실패, 레거시 모드 시도: {e}")
            # 파이프라인 자체에 문제가 있으면 레거시 모드로 폴백
            return await self._execute_legacy(
                rfp_path=rfp_path,
                company_data_path=company_data_path,
                project_name=project_name,
                client_name=client_name,
                submission_date=submission_date,
                proposal_type=proposal_type,
                progress_callback=progress_callback,
            )

    async def _execute_pipeline(
        self,
        rfp_path: Path,
        company_data_path: Optional[Path],
        project_name: str,
        client_name: str,
        submission_date: str,
        proposal_type: Optional[str],
        progress_callback: Optional[Callable],
    ) -> ProposalContent:
        """v4.0 파이프라인 엔진 기반 실행"""

        # 파이프라인 구성
        engine = build_default_pipeline(api_key=self.api_key)

        # 초기 컨텍스트 구성
        context = PipelineContext(
            rfp_path=rfp_path,
            company_data_path=company_data_path,
            project_name=project_name,
            client_name=client_name,
            submission_date=submission_date,
            proposal_type=proposal_type,
            api_key=self.api_key,
            progress_callback=progress_callback,
        )

        # 파이프라인 실행
        context = await engine.execute(context)

        if context.proposal_content is None:
            raise ValueError("파이프라인 실행 후 콘텐츠가 생성되지 않았습니다")

        # 슬라이드 수 로깅
        content = context.proposal_content
        total_slides = len(content.teaser.slides) if content.teaser else 0
        total_slides += sum(len(p.slides) for p in content.phases)
        logger.info(f"제안서 콘텐츠 생성 완료: {total_slides}장")

        # 검증 결과 로깅
        if context.validation_result and not context.validation_result.get("is_valid"):
            for issue in context.validation_result.get("issues", []):
                logger.warning(f"[검증 이슈] {issue}")

        return content

    async def _execute_legacy(
        self,
        rfp_path: Path,
        company_data_path: Optional[Path],
        project_name: str,
        client_name: str,
        submission_date: str,
        proposal_type: Optional[str],
        progress_callback: Optional[Callable],
    ) -> ProposalContent:
        """
        v3.6 레거시 실행 (하위 호환)

        pipeline 모듈에 문제가 있을 때의 폴백 경로입니다.
        """
        try:
            # Step 1: 문서 파싱
            if progress_callback:
                progress_callback({
                    "phase": "parsing",
                    "step": 1,
                    "total": 4,
                    "message": "RFP 문서 파싱 중...",
                })

            parsed_rfp = self._parse_document(rfp_path)
            logger.info(f"RFP 파싱 완료: {len(parsed_rfp.get('raw_text', ''))} 문자")

            # Step 2: 회사 데이터 로드
            company_data = {}
            if company_data_path:
                company_data = self._load_company_data(company_data_path)

            # Step 3: RFP 분석 (Claude)
            if progress_callback:
                progress_callback({
                    "phase": "analysis",
                    "step": 2,
                    "total": 4,
                    "message": "RFP 분석 중 (Claude)...",
                })

            rfp_analysis = await self.rfp_analyzer.execute(
                input_data=parsed_rfp,
                progress_callback=lambda p: progress_callback({
                    "phase": "analysis",
                    "sub_step": p["step"],
                    "sub_total": p["total"],
                    "message": p["message"],
                }) if progress_callback else None,
            )

            # 프로젝트명/발주처명 결정
            final_project_name = project_name or rfp_analysis.project_name
            final_client_name = client_name or rfp_analysis.client_name

            logger.info(f"RFP 분석 완료: {final_project_name} ({final_client_name})")

            # Step 4: 콘텐츠 생성 (Claude) - Impact-8 Framework
            if progress_callback:
                progress_callback({
                    "phase": "generation",
                    "step": 3,
                    "total": 4,
                    "message": "제안서 콘텐츠 생성 중 (Impact-8 Framework)...",
                })

            proposal_content = await self.content_generator.execute(
                input_data={
                    "rfp_analysis": rfp_analysis,
                    "company_data": company_data,
                    "project_name": final_project_name,
                    "client_name": final_client_name,
                    "submission_date": submission_date,
                    "proposal_type": proposal_type,
                },
                progress_callback=lambda p: progress_callback({
                    "phase": "generation",
                    "sub_step": p["step"],
                    "sub_total": p["total"],
                    "message": p["message"],
                }) if progress_callback else None,
            )

            if progress_callback:
                progress_callback({
                    "phase": "complete",
                    "step": 4,
                    "total": 4,
                    "message": "콘텐츠 생성 완료!",
                })

            # 슬라이드 수 계산
            total_slides = len(proposal_content.teaser.slides) if proposal_content.teaser else 0
            total_slides += sum(len(p.slides) for p in proposal_content.phases)

            logger.info(f"제안서 콘텐츠 생성 완료 (레거시): {total_slides}장")
            return proposal_content

        except Exception as e:
            logger.error(f"제안서 생성 실패: {e}")
            raise

    def _parse_document(self, file_path: Path) -> Dict[str, Any]:
        """파일 확장자에 따라 적절한 파서 선택"""
        suffix = file_path.suffix.lower()

        if suffix == ".pdf":
            return self.pdf_parser.parse(file_path)
        elif suffix in [".docx", ".doc"]:
            return self.docx_parser.parse(file_path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {suffix}")

    def _load_company_data(self, data_path: Path) -> Dict[str, Any]:
        """회사 데이터 로드"""
        if not data_path.exists():
            logger.warning(f"회사 데이터 파일 없음: {data_path}")
            return {}

        try:
            return json.loads(data_path.read_text(encoding="utf-8"))
        except Exception as e:
            logger.error(f"회사 데이터 로드 실패: {e}")
            return {}

    def save_content_json(
        self, content: ProposalContent, output_path: Path
    ) -> None:
        """콘텐츠를 JSON 파일로 저장"""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            content.model_dump_json(indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        logger.info(f"콘텐츠 JSON 저장: {output_path}")

    def get_proposal_summary(self, content: ProposalContent) -> Dict[str, Any]:
        """제안서 요약 정보 반환"""
        teaser_slides = len(content.teaser.slides) if content.teaser else 0
        phase_slides = {
            f"Phase {p.phase_number}": len(p.slides)
            for p in content.phases
        }
        total_slides = teaser_slides + sum(phase_slides.values())

        return {
            "project_name": content.project_name,
            "client_name": content.client_name,
            "proposal_type": content.proposal_type.value,
            "slogan": content.slogan,
            "one_sentence_pitch": content.one_sentence_pitch,
            "key_differentiators": content.key_differentiators,
            "total_slides": total_slides,
            "teaser_slides": teaser_slides,
            "phase_slides": phase_slides,
            "design_style": content.design_style,
        }
