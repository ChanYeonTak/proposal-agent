"""
파이프라인 엔진 (v4.0)

등록된 단계를 순차 실행하며, 각 단계는 PipelineContext를 통해 데이터를 주고받습니다.
pipeline_config.yaml 에서 enabled: false 인 단계는 자동으로 건너뜁니다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, TYPE_CHECKING

from .config import PipelineConfig, get_pipeline_config
from ..utils.logger import get_logger

if TYPE_CHECKING:
    from .steps import PipelineStep

logger = get_logger("pipeline_engine")


@dataclass
class PipelineContext:
    """
    파이프라인 단계 간 데이터 전달 컨텍스트

    각 단계가 실행되면서 필요한 데이터를 여기에 적재합니다.
    후속 단계는 선행 단계가 적재한 데이터를 활용합니다.
    """

    # ─── 입력 파라미터 ────────────────────────────────────────
    rfp_path: Optional[Path] = None
    company_data_path: Optional[Path] = None
    project_name: str = ""
    client_name: str = ""
    submission_date: str = ""
    proposal_type: Optional[str] = None
    api_key: str = ""

    # ─── 단계별 산출물 ────────────────────────────────────────
    parsed_rfp: Optional[Dict[str, Any]] = None          # document_parsing
    company_data: Dict[str, Any] = field(default_factory=dict)
    rfp_analysis: Any = None                              # rfp_analysis → RFPAnalysis
    ip_research: Any = None                               # ip_research → IPResearchResult (Phase 3)
    similar_references: List[Any] = field(default_factory=list)   # think_tank
    design_patterns: List[Any] = field(default_factory=list)      # think_tank
    content_patterns: List[Any] = field(default_factory=list)     # think_tank
    proposal_content: Any = None                          # content_generation → ProposalContent
    image_map: Dict[str, Path] = field(default_factory=dict)      # image_pipeline
    validation_result: Optional[Dict[str, Any]] = None    # validation
    design_bridge_result: Any = None                      # design_bridge

    # ─── 메타 정보 ────────────────────────────────────────────
    progress_callback: Optional[Callable] = None
    executed_steps: List[str] = field(default_factory=list)
    skipped_steps: List[str] = field(default_factory=list)
    errors: Dict[str, str] = field(default_factory=dict)

    # ─── 확장 데이터 (미래 모듈용) ────────────────────────────
    extras: Dict[str, Any] = field(default_factory=dict)

    def set_extra(self, key: str, value: Any) -> None:
        """확장 데이터 저장"""
        self.extras[key] = value

    def get_extra(self, key: str, default: Any = None) -> Any:
        """확장 데이터 조회"""
        return self.extras.get(key, default)


class PipelineEngine:
    """
    파이프라인 엔진

    단계를 등록하고 설정에 따라 순차 실행합니다.
    """

    def __init__(self, config: Optional[PipelineConfig] = None):
        self.config = config or get_pipeline_config()
        self._steps: List[tuple[str, "PipelineStep"]] = []

    def register_step(self, name: str, step: "PipelineStep") -> "PipelineEngine":
        """
        파이프라인 단계 등록

        Args:
            name: 단계 이름 (pipeline_config.yaml의 step 이름과 매칭)
            step: PipelineStep 인스턴스

        Returns:
            self (체이닝 지원)
        """
        self._steps.append((name, step))
        return self

    async def execute(self, context: PipelineContext) -> PipelineContext:
        """
        등록된 단계를 순차 실행

        enabled된 단계만 실행하며, 각 단계의 결과는 context에 적재됩니다.

        Args:
            context: 초기 컨텍스트

        Returns:
            PipelineContext: 모든 단계 실행 후의 컨텍스트
        """
        total_enabled = sum(
            1 for name, _ in self._steps
            if self.config.is_step_enabled(name)
        )
        current_step = 0

        logger.info(f"파이프라인 시작: {total_enabled}/{len(self._steps)} 단계 활성화")

        for name, step in self._steps:
            if not self.config.is_step_enabled(name):
                context.skipped_steps.append(name)
                logger.info(f"[SKIP] {name} — disabled")
                continue

            current_step += 1
            step_config = self.config.get_step_config(name)

            # 진행 상황 콜백
            if context.progress_callback:
                context.progress_callback({
                    "phase": name,
                    "step": current_step,
                    "total": total_enabled,
                    "message": f"{step.description} ...",
                })

            logger.info(f"[{current_step}/{total_enabled}] {name}: {step.description}")

            try:
                context = await step.execute(context, step_config)
                context.executed_steps.append(name)
                logger.info(f"[DONE] {name}")
            except Exception as e:
                error_msg = f"{name} 실행 실패: {e}"
                logger.error(error_msg)
                context.errors[name] = str(e)

                if step.is_critical:
                    logger.error(f"[CRITICAL] {name} — 파이프라인 중단")
                    raise
                else:
                    logger.warning(f"[WARN] {name} — 비필수 단계, 계속 진행")
                    context.skipped_steps.append(name)

        logger.info(
            f"파이프라인 완료: "
            f"실행 {len(context.executed_steps)}, "
            f"스킵 {len(context.skipped_steps)}, "
            f"오류 {len(context.errors)}"
        )
        return context

    def get_step_names(self) -> List[str]:
        """등록된 단계 이름 목록"""
        return [name for name, _ in self._steps]

    def get_enabled_steps(self) -> List[str]:
        """활성화된 단계 이름 목록"""
        return [
            name for name, _ in self._steps
            if self.config.is_step_enabled(name)
        ]
