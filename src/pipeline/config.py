"""
파이프라인 설정 (v4.0)

pipeline_config.yaml 을 로드하여 각 단계의 ON/OFF 및 옵션을 관리합니다.
환경변수 오버라이드를 지원합니다.
"""

import os
from pathlib import Path
from typing import Any, Dict, Optional

import yaml
from pydantic import BaseModel, Field

from ..utils.logger import get_logger

logger = get_logger("pipeline_config")


class StepConfig(BaseModel):
    """개별 파이프라인 단계 설정"""

    enabled: bool = True
    options: Dict[str, Any] = Field(default_factory=dict)

    def get_option(self, key: str, default: Any = None) -> Any:
        """옵션 값 조회"""
        return self.options.get(key, default)


class ThinkTankConfig(BaseModel):
    """Think Tank 전역 설정"""

    enabled: bool = False
    db_path: str = "think_tank/references.db"


class PipelineConfig(BaseModel):
    """파이프라인 전체 설정"""

    steps: Dict[str, StepConfig] = Field(default_factory=dict)
    think_tank: ThinkTankConfig = Field(default_factory=ThinkTankConfig)

    def is_step_enabled(self, step_name: str) -> bool:
        """특정 단계가 활성화되어 있는지 확인"""
        # 환경변수 오버라이드: PIPELINE_STEP_{NAME}_ENABLED=true/false
        env_key = f"PIPELINE_STEP_{step_name.upper()}_ENABLED"
        env_val = os.getenv(env_key)
        if env_val is not None:
            return env_val.lower() in ("true", "1", "yes")

        step = self.steps.get(step_name)
        if step is None:
            return True  # 설정에 없으면 기본 활성화
        return step.enabled

    def get_step_config(self, step_name: str) -> StepConfig:
        """특정 단계 설정 반환 (없으면 기본값)"""
        return self.steps.get(step_name, StepConfig())

    def get_step_option(self, step_name: str, key: str, default: Any = None) -> Any:
        """특정 단계의 옵션 값 조회"""
        step = self.steps.get(step_name)
        if step is None:
            return default
        return step.get_option(key, default)


def _parse_yaml_step(raw: Dict[str, Any]) -> StepConfig:
    """YAML에서 파싱한 단계 설정을 StepConfig로 변환"""
    enabled = raw.pop("enabled", True)
    # 나머지 키는 모두 options로
    return StepConfig(enabled=enabled, options=raw)


def load_pipeline_config(config_path: Optional[Path] = None) -> PipelineConfig:
    """
    pipeline_config.yaml 로드

    Args:
        config_path: YAML 파일 경로. None이면 기본 경로 사용.

    Returns:
        PipelineConfig: 파싱된 설정
    """
    if config_path is None:
        # 기본 경로: 프로젝트 루트/config/pipeline_config.yaml
        config_path = Path(__file__).parent.parent.parent / "config" / "pipeline_config.yaml"

    if not config_path.exists():
        logger.info(f"파이프라인 설정 파일 없음: {config_path} — 기본 설정 사용")
        return _get_default_config()

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        if not raw or "pipeline" not in raw:
            logger.warning("pipeline_config.yaml 형식 오류 — 기본 설정 사용")
            return _get_default_config()

        pipeline_raw = raw["pipeline"]

        # Steps 파싱
        steps = {}
        for step_name, step_raw in pipeline_raw.get("steps", {}).items():
            if isinstance(step_raw, dict):
                steps[step_name] = _parse_yaml_step(dict(step_raw))
            elif isinstance(step_raw, bool):
                steps[step_name] = StepConfig(enabled=step_raw)

        # Think Tank 파싱
        tt_raw = pipeline_raw.get("think_tank", {})
        think_tank = ThinkTankConfig(**tt_raw) if isinstance(tt_raw, dict) else ThinkTankConfig()

        config = PipelineConfig(steps=steps, think_tank=think_tank)
        logger.info(f"파이프라인 설정 로드 완료: {sum(1 for s in steps.values() if s.enabled)}/{len(steps)} 단계 활성화")
        return config

    except Exception as e:
        logger.error(f"파이프라인 설정 로드 실패: {e} — 기본 설정 사용")
        return _get_default_config()


def _get_default_config() -> PipelineConfig:
    """기존 v3.6과 동일한 기본 설정 (하위 호환)"""
    return PipelineConfig(
        steps={
            "document_parsing": StepConfig(enabled=True),
            "rfp_analysis": StepConfig(enabled=True),
            "ip_research": StepConfig(enabled=False),
            "think_tank_retrieval": StepConfig(enabled=False),
            "content_generation": StepConfig(enabled=True),
            "image_pipeline": StepConfig(enabled=False),
            "validation": StepConfig(enabled=True),
            "design_bridge": StepConfig(enabled=False),
        },
        think_tank=ThinkTankConfig(enabled=False),
    )


# ─── 싱글톤 ────────────────────────────────────────────────
_pipeline_config: Optional[PipelineConfig] = None


def get_pipeline_config(config_path: Optional[Path] = None) -> PipelineConfig:
    """싱글톤 파이프라인 설정 반환"""
    global _pipeline_config
    if _pipeline_config is None:
        _pipeline_config = load_pipeline_config(config_path)
    return _pipeline_config


def reset_pipeline_config() -> None:
    """파이프라인 설정 리셋 (테스트용)"""
    global _pipeline_config
    _pipeline_config = None
