"""
산업 프로파일 기본 클래스 (v4.0)

모든 산업 프로파일의 베이스. Phase별 콘텐츠 깊이/구조 가이드를 정의합니다.
"""

from __future__ import annotations

from typing import Dict, List, Optional

from pydantic import BaseModel, Field


class PhaseGuide(BaseModel):
    """Phase별 콘텐츠 생성 가이드"""
    phase_number: int = 0
    structure_type: str = ""          # channel, pack, module, section 등
    depth_guide: str = ""             # 콘텐츠 깊이 지침
    required_elements: List[str] = Field(default_factory=list)
    slides_per_item: int = 1          # 항목당 슬라이드 수
    visual_requirements: List[str] = Field(default_factory=list)
    persuasion_framework: str = ""    # 설득 프레임워크 (CHECK POINT 등)


class IndustryProfile(BaseModel):
    """산업별 콘텐츠 프로파일"""

    # 기본 정보
    name: str = ""
    description: str = ""
    industry_type: str = ""           # game_event, marketing_pr, it_system, etc.

    # Phase별 가이드
    phase_guides: Dict[int, PhaseGuide] = Field(default_factory=dict)

    # 핵심 용어/키워드
    industry_keywords: List[str] = Field(default_factory=list)

    # 설득 구조
    default_persuasion: str = "CEI"   # Claim-Evidence-Impact

    # Action Plan 분해 방식
    action_plan_structure: str = ""   # pack, channel, module, phase 등
    action_plan_items: List[str] = Field(default_factory=list)

    # 콘텐츠 깊이 지침
    content_depth_guide: str = ""

    # 시각 요소 가이드
    visual_guide: str = ""

    def get_phase_guide(self, phase_number: int) -> Optional[PhaseGuide]:
        """Phase별 가이드 반환"""
        return self.phase_guides.get(phase_number)

    def get_action_plan_context(self) -> str:
        """Action Plan 프롬프트 컨텍스트 생성"""
        if not self.action_plan_items:
            return ""

        lines = [
            f"\n## 산업 프로파일: {self.name}",
            f"Action Plan 구조: {self.action_plan_structure}",
            f"\n### 필수 항목 ({len(self.action_plan_items)}개):",
        ]
        for item in self.action_plan_items:
            lines.append(f"- {item}")

        if self.content_depth_guide:
            lines.append(f"\n### 콘텐츠 깊이 지침:")
            lines.append(self.content_depth_guide)

        return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════
# 프로파일 레지스트리
# ═══════════════════════════════════════════════════════════════

_PROFILE_REGISTRY: Dict[str, IndustryProfile] = {}


def register_profile(profile: IndustryProfile) -> None:
    """프로파일 등록"""
    _PROFILE_REGISTRY[profile.industry_type] = profile


def get_industry_profile(industry_type: str) -> Optional[IndustryProfile]:
    """산업 프로파일 반환 (없으면 None)"""
    # 지연 로딩
    if not _PROFILE_REGISTRY:
        _load_all_profiles()
    return _PROFILE_REGISTRY.get(industry_type)


def list_profiles() -> List[str]:
    """등록된 프로파일 목록"""
    if not _PROFILE_REGISTRY:
        _load_all_profiles()
    return list(_PROFILE_REGISTRY.keys())


def _load_all_profiles():
    """모든 프로파일 로드"""
    from .game_event import GAME_EVENT_PROFILE
    from .marketing_pr import MARKETING_PR_PROFILE
    from .it_system import IT_SYSTEM_PROFILE
    from .public_project import PUBLIC_PROJECT_PROFILE

    register_profile(GAME_EVENT_PROFILE)
    register_profile(MARKETING_PR_PROFILE)
    register_profile(IT_SYSTEM_PROFILE)
    register_profile(PUBLIC_PROJECT_PROFILE)
