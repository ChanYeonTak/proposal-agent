"""
디자인 에이전트 스키마 (v2.0)

싱크탱크 레퍼런스 + Gamma 테마 추천을 병합한 디자인 결정 데이터 모델.
DesignAgent가 생성하고, 생성 스크립트가 소비합니다.

v2.0: ContentToneBrief 추가 — 콘텐츠 톤/감성/IP 깊이 가이드
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from pydantic import BaseModel, Field


class ThemeRecommendation(BaseModel):
    """Gamma 테마 추천 후보"""
    theme_id: str = ""
    theme_name: str = ""
    tone_keywords: List[str] = Field(default_factory=list)     # ["modern", "bold", "dark"]
    color_keywords: List[str] = Field(default_factory=list)    # ["navy", "blue", "gold"]
    match_reason: str = ""                                      # 선정 이유
    match_score: float = 0.0                                    # 매칭 점수 (0~1)


class ImageStyleGuide(BaseModel):
    """이미지 스타일 가이드"""
    primary_style: str = "photorealistic"   # photorealistic / illustration / minimal / mixed
    keywords: List[str] = Field(default_factory=list)  # ["게임", "부스", "AGF"]
    source_preference: Dict[str, str] = Field(default_factory=dict)
    # {"photo": "pexels", "illustration": "gamma_ai", "diagram": "renderer"}


class ContentToneBrief(BaseModel):
    """
    콘텐츠 톤 브리프 — 생성 스크립트가 참조하는 글쓰기 톤 가이드

    싱크탱크 레퍼런스 분석 결과를 기반으로 생성됩니다.
    어떤 RFP든 해당 산업/유형에 맞는 톤 규칙을 자동 적용합니다.
    """

    # 감성 톤 레벨 (1~5)
    emotional_tone_level: int = 3
    # 1=사무적/기능적, 2=간결 전문가, 3=중립, 4=감성+전문, 5=풀 내러티브

    # 내러티브 프레이밍
    narrative_framing_style: str = "hybrid"
    # worldview_based / data_driven / emotion_led / hybrid
    core_metaphor: str = ""          # 핵심 메타포 (있을 경우)
    entry_hook_pattern: str = ""     # 진입 후크 패턴
    recurring_motif: str = ""        # 반복 모티프

    # IP/브랜드 깊이
    ip_depth_score: float = 0.0      # 0~1
    ip_lore_terms: List[str] = Field(default_factory=list)
    ip_community_terms: List[str] = Field(default_factory=list)

    # 프로그램 네이밍 스타일
    program_naming_style: str = "functional"
    # ip_narrative / functional / branded / hybrid
    program_naming_examples: List[str] = Field(default_factory=list)

    # Win Theme 스타일
    win_theme_style: str = "keyword_functional"
    # ip_worldview / keyword_functional / emotional_hook
    win_theme_examples: List[str] = Field(default_factory=list)

    # 텍스트 밀도 스타일
    text_density_style: str = "balanced"
    # minimal / balanced / rich / dense

    # 핵심: 톤 규칙 (생성 스크립트가 직접 참조)
    tone_rules: List[str] = Field(default_factory=list)
    # 예:
    # - "게임 IP의 세계관과 캐릭터를 제안서 전체에 자연스럽게 녹여라"
    # - "프로그램명에 IP 세계관 용어를 삽입하라"
    # - "제안서 전체에 하나의 내러티브 메타포를 관통시켜라"

    # 분석 근거
    source_analysis: str = ""


class MergedDesignBrief(BaseModel):
    """
    싱크탱크 + Gamma 병합 디자인 브리프

    싱크탱크 레퍼런스가 1순위, Gamma 추천이 2순위 보완.
    생성 스크립트가 이 브리프를 참조하여 디자인 결정을 내립니다.
    """

    # ── 프로젝트 메타 ──
    project_name: str = ""
    project_type: str = ""         # marketing_pr / event / game_event / it_system
    industry: str = ""             # think_tank.models.Industry 값

    # ── 싱크탱크 기반 (레퍼런스 데이터) ──
    section_weights: Dict[str, int] = Field(default_factory=dict)
    # {"HOOK": 4, "INSIGHT": 6, "CONCEPT": 8, "ACTION": 35, ...}

    component_targets: Dict[str, int] = Field(default_factory=dict)
    # {"HERO_IMAGE": 5, "IMG_PH": 18, "COLS": 8, "FLOW": 5, ...}

    layout_distribution: Dict[str, float] = Field(default_factory=dict)
    # {"complex_diagram": 0.42, "image_focused": 0.20, ...}

    background_schedule: List[str] = Field(default_factory=list)
    # ["gradient_dark", "white", "white", "dark", "white", "light", ...]

    visual_density_targets: Dict[str, float] = Field(default_factory=dict)
    # {"image_slides_pct": 0.28, "diagram_slides_pct": 0.38, "text_only_max_pct": 0.15}

    content_patterns: List[Dict[str, Any]] = Field(default_factory=list)
    # [{"type": "show_dont_tell", "slides": 3}, {"type": "data_narrative", "slides": 5}]

    program_templates: List[Dict[str, Any]] = Field(default_factory=list)
    # [{"name": "booth", "category": "booth_design", "mechanism": "..."}]

    # ── 콘텐츠 톤 (v2.0) ──
    content_tone: ContentToneBrief = Field(default_factory=ContentToneBrief)
    # 싱크탱크 레퍼런스 분석 기반 콘텐츠 톤 가이드

    # ── Gamma 보완 (테마/컬러/이미지) ──
    theme_name: str = "default_blue"
    colors: Dict[str, Tuple[int, int, int]] = Field(default_factory=dict)
    # {"primary": (0, 44, 95), "secondary": (0, 170, 210), "teal": (0, 161, 156), ...}

    font_weight_primary: str = "bold"     # bold / semibold / medium
    image_style: ImageStyleGuide = Field(default_factory=ImageStyleGuide)

    gamma_recommendations: List[ThemeRecommendation] = Field(default_factory=list)
    # Gamma 테마 후보 3개 (참고용)

    # ── 메타 ──
    source_references: List[str] = Field(default_factory=list)
    # 참조한 레퍼런스 문서명 목록

    confidence: float = 0.0
    # 레퍼런스 매칭 신뢰도 (0~1). 0 = 레퍼런스 없음 (Gamma only), 1 = 정확한 유형 매칭

    notes: str = ""
    # 디자인 결정 근거 메모


# ── 기본값 팩토리 ─────────────────────────────────

DEFAULT_COLORS = {
    "primary":   (0, 44, 95),       # #002C5F
    "secondary": (0, 170, 210),     # #00AAD2
    "teal":      (0, 161, 156),     # #00A19C
    "accent":    (230, 51, 18),     # #E63312
    "dark":      (33, 33, 33),      # #212121
    "light":     (245, 245, 245),   # #F5F5F5
}

DEFAULT_BACKGROUND_SCHEDULE = [
    "gradient_dark",   # 표지
    "white", "white",  # HOOK
    "dark",            # 섹션 구분자
    "white", "white", "light", "white",  # INSIGHT
    "dark",            # 섹션 구분자
    "dark", "white", "white",  # CONCEPT
    "dark",            # 섹션 구분자
    "white", "light", "white", "white", "light",  # ACTION
    "white", "white", "white", "light", "white",
    "dark",            # 섹션 구분자
    "white", "white", "light",  # MANAGEMENT
    "dark",            # 섹션 구분자
    "white", "white", "light",  # WHY US
    "dark",            # 섹션 구분자
    "white", "white",  # INVESTMENT
    "gradient_dark",   # 클로징
]


def default_design_brief(
    project_name: str = "",
    project_type: str = "event",
) -> MergedDesignBrief:
    """디자인 브리프 기본값 생성 (싱크탱크/Gamma 없이도 동작)"""
    return MergedDesignBrief(
        project_name=project_name,
        project_type=project_type,
        theme_name="default_blue",
        colors=DEFAULT_COLORS,
        background_schedule=DEFAULT_BACKGROUND_SCHEDULE,
        component_targets={
            "HERO_IMAGE": 4, "IMG_PH": 15, "COLS": 8,
            "FLOW": 5, "HIGHLIGHT": 6, "TABLE": 3,
            "METRIC_CARD": 8, "STAT_ROW": 3, "KPIS": 2,
        },
        visual_density_targets={
            "image_slides_pct": 0.25,
            "diagram_slides_pct": 0.35,
            "text_only_max_pct": 0.20,
        },
        content_tone=ContentToneBrief(
            emotional_tone_level=3,
            narrative_framing_style="hybrid",
            program_naming_style="functional",
            win_theme_style="keyword_functional",
            text_density_style="balanced",
            tone_rules=[
                "프로젝트 특성에 맞는 전문 용어와 감성 코드를 활용하라",
                "데이터와 감성을 6:4 비율로 균형있게 배합하라",
                "핵심 메시지는 감성적으로, 뒷받침은 데이터로 구성하라",
            ],
            source_analysis="기본값 (레퍼런스 미사용)",
        ),
        confidence=0.0,
        notes="기본값 (레퍼런스/Gamma 미사용)",
    )
