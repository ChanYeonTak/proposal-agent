"""
Think Tank DB 모델 (v5.0)

레퍼런스 문서에서 추출한 구조화된 데이터 모델.
v5.0: ContentToneProfile 추가 — 콘텐츠 톤/감성/IP 깊이 분석
"""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


# ═══════════════════════════════════════════════════════════════
# Enum
# ═══════════════════════════════════════════════════════════════

class DocType(str, Enum):
    """문서 유형"""
    PROPOSAL = "proposal"         # 제안서
    MANUAL = "manual"             # 매뉴얼/가이드
    REPORT = "report"             # 보고서
    PLAN = "plan"                 # 기획서
    OTHER = "other"


class Industry(str, Enum):
    """산업 분류"""
    GAME_EVENT = "game_event"     # 게임 이벤트/부스
    MARKETING_PR = "marketing_pr" # 마케팅/PR
    EVENT = "event"               # 일반 이벤트
    IT_SYSTEM = "it_system"       # IT/시스템
    PUBLIC = "public"             # 공공
    CONSULTING = "consulting"     # 컨설팅
    FINANCE = "finance"           # 금융
    EDUCATION = "education"       # 교육
    HEALTHCARE = "healthcare"     # 헬스케어
    OTHER = "other"


# ═══════════════════════════════════════════════════════════════
# 디자인 프로파일
# ═══════════════════════════════════════════════════════════════

class ColorInfo(BaseModel):
    """색상 정보"""
    hex: str = ""
    usage: str = ""          # primary, secondary, accent, background, text
    frequency: float = 0.0   # 사용 빈도 (0~1)


class FontInfo(BaseModel):
    """폰트 정보"""
    name: str = ""
    size_pt: float = 0.0
    bold: bool = False
    usage: str = ""          # title, subtitle, body, caption


class LayoutPattern(BaseModel):
    """레이아웃 패턴"""
    pattern_type: str = ""    # full_bleed, two_column, grid, centered, etc.
    frequency: float = 0.0    # 출현 빈도 (0~1)
    description: str = ""


class DesignProfile(BaseModel):
    """문서의 디자인 프로파일"""
    colors: List[ColorInfo] = Field(default_factory=list)
    fonts: List[FontInfo] = Field(default_factory=list)
    font_hierarchy: Dict[str, str] = Field(default_factory=dict)  # {"title": "Pretendard Bold 36pt", ...}
    layout_patterns: List[LayoutPattern] = Field(default_factory=list)
    bg_style: str = ""                        # dark, light, gradient, image
    aspect_ratio: str = "16:9"
    slide_dimensions: Dict[str, float] = Field(default_factory=dict)  # {"width": 13.33, "height": 7.5}


# ═══════════════════════════════════════════════════════════════
# 콘텐츠 패턴
# ═══════════════════════════════════════════════════════════════

class SectionStructure(BaseModel):
    """섹션 구조"""
    name: str = ""
    slide_count: int = 0
    weight_pct: float = 0.0
    subsections: List[str] = Field(default_factory=list)


class ContentPattern(BaseModel):
    """콘텐츠 패턴"""
    pattern_type: str = ""       # narrative_arc, data_driven, case_study, comparison, etc.
    section_context: str = ""    # 어떤 섹션에서 사용되는 패턴인지
    structure: str = ""          # 패턴의 구조 설명
    slide_count: int = 0         # 이 패턴의 일반적 슬라이드 수
    examples: List[str] = Field(default_factory=list)


class ProgramTemplate(BaseModel):
    """프로그램/이벤트 템플릿"""
    name: str = ""
    category: str = ""           # booth_design, event_pack, interaction, stage, campaign, etc.
    mechanism: str = ""          # 프로그램 메커니즘 설명
    reward_structure: str = ""
    operation_plan: str = ""     # 운영 계획 요약
    slide_count: int = 0         # 이 프로그램에 할당된 슬라이드 수
    visual_elements: List[str] = Field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# 콘텐츠 톤 프로파일 (v5.0)
# ═══════════════════════════════════════════════════════════════


class NarrativeFraming(BaseModel):
    """내러티브 프레이밍 스타일"""
    style: str = ""              # worldview_based / data_driven / emotion_led / hybrid
    core_metaphor: str = ""      # 핵심 메타포 ("잔혹 동화", "지상 작전" 등)
    entry_hook: str = ""         # 진입 후크 패턴 ("~를 허가합니다", "~에 오신 것을 환영합니다")
    recurring_motif: str = ""    # 반복 모티프 ("지휘관", "방주", "도파민")
    description: str = ""        # 프레이밍 방식 상세 설명


class ContentToneProfile(BaseModel):
    """
    콘텐츠 톤 프로파일 — 레퍼런스 문서의 글쓰기 스타일과 감성적 톤을 분석

    BD2 수주 성공 제안서 분석에서 도출된 패턴:
    - 세계관 기반 프레이밍: IP 내러티브를 제안서 전체 구조에 적용
    - 감성적 텍스트: 기능 나열이 아닌 스토리텔링 중심
    - IP 깊이감: 캐릭터/로어/커뮤니티 용어의 자연스러운 사용
    - 프로그램 네이밍: 기능 설명형이 아닌 IP 세계관형
    """

    # 감성 톤 레벨 (1~5, 높을수록 감성적)
    # 1=사무적/기능적, 2=간결 전문가, 3=중립, 4=감성+전문, 5=풀 내러티브
    emotional_tone_level: int = 3

    # 내러티브 프레이밍
    narrative_framing: NarrativeFraming = Field(default_factory=NarrativeFraming)

    # IP/브랜드 깊이 (해당 산업 지식의 깊이를 측정)
    ip_depth_score: float = 0.0      # 0~1 (0=일반적, 1=전문가 수준)
    ip_character_count: int = 0       # 언급된 고유명사/캐릭터 수
    ip_lore_terms: List[str] = Field(default_factory=list)
    # 세계관 고유 용어 (예: "방주", "래프처", "니케")
    ip_community_terms: List[str] = Field(default_factory=list)
    # 커뮤니티 은어 (예: "쁘더", "쁘린이", "커피콩")

    # 프로그램/이벤트 네이밍 스타일
    program_naming_style: str = ""   # ip_narrative / functional / branded / hybrid
    # ip_narrative: "지휘관의 첫 번째 작전: REAL RECRUIT"
    # functional:   "부스 체험 프로그램 A"
    # branded:      "NIKKE GATE ZONE"
    # hybrid:       IP 용어 + 기능 설명 혼합
    program_naming_examples: List[str] = Field(default_factory=list)

    # Win Theme 스타일
    win_theme_style: str = ""        # ip_worldview / keyword_functional / emotional_hook
    # ip_worldview:      "WELCOME TO THE ARK" — IP 세계관 기반
    # keyword_functional: "데이터 기반 타겟 마케팅" — 기능적 키워드
    # emotional_hook:     "도파민이 터지는 순간" — 감성 후크
    win_theme_examples: List[str] = Field(default_factory=list)

    # 텍스트 밀도 특성
    text_density_style: str = ""     # minimal / balanced / rich / dense
    # minimal:  짧은 문장, 키워드 중심 (이미지 75%+ 슬라이드)
    # balanced: 적절한 텍스트 + 시각 요소
    # rich:     상세 설명 + 데이터
    # dense:    텍스트 중심
    image_slide_ratio: float = 0.0   # 이미지 포함 슬라이드 비율 (0~1)
    text_only_ratio: float = 0.0     # 텍스트만 있는 슬라이드 비율 (0~1)

    # 산업별 톤 가이드 (분석된 규칙)
    tone_rules: List[str] = Field(default_factory=list)
    # 예: [
    #   "게임 IP는 커뮤니티 은어를 자연스럽게 사용하라",
    #   "프로그램명에 IP 세계관 용어를 삽입하라",
    #   "제안서 전체에 하나의 내러티브 메타포를 관통시켜라",
    # ]

    # 참조 레퍼런스 요약
    source_analysis: str = ""        # 분석 근거 요약


# ═══════════════════════════════════════════════════════════════
# 레퍼런스 문서 (최상위 모델)
# ═══════════════════════════════════════════════════════════════

class ReferenceDocument(BaseModel):
    """
    레퍼런스 문서 — Think Tank DB의 핵심 모델

    하나의 참조 문서(제안서, 매뉴얼 등)에서 추출한 모든 정보를 담습니다.
    """

    # 기본 메타
    id: Optional[int] = None
    file_path: str = ""
    file_hash: str = ""                        # SHA-256 (중복 방지)
    file_name: str = ""
    file_size: int = 0

    # 분류
    doc_type: DocType = DocType.OTHER
    industry: Industry = Industry.OTHER
    project_type: str = ""                     # marketing_pr, event, it_system 등
    won_bid: bool = False                      # 수주 성공 여부

    # 구조
    total_pages: int = 0
    sections: List[SectionStructure] = Field(default_factory=list)
    table_of_contents: List[str] = Field(default_factory=list)

    # 디자인
    design_profile: DesignProfile = Field(default_factory=DesignProfile)

    # 콘텐츠 패턴
    content_patterns: List[ContentPattern] = Field(default_factory=list)
    program_templates: List[ProgramTemplate] = Field(default_factory=list)

    # 콘텐츠 톤 (v5.0)
    content_tone: ContentToneProfile = Field(default_factory=ContentToneProfile)

    # 텍스트 (검색용)
    full_text: str = ""                        # 전체 텍스트 (검색용, 저장 시 압축)
    summary: str = ""                          # AI가 생성한 요약

    # 메타 정보
    ingested_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    tags: List[str] = Field(default_factory=list)
    notes: str = ""


# ═══════════════════════════════════════════════════════════════
# 검색 결과
# ═══════════════════════════════════════════════════════════════

class SearchResult(BaseModel):
    """검색 결과"""
    document: ReferenceDocument
    relevance_score: float = 0.0
    match_reason: str = ""
