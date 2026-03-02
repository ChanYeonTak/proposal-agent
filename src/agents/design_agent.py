"""
디자인 에이전트 (v2.0)

싱크탱크 레퍼런스 데이터 + Gamma 테마 추천을 병합하여
MergedDesignBrief를 생성합니다.

v2.0: 콘텐츠 톤 병합 추가 — content_tone 필드로 감성/IP깊이/네이밍 가이드 제공.

워크플로우:
    1. 싱크탱크 DB에서 유사 수주 레퍼런스 검색
    2. DesignBriefBuilder로 레퍼런스 기반 디자인 브리프 생성 (content_tone 포함)
    3. Gamma get_themes() 결과를 전달받아 테마 후보 선정
    4. 양쪽을 병합하여 MergedDesignBrief 반환

사용법:
    agent = DesignAgent()

    # 1. 싱크탱크 기반 브리프 조회
    tt_brief = agent.get_think_tank_brief("event", "game_event", target_slides=70)

    # 2. Gamma 테마 해석 (Claude가 MCP get_themes() 호출 후 결과 전달)
    gamma_recs = agent.interpret_gamma_themes(gamma_themes_data, project_keywords)

    # 3. 병합 (content_tone 자동 포함)
    merged = agent.merge(tt_brief, gamma_recs, project_name="NIKKE AGF 2025")
    # merged.content_tone.tone_rules → 생성 스크립트가 참조하는 톤 규칙
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from src.schemas.design_schema import (
    ContentToneBrief,
    DEFAULT_COLORS,
    ImageStyleGuide,
    MergedDesignBrief,
    ThemeRecommendation,
    default_design_brief,
)
from src.utils.logger import get_logger

logger = get_logger("design_agent")


class DesignAgent:
    """싱크탱크 + Gamma 병합 디자인 에이전트"""

    def __init__(self):
        self._tt_builder = None  # lazy init

    # ── 1. 싱크탱크 조회 ──────────────────────────────────────

    def get_think_tank_brief(
        self,
        project_type: str = "event",
        industry: str = "game_event",
        target_slides: int = 70,
    ) -> Optional[Dict]:
        """
        싱크탱크 DB에서 유사 레퍼런스 검색 → DesignBrief 반환.

        Returns:
            DesignBrief.to_dict() 또는 None (싱크탱크 사용 불가 시)
        """
        try:
            from think_tank.design_brief import DesignBriefBuilder
            builder = DesignBriefBuilder()
            brief = builder.build(
                project_type=project_type,
                industry=industry,
                target_slides=target_slides,
            )
            logger.info(f"싱크탱크 브리프 생성 완료:\n{brief.summary()}")
            return brief.to_dict()
        except ImportError:
            logger.warning("think_tank 모듈 미설치 — 싱크탱크 없이 진행")
            return None
        except Exception as e:
            logger.error(f"싱크탱크 브리프 생성 실패: {e}")
            return None

    # ── 2. Gamma 테마 해석 ─────────────────────────────────────

    def interpret_gamma_themes(
        self,
        gamma_themes_data: List[Dict[str, Any]],
        project_keywords: List[str] = None,
    ) -> List[ThemeRecommendation]:
        """
        Gamma get_themes() 결과를 해석하여 상위 3개 후보 반환.

        Claude가 Gamma MCP get_themes()를 호출한 결과를
        이 메서드에 전달합니다.

        Args:
            gamma_themes_data: Gamma get_themes() 반환값 (list of dicts)
            project_keywords: 프로젝트 관련 키워드 (매칭용)

        Returns:
            List[ThemeRecommendation]: 상위 3개 테마 후보
        """
        if not gamma_themes_data:
            return []

        keywords = [kw.lower() for kw in (project_keywords or [])]
        scored = []

        for theme in gamma_themes_data:
            score = 0.0
            reasons = []

            theme_name = str(theme.get("name", "")).lower()
            theme_desc = str(theme.get("description", "")).lower()
            tone = [t.lower() for t in theme.get("tone_keywords", [])]
            colors = [c.lower() for c in theme.get("color_keywords", [])]

            # 키워드 매칭 점수
            for kw in keywords:
                if kw in theme_name or kw in theme_desc:
                    score += 0.3
                    reasons.append(f"이름/설명에 '{kw}' 포함")
                for t in tone:
                    if kw in t or t in kw:
                        score += 0.2
                        reasons.append(f"톤 '{t}' 매칭")
                for c in colors:
                    if kw in c or c in kw:
                        score += 0.15
                        reasons.append(f"컬러 '{c}' 매칭")

            # 프로페셔널/모던 계열 보너스
            for t in tone:
                if t in ("professional", "modern", "bold", "dark", "sleek"):
                    score += 0.1

            scored.append(ThemeRecommendation(
                theme_id=str(theme.get("id", "")),
                theme_name=theme.get("name", ""),
                tone_keywords=tone,
                color_keywords=colors,
                match_reason="; ".join(reasons[:3]) if reasons else "일반 매칭",
                match_score=min(score, 1.0),
            ))

        # 점수순 정렬 → 상위 3개
        scored.sort(key=lambda x: x.match_score, reverse=True)
        top3 = scored[:3]

        for rec in top3:
            logger.info(f"Gamma 테마 후보: {rec.theme_name} (score={rec.match_score:.2f}, {rec.match_reason})")

        return top3

    # ── 3. 병합 ───────────────────────────────────────────────

    def merge(
        self,
        tt_brief: Optional[Dict] = None,
        gamma_recs: Optional[List[ThemeRecommendation]] = None,
        project_name: str = "",
        project_type: str = "event",
        industry: str = "",
        custom_colors: Optional[Dict[str, Tuple[int, int, int]]] = None,
    ) -> MergedDesignBrief:
        """
        싱크탱크 + Gamma 데이터를 병합하여 MergedDesignBrief 생성.

        우선순위:
            1. custom_colors (사용자 직접 지정)
            2. 싱크탱크 레퍼런스 컬러
            3. Gamma 테마 컬러
            4. 기본값 (DEFAULT_COLORS)

        Args:
            tt_brief: 싱크탱크 DesignBrief.to_dict() 결과
            gamma_recs: Gamma 테마 추천 후보
            project_name: 프로젝트명
            project_type: 프로젝트 유형
            industry: 산업 분류
            custom_colors: 사용자 지정 컬러 {"primary": (r,g,b), ...}

        Returns:
            MergedDesignBrief
        """
        merged = MergedDesignBrief(
            project_name=project_name,
            project_type=project_type,
            industry=industry,
        )

        # ── 싱크탱크 데이터 반영 (1순위) ──
        if tt_brief:
            merged.section_weights = tt_brief.get("section_weights", {})
            merged.component_targets = tt_brief.get("component_targets", {})
            merged.layout_distribution = tt_brief.get("layout_distribution", {})
            merged.background_schedule = tt_brief.get("background_schedule", [])
            merged.visual_density_targets = tt_brief.get("visual_density_targets", {})
            merged.content_patterns = tt_brief.get("content_patterns", [])
            merged.program_templates = tt_brief.get("program_templates", [])
            merged.source_references = [
                r.get("file", "") for r in tt_brief.get("source_references", [])
            ]

            # 싱크탱크 컬러 추출 (hex → RGB tuple)
            design_ref = tt_brief.get("design_reference", {})
            ref_colors = design_ref.get("colors", {})
            if ref_colors:
                tt_colors = {}
                for usage, hex_val in ref_colors.items():
                    rgb = _hex_to_rgb(hex_val)
                    if rgb:
                        tt_colors[usage] = rgb
                if tt_colors:
                    merged.colors = {**DEFAULT_COLORS, **tt_colors}
                    merged.confidence = 0.8
                    merged.notes += "컬러: 싱크탱크 레퍼런스 기반. "
                else:
                    merged.colors = dict(DEFAULT_COLORS)
                    merged.confidence = 0.5
                    merged.notes += "컬러: 레퍼런스 컬러 파싱 실패, 기본값 사용. "
            else:
                merged.colors = dict(DEFAULT_COLORS)
                merged.confidence = 0.3
                merged.notes += "컬러: 레퍼런스에 컬러 데이터 없음. "
        else:
            # 싱크탱크 없음 → 기본값
            merged.confidence = 0.0
            merged.colors = dict(DEFAULT_COLORS)
            merged.notes += "싱크탱크 미사용. "

        # ── Gamma 테마 보완 (2순위) ──
        if gamma_recs:
            merged.gamma_recommendations = gamma_recs

            # 싱크탱크 컬러가 약하면 (confidence < 0.5) Gamma 테마 반영
            if merged.confidence < 0.5 and gamma_recs:
                top = gamma_recs[0]
                merged.theme_name = f"gamma_{top.theme_name.lower().replace(' ', '_')}"
                merged.notes += f"테마: Gamma '{top.theme_name}' 기반. "

                # Gamma 컬러 키워드에서 색상 추론
                gamma_colors = _infer_colors_from_keywords(top.color_keywords)
                if gamma_colors:
                    merged.colors = {**merged.colors, **gamma_colors}
            else:
                merged.theme_name = "reference_based"
                merged.notes += "테마: 싱크탱크 레퍼런스 기반. "

        # ── 사용자 지정 컬러 최우선 ──
        if custom_colors:
            merged.colors = {**merged.colors, **custom_colors}
            merged.notes += "컬러: 사용자 직접 지정 반영. "

        # ── 이미지 스타일 가이드 ──
        merged.image_style = _build_image_style_guide(
            project_type, industry, gamma_recs
        )

        # ── 콘텐츠 톤 반영 (v2.0) ──
        content_tone_data = tt_brief.get("content_tone", {}) if tt_brief else {}
        merged.content_tone = _build_content_tone_brief(content_tone_data)

        logger.info(
            f"MergedDesignBrief 생성: confidence={merged.confidence:.2f}, "
            f"theme={merged.theme_name}, "
            f"tone_level={merged.content_tone.emotional_tone_level}/5, "
            f"tone_rules={len(merged.content_tone.tone_rules)}건, "
            f"refs={len(merged.source_references)}건"
        )

        return merged

    # ── 편의 메서드: 한 번에 전체 수행 ──

    def generate_full_brief(
        self,
        project_name: str = "",
        project_type: str = "event",
        industry: str = "game_event",
        target_slides: int = 70,
        gamma_themes_data: Optional[List[Dict]] = None,
        project_keywords: Optional[List[str]] = None,
        custom_colors: Optional[Dict[str, Tuple[int, int, int]]] = None,
    ) -> MergedDesignBrief:
        """
        전체 디자인 브리프 생성 (one-shot).

        싱크탱크 조회 + Gamma 해석 + 병합을 한 번에 수행합니다.
        """
        # 1. 싱크탱크
        tt_brief = self.get_think_tank_brief(project_type, industry, target_slides)

        # 2. Gamma
        gamma_recs = []
        if gamma_themes_data:
            gamma_recs = self.interpret_gamma_themes(
                gamma_themes_data, project_keywords or []
            )

        # 3. 병합
        return self.merge(
            tt_brief=tt_brief,
            gamma_recs=gamma_recs,
            project_name=project_name,
            project_type=project_type,
            industry=industry,
            custom_colors=custom_colors,
        )

    # ── slide_kit 테마 등록 헬퍼 ──

    @staticmethod
    def register_to_slide_kit(brief: MergedDesignBrief) -> str:
        """
        MergedDesignBrief의 컬러를 slide_kit에 동적 테마로 등록.

        Returns:
            등록된 테마 이름
        """
        from src.generators.slide_kit import register_theme, apply_theme

        theme_name = brief.theme_name or "design_agent_custom"
        register_theme(theme_name, brief.colors)
        apply_theme(theme_name)

        logger.info(f"slide_kit 테마 등록 및 적용: {theme_name}")
        return theme_name


# ═══════════════════════════════════════════════════════════════
#  헬퍼 함수
# ═══════════════════════════════════════════════════════════════

def _hex_to_rgb(hex_str: str) -> Optional[Tuple[int, int, int]]:
    """HEX 문자열 → RGB 튜플. 실패 시 None."""
    if not hex_str:
        return None
    hex_str = hex_str.strip().lstrip("#")
    if len(hex_str) != 6:
        return None
    try:
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
        return (r, g, b)
    except ValueError:
        return None


def _infer_colors_from_keywords(
    color_keywords: List[str],
) -> Dict[str, Tuple[int, int, int]]:
    """Gamma 테마의 컬러 키워드에서 slide_kit C[] 컬러 추론.

    완벽하지 않지만, 레퍼런스가 없을 때 차선의 컬러 가이드를 제공합니다.
    """
    keyword_map = {
        # 다크 계열
        "dark": {"dark": (25, 25, 25)},
        "navy": {"primary": (0, 44, 95)},
        "black": {"dark": (15, 15, 15)},
        # 블루 계열
        "blue": {"primary": (0, 70, 150)},
        "sky": {"secondary": (0, 170, 210)},
        "ocean": {"primary": (0, 60, 120)},
        "teal": {"teal": (0, 161, 156)},
        "cyan": {"secondary": (0, 190, 210)},
        # 레드/웜 계열
        "red": {"accent": (230, 51, 18)},
        "coral": {"accent": (255, 127, 80)},
        "warm": {"accent": (210, 105, 30), "light": (255, 248, 240)},
        "orange": {"accent": (245, 166, 35)},
        # 그린 계열
        "green": {"teal": (46, 125, 50)},
        "forest": {"primary": (27, 94, 32)},
        "mint": {"secondary": (0, 200, 170)},
        # 퍼플 계열
        "purple": {"primary": (74, 20, 140)},
        "violet": {"primary": (100, 50, 180)},
        # 골드/프리미엄
        "gold": {"accent": (197, 151, 62)},
        "premium": {"primary": (30, 30, 50), "accent": (197, 151, 62)},
        # 밝은 계열
        "light": {"light": (248, 248, 250)},
        "white": {"light": (252, 252, 252)},
        "pastel": {"light": (240, 245, 250)},
        # 스타일
        "modern": {},
        "bold": {},
        "minimal": {"light": (250, 250, 250)},
    }

    result = {}
    for kw in color_keywords:
        kw_lower = kw.lower().strip()
        for map_key, colors in keyword_map.items():
            if map_key in kw_lower:
                result.update(colors)

    return result


def _build_image_style_guide(
    project_type: str,
    industry: str,
    gamma_recs: Optional[List[ThemeRecommendation]] = None,
) -> ImageStyleGuide:
    """프로젝트 유형 + Gamma 추천 기반 이미지 스타일 가이드"""
    # 프로젝트 유형별 기본 스타일
    type_styles = {
        "game_event": ImageStyleGuide(
            primary_style="photorealistic",
            keywords=["게임", "부스", "이벤트", "코스프레", "AGF"],
            source_preference={
                "photo": "pexels",
                "illustration": "gamma_ai",
                "diagram": "renderer",
            },
        ),
        "event": ImageStyleGuide(
            primary_style="photorealistic",
            keywords=["이벤트", "행사", "부스", "전시"],
            source_preference={
                "photo": "pexels",
                "illustration": "gamma_ai",
                "diagram": "renderer",
            },
        ),
        "marketing_pr": ImageStyleGuide(
            primary_style="mixed",
            keywords=["마케팅", "SNS", "캠페인", "브랜드"],
            source_preference={
                "photo": "pexels",
                "illustration": "gamma_ai",
                "diagram": "renderer",
            },
        ),
        "it_system": ImageStyleGuide(
            primary_style="minimal",
            keywords=["시스템", "인프라", "아키텍처", "대시보드"],
            source_preference={
                "photo": "pexels",
                "illustration": "gamma_ai",
                "diagram": "renderer",
            },
        ),
    }

    guide = type_styles.get(industry, type_styles.get(project_type, ImageStyleGuide()))

    # Gamma 추천이 있으면 이미지 키워드 보강
    if gamma_recs:
        top = gamma_recs[0]
        if top.tone_keywords:
            guide.keywords.extend(top.tone_keywords[:2])

    return guide


def _build_content_tone_brief(
    tone_data: Dict,
) -> ContentToneBrief:
    """
    싱크탱크 DesignBrief.content_tone 딕셔너리 → ContentToneBrief 변환.

    DesignBriefBuilder._analyze_content_tone()가 반환하는 dict를
    MergedDesignBrief에 포함되는 ContentToneBrief Pydantic 모델로 변환합니다.
    """
    if not tone_data:
        return ContentToneBrief()

    framing = tone_data.get("narrative_framing", {})

    return ContentToneBrief(
        emotional_tone_level=tone_data.get("emotional_tone_level", 3),
        narrative_framing_style=framing.get("style", "hybrid"),
        core_metaphor=framing.get("core_metaphor", ""),
        entry_hook_pattern=framing.get("entry_hook", ""),
        recurring_motif=framing.get("recurring_motif", ""),
        ip_depth_score=tone_data.get("ip_depth_score", 0.0),
        ip_lore_terms=tone_data.get("ip_lore_terms", []),
        ip_community_terms=tone_data.get("ip_community_terms", []),
        program_naming_style=tone_data.get("program_naming_style", "functional"),
        program_naming_examples=tone_data.get("program_naming_examples", []),
        win_theme_style=tone_data.get("win_theme_style", "keyword_functional"),
        win_theme_examples=tone_data.get("win_theme_examples", []),
        text_density_style=tone_data.get("text_density_style", "balanced"),
        tone_rules=tone_data.get("tone_rules", []),
        source_analysis=tone_data.get("source_analysis", ""),
    )
