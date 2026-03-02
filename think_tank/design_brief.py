"""
디자인 브리프 생성기 (v5.0)

Think Tank 레퍼런스 데이터를 분석하여 제안서 생성 스크립트가 참조할
구체적 가이드를 생성합니다.

v5.0: 콘텐츠 톤 분석 추가 — 감성적 글쓰기, IP 깊이감, 프로그램 네이밍 스타일 분석

Usage:
    from think_tank.design_brief import DesignBriefBuilder
    builder = DesignBriefBuilder()
    brief = builder.build(project_type="event", industry="game_event")
    # brief.content_tone → 콘텐츠 톤 가이드
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List, Optional

from .db import ThinkTankDB
from .models import (
    ContentPattern,
    ContentToneProfile,
    DesignProfile,
    LayoutPattern,
    NarrativeFraming,
    ProgramTemplate,
    ReferenceDocument,
    SectionStructure,
)
from .retrieval import ThinkTankRetrieval
from src.utils.logger import get_logger

logger = get_logger("design_brief")


class DesignBrief:
    """생성 스크립트가 참조하는 디자인/콘텐츠 가이드"""

    def __init__(self):
        # 섹션별 목표 슬라이드 수 (레퍼런스 기반)
        self.section_weights: Dict[str, int] = {}

        # 컴포넌트 목표 사용 횟수
        self.component_targets: Dict[str, int] = {}

        # 레이아웃 분포 목표 (유형 → 비율 0~1)
        self.layout_distribution: Dict[str, float] = {}

        # 배경 스케줄 (슬라이드별 bg preset 권장)
        self.background_schedule: List[str] = []

        # 콘텐츠 패턴 (참조할 스토리 구조)
        self.content_patterns: List[Dict] = []

        # 프로그램 템플릿 (참조할 프로그램 구조)
        self.program_templates: List[Dict] = []

        # 디자인 프로파일 (컬러/폰트/스타일 참조)
        self.design_reference: Dict = {}

        # 시각 밀도 목표
        self.visual_density_targets: Dict = {}

        # 콘텐츠 톤 가이드 (v5.0)
        self.content_tone: Dict = {}

        # 소스 레퍼런스 요약
        self.source_references: List[Dict] = []

    def to_dict(self) -> dict:
        """딕셔너리로 변환 (스크립트 전달용)"""
        return {
            "section_weights": self.section_weights,
            "component_targets": self.component_targets,
            "layout_distribution": self.layout_distribution,
            "background_schedule": self.background_schedule,
            "content_patterns": self.content_patterns,
            "program_templates": self.program_templates,
            "design_reference": self.design_reference,
            "visual_density_targets": self.visual_density_targets,
            "content_tone": self.content_tone,
            "source_references": self.source_references,
        }

    def summary(self) -> str:
        """브리프 요약 (로그/디버그용)"""
        tone = self.content_tone
        tone_info = ""
        if tone:
            tone_info = (
                f"\n  감성 톤 레벨: {tone.get('emotional_tone_level', '?')}/5"
                f"\n  프레이밍: {tone.get('narrative_framing', {}).get('style', '?')}"
                f"\n  IP 깊이: {tone.get('ip_depth_score', 0):.1f}"
                f"\n  프로그램 네이밍: {tone.get('program_naming_style', '?')}"
                f"\n  톤 규칙: {len(tone.get('tone_rules', []))}건"
            )
        lines = [
            "=== Design Brief ===",
            f"섹션 비중: {len(self.section_weights)}개 섹션",
            f"컴포넌트 목표: {sum(self.component_targets.values())}개 배치",
            f"레이아웃 분포: {len(self.layout_distribution)}개 유형",
            f"배경 스케줄: {len(self.background_schedule)}슬라이드",
            f"콘텐츠 패턴: {len(self.content_patterns)}건 참조",
            f"프로그램 템플릿: {len(self.program_templates)}건 참조",
            f"콘텐츠 톤: {'설정됨' if tone else '미설정'}{tone_info}",
            f"소스 레퍼런스: {len(self.source_references)}건",
        ]
        return "\n".join(lines)


class DesignBriefBuilder:
    """레퍼런스 데이터 → DesignBrief 변환"""

    def __init__(self, db: Optional[ThinkTankDB] = None):
        self.db = db or ThinkTankDB()
        self.retrieval = ThinkTankRetrieval(db=self.db)

    def build(
        self,
        project_type: Optional[str] = None,
        industry: Optional[str] = None,
        target_slides: int = 70,
    ) -> DesignBrief:
        """
        레퍼런스 기반 디자인 브리프 생성

        Args:
            project_type: 프로젝트 유형 (event, marketing_pr 등)
            industry: 산업 분류 (game_event 등)
            target_slides: 목표 슬라이드 수

        Returns:
            DesignBrief: 생성 가이드
        """
        brief = DesignBrief()

        # 레퍼런스 검색
        results = self.retrieval.search_similar(
            project_type=project_type,
            industry=industry,
            won_bid_only=True,
            top_k=5,
        )
        if not results:
            results = self.retrieval.search_similar(
                project_type=project_type,
                industry=industry,
                won_bid_only=False,
                top_k=5,
            )

        if not results:
            logger.warning("레퍼런스 없음: 기본 브리프 반환")
            brief = self._build_default_brief(target_slides, industry=industry)
            return brief

        docs = [r.document for r in results]

        # 1. 섹션 비중 계산
        brief.section_weights = self._calc_section_weights(docs, target_slides)

        # 2. 레이아웃 분포 분석
        brief.layout_distribution = self._calc_layout_distribution(docs)

        # 3. 컴포넌트 목표 산출
        brief.component_targets = self._calc_component_targets(
            brief.layout_distribution, target_slides
        )

        # 4. 배경 스케줄 생성
        brief.background_schedule = self._build_bg_schedule(
            docs, target_slides
        )

        # 5. 콘텐츠 패턴 수집
        brief.content_patterns = self._collect_content_patterns(docs)

        # 6. 프로그램 템플릿 수집
        brief.program_templates = self._collect_program_templates(docs)

        # 7. 디자인 프로파일 종합
        brief.design_reference = self._merge_design_profiles(docs)

        # 8. 시각 밀도 목표
        brief.visual_density_targets = self._calc_visual_density(docs)

        # 9. 콘텐츠 톤 분석 (v5.0)
        # 콘텐츠 톤은 제안서뿐 아니라 같은 산업의 모든 문서(계획안, 매뉴얼, 보고서)를 참조
        all_industry_docs = self._get_all_industry_docs(industry, docs)
        brief.content_tone = self._analyze_content_tone(all_industry_docs)

        # 10. 소스 정보
        brief.source_references = [
            {
                "file": d.file_name,
                "pages": d.total_pages,
                "won_bid": d.won_bid,
                "industry": d.industry.value if d.industry else "",
            }
            for d in docs
        ]

        logger.info(
            f"디자인 브리프 생성 완료 (레퍼런스 {len(docs)}건, "
            f"목표 {target_slides}슬라이드)"
        )
        return brief

    # ── Private: 섹션 비중 ──

    def _calc_section_weights(
        self, docs: List[ReferenceDocument], target_slides: int
    ) -> Dict[str, int]:
        """레퍼런스 섹션 구조를 기반으로 목표 슬라이드 수 산출"""
        section_pcts: Dict[str, List[float]] = {}

        for doc in docs:
            total = sum(s.slide_count for s in doc.sections) or 1
            for section in doc.sections:
                name = section.name.upper().strip()
                pct = section.slide_count / total
                if name not in section_pcts:
                    section_pcts[name] = []
                section_pcts[name].append(pct)

        weights = {}
        for name, pcts in section_pcts.items():
            avg_pct = sum(pcts) / len(pcts)
            slide_count = max(1, round(avg_pct * target_slides))
            weights[name] = slide_count

        # 총합 조정
        total_assigned = sum(weights.values())
        if total_assigned != target_slides and weights:
            largest = max(weights, key=weights.get)
            weights[largest] += (target_slides - total_assigned)

        return weights

    # ── Private: 레이아웃 분포 ──

    def _calc_layout_distribution(
        self, docs: List[ReferenceDocument]
    ) -> Dict[str, float]:
        """레퍼런스의 레이아웃 패턴 분포 평균"""
        combined: Dict[str, List[float]] = {}

        for doc in docs:
            if doc.design_profile and doc.design_profile.layout_patterns:
                for lp in doc.design_profile.layout_patterns:
                    if lp.pattern_type not in combined:
                        combined[lp.pattern_type] = []
                    combined[lp.pattern_type].append(lp.frequency)

        distribution = {}
        for pt, freqs in combined.items():
            distribution[pt] = round(sum(freqs) / len(freqs), 3)

        return distribution

    # ── Private: 컴포넌트 목표 ──

    def _calc_component_targets(
        self, layout_dist: Dict[str, float], target_slides: int
    ) -> Dict[str, int]:
        """레이아웃 분포 → slide_kit 컴포넌트 목표 매핑"""
        # 레이아웃 유형 → slide_kit 컴포넌트 매핑
        layout_to_component = {
            "full_bleed_image": "HERO_IMAGE",
            "split_image_text": "SPLIT_VISUAL",
            "multi_image_grid": "MOOD_BOARD",
            "image_focused": "IMG_PH",
            "table_based": "TABLE",
            "hero_typography": "HIGHLIGHT",
            "data_card_array": "METRIC_CARD",
            "process_flow": "FLOW",
            "layered_composition": "COLS",
            "complex_diagram": "COLS",
            "multi_column": "COLS",
            "content_standard": "MT",
        }

        targets = {}
        for layout_type, freq in layout_dist.items():
            component = layout_to_component.get(layout_type, "COLS")
            count = max(1, round(freq * target_slides))
            if component in targets:
                targets[component] += count
            else:
                targets[component] = count

        # 최소 보장: 주요 비주얼 컴포넌트
        minimums = {
            "HERO_IMAGE": 4,
            "SPLIT_VISUAL": 3,
            "IMG_PH": 15,
            "SECTION_BRIDGE": 5,
            "ZONE_MAP": 2,
        }
        for comp, minimum in minimums.items():
            targets[comp] = max(targets.get(comp, 0), minimum)

        return targets

    # ── Private: 배경 스케줄 ──

    def _build_bg_schedule(
        self, docs: List[ReferenceDocument], target_slides: int
    ) -> List[str]:
        """레퍼런스 배경 스타일 분석 → 슬라이드별 배경 프리셋 스케줄"""
        # 가장 큰 레퍼런스의 bg_style 참조
        bg_style = "light"
        for doc in docs:
            if doc.design_profile:
                bg_style = doc.design_profile.bg_style
                break

        # 기본 패턴: 섹션 구분자=dark, 본문=white/light 교차, 포인트=gradient
        schedule = []
        for i in range(target_slides):
            pos_ratio = i / max(target_slides - 1, 1)

            # 매 8~10슬라이드마다 섹션 구분자 (dark)
            if i == 0:
                schedule.append("gradient_dark")  # 표지
            elif i == target_slides - 1:
                schedule.append("gradient_teal")  # 클로징
            elif i % 9 == 0:
                schedule.append("dark")  # 섹션 구분자
            elif i % 9 == 1:
                schedule.append("subtle_blue")  # 섹션 직후 포인트
            elif i % 5 == 0:
                schedule.append("light")  # 변화
            elif i % 15 == 7:
                schedule.append("warm_light")  # 간헐적 따뜻한 톤
            else:
                schedule.append("white")  # 기본

        return schedule

    # ── Private: 콘텐츠 패턴 수집 ──

    def _collect_content_patterns(
        self, docs: List[ReferenceDocument]
    ) -> List[Dict]:
        """모든 레퍼런스에서 콘텐츠 패턴 수집"""
        patterns = []
        for doc in docs:
            for cp in doc.content_patterns:
                patterns.append({
                    "type": cp.pattern_type,
                    "section": cp.section_context,
                    "structure": cp.structure,
                    "slides": cp.slide_count,
                    "source": doc.file_name,
                })
        return patterns

    # ── Private: 프로그램 템플릿 수집 ──

    def _collect_program_templates(
        self, docs: List[ReferenceDocument]
    ) -> List[Dict]:
        """모든 레퍼런스에서 프로그램 템플릿 수집"""
        templates = []
        for doc in docs:
            for pt in doc.program_templates:
                templates.append({
                    "name": pt.name,
                    "category": pt.category,
                    "mechanism": pt.mechanism,
                    "reward": pt.reward_structure,
                    "slides": pt.slide_count,
                    "visuals": pt.visual_elements,
                    "source": doc.file_name,
                })
        return templates

    # ── Private: 디자인 프로파일 종합 ──

    def _merge_design_profiles(
        self, docs: List[ReferenceDocument]
    ) -> Dict:
        """다중 레퍼런스 디자인 프로파일 병합"""
        if not docs:
            return {}

        # 첫 번째 수주 성공 레퍼런스 우선
        primary = None
        for doc in docs:
            if doc.won_bid and doc.design_profile:
                primary = doc.design_profile
                break
        if not primary and docs[0].design_profile:
            primary = docs[0].design_profile

        if not primary:
            return {}

        # 컬러 팔레트 (primary/secondary/accent만 추출)
        color_map = {}
        for c in primary.colors:
            if c.usage in ("primary", "secondary", "accent") and c.usage not in color_map:
                color_map[c.usage] = c.hex

        return {
            "colors": color_map,
            "font_hierarchy": primary.font_hierarchy,
            "bg_style": primary.bg_style,
            "aspect_ratio": primary.aspect_ratio,
            "slide_dimensions": primary.slide_dimensions,
            "top_layout_patterns": [
                {"type": lp.pattern_type, "freq": lp.frequency}
                for lp in (primary.layout_patterns or [])[:5]
            ],
        }

    # ── Private: 시각 밀도 목표 ──

    def _calc_visual_density(
        self, docs: List[ReferenceDocument]
    ) -> Dict:
        """레퍼런스 시각 밀도 분석 → 목표 설정"""
        image_pcts = []
        diagram_pcts = []

        for doc in docs:
            if doc.design_profile and doc.design_profile.layout_patterns:
                for lp in doc.design_profile.layout_patterns:
                    if lp.pattern_type in (
                        "full_bleed_image", "split_image_text",
                        "multi_image_grid", "image_focused"
                    ):
                        image_pcts.append(lp.frequency)
                    if lp.pattern_type in (
                        "complex_diagram", "process_flow",
                        "data_card_array", "layered_composition",
                        "hierarchical_diagram"
                    ):
                        diagram_pcts.append(lp.frequency)

        image_target = sum(image_pcts) / len(image_pcts) if image_pcts else 0.20
        diagram_target = sum(diagram_pcts) / len(diagram_pcts) if diagram_pcts else 0.30

        return {
            "image_slides_pct": round(image_target, 2),
            "diagram_slides_pct": round(diagram_target, 2),
            "text_only_max_pct": round(max(0, 1.0 - image_target - diagram_target - 0.15), 2),
            "bg_variety_target": "60% white, 20% light, 10% dark, 10% gradient",
        }

    # ── Private: 전체 산업 문서 조회 (v5.1) ──

    def _get_all_industry_docs(
        self,
        industry: Optional[str],
        proposal_docs: List[ReferenceDocument],
    ) -> List[ReferenceDocument]:
        """
        콘텐츠 톤 분석용: 같은 산업의 모든 문서 유형 조회

        제안서뿐 아니라 운영계획안, 매뉴얼, 보고서 등도 포함하여
        풍부한 텍스트 데이터로 콘텐츠 톤을 분석합니다.

        Args:
            industry: 산업 분류
            proposal_docs: 이미 검색된 제안서 문서

        Returns:
            제안서 + 같은 산업의 추가 문서 (중복 제거)
        """
        if not industry:
            return proposal_docs

        # 같은 산업의 모든 문서 유형 (plan, manual, report) 추가 검색
        all_docs = list(proposal_docs)
        existing_hashes = {d.file_hash for d in all_docs}

        for doc_type_str in ["plan", "manual", "report"]:
            extra = self.db.search_by_type(
                doc_type=doc_type_str,
                industry=industry,
                won_bid_only=False,
                limit=10,
            )
            for d in extra:
                if d.file_hash not in existing_hashes:
                    all_docs.append(d)
                    existing_hashes.add(d.file_hash)

        if len(all_docs) > len(proposal_docs):
            logger.info(
                f"콘텐츠 톤 분석 확장: 제안서 {len(proposal_docs)}건 + "
                f"추가 {len(all_docs) - len(proposal_docs)}건 "
                f"(계획안/매뉴얼/보고서)"
            )

        return all_docs

    # ── Private: 콘텐츠 톤 분석 (v5.0) ──

    def _analyze_content_tone(
        self, docs: List[ReferenceDocument]
    ) -> Dict:
        """
        레퍼런스 문서의 콘텐츠 톤 분석

        분석 항목:
        1. 감성 톤 레벨 (1~5)
        2. 내러티브 프레이밍 스타일
        3. IP/브랜드 깊이
        4. 프로그램 네이밍 스타일
        5. Win Theme 스타일
        6. 텍스트 밀도
        7. 톤 규칙 생성
        """
        if not docs:
            return self._default_content_tone()

        # 수주 성공 레퍼런스 우선 분석
        primary_docs = [d for d in docs if d.won_bid]
        if not primary_docs:
            primary_docs = docs

        # 기존 content_tone 데이터가 "실질적으로" 설정되어 있는지 확인
        # (tone_rules가 있거나, ip_depth_score > 0이면 실질적 데이터로 간주)
        existing_tones = [
            d.content_tone for d in primary_docs
            if d.content_tone and (
                d.content_tone.tone_rules  # 톤 규칙이 설정됨
                or d.content_tone.ip_depth_score > 0  # IP 깊이가 분석됨
                or d.content_tone.emotional_tone_level != 3  # 기본값(3)이 아님
                or d.content_tone.ip_lore_terms  # IP 용어가 추출됨
            )
        ]
        if existing_tones:
            return self._merge_existing_tones(existing_tones, primary_docs)

        # content_tone 데이터가 없거나 기본값이면 텍스트/패턴에서 추론
        return self._infer_content_tone(primary_docs)

    def _merge_existing_tones(
        self,
        tones: List[ContentToneProfile],
        docs: List[ReferenceDocument],
    ) -> Dict:
        """이미 분석된 ContentToneProfile 데이터 병합"""
        # 가장 감성 톤이 높은 (수주 성공한) 레퍼런스 기준
        best = max(tones, key=lambda t: t.emotional_tone_level)

        # 모든 톤 규칙 수집
        all_rules = []
        all_lore_terms = []
        all_community_terms = []
        all_naming_examples = []
        all_wt_examples = []

        for t in tones:
            all_rules.extend(t.tone_rules)
            all_lore_terms.extend(t.ip_lore_terms)
            all_community_terms.extend(t.ip_community_terms)
            all_naming_examples.extend(t.program_naming_examples)
            all_wt_examples.extend(t.win_theme_examples)

        # 노이즈 필터링: 줄바꿈/긴 텍스트 파편 제거
        def _clean(items: List[str], max_len: int = 40) -> List[str]:
            cleaned = []
            for item in items:
                if not item:
                    continue
                # 줄바꿈 포함 → 노이즈
                if '\n' in item or '\r' in item:
                    continue
                # 너무 긴 항목 → 텍스트 파편
                if len(item) > max_len:
                    continue
                # 한글 조사로 시작 → 문장 파편
                if re.match(r'^[은는이가을를의에서로와과에게]', item):
                    continue
                # 한글 조사로 끝남 → 문장 파편
                if re.search(r'[은는을를의에과와]$', item):
                    continue
                # 공백 3개 이상 → 문장 파편
                if item.count(' ') > 3:
                    continue
                # 일반적인 문장 파편 키워드
                if any(g in item for g in [
                    "유저", "신규", "특별한", "추억", "설치",
                    "팬들에게", "안내", "선사",
                ]):
                    continue
                if item not in cleaned:
                    cleaned.append(item)
            return cleaned

        all_naming_examples = _clean(all_naming_examples, max_len=30)
        all_wt_examples = _clean(all_wt_examples, max_len=40)

        # 평균 감성 톤
        avg_tone = round(sum(t.emotional_tone_level for t in tones) / len(tones))

        return {
            "emotional_tone_level": max(avg_tone, best.emotional_tone_level),
            "narrative_framing": {
                "style": best.narrative_framing.style or "hybrid",
                "core_metaphor": best.narrative_framing.core_metaphor,
                "entry_hook": best.narrative_framing.entry_hook,
                "recurring_motif": best.narrative_framing.recurring_motif,
                "description": best.narrative_framing.description,
            },
            "ip_depth_score": max(t.ip_depth_score for t in tones),
            "ip_character_count": max(t.ip_character_count for t in tones),
            "ip_lore_terms": list(dict.fromkeys(all_lore_terms))[:20],
            "ip_community_terms": list(dict.fromkeys(all_community_terms))[:15],
            "program_naming_style": best.program_naming_style or "hybrid",
            "program_naming_examples": list(dict.fromkeys(all_naming_examples))[:10],
            "win_theme_style": best.win_theme_style or "emotional_hook",
            "win_theme_examples": list(dict.fromkeys(all_wt_examples))[:6],
            "text_density_style": best.text_density_style or "balanced",
            "image_slide_ratio": best.image_slide_ratio,
            "text_only_ratio": best.text_only_ratio,
            "tone_rules": list(dict.fromkeys(all_rules))[:15],
            "source_analysis": (
                f"ContentToneProfile 병합: {len(tones)}건 레퍼런스 "
                f"(최고 톤 레벨: {best.emotional_tone_level}/5)"
            ),
        }

    def _infer_content_tone(
        self, docs: List[ReferenceDocument]
    ) -> Dict:
        """텍스트와 패턴에서 콘텐츠 톤 추론 (content_tone 미설정 시 fallback)"""
        all_text = " ".join(d.full_text for d in docs if d.full_text)
        all_patterns = []
        for d in docs:
            all_patterns.extend(d.content_patterns)
        all_programs = []
        for d in docs:
            all_programs.extend(d.program_templates)

        # 1. 감성 톤 레벨 추론
        emotional_level = self._infer_emotional_level(all_text, all_patterns)

        # 2. 내러티브 프레이밍 추론
        framing = self._infer_narrative_framing(all_text, all_patterns)

        # 3. IP 깊이 추론
        ip_depth = self._infer_ip_depth(all_text, docs)

        # 4. 프로그램 네이밍 스타일 추론
        naming = self._infer_program_naming(all_programs, all_text)

        # 5. Win Theme 스타일 추론
        wt_style = self._infer_win_theme_style(all_patterns, all_text)

        # 6. 텍스트 밀도
        density = self._infer_text_density(docs)

        # 7. 산업별 톤 규칙 생성
        industry = docs[0].industry.value if docs[0].industry else "other"
        tone_rules = self._generate_tone_rules(
            industry, emotional_level, framing, ip_depth, naming, density
        )

        return {
            "emotional_tone_level": emotional_level,
            "narrative_framing": framing,
            "ip_depth_score": ip_depth.get("score", 0.0),
            "ip_character_count": ip_depth.get("character_count", 0),
            "ip_lore_terms": ip_depth.get("lore_terms", []),
            "ip_community_terms": ip_depth.get("community_terms", []),
            "program_naming_style": naming.get("style", "functional"),
            "program_naming_examples": naming.get("examples", []),
            "win_theme_style": wt_style.get("style", "keyword_functional"),
            "win_theme_examples": wt_style.get("examples", []),
            "text_density_style": density.get("style", "balanced"),
            "image_slide_ratio": density.get("image_ratio", 0.0),
            "text_only_ratio": density.get("text_only_ratio", 0.0),
            "tone_rules": tone_rules,
            "source_analysis": (
                f"텍스트 추론 기반: {len(docs)}건 레퍼런스, "
                f"총 {len(all_text)}자 분석"
            ),
        }

    def _infer_emotional_level(
        self, text: str, patterns: List[ContentPattern]
    ) -> int:
        """텍스트에서 감성 톤 레벨 추론 (1~5)"""
        if not text:
            return 3  # 기본 중립

        score = 3.0  # 기본 중립 시작

        # 감성 지표 키워드 탐지
        emotional_markers = {
            # 강한 감성 (+0.3 each)
            "high": [
                "도파민", "전율", "심장", "감동", "열광", "열정", "흥분",
                "두근", "짜릿", "환호", "떨림", "몰입", "광기", "전설",
                "영웅", "운명", "부활", "귀환", "각성", "해방",
            ],
            # 중간 감성 (+0.15 each)
            "mid": [
                "특별한", "놀라운", "감각적", "독보적", "압도적",
                "혁신적", "궁극의", "완벽한", "최초의", "유일한",
                "경험", "체험", "순간", "기억", "추억",
            ],
            # 사무적/기능적 (-0.1 each)
            "functional": [
                "효율적", "체계적", "합리적", "안정적", "지속가능",
                "최적화", "관리", "운영", "프로세스", "시스템",
                "분석", "평가", "보고", "검토", "점검",
            ],
        }

        text_lower = text.lower()
        for word in emotional_markers["high"]:
            count = text_lower.count(word)
            score += min(count * 0.3, 0.6)

        for word in emotional_markers["mid"]:
            count = text_lower.count(word)
            score += min(count * 0.15, 0.3)

        for word in emotional_markers["functional"]:
            count = text_lower.count(word)
            score -= min(count * 0.1, 0.3)

        # 콘텐츠 패턴에서 감성 지표
        for p in patterns:
            ptype = p.pattern_type.lower()
            if any(k in ptype for k in ["narrative", "story", "emotion", "hero", "immersive"]):
                score += 0.3
            if any(k in ptype for k in ["show_dont_tell", "cinematic", "reveal"]):
                score += 0.4

        return max(1, min(5, round(score)))

    def _infer_narrative_framing(
        self, text: str, patterns: List[ContentPattern]
    ) -> Dict:
        """내러티브 프레이밍 스타일 추론"""
        framing = {
            "style": "hybrid",
            "core_metaphor": "",
            "entry_hook": "",
            "recurring_motif": "",
            "description": "",
        }

        if not text:
            return framing

        text_lower = text.lower()

        # 세계관 기반 프레이밍 지표
        worldview_score = 0
        for term in ["세계관", "로어", "스토리", "캐릭터", "퀘스트",
                      "미션", "작전", "모험", "마법", "판타지", "sf",
                      "방주", "전쟁", "왕국", "영웅", "마왕", "던전"]:
            if term in text_lower:
                worldview_score += 1

        # 데이터 기반 프레이밍 지표
        data_score = 0
        for term in ["데이터", "분석", "통계", "roi", "kpi", "ctr",
                      "cvr", "reach", "impression", "engagement",
                      "증가율", "성장률", "전환율", "도달률"]:
            if term in text_lower:
                data_score += 1

        # 감성 주도 프레이밍 지표
        emotion_score = 0
        for term in ["감동", "감성", "경험", "체험", "추억", "기억",
                      "특별한", "잊지 못할", "순간", "환호"]:
            if term in text_lower:
                emotion_score += 1

        # 스타일 결정
        scores = {
            "worldview_based": worldview_score,
            "data_driven": data_score,
            "emotion_led": emotion_score,
        }
        top_style = max(scores, key=scores.get)
        if scores[top_style] < 3:
            framing["style"] = "hybrid"
        else:
            framing["style"] = top_style

        # 콘텐츠 패턴에서 내러티브 패턴 추출
        for p in patterns:
            desc = (p.structure or "").lower()
            if "메타포" in desc or "컨셉" in desc:
                framing["core_metaphor"] = p.structure[:100]
            if "후크" in desc or "오프닝" in desc or "인트로" in desc:
                framing["entry_hook"] = p.structure[:100]
            if "반복" in desc or "모티프" in desc or "관통" in desc:
                framing["recurring_motif"] = p.structure[:100]

        framing["description"] = (
            f"프레이밍 스타일: {framing['style']} "
            f"(세계관={worldview_score}, 데이터={data_score}, 감성={emotion_score})"
        )

        return framing

    def _infer_ip_depth(
        self, text: str, docs: List[ReferenceDocument]
    ) -> Dict:
        """IP/브랜드 깊이 추론"""
        result = {
            "score": 0.0,
            "character_count": 0,
            "lore_terms": [],
            "community_terms": [],
        }

        if not text:
            return result

        # 고유명사 밀도 측정 (한글 2글자 이상 대문자/고유명사 패턴)
        # 영문 고유명사 (대문자 시작 단어)
        english_proper = set(re.findall(r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b', text))
        # 영문 대문자 약어 (NIKKE, AGF 등)
        english_acronyms = set(re.findall(r'\b[A-Z]{2,}\b', text))

        # 프로그램 템플릿에서 IP 용어 추출
        ip_terms = set()
        for doc in docs:
            for tmpl in doc.program_templates:
                # 프로그램명에서 IP 고유명사 추출
                words = re.findall(r'[가-힣]{2,}|[A-Z][a-z]+|[A-Z]{2,}', tmpl.name)
                ip_terms.update(words)

        # 커뮤니티 용어 패턴 (비공식 줄임말, 밈 등)
        community_patterns = re.findall(
            r'[가-힣]{2,3}(?:러|이|충|린이|겜|갤|방)',
            text,
        )

        # IP 깊이 점수 계산
        proper_noun_density = len(english_proper) / max(len(text) / 1000, 1)
        has_lore_depth = len(ip_terms) > 5
        has_community = len(community_patterns) > 0

        score = min(1.0, proper_noun_density * 0.1)
        if has_lore_depth:
            score = min(1.0, score + 0.3)
        if has_community:
            score = min(1.0, score + 0.2)

        # 콘텐츠 패턴에서 IP 깊이 지표
        for doc in docs:
            for p in doc.content_patterns:
                if any(k in p.pattern_type.lower() for k in
                       ["character", "lore", "worldview", "ip_deep", "fandom"]):
                    score = min(1.0, score + 0.15)

        result["score"] = round(score, 2)
        result["character_count"] = len(english_proper)
        result["lore_terms"] = list(ip_terms)[:20]
        result["community_terms"] = list(set(community_patterns))[:10]

        return result

    def _infer_program_naming(
        self, programs: List[ProgramTemplate], text: str
    ) -> Dict:
        """프로그램 네이밍 스타일 추론"""
        result = {"style": "functional", "examples": []}

        if not programs:
            return result

        ip_count = 0
        func_count = 0
        branded_count = 0

        for p in programs:
            name = p.name
            result["examples"].append(name)

            # IP 세계관형: 게임/IP 용어 + 영문 브랜드 혼합
            has_ip = bool(re.search(r'[A-Z]{2,}', name))
            has_narrative = any(
                k in name for k in ["작전", "미션", "퀘스트", "전투",
                                     "탐험", "보스", "레이드", "소환"]
            )
            has_colon = ":" in name or "–" in name or "-" in name

            if has_ip and (has_narrative or has_colon):
                ip_count += 1
            elif has_ip:
                branded_count += 1
            else:
                func_count += 1

        total = len(programs)
        if ip_count / total > 0.4:
            result["style"] = "ip_narrative"
        elif branded_count / total > 0.4:
            result["style"] = "branded"
        elif func_count / total > 0.6:
            result["style"] = "functional"
        else:
            result["style"] = "hybrid"

        result["examples"] = result["examples"][:8]
        return result

    def _infer_win_theme_style(
        self, patterns: List[ContentPattern], text: str
    ) -> Dict:
        """Win Theme 스타일 추론"""
        result = {"style": "keyword_functional", "examples": []}

        # 콘텐츠 패턴에서 Win Theme 관련 패턴 추출
        wt_patterns = [
            p for p in patterns
            if any(k in (p.section_context or "").lower()
                   for k in ["concept", "strategy", "theme", "hook"])
        ]

        for p in wt_patterns:
            for ex in (p.examples or []):
                result["examples"].append(ex)

        # 텍스트에서 Win Theme 스타일 추론
        if text:
            text_lower = text.lower()
            has_worldview = any(
                k in text_lower for k in ["세계관", "모험", "여정", "전설"]
            )
            has_emotional = any(
                k in text_lower for k in ["도파민", "전율", "심장", "감동"]
            )

            if has_worldview:
                result["style"] = "ip_worldview"
            elif has_emotional:
                result["style"] = "emotional_hook"

        result["examples"] = result["examples"][:6]
        return result

    def _infer_text_density(self, docs: List[ReferenceDocument]) -> Dict:
        """텍스트 밀도 스타일 추론"""
        result = {
            "style": "balanced",
            "image_ratio": 0.0,
            "text_only_ratio": 0.0,
        }

        image_freqs = []
        text_freqs = []

        for doc in docs:
            if not doc.design_profile:
                continue
            for lp in doc.design_profile.layout_patterns:
                pt = lp.pattern_type.lower()
                if "image" in pt or "full_bleed" in pt:
                    image_freqs.append(lp.frequency)
                elif "content_standard" in pt or "text" in pt:
                    text_freqs.append(lp.frequency)

        img_ratio = sum(image_freqs) if image_freqs else 0.2
        text_ratio = sum(text_freqs) if text_freqs else 0.15

        result["image_ratio"] = round(min(1.0, img_ratio), 2)
        result["text_only_ratio"] = round(min(1.0, text_ratio), 2)

        if img_ratio > 0.6:
            result["style"] = "minimal"
        elif img_ratio > 0.4:
            result["style"] = "balanced"
        elif text_ratio > 0.3:
            result["style"] = "rich"
        else:
            result["style"] = "balanced"

        return result

    def _generate_tone_rules(
        self,
        industry: str,
        emotional_level: int,
        framing: Dict,
        ip_depth: Dict,
        naming: Dict,
        density: Dict,
    ) -> List[str]:
        """산업별 + 분석 결과 기반 콘텐츠 톤 규칙 생성"""
        rules = []

        # ── 산업별 기본 규칙 ──
        industry_rules = {
            "game_event": [
                "게임 IP의 세계관과 캐릭터를 제안서 전체에 자연스럽게 녹여라",
                "커뮤니티에서 사용하는 팬 용어/밈을 적극 활용하여 '우리편'임을 보여라",
                "프로그램명에 IP 세계관 용어를 삽입하라 (기능 설명형 네이밍 지양)",
                "게임 유저의 감성 언어를 사용하라 (도파민, 파밍, 가챠 등)",
                "IP 캐릭터의 고유 대사/성격을 프로그램 기획에 반영하라",
            ],
            "marketing_pr": [
                "타겟 오디언스의 언어와 감성 코드로 소통하라",
                "데이터 인사이트를 감성적 스토리로 포장하라 (숫자만 나열하지 말라)",
                "브랜드 보이스와 일관된 톤앤매너를 유지하라",
                "트렌드 용어를 자연스럽게 활용하되 과하지 않게 하라",
            ],
            "event": [
                "참여자의 감정 여정(기대→체험→감동→공유)을 설계하라",
                "현장감이 느껴지는 생생한 묘사를 포함하라",
                "공간과 동선의 경험을 스토리텔링으로 전달하라",
            ],
            "it_system": [
                "기술적 정확성을 우선하되 비전문가도 이해 가능한 표현을 사용하라",
                "아키텍처와 프로세스를 시각적 도식으로 설명하라",
                "기술 도입의 비즈니스 임팩트를 명확히 연결하라",
            ],
            "public": [
                "시민/주민 관점에서 체감 가능한 효과를 강조하라",
                "정책 목표와의 정렬을 명확히 보여라",
                "쉬운 언어로 전문 내용을 설명하라",
            ],
            "consulting": [
                "프레임워크와 방법론의 차별성을 강조하라",
                "유사 프로젝트 성과를 구체적 수치로 제시하라",
                "분석 결과를 인사이트로 전환하여 전달하라",
                "As-Is → To-Be 구조로 변화의 필요성과 방향성을 설계하라",
                "워크숍/인터뷰 등 참여형 방법론을 통해 현장 밀착감을 보여라",
            ],
            "finance": [
                "규제 준수와 리스크 관리 역량을 명확히 보여라",
                "수치와 데이터 기반의 정량적 논거를 핵심으로 배치하라",
                "금융 전문 용어는 정확하게 사용하되 의사결정자 관점에서 설명하라",
            ],
            "education": [
                "학습자 관점에서 교육 효과와 경험의 변화를 생생하게 전달하라",
                "커리큘럼 설계의 체계성과 교수학습 방법론의 근거를 제시하라",
                "EdTech 도구 도입의 학습 성과 연결고리를 데이터로 뒷받침하라",
            ],
            "healthcare": [
                "임상 근거와 규제 요건을 명확히 연결하여 전달하라",
                "환자/사용자 관점에서 체감 가능한 효과를 강조하라",
                "인허가 요건과 품질관리 체계를 구체적으로 제시하라",
            ],
        }

        rules.extend(industry_rules.get(industry, [
            "프로젝트 특성에 맞는 전문 용어와 감성 코드를 활용하라",
        ]))

        # ── 감성 톤 레벨 기반 규칙 ──
        if emotional_level >= 4:
            rules.extend([
                "기능 나열보다 감성적 스토리텔링을 우선하라",
                "제안서 전체에 하나의 내러티브 메타포를 관통시켜라",
                "Action Title에 감성적 후크를 포함하라 (사실 전달 < 감정 유발)",
            ])
        elif emotional_level >= 3:
            rules.extend([
                "데이터와 감성을 6:4 비율로 균형있게 배합하라",
                "핵심 메시지는 감성적으로, 뒷받침은 데이터로 구성하라",
            ])
        else:
            rules.extend([
                "간결하고 사실 중심의 전달을 유지하라",
                "감성 표현은 HOOK/CLOSING에만 제한적으로 사용하라",
            ])

        # ── IP 깊이 기반 규칙 ──
        if ip_depth.get("score", 0) > 0.5:
            rules.extend([
                "IP 고유명사와 세계관 용어를 제안서 전체에 일관되게 사용하라",
                "캐릭터 이름, 로어 키워드를 프로그램 설계에 직접 반영하라",
            ])
            if ip_depth.get("community_terms"):
                rules.append(
                    "팬 커뮤니티 용어를 자연스럽게 섞어 '인사이더' 느낌을 주라"
                )

        # ── 프레이밍 기반 규칙 ──
        framing_style = framing.get("style", "hybrid")
        if framing_style == "worldview_based":
            rules.append("IP/브랜드 세계관을 제안서 구조의 뼈대로 사용하라")
        elif framing_style == "emotion_led":
            rules.append("감정의 기승전결 구조로 제안서를 설계하라")

        # ── 네이밍 스타일 규칙 ──
        naming_style = naming.get("style", "functional")
        if naming_style == "ip_narrative":
            rules.append(
                "모든 프로그램/이벤트명에 IP 세계관 키워드를 포함하라"
            )
        elif naming_style == "branded":
            rules.append(
                "프로그램명은 브랜드 영문명 + 한글 설명 조합으로 구성하라"
            )

        # ── 텍스트 밀도 규칙 ──
        density_style = density.get("style", "balanced")
        if density_style == "minimal":
            rules.extend([
                "텍스트를 최소화하고 이미지/비주얼 비중을 70% 이상 유지하라",
                "한 슬라이드에 핵심 메시지 1개만 담아라 (Show, Don't Tell)",
            ])
        elif density_style == "rich":
            rules.append(
                "충분한 텍스트로 깊이감을 보여주되 불릿 구조로 가독성을 확보하라"
            )

        return rules

    def _default_content_tone(self, industry: Optional[str] = None) -> Dict:
        """레퍼런스 없을 때 산업별 기본 콘텐츠 톤"""
        # 산업별 기본 설정
        industry_defaults = {
            "game_event": {
                "emotional_tone_level": 4,
                "narrative_framing_style": "worldview_based",
                "program_naming_style": "ip_narrative",
                "win_theme_style": "ip_worldview",
                "text_density_style": "minimal",
            },
            "marketing_pr": {
                "emotional_tone_level": 4,
                "narrative_framing_style": "emotion_led",
                "program_naming_style": "branded",
                "win_theme_style": "emotional_hook",
                "text_density_style": "balanced",
            },
            "event": {
                "emotional_tone_level": 3,
                "narrative_framing_style": "emotion_led",
                "program_naming_style": "branded",
                "win_theme_style": "emotional_hook",
                "text_density_style": "balanced",
            },
            "it_system": {
                "emotional_tone_level": 2,
                "narrative_framing_style": "data_driven",
                "program_naming_style": "functional",
                "win_theme_style": "keyword_functional",
                "text_density_style": "rich",
            },
            "public": {
                "emotional_tone_level": 2,
                "narrative_framing_style": "data_driven",
                "program_naming_style": "functional",
                "win_theme_style": "keyword_functional",
                "text_density_style": "rich",
            },
            "consulting": {
                "emotional_tone_level": 2,
                "narrative_framing_style": "data_driven",
                "program_naming_style": "functional",
                "win_theme_style": "keyword_functional",
                "text_density_style": "rich",
            },
            "finance": {
                "emotional_tone_level": 2,
                "narrative_framing_style": "data_driven",
                "program_naming_style": "functional",
                "win_theme_style": "keyword_functional",
                "text_density_style": "rich",
            },
            "education": {
                "emotional_tone_level": 3,
                "narrative_framing_style": "emotion_led",
                "program_naming_style": "branded",
                "win_theme_style": "emotional_hook",
                "text_density_style": "balanced",
            },
            "healthcare": {
                "emotional_tone_level": 2,
                "narrative_framing_style": "data_driven",
                "program_naming_style": "functional",
                "win_theme_style": "keyword_functional",
                "text_density_style": "rich",
            },
        }

        defaults = industry_defaults.get(industry or "", {})
        emotional_level = defaults.get("emotional_tone_level", 3)
        framing_style = defaults.get("narrative_framing_style", "hybrid")
        naming_style = defaults.get("program_naming_style", "functional")
        wt_style = defaults.get("win_theme_style", "keyword_functional")
        density_style = defaults.get("text_density_style", "balanced")

        # 산업별 톤 규칙 생성
        tone_rules = self._generate_tone_rules(
            industry=industry or "other",
            emotional_level=emotional_level,
            framing={"style": framing_style},
            ip_depth={"score": 0.0},
            naming={"style": naming_style},
            density={"style": density_style},
        )

        return {
            "emotional_tone_level": emotional_level,
            "narrative_framing": {
                "style": framing_style,
                "core_metaphor": "",
                "entry_hook": "",
                "recurring_motif": "",
                "description": f"기본 톤 ({industry or 'other'} 산업 기본값)",
            },
            "ip_depth_score": 0.0,
            "ip_character_count": 0,
            "ip_lore_terms": [],
            "ip_community_terms": [],
            "program_naming_style": naming_style,
            "program_naming_examples": [],
            "win_theme_style": wt_style,
            "win_theme_examples": [],
            "text_density_style": density_style,
            "image_slide_ratio": 0.25,
            "text_only_ratio": 0.15,
            "tone_rules": tone_rules,
            "source_analysis": f"기본 톤 ({industry or 'other'} 산업, 레퍼런스 없음)",
        }

    # ── Default brief ──

    def _build_default_brief(
        self, target_slides: int, industry: Optional[str] = None,
    ) -> DesignBrief:
        """레퍼런스 없을 때 기본 브리프"""
        brief = DesignBrief()
        brief.section_weights = {
            "HOOK": 4,
            "EXECUTIVE SUMMARY": 2,
            "INSIGHT": 6,
            "CONCEPT & STRATEGY": 8,
            "ACTION PLAN": round(target_slides * 0.35),
            "MANAGEMENT": 5,
            "WHY US": 4,
            "BUDGET & TIMELINE": 4,
            "CLOSING": 3,
        }
        brief.component_targets = {
            "HERO_IMAGE": 5,
            "SPLIT_VISUAL": 4,
            "MOOD_BOARD": 2,
            "IMG_PH": 18,
            "COLS": 8,
            "HIGHLIGHT": 6,
            "TABLE": 3,
            "FLOW": 3,
            "METRIC_CARD": 8,
            "SECTION_BRIDGE": 5,
            "ZONE_MAP": 2,
        }
        brief.layout_distribution = {
            "complex_diagram": 0.35,
            "image_focused": 0.20,
            "hero_typography": 0.10,
            "content_standard": 0.15,
            "split_image_text": 0.10,
            "table_based": 0.05,
            "other": 0.05,
        }
        brief.background_schedule = self._build_bg_schedule([], target_slides)
        brief.visual_density_targets = {
            "image_slides_pct": 0.25,
            "diagram_slides_pct": 0.35,
            "text_only_max_pct": 0.20,
            "bg_variety_target": "60% white, 20% light, 10% dark, 10% gradient",
        }
        brief.content_tone = self._default_content_tone(industry=industry)
        return brief
