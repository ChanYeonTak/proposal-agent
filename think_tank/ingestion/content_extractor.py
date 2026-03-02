"""
콘텐츠 패턴 & 프로그램 템플릿 추출기 (v4.0)

PPTX 슬라이드 시퀀스를 분석하여 ContentPattern, ProgramTemplate을 추출합니다.
design_extractor가 시각적 레이아웃을 분석한다면,
content_extractor는 콘텐츠 흐름과 스토리 구조를 분석합니다.
"""

from __future__ import annotations

from collections import Counter
from typing import Dict, List, Optional

from ..models import ContentPattern, ProgramTemplate, SectionStructure
from src.utils.logger import get_logger

logger = get_logger("content_extractor")

# 프로그램/이벤트 관련 키워드 맵
PROGRAM_KEYWORDS = {
    "booth_design": ["부스", "booth", "design", "시공", "디자인", "설치", "조형물", "조명",
                     "평면도", "floor", "layout", "인포데스크", "desk"],
    "stage": ["스테이지", "stage", "무대", "공연", "밴드", "라이브", "live", "MC",
              "코스프레", "cosplay", "토크쇼", "talk", "DJ", "퍼포먼스"],
    "game_program": ["게임", "game", "program", "미니게임", "슈팅", "가챠", "gacha",
                     "캡슐", "머신", "아케이드", "arcade", "체험"],
    "interaction": ["인터랙션", "interaction", "미션", "mission", "스탬프", "stamp",
                    "투어", "tour", "QR", "AR", "VR", "포토존", "photo"],
    "campaign": ["캠페인", "campaign", "SNS", "바이럴", "viral", "해시태그", "hashtag",
                 "KOL", "인플루언서", "influencer", "팬아트", "fanart", "프로모션"],
    "goods_md": ["굿즈", "goods", "MD", "머천다이즈", "merchandise", "구매",
                 "한정판", "limited", "판매", "sale"],
    "operation": ["운영", "operation", "인력", "staffing", "스케줄", "schedule",
                  "매뉴얼", "manual", "비상", "emergency", "안전", "safety"],
    "event_pack": ["보상", "reward", "경품", "prize", "이벤트", "event",
                   "추첨", "draw", "럭키", "lucky", "참여"],
}


class ContentExtractor:
    """PPTX 콘텐츠 패턴 및 프로그램 템플릿 추출"""

    def extract_content_patterns(
        self,
        prs,
        slide_texts: List[Dict],
        sections: List[SectionStructure],
    ) -> List[ContentPattern]:
        """슬라이드 시퀀스를 분석하여 콘텐츠 패턴 식별

        Args:
            prs: python-pptx Presentation 객체
            slide_texts: 슬라이드별 텍스트 딕셔너리 리스트
            sections: 섹션 구조 리스트

        Returns:
            List[ContentPattern]: 식별된 콘텐츠 패턴
        """
        patterns = []

        # 1. 섹션별 패턴 분석
        section_start = 0
        for section in sections:
            end = section_start + section.slide_count
            section_slides = slide_texts[section_start:end]
            section_patterns = self._analyze_section_patterns(
                prs, section_slides, section.name, section_start
            )
            patterns.extend(section_patterns)
            section_start = end

        # 2. 전역 패턴 분석 (슬라이드 간 시퀀스)
        global_patterns = self._analyze_global_patterns(prs, slide_texts)
        patterns.extend(global_patterns)

        logger.info(f"콘텐츠 패턴 추출 완료: {len(patterns)}건")
        return patterns

    def extract_program_templates(
        self,
        prs,
        slide_texts: List[Dict],
        sections: List[SectionStructure],
    ) -> List[ProgramTemplate]:
        """프로그램/이벤트 템플릿 추출

        슬라이드 텍스트에서 부스, 스테이지, 게임, 캠페인 등의
        프로그램 구조를 식별하여 ProgramTemplate으로 추출.
        """
        templates = []
        section_start = 0

        for section in sections:
            end = section_start + section.slide_count
            section_slides = slide_texts[section_start:end]

            # 섹션 전체 텍스트
            section_full = " ".join(
                st["full_text"] for st in section_slides
            ).lower()

            # 카테고리 매칭 (키워드 빈도 기반)
            category_scores = {}
            for cat, keywords in PROGRAM_KEYWORDS.items():
                score = sum(section_full.count(kw.lower()) for kw in keywords)
                if score > 2:  # 최소 3회 이상 등장
                    category_scores[cat] = score

            if category_scores:
                best_cat = max(category_scores, key=category_scores.get)
                # 메커니즘/보상/운영 텍스트 추출
                mechanism = self._extract_mechanism(section_slides)
                reward = self._extract_reward_info(section_slides)
                visual_elements = self._extract_visual_elements(
                    prs, section_start, end
                )

                templates.append(ProgramTemplate(
                    name=section.name,
                    category=best_cat,
                    mechanism=mechanism,
                    reward_structure=reward,
                    operation_plan="",
                    slide_count=section.slide_count,
                    visual_elements=visual_elements,
                ))

            section_start = end

        logger.info(f"프로그램 템플릿 추출 완료: {len(templates)}건")
        return templates

    def extract_slide_compositions(self, prs) -> List[Dict]:
        """슬라이드별 시각 구성 분석 (shape 기반)

        Returns:
            List[Dict]: 각 슬라이드의 구성 정보
                - shape_count, text_count, image_count, table_count
                - text_density (chars per slide area)
                - image_coverage (이미지가 차지하는 면적 비율)
                - composition_type (분류)
                - background_brightness ("dark" / "light")
        """
        slide_w = prs.slide_width / 914400 if prs.slide_width else 13.333
        slide_h = prs.slide_height / 914400 if prs.slide_height else 7.5
        slide_area = slide_w * slide_h
        compositions = []

        for i, slide in enumerate(prs.slides):
            shapes = list(slide.shapes)
            text_shapes = [s for s in shapes if s.has_text_frame]
            image_shapes = [s for s in shapes if s.shape_type == 13]
            table_shapes = [s for s in shapes if s.has_table]
            auto_shapes = [s for s in shapes
                           if hasattr(s, 'shape_type') and s.shape_type not in (13, 19)]

            # 텍스트 밀도
            total_chars = sum(
                len(s.text_frame.text) for s in text_shapes
            )
            text_density = total_chars / slide_area

            # 이미지 커버리지
            image_area = 0.0
            for img in image_shapes:
                if img.width and img.height:
                    iw = img.width / 914400
                    ih = img.height / 914400
                    image_area += iw * ih
            image_coverage = image_area / slide_area

            # 도형 공간 분포 분석
            shape_rects = []
            for s in shapes:
                if s.left is not None and s.top is not None:
                    shape_rects.append((
                        s.left / 914400, s.top / 914400,
                        (s.width or 0) / 914400, (s.height or 0) / 914400,
                    ))

            comp_type = self._classify_composition(
                len(shapes), len(text_shapes), len(image_shapes),
                len(table_shapes), len(auto_shapes),
                text_density, image_coverage, shape_rects
            )

            # 배경 밝기
            bg_brightness = self._detect_slide_brightness(slide)

            compositions.append({
                "slide_number": i + 1,
                "shape_count": len(shapes),
                "text_count": len(text_shapes),
                "image_count": len(image_shapes),
                "table_count": len(table_shapes),
                "auto_shape_count": len(auto_shapes),
                "text_density": round(text_density, 2),
                "image_coverage": round(image_coverage, 3),
                "composition_type": comp_type,
                "background": bg_brightness,
            })

        return compositions

    # ── Private helpers ──

    def _analyze_section_patterns(
        self, prs, section_slides, section_name, start_idx
    ) -> List[ContentPattern]:
        """섹션 내 슬라이드 시퀀스에서 패턴 식별"""
        patterns = []

        if not section_slides:
            return patterns

        # 슬라이드 구성 분석
        compositions = []
        for i, st in enumerate(section_slides):
            slide_idx = start_idx + i
            if slide_idx < len(prs.slides):
                slide = prs.slides[slide_idx]
                shapes = list(slide.shapes)
                n_shapes = len(shapes)
                n_images = sum(1 for s in shapes if s.shape_type == 13)
                n_text = sum(1 for s in shapes if s.has_text_frame)
                n_table = sum(1 for s in shapes if s.has_table)
                total_chars = sum(
                    len(s.text_frame.text) for s in shapes if s.has_text_frame
                )
                compositions.append({
                    "shapes": n_shapes, "images": n_images,
                    "text": n_text, "tables": n_table,
                    "chars": total_chars,
                })

        # 패턴 1: 이미지 중심 시퀀스 (이미지 비율 높은 연속 슬라이드)
        image_heavy = sum(
            1 for c in compositions if c["images"] > 0
        )
        if image_heavy >= 2:
            patterns.append(ContentPattern(
                pattern_type="show_dont_tell",
                section_context=section_name,
                structure=f"이미지 중심 {image_heavy}슬라이드 시퀀스 (섹션 {len(compositions)}p 중 {image_heavy}p가 이미지 포함)",
                slide_count=image_heavy,
            ))

        # 패턴 2: 데이터 집약 시퀀스 (테이블/차트 연속)
        data_heavy = sum(
            1 for c in compositions
            if c["tables"] > 0 or c["shapes"] > 10
        )
        if data_heavy >= 2:
            patterns.append(ContentPattern(
                pattern_type="data_narrative",
                section_context=section_name,
                structure=f"데이터 중심 {data_heavy}슬라이드 (테이블/다이어그램 활용)",
                slide_count=data_heavy,
            ))

        # 패턴 3: 도식 중심 (많은 도형)
        diagram_heavy = sum(
            1 for c in compositions if c["shapes"] > 8 and c["images"] == 0
        )
        if diagram_heavy >= 2:
            patterns.append(ContentPattern(
                pattern_type="complex_diagram",
                section_context=section_name,
                structure=f"복합 도식 {diagram_heavy}슬라이드 (도형 8+개 활용, 플로우/구조도)",
                slide_count=diagram_heavy,
            ))

        # 패턴 4: 텍스트 라이트 (간결한 슬라이드)
        light_slides = sum(
            1 for c in compositions if c["chars"] < 100 and c["shapes"] <= 5
        )
        if light_slides >= 2:
            patterns.append(ContentPattern(
                pattern_type="hero_statement",
                section_context=section_name,
                structure=f"간결한 히어로 {light_slides}슬라이드 (100자 미만, 도형 5개 이하)",
                slide_count=light_slides,
            ))

        return patterns

    def _analyze_global_patterns(self, prs, slide_texts) -> List[ContentPattern]:
        """전체 프레젠테이션 수준 패턴 분석"""
        patterns = []
        total = len(prs.slides)

        # 전체 시각 밀도 통계
        image_slides = 0
        diagram_slides = 0
        text_only_slides = 0

        for slide in prs.slides:
            shapes = list(slide.shapes)
            has_image = any(s.shape_type == 13 for s in shapes)
            is_complex = len(shapes) > 10
            is_text_only = all(
                s.has_text_frame or s.has_table for s in shapes
            ) and len(shapes) <= 5

            if has_image:
                image_slides += 1
            if is_complex:
                diagram_slides += 1
            if is_text_only:
                text_only_slides += 1

        if total > 0:
            patterns.append(ContentPattern(
                pattern_type="visual_density",
                section_context="GLOBAL",
                structure=(
                    f"전체 {total}p 중: "
                    f"이미지 포함 {image_slides}p ({image_slides*100//total}%), "
                    f"복합 도식 {diagram_slides}p ({diagram_slides*100//total}%), "
                    f"텍스트 위주 {text_only_slides}p ({text_only_slides*100//total}%)"
                ),
                slide_count=total,
            ))

        return patterns

    def _extract_mechanism(self, section_slides: List[Dict]) -> str:
        """프로그램 메커니즘 텍스트 추출 (첫 2슬라이드 요약)"""
        texts = []
        for st in section_slides[:3]:
            for t in st["texts"][:3]:
                if len(t) > 10:
                    texts.append(t)
        return " / ".join(texts[:5])[:300]

    def _extract_reward_info(self, section_slides: List[Dict]) -> str:
        """보상/경품 관련 텍스트 추출"""
        reward_keywords = ["보상", "경품", "reward", "prize", "상품", "특전",
                           "포인트", "쿠폰", "교환", "한정"]
        for st in section_slides:
            for t in st["texts"]:
                if any(kw in t.lower() for kw in reward_keywords):
                    return t[:200]
        return ""

    def _extract_visual_elements(self, prs, start, end) -> List[str]:
        """슬라이드 범위에서 시각 요소 유형 추출"""
        elements = set()
        for i in range(start, min(end, len(prs.slides))):
            slide = prs.slides[i]
            shapes = list(slide.shapes)
            if any(s.shape_type == 13 for s in shapes):
                elements.add("image")
            if any(s.has_table for s in shapes):
                elements.add("table")
            if len(shapes) > 10:
                elements.add("complex_diagram")
            if len(shapes) <= 4:
                elements.add("minimal_layout")
        return sorted(elements)

    def _classify_composition(
        self, n_shapes, n_text, n_images, n_tables, n_auto,
        text_density, image_coverage, rects
    ) -> str:
        """슬라이드 구성 유형 분류 (12 유형)"""
        # full_bleed_image: 이미지가 슬라이드 대부분 차지
        if image_coverage > 0.5:
            return "full_bleed_image"
        # split_image_text: 이미지 + 텍스트 혼합
        if n_images > 0 and n_text > 2:
            return "split_image_text"
        # multi_image_grid: 다중 이미지
        if n_images >= 3:
            return "multi_image_grid"
        # hero_typography: 1~2 텍스트 요소, 최소 도형
        if n_text <= 2 and n_shapes <= 4:
            return "hero_typography"
        # data_card_array: 많은 소형 도형 (카드 배열)
        if n_auto >= 6 and n_tables == 0:
            return "data_card_array"
        # table_data: 테이블 포함
        if n_tables > 0:
            return "table_data"
        # process_flow: 중간 도형 수 + 텍스트
        if 5 <= n_shapes <= 12 and n_auto >= 3:
            return "process_flow"
        # layered_composition: 많은 겹침 (고밀도 도형)
        if n_shapes > 15:
            return "layered_composition"
        # hierarchical_diagram: 다단 구조
        if n_shapes > 10:
            return "hierarchical_diagram"
        # content_standard: 일반
        if n_text > 2:
            return "content_standard"
        # minimal: 최소 요소
        if n_shapes <= 3:
            return "minimal"

        return "other"

    def _detect_slide_brightness(self, slide) -> str:
        """슬라이드 배경 밝기 감지"""
        try:
            bg_fill = slide.background.fill
            if bg_fill and bg_fill.type is not None:
                if hasattr(bg_fill, 'fore_color') and bg_fill.fore_color:
                    rgb = str(bg_fill.fore_color.rgb)
                    r = int(rgb[0:2], 16)
                    g = int(rgb[2:4], 16)
                    b = int(rgb[4:6], 16)
                    brightness = (r + g + b) / 3
                    return "dark" if brightness < 128 else "light"
        except Exception:
            pass
        return "light"
