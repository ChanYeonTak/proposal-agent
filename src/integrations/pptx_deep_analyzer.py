"""PPTX 레퍼런스 심층 분석기.

디자인 업그레이드를 위해 원본 PPTX에서 다음을 추출:
- 폰트 시스템 (이름/크기/웨이트/사용 빈도)
- 컬러 팔레트 (솔리드/그라디언트, 빈도순)
- 도형 타입 및 이펙트 (라운드/그림자/그라디언트/아웃라인)
- 타이포그래피 위계 (타이틀/본문/캡션)
- 레이아웃 패턴 (슬라이드 유형 분류)
- 시각 효과 (블러/투명도/이미지 처리)
"""
from __future__ import annotations

import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


@dataclass
class Analysis:
    typefaces: Counter = field(default_factory=Counter)
    font_sizes: Counter = field(default_factory=Counter)
    bold_count: int = 0
    italic_count: int = 0
    underline_count: int = 0
    char_spacing: Counter = field(default_factory=Counter)

    solid_colors: Counter = field(default_factory=Counter)  # rgb hex
    scheme_colors: Counter = field(default_factory=Counter)
    gradient_count: int = 0
    gradient_samples: list = field(default_factory=list)

    shape_types: Counter = field(default_factory=Counter)  # preset geom
    rounded_rect_count: int = 0
    shadow_count: int = 0
    outline_widths: Counter = field(default_factory=Counter)
    fill_none: int = 0

    alpha_samples: Counter = field(default_factory=Counter)
    blur_count: int = 0

    image_refs: int = 0
    group_count: int = 0

    text_box_per_slide: list = field(default_factory=list)
    slide_background_types: Counter = field(default_factory=Counter)


def _rgb_hex(val: str) -> str:
    val = val.strip("#").upper()
    if len(val) == 6:
        return val
    return val


def analyze(pptx_path: Path) -> Analysis:
    a = Analysis()

    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_files = [
            n for n in z.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        ]
        for name in slide_files:
            data = z.read(name).decode("utf-8", errors="replace")
            _parse_slide(data, a)

        # 테마 파일도 보기
        for name in z.namelist():
            if name.startswith("ppt/theme/") and name.endswith(".xml"):
                data = z.read(name).decode("utf-8", errors="replace")
                _parse_theme(data, a)

    return a


def _parse_slide(xml: str, a: Analysis) -> None:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return

    # 타이프페이스 (모든 위치)
    for node in root.iter():
        tag = node.tag.split("}")[-1]

        if tag in ("latin", "ea", "cs", "font"):
            tf = node.get("typeface")
            if tf:
                a.typefaces[tf] += 1

        if tag == "rPr":
            sz = node.get("sz")
            if sz:
                try:
                    pt = int(sz) / 100.0
                    a.font_sizes[round(pt)] += 1
                except ValueError:
                    pass
            if node.get("b") == "1":
                a.bold_count += 1
            if node.get("i") == "1":
                a.italic_count += 1
            if node.get("u") and node.get("u") != "none":
                a.underline_count += 1
            spc = node.get("spc")
            if spc:
                try:
                    a.char_spacing[int(spc)] += 1
                except ValueError:
                    pass

        # 컬러
        if tag == "srgbClr":
            val = node.get("val")
            if val:
                # alpha modifier?
                alpha = None
                for child in node:
                    if child.tag.endswith("}alpha"):
                        alpha = child.get("val")
                if alpha:
                    try:
                        a.alpha_samples[int(alpha) // 1000] += 1  # % 단위 근사
                    except ValueError:
                        pass
                a.solid_colors[_rgb_hex(val)] += 1

        if tag == "schemeClr":
            val = node.get("val")
            if val:
                a.scheme_colors[val] += 1

        # 그라디언트
        if tag == "gradFill":
            a.gradient_count += 1
            stops = []
            for gs in node.iter():
                if gs.tag.endswith("}gs"):
                    pos = gs.get("pos")
                    for c in gs:
                        if c.tag.endswith("}srgbClr"):
                            stops.append((pos, c.get("val")))
            if stops and len(a.gradient_samples) < 20:
                a.gradient_samples.append(stops)

        # 도형 지오메트리
        if tag == "prstGeom":
            prst = node.get("prst")
            if prst:
                a.shape_types[prst] += 1
                if "roundRect" in prst or prst == "round2SameRect":
                    a.rounded_rect_count += 1

        # 그림자
        if tag in ("outerShdw", "innerShdw"):
            a.shadow_count += 1

        # 아웃라인
        if tag == "ln":
            w = node.get("w")
            if w:
                try:
                    pt = int(w) / 12700
                    a.outline_widths[round(pt, 1)] += 1
                except ValueError:
                    pass
            # noFill 확인
            for child in node:
                if child.tag.endswith("}noFill"):
                    a.fill_none += 1

        # 블러
        if tag == "blur":
            a.blur_count += 1

        # 이미지
        if tag == "blip":
            a.image_refs += 1

        # 그룹
        if tag == "grpSp":
            a.group_count += 1

    # 슬라이드당 텍스트 박스 수
    sp_count = sum(1 for _ in root.iter() if _.tag.endswith("}sp"))
    a.text_box_per_slide.append(sp_count)

    # 배경 타입
    for bg in root.iter():
        if bg.tag.endswith("}bgPr"):
            for child in bg:
                t = child.tag.split("}")[-1]
                a.slide_background_types[t] += 1
        if bg.tag.endswith("}bgRef"):
            a.slide_background_types["themeRef"] += 1


def _parse_theme(xml: str, a: Analysis) -> None:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return
    for node in root.iter():
        tag = node.tag.split("}")[-1]
        if tag in ("latin", "ea", "cs", "font"):
            tf = node.get("typeface")
            if tf:
                a.typefaces[f"[theme] {tf}"] += 1


def print_analysis(a: Analysis) -> None:
    print("\n" + "=" * 70)
    print("PPTX 심층 분석 결과")
    print("=" * 70)

    print(f"\n[타이포그래피]")
    print(f"  굵게 사용 : {a.bold_count}회")
    print(f"  기울임    : {a.italic_count}회")
    print(f"  밑줄      : {a.underline_count}회")
    print(f"  상위 폰트 (top 15):")
    for name, cnt in a.typefaces.most_common(15):
        print(f"    {cnt:>5}  {name}")
    print(f"  상위 폰트 크기 (pt, top 10):")
    for sz, cnt in sorted(a.font_sizes.most_common(10), key=lambda x: -x[1]):
        print(f"    {cnt:>5}  {sz}pt")
    print(f"  자간(spc) 샘플 (top 5):")
    for spc, cnt in a.char_spacing.most_common(5):
        print(f"    {cnt:>5}  spc={spc}")

    print(f"\n[컬러]")
    print(f"  고유 RGB 색상 : {len(a.solid_colors)}개")
    print(f"  상위 20색 (hex : 빈도):")
    for hx, cnt in a.solid_colors.most_common(20):
        print(f"    {cnt:>5}  #{hx}")
    if a.scheme_colors:
        print(f"  스킴 컬러 사용:")
        for sc, cnt in a.scheme_colors.most_common(10):
            print(f"    {cnt:>5}  {sc}")

    print(f"\n[그라디언트]")
    print(f"  gradFill 총 : {a.gradient_count}회")
    for i, stops in enumerate(a.gradient_samples[:5]):
        print(f"  샘플 {i+1}: {stops}")

    print(f"\n[도형 & 이펙트]")
    print(f"  도형 지오메트리 (top 15):")
    for geo, cnt in a.shape_types.most_common(15):
        print(f"    {cnt:>5}  {geo}")
    print(f"  라운드 사각형 : {a.rounded_rect_count}회")
    print(f"  그림자 효과   : {a.shadow_count}회")
    print(f"  블러 효과     : {a.blur_count}회")
    print(f"  그룹화        : {a.group_count}회")
    print(f"  아웃라인 두께 (pt, top 8):")
    for w, cnt in a.outline_widths.most_common(8):
        print(f"    {cnt:>5}  {w}pt")

    print(f"\n[투명도/알파]")
    if a.alpha_samples:
        for ap, cnt in sorted(a.alpha_samples.most_common(10), key=lambda x: x[0]):
            print(f"    {cnt:>5}  α~{ap}%")
    else:
        print("    (알파 값 없음 또는 추출 실패)")

    print(f"\n[레이아웃 밀도]")
    if a.text_box_per_slide:
        avg = sum(a.text_box_per_slide) / len(a.text_box_per_slide)
        print(f"  슬라이드당 도형 수 (sp): 평균 {avg:.1f}, 최대 {max(a.text_box_per_slide)}, 최소 {min(a.text_box_per_slide)}")

    print(f"\n[기타]")
    print(f"  이미지 참조(blip) : {a.image_refs}회")
    print(f"  배경 타입:")
    for t, cnt in a.slide_background_types.most_common():
        print(f"    {cnt:>5}  {t}")

    print("\n" + "=" * 70)


if __name__ == "__main__":
    import sys
    path = Path(sys.argv[1])
    result = analyze(path)
    print_analysis(result)
