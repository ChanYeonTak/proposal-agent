"""마비노기 모바일 2026년 6월 쇼케이스 프로모션 대행 제안서 — BIG CAMPFIRE.

RFP (2026.04.20):
- 타이틀: 마비노기 모바일 6월 업데이트 쇼케이스 (가칭; 빅 캠프파이어 쇼케이스)
- 진행 일정: 2026년 6월 중순
- 타겟: 마비노기 모바일 액티브 유저
- 방식: 마비노기 모바일 공식 유튜브 채널 라이브 스트리밍 (오프라인 없음)
- 주요 안내: 신규 시즌, 업데이트 핵심 콘텐츠, 개선 사항, 운영 방향, Q&A
- 예산: Max 5억 (VAT별도), 적정성 평가 포함 (100% 소진 아님)
- 기간: 2026.05.01 ~ 07.31 (3개월)
- 평가: 프로그램+견적 45% / 제작 역량 30% / 유사 경험 15% / 인력 10%

레퍼런스 흡수:
- VAETKI Commerce 쇼케이스 (수주작, ID 1) — 파스텔 에디토리얼
- 메커톤 메이플 월드 (수주작, ID 2) — 다크 에디토리얼
- 마비노기 IP 캠프파이어 모티브 (유저 모임·이야기 상징)

팔레트: fantasy_mystic
- bg: #14102A (딥 네이비-퍼플, 캠프파이어 밤)
- key: #9D7DFF (마법 퍼플)
- sub1: #D4AF66 (골드, 모닥불 빛)
- sub2: #F0C4E0 (소프트 핑크, 벚꽃)

디자인 표준 준수:
- 헤더 3요소 16/14/22, 본문 14/12/10
- 3단 카드 중앙정렬, 표 상하좌우 중앙
- 섹션 디바이더 60pt, 캔버스 10"×5.625"
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from src.generators.slide_kit import *
import src.generators.slide_kit as sk


def build():
    # ── 표준 셋업 ──────────────────────────────────────────────────
    set_slide_size(10.0, 5.625, margin_in=0.4)
    apply_from_library("fantasy_mystic")
    prs = new_presentation()
    setup_editorial_deck(prs, bg_color=tok("surface/darker"), prune=True)

    ML = sk.ML_IN
    CW = sk.CW_IN
    SHI = float(sk.SH / 914400)
    SWI = float(sk.SW / 914400)

    # 자주 쓰는 색
    GOLD    = RGBColor(212, 175, 102)
    PURPLE  = RGBColor(157, 125, 255)
    PINK    = RGBColor(240, 196, 224)
    IVORY   = RGBColor(245, 240, 255)
    MUTED   = RGBColor(200, 190, 230)
    DIM     = RGBColor(150, 145, 180)
    CARD_BG = RGBColor(35, 28, 65)

    # ═══════════════════════════════════════════════════════════════
    # [1] 표지
    # ═══════════════════════════════════════════════════════════════
    s = new_slide(prs)
    gradient_headline(s, ML, SHI * 0.30, CW, 1.2,
                       "BIG CAMPFIRE",
                       c1=PURPLE, c2=PINK,
                       sz_pt=60, align="center", font_weight="black")
    T(s, Inches(ML), Inches(SHI * 0.52), Inches(CW), Inches(0.35),
      "마비노기 모바일 2026년 6월 업데이트 쇼케이스",
      sz=14, c=GOLD, b=True,
      al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(SHI * 0.60), Inches(CW), Inches(0.35),
      "프로모션 대행 제안서", sz=22, c=IVORY,
      b=True, al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(SHI * 0.82), Inches(CW * 0.4), Inches(0.3),
      "2026. 4. 26", sz=12, c=MUTED,
      al=PP_ALIGN.LEFT, fn=FONT_W["regular"])
    T(s, Inches(ML + CW * 0.5), Inches(SHI * 0.82),
      Inches(CW * 0.5), Inches(0.3),
      "LAON MARKETING COMPANY", sz=12,
      c=GOLD, b=True, al=PP_ALIGN.RIGHT, fn=FONT_W["bold"])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 0 — HOOK
    # ═══════════════════════════════════════════════════════════════

    # [2] Hook
    s = new_slide(prs)
    T(s, Inches(ML), Inches(SHI * 0.12), Inches(CW), Inches(0.3),
      "WHY BIG CAMPFIRE", sz=12,
      c=GOLD, b=True, al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(SHI * 0.20), Inches(CW), Inches(0.4),
      "마비노기에서 '캠프파이어'는 모임·휴식·이야기의 상징입니다",
      sz=14, c=MUTED,
      al=PP_ALIGN.CENTER, fn=FONT_W["regular"])
    gradient_headline(s, ML, SHI * 0.30, CW, 0.8,
                       "6월의 쇼케이스는 '함께 둘러앉는 밤'이 됩니다",
                       sz_pt=22, align="center", font_weight="bold")
    STAT_ROW_HERO(s, y_in=SHI * 0.54, h_in=SHI * 0.38, on_dark=True,
        show_dividers=True,
        items=[
            {"value": "Live",  "label": "공식 유튜브 송출",
             "desc": "마비노기 모바일 공식 채널"},
            {"value": "60분+", "label": "쇼케이스 구성",
             "desc": "개막 + 발표 + Q&A + 피날레"},
            {"value": "액티브", "label": "타겟 유저",
             "desc": "6월 메이저 업데이트 기대 유저"},
        ])

    # [3] 제안의 3가지 핵심
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="OUR PROPOSAL",
        pre="온라인 라이브 쇼케이스를 '경험'으로 만드는 3가지 축",
        headline="세트 몰입 · 스토리텔링 · 시청자 참여",
        gradient_headline_text=True)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "IMMERSIVE SET",
             "title": "캠프파이어\n몰입형 스튜디오",
             "body": "따뜻한 불꽃 연출 · 물리 세트 + LED 하이브리드"},
            {"label": "NARRATIVE SHOW",
             "title": "드라마처럼 펼쳐지는\n4막 구성 방송",
             "body": "오프닝 시네마틱 → 대발표 → Q&A → 피날레"},
            {"label": "LIVE ENGAGE",
             "title": "시청자가 주인공인\n실시간 참여 방송",
             "body": "채팅 반영 · 투표 · 시청자 질문 무대 투입"},
        ])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 2 — INSIGHT
    # ═══════════════════════════════════════════════════════════════

    # [4] INSIGHT 디바이더
    slide_divider_hero(prs,
        eng_title="INSIGHT",
        kr_subtitle="마비노기 모바일 현재 좌표",
        tagline="1주년 이후, 팬덤은 '다음 챕터'를 기다린다", pg=4)

    # [5] 게임 현황
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="GAME STATUS",
        pre="넥슨 데브캣이 풀어낸 판타지 라이프 MMORPG",
        headline="1주년 업데이트 직후, 기대감이 정점에 도달한 시점",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        items=[
            {"label": "LAUNCH",
             "title": "2025.02\n정식 출시",
             "body": "데브캣 오리지널 캐릭터 집결 MMORPG"},
            {"label": "1ST ANNIVERSARY",
             "title": "2026.03.27\n1주년 대규모 업데이트",
             "body": "여신강림 4장 · 신규 심층던전 · 룬 평가 시스템"},
            {"label": "NEXT SHOWCASE",
             "title": "2026.06\n빅 캠프파이어",
             "body": "신규 시즌 · 핵심 업데이트 · 운영 방향 공개"},
        ])
    T(s, Inches(ML), Inches(SHI - 0.28), Inches(CW), Inches(0.18),
      "출처: 넥슨 공식 공지 (mabinogimobile.nexon.com), 게임메카 2026.03.27",
      sz=10, c=DIM, al=PP_ALIGN.CENTER, raw_sz=False)

    # [6] 라이브 쇼케이스 트렌드
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="LIVE SHOWCASE TREND",
        pre="게임 라이브 쇼케이스는 '발표회'가 아닌 '팬덤 행사'로 진화 중",
        headline="스튜디오 몰입감 + 실시간 참여가 승부를 결정한다",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "SET DESIGN",
             "title": "물리 세트 × LED\n하이브리드",
             "body": "IP 세계관 시각화 · 카메라 앵글 다양성"},
            {"label": "STORYTELLING",
             "title": "극장형\n방송 구성",
             "body": "Act 1-4 드라마틱 전개 · 중간 시네마틱"},
            {"label": "INTERACTION",
             "title": "실시간 채팅\n반영 구조",
             "body": "시청자 질문 → 무대 · 투표 · 이스터에그"},
        ])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 3 — CONCEPT
    # ═══════════════════════════════════════════════════════════════

    # [7] CONCEPT 디바이더
    slide_divider_hero(prs,
        eng_title="THE CONCEPT",
        kr_subtitle="BIG CAMPFIRE",
        tagline="모닥불 옆에 둘러앉아 나누는 밤의 이야기", pg=7)

    # [8] Big Idea
    s = new_slide(prs)
    T(s, Inches(ML), Inches(SHI * 0.10), Inches(CW), Inches(0.3),
      "THE BIG IDEA", sz=12,
      c=GOLD, b=True, al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    gradient_headline(s, ML, SHI * 0.22, CW, 1.3,
                       '"둘러앉는 시간"',
                       sz_pt=60, align="center", font_weight="black")
    T(s, Inches(ML), Inches(SHI * 0.55), Inches(CW), Inches(0.35),
      "COME SIT BY THE FIRE.", sz=14,
      c=GOLD, b=True, al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(SHI * 0.65), Inches(CW), Inches(0.35),
      "유저와 개발진이 같은 캠프파이어를 둘러싸고 나누는 60분",
      sz=22, c=IVORY, b=True,
      al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(SHI * 0.80), Inches(CW), Inches(0.45),
      "마비노기에서 캠프파이어는 유저들이 모이는 장소입니다.\n"
      "6월의 쇼케이스도 그 장면 그대로 재현합니다.",
      sz=14, c=MUTED, al=PP_ALIGN.CENTER, fn=FONT_W["regular"])

    # [9] 3 Win Themes
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="3 WIN THEMES",
        pre="캠프파이어 컨셉을 지탱하는 3개의 기둥",
        headline="세트 · 구성 · 참여 — 3축 설계",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "① WARM SET",
             "title": "따뜻한 불빛의\n캠프파이어 스튜디오",
             "body": "물리 세트 + LED · 카메라 다중 앵글 구성"},
            {"label": "② STORY ARC",
             "title": "시네마틱 오프닝부터\n피날레까지의 4막",
             "body": "웰컴 → 회고 → ★대발표 → Q&A → 마무리"},
            {"label": "③ LIVE BRIDGE",
             "title": "실시간 채팅을\n무대 위로 끌어올림",
             "body": "유저 질문 · 투표 · 반응을 방송에 즉시 반영"},
        ])

    # [10] 시청자 경험 Journey
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="VIEWER JOURNEY",
        pre="시청자가 스트리밍 1시간 동안 경험하는 감정선",
        headline="알림 → 접속 → 몰입 → 반응 → 공유",
        gradient_headline_text=False)
    circle_d = min(1.0, (CW / 5) * 0.7)
    CIRCULAR_PHOTO_FLOW(s, y_in=y_end + 0.5, circle_d=circle_d,
        arrow_color=GOLD,
        items=[
            {"stage": "01", "title": "예고",   "time": "티저·카운트다운"},
            {"stage": "02", "title": "접속",   "time": "라이브 시작 전 웰컴"},
            {"stage": "03", "title": "몰입",   "time": "4막 방송 본편"},
            {"stage": "04", "title": "참여",   "time": "채팅·Q&A·투표"},
            {"stage": "05", "title": "공유",   "time": "클립·숏폼 확산"},
        ])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 4 — ACTION PLAN
    # ═══════════════════════════════════════════════════════════════

    # [11] ACTION PLAN 디바이더
    slide_divider_hero(prs,
        eng_title="ACTION PLAN",
        kr_subtitle="세트 · 구성 · 제작물 상세 실행",
        tagline="기획부터 후속 운영까지, 3개월을 설계합니다", pg=11)

    # [12] 3-Month Roadmap
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="3-MONTH ROADMAP",
        pre="계약 체결부터 결과 리포트까지의 전체 타임라인",
        headline="5월 준비 · 6월 실행 · 7월 확산",
        gradient_headline_text=False)
    rows = [
        ["Week 1",    "5/07~5/13",  "킥오프 & 컨셉",     "방향성 확정 · 스튜디오 로케이션 선정"],
        ["Week 2-3",  "5/14~5/27",  "기획 & 대본",       "방송 시나리오 · 게스트 섭외"],
        ["Week 4-5",  "5/28~6/10",  "제작 & 세트",       "OAP/타이틀/루핑 · 세트 제작"],
        ["Week 6",    "6/11~6/14",  "리허설",            "기술 리허설 2회 · 최종 점검"],
        ["Week 7",    "6/15 중",    "★ 쇼케이스 실행",    "라이브 방송 송출"],
        ["Week 8-9",  "6/16~6/30",  "후속 콘텐츠",       "숏폼 · 클립 · 요약 영상 배포"],
        ["Week 10-13","7/01~7/31",  "결과 리포트",       "KPI 분석 · 정산 · 인사이트"],
    ]
    avail_h = SHI - y_end - 0.4
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["기간", "일자", "단계", "주요 Task"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140), highlight_col=0)

    # [13] 방송 당일 Timeline
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="SHOWCASE TIMELINE",
        pre="쇼케이스 당일 약 60분, 분 단위 방송 구성",
        headline="오프닝부터 피날레까지 매끄러운 흐름",
        gradient_headline_text=False)
    rows = [
        ["-30분",      "방송 준비·웰컴 스탠바이",   "BGM + 카운트다운 루핑"],
        ["00:00~03:00", "오프닝 시네마틱",            "타이틀 + OAP 재생"],
        ["03:00~10:00", "Act 1 — Welcome",           "호스트 인사 · 쇼케이스 안내"],
        ["10:00~18:00", "Act 2 — Story So Far",       "1주년 회고 · 커뮤니티 하이라이트"],
        ["18:00~38:00", "Act 3 — ★ Big Reveal",       "★ 6월 업데이트 핵심 발표"],
        ["38:00~50:00", "Act 4 — Around the Fire",    "개발진 Q&A · 시청자 채팅 반영"],
        ["50:00~58:00", "Special Program",            "게스트 라이브 또는 시연"],
        ["58:00~60:00", "피날레 & 예고",              "커밍순 티저 + 클로징"],
    ]
    avail_h = SHI - y_end - 0.4
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["시간", "프로그램", "구성"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140), highlight_col=0)

    # [14] 방송 구성 4막
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="SHOW STRUCTURE",
        pre="'발표 나열'이 아닌 '이야기 전개'로 설계",
        headline="4막 극장형 구성 — 몰입과 환기의 리듬",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        items=[
            {"label": "ACT 1",
             "title": "Welcome\n캠프파이어 점화",
             "body": "시네마틱 오프닝 + 호스트 웰컴 · 분위기 형성"},
            {"label": "ACT 2-3",
             "title": "Story + Reveal\n회고 → 대발표",
             "body": "1주년 회고 → ★ 6월 업데이트 핵심 공개"},
            {"label": "ACT 4",
             "title": "Around the Fire\n개발진·유저 대화",
             "body": "개발진 Q&A + 시청자 채팅 실시간 반영"},
        ])

    # [15] 스튜디오 세트 디자인
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="SET DESIGN",
        pre="모닥불·나무·별빛 — 마비노기 캠프파이어 감성 그대로",
        headline="물리 세트 + LED 하이브리드 스튜디오",
        gradient_headline_text=True)
    # 3 Zone
    pw = 2.8
    ph = 1.3
    gap = 0.3
    total = pw * 3 + gap * 2
    sx = (CW - total) / 2 + ML
    py = y_end + 0.25
    zones = [
        ("MAIN STAGE",     "호스트 메인",      PURPLE),
        ("CAMPFIRE PIT",   "캠프파이어 중앙",   GOLD),
        ("INTERACT LED",   "실시간 채팅",       PINK),
    ]
    for i, (eng, kr, color) in enumerate(zones):
        x = sx + i * (pw + gap)
        PARALLELOGRAM_ZONE(s, x, py, pw, ph, eng,
                            color=color, text_color=IVORY, sz_pt=16)
        T(s, Inches(x), Inches(py + ph + 0.15),
          Inches(pw), Inches(0.3),
          kr, sz=22, c=IVORY, b=True,
          al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML), Inches(py + ph + 0.65),
      Inches(CW), Inches(0.4),
      "호스트 메인 공간 · 모닥불 중앙 프로젝션 · 시청자 채팅 LED 3면 구성",
      sz=14, c=MUTED, al=PP_ALIGN.CENTER, fn=FONT_W["regular"])

    # [16] 세트 상세 (RENDER_CAPTION)
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="STUDIO RENDER",
        pre="캠프파이어 스튜디오 컨셉 비주얼",
        headline="따뜻한 불빛·소재·공간감의 밤 설계",
        gradient_headline_text=False)
    RENDER_CAPTION(s,
        title="빅 캠프파이어 스튜디오 예상 구성",
        caption="프리셋 키비주얼은 최종 프로덕션 단계에서 확정",
        image_area=(ML, y_end + 0.2, CW, SHI - y_end - 1.0),
        on_dark=True)

    # [17] 제작물 전체
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="PRODUCTION LIST",
        pre="RFP 요구사항 + 쇼케이스 완성도에 필요한 제작물 전반",
        headline="영상 4종 · 세트 · 방송 그래픽 — 원스톱 체계",
        gradient_headline_text=False)
    rows = [
        ["OAP",           "On-Air Promotion",   "3~5분 티저 · 본방 예고 · 애프터 영상"],
        ["타이틀",         "쇼케이스 타이틀",     "메인 로고 모션 · 챕터별 전환 타이틀"],
        ["루핑 영상",      "대기 시간 루핑",      "카운트다운 + IP 배경 순환 영상"],
        ["방송 그래픽",    "실시간 송출 그래픽",  "G/B · Lower Third · 이벤트 애니메이션"],
        ["공간·세트",      "스튜디오 물리 세트",   "메인 + 캠프파이어 + LED 월"],
        ["사후 콘텐츠",    "숏폼 / 클립",         "하이라이트 5종 · 요약 풀영상"],
    ]
    avail_h = SHI - y_end - 0.4
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["분류", "항목", "세부 구성"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140))

    # [18] OAP / 타이틀 / 루핑 영상
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="MOTION ASSETS",
        pre="라이브 방송의 '첫인상·전환·여운'을 만드는 영상 자산",
        headline="OAP · 타이틀 · 루핑 영상 제작 계획",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "OAP",
             "title": "On-Air Promotion\n3종 × 60-90초",
             "body": "티저 · 카운트다운 · 피날레 · 감성적 톤"},
            {"label": "TITLE",
             "title": "쇼케이스 타이틀\n모션 그래픽",
             "body": "메인 로고 + 챕터 전환 4종 · 시네마틱 톤"},
            {"label": "LOOP",
             "title": "루핑 영상\n대기 시간용",
             "body": "IP 배경 + 카운트다운 · 15분 반복"},
        ])

    # [19] 방송 그래픽
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="BROADCAST GRAPHICS",
        pre="라이브 송출 중 실시간으로 움직이는 그래픽 시스템",
        headline="G/B · Lower Third · 애니메이션 — 통일된 VI",
        gradient_headline_text=False)
    rows = [
        ["G/B 월",          "메인 배경 LED 그래픽",    "캠프파이어 + 별빛 배경 루프"],
        ["Lower Third",     "호스트·게스트 이름표",    "등장 애니메이션 + 직함 표기"],
        ["Callout",         "강조 팝업",               "업데이트 명 · 수치 강조 10종"],
        ["Transition",      "챕터 전환",               "Act 1→2→3→4 전환 4종"],
        ["Overlay",         "실시간 채팅 프레임",       "시청자 댓글 하단 노출 UI"],
        ["Endcard",         "클로징 그래픽",           "커밍순 + 공식 채널 안내"],
    ]
    avail_h = SHI - y_end - 0.4
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["구분", "내용", "비고"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140))

    # [20] 부가 프로그램
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="SPECIAL PROGRAM",
        pre="'업데이트 브리핑'에서 '팬덤 행사'로 격상하는 핵심",
        headline="공연 · 이벤트 · 시청자 참여 프로그램 검토",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "LIVE MUSIC",
             "title": "OST 아티스트\n어쿠스틱 공연",
             "body": "마비노기 OST 기반 3-4곡 · 감성 마무리"},
            {"label": "SPECIAL GUEST",
             "title": "모델·인플루언서\n라이브 시연",
             "body": "RPG 전문 스트리머 · 신규 콘텐츠 플레이"},
            {"label": "VIEWER PARTY",
             "title": "시청자 이벤트\n참여형 코너",
             "body": "실시간 투표 · 퀴즈 · 한정 코드 증정"},
        ])

    # [21] 시청자 인터랙션
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="VIEWER INTERACTION",
        pre="유튜브 라이브 채팅을 방송 연출의 일부로",
        headline="실시간 채팅을 무대 위로 끌어올리는 3가지 장치",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        items=[
            {"label": "CHAT PICK",
             "title": "실시간 채팅\n하이라이트 반영",
             "body": "시청자 댓글 5-10건 무대 스크린 투사"},
            {"label": "LIVE POLL",
             "title": "즉석 투표\n콘텐츠 선택",
             "body": "차기 업데이트 방향 2-3개 즉석 투표"},
            {"label": "Q&A BRIDGE",
             "title": "Q&A\n채팅 → 개발진",
             "body": "사전 질문 + 실시간 선별 질문 혼합 응답"},
        ])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 5 — MANAGEMENT
    # ═══════════════════════════════════════════════════════════════

    # [22] MANAGEMENT 디바이더
    slide_divider_hero(prs,
        eng_title="MANAGEMENT",
        kr_subtitle="운영 관리 체계",
        tagline="라이브 방송의 '0 지연·0 사고'를 위한 조직", pg=22)

    # [23] Project Organization
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="PROJECT ORGANIZATION",
        pre="총 13인 전담 TF · 4개 본부 구조",
        headline="기획 · 제작 · 방송 · 운영 4분할",
        gradient_headline_text=False)
    col_w = CW / 4
    col_gap = 0.15
    col_data_w = col_w - col_gap
    badge_w = col_data_w * 0.42
    name_w = col_data_w * 0.55
    row_h = 0.35
    row_gap = 0.08
    content_y = y_end + 0.4
    groups = [
        ("기획",    [("총괄", "최승진"), ("PM", "탁찬연"),
                    ("플래너", "이신지")]),
        ("제작",    [("디자인", "구유빈"), ("영상 4종", "외부 3"),
                    ("세트", "외주 협력")]),
        ("방송",    [("방송 PD", "공만진"), ("스위처", "1명"),
                    ("그래픽 OP", "1명")]),
        ("운영",    [("현장 운영", "박은원"), ("MC 섭외", "1명"),
                    ("인터랙션", "1명")]),
    ]
    for gi, (title, rows) in enumerate(groups):
        cx = ML + gi * col_w
        T(s, Inches(cx), Inches(content_y - 0.25),
          Inches(col_data_w), Inches(0.22),
          title, sz=10, c=MUTED,
          al=PP_ALIGN.CENTER, fn=FONT_W["medium"])
        for ri, (role, name) in enumerate(rows):
            row_y = content_y + ri * (row_h + row_gap)
            BADGE(s, cx, row_y, badge_w, row_h, role,
                  fill=RGBColor(80, 60, 140), sz_pt=12)
            T(s, Inches(cx + badge_w + 0.05), Inches(row_y),
              Inches(name_w - 0.05), Inches(row_h),
              name, sz=14, c=IVORY,
              al=PP_ALIGN.LEFT, fn=FONT_W["regular"])

    # [24] Risk Plan
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="RISK PLAN",
        pre="라이브 방송 고유 리스크에 대한 4단 대응 체계",
        headline="기술·콘텐츠·커뮤니케이션 3영역 매뉴얼",
        gradient_headline_text=False)
    rows = [
        ["기술",   "스트리밍 끊김",            "2중 인코더 · 회선 이중화 · 엔지니어 상주"],
        ["기술",   "음향·영상 싱크 오류",       "리허설 2회 · 실시간 모니터링"],
        ["콘텐츠", "대본 지연·시간 초과",       "진행 큐시트 · 3단 지연 대응 시나리오"],
        ["콘텐츠", "스포일러 노출",             "NDA + 화면 워터마크 + 편집자 승인"],
        ["소통",   "부정적 채팅 폭주",           "실시간 모더레이션 + 긴정 채팅 모드"],
    ]
    avail_h = SHI - y_end - 0.4
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["영역", "리스크", "대응 방안"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140), highlight_col=0)

    # [25] Result Report
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="RESULT REPORT",
        pre="7월 말까지 제출하는 결과 리포트",
        headline="데이터 · 콘텐츠 · 인사이트 3파트",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        items=[
            {"label": "DATA",
             "title": "정량 데이터\n대시보드",
             "body": "시청·체류·채팅 수 · 피크 동접 · 지역 분포"},
            {"label": "CONTENT",
             "title": "사후 활용\n자산 아카이브",
             "body": "영상·숏폼·그래픽 · 후속 PR 소재 정리"},
            {"label": "INSIGHT",
             "title": "다음 쇼케이스\n개선 제안",
             "body": "관객 반응 분석 + 차기 기획 인사이트 정리"},
        ])

    # ═══════════════════════════════════════════════════════════════
    # PHASE 6 — WHY US
    # ═══════════════════════════════════════════════════════════════

    # [26] WHY US 디바이더
    slide_divider_hero(prs,
        eng_title="WHY LAON",
        kr_subtitle="라이브 방송 · 게임 IP 전문성",
        tagline="이미 해왔고, 더 잘할 수 있는 이유", pg=26)

    # [27] Credentials
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="CREDENTIALS",
        pre="넥슨 IP · 대형 쇼케이스 · 라이브 방송 실적",
        headline="2024-2026, 라온의 연속 수주 및 수행 기록",
        gradient_headline_text=False)
    CREDENTIAL_STAGE(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        stage_colors=[PURPLE, GOLD, PINK],
        items=[
            {"stage": "2024~",   "title": "NYPC 넥슨 청소년 코딩",
             "body": "넥슨 사옥 대회 운영 · 10년 파트너십"},
            {"stage": "2025.03", "title": "DATADOG SUMMIT SEOUL",
             "body": "대형 행사 라이브 방송 · 그래픽 운영"},
            {"stage": "2026",    "title": "VARCO · 메커톤 수주",
             "body": "NC AI 쇼케이스 + 메이플 메커톤 2건 수주"},
        ])

    # [28] Core Capability
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="CORE CAPABILITY",
        pre="이 쇼케이스에 적합한 3가지 차별 역량",
        headline="게임 IP · 라이브 방송 · 콘텐츠 확산",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "① IP UNDERSTANDING",
             "title": "판타지 게임 IP\n감성 이해",
             "body": "IP 세계관 해석 · 팬덤 감정선 케어 경험"},
            {"label": "② LIVE BROADCAST",
             "title": "4K 멀티캠\n라이브 연출",
             "body": "전담 방송 PD · 스위처 · 그래픽 OP 팀"},
            {"label": "③ CONTENT AMP",
             "title": "후속 콘텐츠\n자산화 운영",
             "body": "숏폼·클립·요약본 동시 생성 · 2주 운영"},
        ])

    # [29] Portfolio 간략
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="KEY PORTFOLIO",
        pre="라이브 방송·쇼케이스 실제 제작물 예시",
        headline="OAP · 세트 · 그래픽 포트폴리오 링크",
        gradient_headline_text=False)
    T(s, Inches(ML), Inches(y_end + 0.3), Inches(CW), Inches(0.4),
      "상세 포트폴리오는 별도 웹하드/링크로 제공", sz=14, c=MUTED,
      al=PP_ALIGN.CENTER, fn=FONT_W["regular"])
    rows = [
        ["DATADOG SUMMIT 2025",     "라이브 방송 · 그래픽",     "4K 멀티캠 · G/B 제작"],
        ["NYPC 2024 현장",          "오프라인 + 온라인 중계",   "시네마틱 영상 + 현장 송출"],
        ["VARCO Commerce 2026",     "쇼케이스 운영",           "인플루언서 150명 수주"],
        ["메이플 메커톤 2026",       "게임 이벤트 운영",        "제안서 수주 (2026.04)"],
    ]
    avail_h = SHI - y_end - 1.0
    row_h = avail_h / (len(rows) + 1)
    DATA_TABLE_DARK(s,
        headers=["프로젝트", "유형", "핵심 제작"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.8, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140), highlight_col=0)

    # ═══════════════════════════════════════════════════════════════
    # PHASE 7 — BUDGET
    # ═══════════════════════════════════════════════════════════════

    # [30] BUDGET 디바이더
    slide_divider_hero(prs,
        eng_title="BUDGET",
        kr_subtitle="5억 예산 가견적서",
        tagline="적정성 기반 예산 구성 — 100% 소진 아닌 최적", pg=30)

    # [31] Budget Breakdown
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="BUDGET BREAKDOWN",
        pre="가견적 총 4억 8,000만원 (VAT별도) · 예산 2천만원 여유",
        headline="제작 46% · 방송 17% · 운영 10% · 부가 10%",
        gradient_headline_text=False)
    rows = [
        ["공간·세트 제작·구축",  "120,000,000", "25%", "메인+캠프파이어+LED 월"],
        ["영상 제작",            "100,000,000", "21%", "OAP·타이틀·루핑·후속"],
        ["방송 기술·송출",       "80,000,000",  "17%", "4K 멀티캠·엔지니어링"],
        ["운영 인력",            "50,000,000",  "10%", "13인 TF·리허설 포함"],
        ["부가 프로그램",        "50,000,000",  "10%", "게스트·공연·이벤트"],
        ["방송 그래픽",          "30,000,000",  "6%",  "G/B·Lower Third·전환"],
        ["콘텐츠·운영·예비",      "50,000,000",  "11%",  "채팅 모더레이션·돌발 대응"],
    ]
    # 합계 박스 포함 높이 계산 — 여유 충분히
    avail_h = SHI - y_end - 0.35
    row_h = avail_h / (len(rows) + 3)    # header + body + 합계 + 버퍼
    DATA_TABLE_DARK(s,
        headers=["항목", "금액 (원)", "비중", "세부"],
        rows=rows,
        x_in=ML, y_in=y_end + 0.15, w_in=CW, row_h_in=row_h,
        header_color=RGBColor(80, 60, 140), highlight_col=0)
    # 합계 박스 — 테이블 하단에 여유 있게 배치
    total_y = y_end + 0.15 + row_h * (len(rows) + 1) + 0.18
    total_h = row_h * 0.95
    R(s, Inches(ML), Inches(total_y), Inches(CW), Inches(total_h),
      f=RGBColor(10, 8, 25), lc=GOLD)
    T(s, Inches(ML + 0.25), Inches(total_y),
      Inches(CW * 0.2), Inches(total_h),
      "소계 (VAT 별도)", sz=14, c=GOLD, b=True,
      al=PP_ALIGN.LEFT, fn=FONT_W["bold"])
    T(s, Inches(ML + CW * 0.30), Inches(total_y),
      Inches(CW * 0.35), Inches(total_h),
      "480,000,000 원", sz=22, c=IVORY, b=True,
      al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    T(s, Inches(ML + CW * 0.65), Inches(total_y),
      Inches(CW * 0.33), Inches(total_h),
      "예산 5억 내 · 2천만원 여유 확보", sz=12,
      c=MUTED, al=PP_ALIGN.RIGHT, fn=FONT_W["regular"])

    # [32] 예산 구성 원칙
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="BUDGET PRINCIPLES",
        pre="RFP '예산 적정성' 평가 반영 — 100% 소진이 목표가 아닙니다",
        headline="적정·투명·선택과 집중 3원칙",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        label_colors=[PURPLE, GOLD, PINK],
        items=[
            {"label": "① FIT",
             "title": "적정\n시장 평균 기반",
             "body": "항목별 시장가 기준 · 부풀리기 없음"},
            {"label": "② TRANSPARENT",
             "title": "투명\n산출 근거 명시",
             "body": "항목별 단가 · 수량 · 총액 명확 기재"},
            {"label": "③ FOCUS",
             "title": "선택과 집중\n핵심에 배분",
             "body": "제작·방송에 62% 집중 · 여유 2천만원 확보"},
        ])

    # [33] 기대 효과
    s = new_slide(prs)
    y_end = PAGE_HEADER_LIGHT(s,
        page_title="EXPECTED OUTCOMES",
        pre="쇼케이스가 남겨야 할 정성·정량 가치",
        headline="시청·참여·자산 3단계 효과",
        gradient_headline_text=False)
    PHOTO_CARD_TRIO(s, y_in=y_end + 0.2, h_in=SHI - y_end - 0.4,
        on_dark=True,
        items=[
            {"label": "VIEWING",
             "title": "라이브 동시 시청\n높은 체류 시간",
             "body": "평균 체류 40분+ 목표 · 피크 동접 극대화"},
            {"label": "ENGAGE",
             "title": "실시간 채팅\n활성도",
             "body": "채팅 CPM · 시청자 질문 반영률 지표화"},
            {"label": "ASSET",
             "title": "2주간 후속\n프로모션 연료",
             "body": "숏폼 5종+ · 요약 영상 · PR 소재 팩키지"},
        ])
    T(s, Inches(ML), Inches(SHI - 0.3), Inches(CW), Inches(0.18),
      "정확한 KPI는 클라이언트 합의 후 확정. 구체적 수치는 NDA 협의 하 별도 제안.",
      sz=10, c=DIM, al=PP_ALIGN.CENTER, fn=FONT_W["regular"])

    # ═══════════════════════════════════════════════════════════════
    # CLOSING
    # ═══════════════════════════════════════════════════════════════

    # [34] Next Step
    s = new_slide(prs)
    T(s, Inches(ML), Inches(SHI * 0.10), Inches(CW), Inches(0.3),
      "NEXT STEP", sz=12,
      c=GOLD, b=True, al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
    gradient_headline(s, ML, SHI * 0.22, CW, 0.9,
                       "다음은 함께 둘러앉을 준비입니다",
                       sz_pt=22, align="center", font_weight="bold")
    steps = [
        ("STEP 1", "제안 심사",    "4/26~5/01 결과 발표"),
        ("STEP 2", "계약 체결",    "5/07 공식 착수"),
        ("STEP 3", "★ 캠프파이어", "6월 중순 방송"),
    ]
    sw = CW / 3 - 0.2
    sy = SHI * 0.48
    for i, (step, title, date) in enumerate(steps):
        x = ML + i * (sw + 0.3)
        R(s, Inches(x), Inches(sy), Inches(sw), Inches(1.5),
          f=CARD_BG, lc=PURPLE)
        T(s, Inches(x), Inches(sy + 0.25), Inches(sw), Inches(0.3),
          step, sz=12, c=GOLD, b=True,
          al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
        T(s, Inches(x), Inches(sy + 0.65), Inches(sw), Inches(0.4),
          title, sz=22, c=IVORY, b=True,
          al=PP_ALIGN.CENTER, fn=FONT_W["bold"])
        T(s, Inches(x), Inches(sy + 1.10), Inches(sw), Inches(0.3),
          date, sz=12, c=MUTED,
          al=PP_ALIGN.CENTER, fn=FONT_W["regular"])

    # [35] E.O.D
    s = new_slide(prs)
    gradient_headline(s, ML, SHI * 0.35, CW, 1.0,
                       "E.O.D",
                       sz_pt=72, align="center", font_weight="black")
    T(s, Inches(ML), Inches(SHI * 0.65), Inches(CW), Inches(0.3),
      "ALL RIGHTS RESERVED BY LAON MARKETING COMPANY.",
      sz=10, c=MUTED, al=PP_ALIGN.CENTER, fn=FONT_W["regular"])
    T(s, Inches(ML), Inches(SHI * 0.71), Inches(CW), Inches(0.3),
      "이 제안서는 넥슨과 라온 간 협의용이며, 무단 배포를 금합니다.",
      sz=10, c=MUTED, al=PP_ALIGN.CENTER, fn=FONT_W["regular"])

    # ── 저장 + 검증 ─────────────────────────────────────────────────
    out = Path(__file__).parent / "mabinogi_big_campfire_proposal_v2.pptx"
    save_pptx(prs, str(out))
    print(f"[OK] BIG CAMPFIRE 제안서 생성: {out}")
    print(f"     슬라이드: {len(prs.slides)}")
    print(f"     팔레트: Fantasy Mystic")

    # 자동 검증
    print(f"\n=== 자동 레이아웃 검증 ===")
    issues = validate_deck(prs, verbose=True, check_overlaps=True)
    if issues:
        print(f"→ 자동 수정 시도...")
        fixed, remaining = auto_fix_overflow(prs)
        print(f"→ {fixed}개 폰트 축소됨, 남은 이슈: {len(remaining)}건")
        if fixed > 0:
            save_pptx(prs, str(out))


if __name__ == "__main__":
    build()
