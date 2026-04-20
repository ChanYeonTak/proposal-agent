---
name: slide-kit-design-standard
description: |
  slide_kit v4.1+ 제안서 생성 시 적용하는 **공식 디자인 표준**. 폰트 크기 체계,
  정렬 규칙, 색상 팔레트, 컴포넌트 기본값, 검증 파이프라인을 정의한다.
  모든 신규 제안서는 이 표준을 준수한다. 예외는 사용자 명시 요청 시에만.
---

# slide_kit 디자인 표준 (2026-04-17 확정)

**이 문서가 최종 표준**이다. 신규 제안서는 반드시 이 규칙으로 작성.
코드는 `src/generators/slide_kit.py`에 구현되어 있고 자동 강제됨.

## 🔤 폰트 크기 체계

### 허용 스케일 (FONT_SCALE) — 14종 고정
```
10 / 12 / 14 / 16 / 18 / 22 / 28 / 32 / 40 / 48 / 54 / 60 / 72 / 96
```

**자동 스냅**: `T()` / `Pt()` / `.font.size` 어디든 비표준 값 입력 시 가장 가까운 허용값으로 자동 변환. 개발 중 `sz=27` 작성해도 런타임에 28로 스냅됨.

### 시맨틱 키 → 크기 매핑

#### 헤더 3요소 (PAGE_HEADER_LIGHT 기본값)
| 역할 | SZ 키 | pt | 정렬 |
|------|------|----|------|
| 좌상단 페이지 제목 | `page_label` | **16** | Left, Bold |
| 중앙 상부 소제목 | `pre_title` | **14** | Center, Regular |
| 중앙 상부 대제목 | `main_title` | **22** | Center, Bold |

#### 본문 3요소 (콘텐츠 텍스트)
| 역할 | SZ 키 | pt |
|------|------|----|
| 한 줄 강조 본문 | `body_emphasis` | **14** |
| 중요 텍스트 | `body_bold` | **12** (Bold) |
| 일반 본문 | `body` | **10** |

#### 카드/섹션 레벨
| 역할 | SZ 키 | pt |
|------|------|----|
| 카드 내 타이틀 | `sub_headline` | 22 |
| 섹션 마커 (대문자) | `eyebrow` | 12 Bold |
| 섹션 디바이더 영문 | `section_hero` | **60** (이전 96 → 축소) |
| E.O.D 최종 | `eod` | **72** (이전 96 → 축소) |

#### 거대 수치
| 역할 | SZ 키 | pt |
|------|------|----|
| 3-열 stat row | `stat_big` | 72 |
| 단일 거대 수치 | `stat_hero_v41` | 96 |

#### Legacy 별칭 (자동 매핑)
```
body_sm = body_reading = caption = caption_sm = micro → 해당 tier로 자동
label = body_bold, fine = body_emphasis
pre_headline = pre_title, headline = main_title
```

## 🎨 정렬 규칙

### 3단 포토 카드 (`PHOTO_CARD_TRIO`)
- **모든 텍스트 중앙 정렬** (라벨/타이틀/본문)
- 하단 텍스트 블록은 각 이미지 컬럼의 center axis를 기준

### 표 (`DATA_TABLE_DARK`)
- **모든 셀 상하좌우 중앙 정렬**
- `vertical_anchor = MSO_ANCHOR.MIDDLE` 강제
- 마진: L/R 0.1", T/B 0.04" (타이트)

### 컬러 배경 위 텍스트 (`BADGE`)
- 도형 자체의 text_frame 사용
- `vertical_anchor = MIDDLE` + `alignment = CENTER`
- R() + T() 조합 금지 (중앙정렬 어긋남)

### PAGE_HEADER 계열
- page_title: **좌상단 좌정렬**
- pre + headline: **중앙 정렬**

## 🌈 팔레트 체계

### 5요소 시맨틱
모든 팔레트는 5개 슬롯으로 구성:
```python
{
    "bg":   "#XXX",   # 배경
    "text": "#XXX",   # 글자
    "key":  "#XXX",   # 키컬러 (악센트, 뱃지)
    "sub1": "#XXX",   # 서브1 (라벨 두번째)
    "sub2": "#XXX",   # 서브2 (eyebrow/태그)
}
```

### 라이브러리 (21종)
- **Dark**: `editorial_dark`, `cyberpunk_neon`, `midnight_forest`, `deep_luxury`, `finance_navy`, `fantasy_mystic`
- **Light**: `minimal_light`, `paper_warm`, `nordic_cool`, `healthcare_mint`, `corporate_blue`, `vaetki_pastel` (수주작)
- **Vibrant**: `tech_gradient`, `sunset_coral`, `youth_pop`, `food_warm`
- **전문**: `heritage_gold`, `industrial_steel`, `mono_elegant`, `fintech_purple`, `event_gala`, `eco_fresh`

### 프로젝트 → 팔레트 추천
```python
recs = recommend_palettes(
    project_type="event" | "marketing_pr" | "it_system" | "public" | "consulting",
    industry="...",
    keywords=[...],
    mood="...",
)
apply_from_library(recs[0]["key"])
```

## 📐 캔버스 표준

**기본 캔버스**: 10" × 5.625" (25.4cm × 14.288cm, Google Slides 기본)
```python
set_slide_size(10.0, 5.625, margin_in=0.4)
# scale_fonts=False (기본) — pt는 캔버스 무관 일관성 유지
```

- 좌우 마진: 0.4" (ML_IN = 0.4, CW_IN = 9.2)
- 세이프 존 상하단 0.25" 여유

## 🏗️ 마스터 배경 & 레이아웃

```python
prs = new_presentation()
setup_editorial_deck(prs, bg_color=tok("surface/darker"), prune=True)
# → 마스터 배경 설정 + Blank 외 레이아웃 모두 삭제
```
- 모든 슬라이드 `new_slide(prs)` 호출 시 마스터 배경 자동 상속
- 개별 슬라이드는 `bg()` / `bg_pastel_gradient()`로 덮어쓰기 가능

## 🧩 핵심 컴포넌트 요약

| 컴포넌트 | 용도 |
|---------|------|
| `PAGE_HEADER_LIGHT(s, page_title=, pre=, headline=, gradient_headline_text=)` | 표준 헤더 (좌상단 제목 + 중앙 소제목/대제목) |
| `PHOTO_CARD_TRIO(s, y_in, h_in, items)` | 3단 이미지+텍스트 카드 (메인 워크호스) |
| `STAT_ROW_HERO(s, y_in, h_in, items)` | 3-4열 거대 수치 (160명/48H 등) |
| `DATA_TABLE_DARK(s, headers, rows)` | 다크 배경 표 (모든 셀 상하좌우 중앙) |
| `BADGE(s, l, t, w, h, text)` | 컬러 배경 텍스트 뱃지 (중앙정렬 보장) |
| `gradient_headline(s, ..., c1, c2)` | 브랜드 그라디언트 텍스트 |
| `bg_pastel_gradient(s, c1, c2, c3)` | 3-stop 파스텔 배경 (VAETKI 스타일) |
| `PARALLELOGRAM_ZONE(s, ..., text)` | 기울어진 공간 구분 라벨 |
| `CIRCULAR_PHOTO_FLOW(s, items)` | 원형 사진 타임라인 |
| `CREDENTIAL_STAGE(s, items)` | STAGE 뱃지 + 실적 카드 |
| `slide_divider_hero(prs, eng_title=, ...)` | 섹션 디바이더 (여백 80%+60pt 영문) |
| `slide_divider_light(prs, eng_title=)` | 라이트 섹션 디바이더 |
| `slide_hook_question(prs, question, stats=)` | HOOK 슬라이드 |
| `slide_cover_editorial(prs, title=, subtitle=)` | 표지 |

## 🔍 자동 검증 파이프라인

모든 제안서 생성 스크립트 마지막에 필수 포함:
```python
save_pptx(prs, out)
issues = validate_deck(prs, check_overlaps=True)
if issues:
    fixed, remaining = auto_fix_overflow(prs)
    if fixed > 0:
        save_pptx(prs, out)  # 수정본 재저장
```

검증 항목:
- 슬라이드 경계 초과 (우/하/좌/상)
- 텍스트 박스 겹침 (90% 포함 관계 오탐 제외)
- 자동 수정: 폰트 축소 시도 (15% 비례)

## 📋 제안서 생성 템플릿 체크리스트

신규 RFP 대응 시:
1. ☐ 캔버스: `set_slide_size(10.0, 5.625, margin_in=0.4)`
2. ☐ 팔레트 추천: `recommend_palettes(...)` → 1순위 적용
3. ☐ 마스터: `setup_editorial_deck(prs)`
4. ☐ 구조: Impact-8 8-Phase (Hook/Summary/Insight/Concept/Action/Management/Why Us/Investment)
5. ☐ 헤더: PAGE_HEADER_LIGHT로 16/14/22 자동 적용
6. ☐ 콘텐츠: PHOTO_CARD_TRIO / STAT_ROW_HERO / DATA_TABLE_DARK 위주
7. ☐ 섹션: slide_divider_hero 60pt 영문
8. ☐ 검증: `validate_deck()` + `auto_fix_overflow()`
9. ☐ 업로드 필요 시: `pptx-google-slides-upload` 스킬 사용

## 🚫 금지 사항

- ❌ 임의 폰트 크기 (자동 스냅되지만 가능하면 시맨틱 키 사용)
- ❌ 좌상단 페이지 제목을 중앙정렬하거나 22pt 이상으로
- ❌ 카드 내 텍스트 좌정렬 (3단 카드는 중앙만)
- ❌ 표 셀 좌정렬 (상하좌우 중앙만)
- ❌ 섹션 디바이더 96pt (60pt로 고정)
- ❌ 수치 데이터 임의 생성 (WebSearch로 검증, 없으면 [데이터 확인 필요])

## 📚 레퍼런스 (think_tank DB)

- ID 1: **VAETKI Commerce 쇼케이스** (수주작) — 파스텔 에디토리얼 라이트 계열
- ID 2: **메이플스토리 월드 메커톤** (수주작) — 다크 에디토리얼 계열

새 프로젝트에서 `ThinkTankRetrieval.search_similar()` 또는 `recommend_palettes()` 경유로 유사 레퍼런스 자동 연결.

## 🔖 이 표준의 버전

**v4.1 (2026-04-17)** — 사용자 피드백 반영 확정
- 커밋 체인: `74272bf` (폰트 스냅) → `d052851` (본문 3tier) → `c9deaad` (정렬 중앙) → `fa511b0` (섹션 60pt)

이 표준을 변경하고 싶으면 **사용자 명시 요청 후** 이 문서를 업데이트한 다음 관련 코드/컴포넌트 수정.
