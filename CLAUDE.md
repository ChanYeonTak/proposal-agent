# 입찰 제안서 자동 생성 에이전트 (v5.1 - Impact-8 + slide_kit v3.8 + Design Agent + Content Tone + Gamma MCP)

## 프로젝트 개요
RFP(제안요청서) 문서를 입력받아 PPTX 형식의 입찰 제안서를 자동 생성하는 Python 에이전트 시스템

**v5.1 아키텍처**: 콘텐츠(slide_kit) / 디자인(DesignAgent + Gamma) / 톤(ContentTone) 분리
- slide_kit = 경량 콘텐츠 엔진 (RFP 분석, 싱크탱크 참조, 인터넷 서칭)
- DesignAgent = 디자인 결정 (싱크탱크 1순위 + Gamma 2순위 보완)
- **ContentTone = 콘텐츠 톤 결정** (감성 레벨, IP 깊이, 프로그램 네이밍, Win Theme 스타일)
- Gamma MCP = 테마 컨설팅 + 디자인 워싱 (선택적)
- ImagePipeline = IMG_PH 플레이스홀더 → 실제 이미지 교체

## ★★★ 제안서 생성 워크플로우 (최우선 규칙)

사용자가 "제안요청서 폴더에 있는 테스트 XX 폴더 내 파일을 분석한 후 제안서를 제작해줘" 라고 요청하면:

### 폴더 구조
```
제안요청서/테스트 XX/    ← RFP 입력 (PDF 문서들)
output/테스트 XX/        ← PPTX 출력 (생성 스크립트 + 결과물)
```

### 실행 단계

**STEP 1: RFP 분석** (제안요청서 폴더 내 PDF 읽기)
- `제안요청서/테스트 XX/` 내 모든 PDF를 분석
- 추출 항목: 프로젝트명, 발주처, 과업 범위, 평가 기준, 예산, 일정, 특이사항
- 프로젝트 유형 판별: marketing_pr / event / it_system / public / consulting

**STEP 2: 콘텐츠 기획** (Impact-8 Phase 구조)
- Phase 0~7 콘텐츠를 RFP 맞춤형으로 설계
- Win Theme 3개 도출
- Action Title (인사이트 기반 문장형 제목) 작성
- KPI + 산출근거 설계
- ★★★ **KPI 데이터 무결성 규칙 (절대 준수)**:
  - 통계 수치, 시장 데이터, 퍼센트, 배수 등 **정량적 데이터를 임의로 생성/조작하지 않는다**
  - 모든 수치는 반드시 **실제 출처(리서치 보고서, 공식 통계, 검증된 매체)에서 검색하여 인용**한다
  - 출처를 찾을 수 없는 수치는 `[데이터 확인 필요]`로 표기하고, 허위 출처명을 만들지 않는다
  - 인터넷 검색(WebSearch)을 통해 실제 데이터를 확보한 후 삽입한다
  - RFP에 명시된 수치(예산, 인원, 일정 등)는 그대로 사용 가능
  - 산출 근거가 자체 추정인 경우 "자체 추정" 또는 "유사 사례 기반 추정"으로 명확히 표기

**STEP 2.5: 디자인 에이전트** (3단계 파이프라인)

디자인 결정을 콘텐츠와 분리하여 처리합니다.
**우선순위**: 사용자 직접 지정 > 사용자 편집 오버라이드 > 싱크탱크 레퍼런스 > Gamma 추천 > 기본값

**2.5a: 싱크탱크 조회**
- `think_tank.retrieval.search_similar(project_type, industry, won_bid_only=True)` → 유사 레퍼런스
- `think_tank.design_brief.DesignBriefBuilder.build()` → 레퍼런스 기반 디자인 브리프
- 수주 성공 제안서의 컬러/레이아웃/시각밀도 패턴 추출

**2.5b: Gamma 테마 조회** (선택적 — Gamma MCP 사용 가능 시)
- Claude Code가 Gamma MCP `get_themes()` 호출 → 테마 목록 수신
- `DesignAgent.interpret_gamma_themes(결과, keywords)` → 상위 3개 후보 선정
- 프로젝트 주제/분위기와 매칭 (dark/bold/modern 등 키워드 기반)

**2.5c: 병합 → MergedDesignBrief 생성**
- `DesignAgent.merge()` → 양쪽 비교 후 최적 조합 선택
- `register_theme()` + `apply_theme()` → slide_kit 컬러 동적 변경
- MergedDesignBrief에 포함: colors, background_schedule, component_targets, layout_distribution, image_style, **content_tone**

**2.5d: 콘텐츠 톤 가이드 적용** (v5.1 신규)
- `brief.content_tone` → 생성 스크립트의 글쓰기 톤 결정
- 감성 톤 레벨 (1~5): 사무적(1) ~ 풀 내러티브(5)
- IP 깊이: 세계관/캐릭터/커뮤니티 용어 활용 수준
- 프로그램 네이밍: ip_narrative / branded / functional / hybrid
- Win Theme 스타일: ip_worldview / keyword_functional / emotional_hook
- `tone_rules[]`: 생성 스크립트가 직접 참조하는 콘텐츠 규칙 목록

```python
# 콘텐츠 톤 적용 예시
tone = brief.content_tone
print(f"감성 톤: {tone.emotional_tone_level}/5")
print(f"프레이밍: {tone.narrative_framing_style}")
print(f"IP 깊이: {tone.ip_depth_score:.1f}")
print(f"네이밍: {tone.program_naming_style}")
print(f"톤 규칙 {len(tone.tone_rules)}개:")
for rule in tone.tone_rules:
    print(f"  - {rule}")
```

```python
# 디자인 에이전트 사용 예시
from src.agents.design_agent import DesignAgent
agent = DesignAgent()
brief = agent.generate_full_brief(
    project_name="NIKKE AGF 2025",
    project_type="event",
    industry="game_event",
    target_slides=70,
    gamma_themes_data=gamma_result,  # Gamma MCP get_themes() 결과
    project_keywords=["게임", "부스", "AGF", "다크", "모던"],
    custom_colors={                  # 사용자/이전 편집 오버라이드
        "primary": (16, 22, 32),
        "secondary": (0, 42, 128),
        "accent": (98, 150, 255),
    },
)
agent.register_to_slide_kit(brief)
# 이후 slide_kit 함수들이 brief.colors 기반으로 동작
```

**2.5e: 사용자 편집 역파싱** (이전 버전 오버라이드 반영)
```python
from src.integrations.design_bridge import GammaMCPBridge
bridge = GammaMCPBridge(project_dir=Path("output/NIKKE_AGF_2025"))
edits = bridge.extract_user_edits(
    original_pptx=Path("v4.pptx"),  # slide_kit 생성 원본
    edited_pptx=Path("v4_edited.pptx"),  # 사용자가 수정한 버전
)
bridge.save_design_overrides(edits, output_dir, "nikke_agf_2025")
# 다음 버전 생성 시 오버라이드 로드하여 custom_colors에 반영
```

**STEP 3: 생성 스크립트 작성**
- `output/테스트 XX/generate_제안서.py` 스크립트 생성
- **반드시 slide_kit.py import** (아래 규칙 참조)
- **LAYOUTS 프리셋 활용** — `get_zones()` 으로 안전 영역 사용
- 목표 분량: 40~80장 (프로젝트 규모에 따라 조정)
- MergedDesignBrief 있으면: background_schedule, component_targets 참조
- **content_tone 참조**: tone_rules 준수, emotional_tone_level에 맞는 글쓰기 스타일 적용
  - 톤 레벨 4~5: 감성적 스토리텔링, IP 세계관 활용, 내러티브 메타포 관통
  - 톤 레벨 3: 데이터+감성 균형, 인사이트 기반 Action Title
  - 톤 레벨 1~2: 사실 중심, 간결한 전달, 데이터 위주

**STEP 4: 실행 및 검증**
- 스크립트 실행하여 PPTX 생성
- 오류 발생 시 즉시 수정 후 재실행
- 최종 파일 경로 안내

**STEP 5: 이미지 삽입** (선택적)
- `ImagePipelineManager(design_brief=brief)` → 디자인 브리프 기반 소스 자동 선택
- 카테고리별 자동 소스: photo→Pexels, illustration→AI(DALL-E), diagram→matplotlib
- `ImageInserter.insert_images(pptx_path, image_map)` → PPTX 내 IMG_PH 도형 교체

```python
from src.image_pipeline.manager import ImagePipelineManager
from src.image_pipeline.inserter import ImageInserter

mgr = ImagePipelineManager(design_brief=brief, ai_generation_enabled=True)
requests = mgr.extract_placeholders_from_content(proposal_content)
results = await mgr.process_requests(requests)
image_map = {r.placeholder_id: r.file_path for r in results.values() if r.success}
ImageInserter.insert_images(pptx_path, image_map)
```

**STEP 6: Gamma 디자인 워싱** (선택적 — GAMMA_API_KEY 보유 시)
- `bridge.run_gamma_pipeline(pptx_path, brief, num_cards=70)` → 준비 원스텝
- Claude Code가 Gamma MCP `generate(**result.params)` 호출
- `bridge.handle_gamma_response(response, result)` → 다운로드 전략 결정
- `bridge.poll_and_download_gamma(generation_id, result)` → API 폴링 + PPTX 자동 다운로드

```python
from src.integrations.design_bridge import GammaMCPBridge
bridge = GammaMCPBridge(project_dir=Path("output/PROJECT"))

# 1. 파이프라인 준비 (텍스트 추출 + 파라미터 빌드)
result = bridge.run_gamma_pipeline(pptx_path, brief, num_cards=70)

# 2. Claude Code: Gamma MCP generate(**result.params) 호출
# → gamma_response = {"generationId": "...", "status": "pending", ...}

# 3. 응답 처리
download_info = bridge.handle_gamma_response(gamma_response, result)

# 4. API 폴링 + 자동 다운로드 (GAMMA_API_KEY 필요)
output_path = bridge.poll_and_download_gamma(result.generation_id, result)
# → output/PROJECT/gamma_pptx/PROJECT_gamma_0302_1350.pptx
```

### 레이아웃 선택 가이드 (내용에 맞게 적용)

| 콘텐츠 유형 | 권장 레이아웃 | slide_kit 함수 |
|------------|-------------|---------------|
| 시장 환경/배경 분석 | `THREE_COL` or `TWO_COL` | `COLS()` or Zone 직접 |
| 핵심 인사이트/메시지 | `HIGHLIGHT_BODY` | `HIGHLIGHT()` |
| 전략 프레임워크 | `PYRAMID_DESC` | `PYRAMID()` |
| 채널/항목 비교 | `COMPARE_LR` | `COMPARE()` |
| 실행 프로세스 | `PROCESS_DESC` | `FLOW()` |
| KPI/성과 목표 | `KPI_GRID` | `KPIS()` |
| 월별 일정 | `GANTT` | `GANTT_CHART()` |
| 조직도 | `ORG_CHART` | `ORG()` |
| 수행 실적 | `GALLERY_3x2` or `GRID` | `GRID()` |
| 리스크 관리 | `RISK_CARD` | Zone 직접 |
| 데이터 비교 | `TABLE_INSIGHT` | `TABLE()` |
| 통계/수치 강조 | `FULL_BODY` | `STAT_ROW()` |
| 타임라인 | `TIMELINE_DESC` | `TIMELINE()` |
| 차별화 포인트 | `FOUR_COL` | `ICON_CARDS()` |
| 우선순위 매트릭스 | `MATRIX_DESC` | `MATRIX()` |
| 프로그램 소개 | `PROGRAM_CARD_3` | Zone 직접 |
| 키비주얼/대표이미지 | `KEY_VISUAL` | `IMG()` + Zone |
| 인용문/핵심 메시지 | `HIGHLIGHT_BODY` | `QUOTE()` |
| 예산/비율 시각화 | `FULL_BODY` | `PIE_CHART()` or `BAR_CHART()` |
| 추세/성장 데이터 | `FULL_BODY` | `LINE_CHART()` |
| 구조화된 항목 | `FULL_BODY` | `NUMBERED_LIST()` |
| 고급 카드 (그림자) | `THREE_COL` or `FOUR_COL` | `CARD()` |

## ★ 필수 규칙: PPTX 생성 스크립트 작성 시

**모든 제안서 생성 스크립트는 반드시 `src/generators/slide_kit.py`를 import하여 사용해야 합니다.**

```python
# 스크립트 상단에 반드시 추가
import sys; sys.path.insert(0, "/path/to/proposal-agent")
from src.generators.slide_kit import *

# 또는 importlib 사용
import importlib.util
spec = importlib.util.spec_from_file_location('slide_kit', '프로젝트경로/src/generators/slide_kit.py')
sk = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sk)
```

### slide_kit이 제공하는 것 (v3.8)

| 카테고리 | 함수 | 설명 |
|---------|------|------|
| **상수** | `C`, `SW`, `SH`, `ML`, `CW`, `SZ`, `FONT` | 컬러(21색), 크기, 폰트 |
| **상수 (v3.6)** | `FONT_W`, `SHADOW`, `GRAD` | 폰트 웨이트, 그림자 프리셋, 그라디언트 프리셋 |
| **컬러 유틸 (v3.6)** | `darken()`, `lighten()` | RGBColor 밝기 조절 유틸 |
| **Zone** | `Z`, `GAP`, `CGAP`, `CW_IN`, `ML_IN` | 표준 영역, 간격 |
| **레이아웃** | `LAYOUTS`, `get_zones()`, `zone_to_inches()`, `list_layouts()` | 20가지 프리셋 |
| **도형 (기본)** | `R()`, `BOX()`, `OBOX()` | 직각 사각형, 텍스트 박스, 아웃라인 |
| **도형 (v3.5)** | `RBOX()`, `ORBOX()`, `CARD()` | 라운드 박스, 라운드 아웃라인, 통합 카드 |
| **텍스트** | `T(fn=)`, `RT()`, `MT()` | 단일(fn=FONT_W 지원)/리치/멀티라인 |
| **이펙트** | `gradient_bg()`, `bg()`, `set_char_spacing()` | 그래디언트, 자간 |
| **이펙트 (v3.6)** | `gradient_shape()`, `add_shadow(preset=)`, `OVERLAY()` | 도형 그라디언트, 프리셋 그림자, 오버레이 |
| **구분/악센트** | `DIVIDER()`, `ACCENT_LINE()` | 구분선 3종, 좌측 악센트 |
| **컴포넌트** | `IMG()`, `PN()`, `TB()`, `SRC()`, `WB()` | 이미지홀더, 페이지번호, 타이틀바, 출처, Win Theme |
| **텍스트 블록** | `QUOTE()`, `NUMBERED_LIST()` | 인용문(modern/box), 번호 리스트 |
| **도식화 (기본)** | `FLOW()`, `COLS(shadow=)`, `PYRAMID()`, `MATRIX()`, `TABLE()`, `HIGHLIGHT(grad=)`, `KPIS(shadow=)`, `COMPARE()`, `TIMELINE()` | 플로우, 컬럼(그림자), 피라미드 등 |
| **도식화 (확장)** | `GRID(shadow=)`, `STAT_ROW(shadow=)`, `GANTT_CHART()`, `ORG()`, `ICON_CARDS()` | 그리드(그림자), 통계, 간트, 조직도, 아이콘카드 |
| **차트** | `BAR_CHART()`, `PIE_CHART()`, `LINE_CHART(smooth=)` | 바(세로/가로), 파이/도넛, 라인(곡선 수정) |
| **시각화 헬퍼** | `IMG_PH()`, `PROGRESS_BAR()`, `METRIC_CARD(shadow=)`, `STEP_ARROW()`, `DONUT_LABEL()` | 이미지홀더, 프로그레스, 메트릭카드(그림자), 스텝화살표, 도넛 |
| **슬라이드** | `slide_cover()`, `slide_section_divider()`, `slide_toc()`, `slide_exec_summary()`, `slide_next_step()`, `slide_closing()` | 표지(그라디언트), 구분자(그라디언트), 목차, 요약, CTA(그라디언트), 마지막(그라디언트) |
| **자동 배치** | `VStack` 클래스 | 자동 Y좌표 계산, 겹침 방지 |
| **테마** | `THEMES`, `apply_theme()`, `reset_theme()`, `list_themes()`, `register_theme()` | 5+N 테마, 동적 등록/변경, 파생컬러 자동 재계산 |
| **검증** | `validate_sequence()` | 레이아웃 시퀀스 단조로움 검증 |
| **유틸** | `new_presentation()`, `new_presentation_from_template()`, `new_slide()`, `save_pptx()`, `_cols()` | 생성, 템플릿, 저장, 컬럼너비 |

### ★★★ 겹침·공백·경계 방지 규칙 (v3.8 — NIKKE AGF 검증 반영)

**0. 슬라이드 경계 (최우선)**
```
슬라이드 높이: 7.5"
모든 요소 하단(y + h) ≤ 7.0" (하단 0.5" 마진: 페이지번호 + 출처 영역)
위반 시 slide_kit.py COLS/SPLIT_VISUAL 자동 클램핑 작동
수동 배치 시에도 반드시 y + h ≤ 7.0" 확인
```

**1. 요소 간 최소 간격 (인치)**
```
HIGHLIGHT(sub 있음)  → 다음 요소:  1.4" (HIGHLIGHT(sub) 높이 = 1.2")
HIGHLIGHT(sub 없음)  → 다음 요소:  1.0" (HIGHLIGHT 높이 = 0.8")
COLS                → 다음 요소:  0.30"
FLOW                → 다음 요소:  0.2"  (FLOW desc 높이 = 0.4")
METRIC_CARD         → 다음 요소:  0.15"
MT(불릿)            → 다음 요소:  0.20"
```

**2. MT(불릿 텍스트) 높이 — 줄 수에 맞춤**
```
3줄=1.1"  4줄=1.4"  5줄=1.7"  6줄=2.0"  8줄=2.8"
❌ 절대 금지: 줄 수와 무관한 고정 높이 (예: 4줄인데 h=3.2")
```

**3. 한글 텍스트 너비 추정**
```
44pt: 0.61"/자 → CW(~11.8") 내 최대 ~18자
36pt: 0.50"/자 → CW 내 최대 ~23자
→ 44pt 제목이 18자 초과 시 반드시 2줄 분리 (별도 T() 호출)
```

**4. 공백 보완 규칙**
- 콘텐츠 하단 공백 > 0.5" → IMG_PH 또는 HIGHLIGHT 추가
- METRIC_CARD 높이 확대 (비율 기반 배치가 자동 대응)
- 섹션 구분자/표지/마지막 슬라이드의 공백은 의도적 → 수정 불필요

**5. 배경색 충돌 방지**
- slide_next_step 배경: C["dark"] (카드가 C["primary"] 등)
- 카드 색상 = 배경 색상이면 반드시 다른 색상으로 변경

**6. Phase 3 필수 컨셉 장표 (3종)**
1. **Concept Reveal** — 다크 배경, 60pt 대형 컨셉 키워드, 4단계 순환 카드
2. **Strategy Synergy Map** — 3대 Win Theme 연결 구조, 순환 흐름도
3. **Big Idea Reveal** — 36pt 중앙 컨셉 + 3-Step 카드

**7. 시각 요소 필수 포함**
```
| 슬라이드 유형 | 필수 시각 요소 |
|-------------|-------------|
| 시장 분석    | METRIC_CARD 4개 + HIGHLIGHT + IMG_PH |
| 컨셉        | Concept Reveal + Synergy Map |
| 시즌 전략    | 좌우 카드 + IMG_PH (캠페인 비주얼) |
| 이벤트 종합  | TABLE + METRIC_CARD + IMG_PH (현장 사진) |
| 운영 프로세스 | COLS + HIGHLIGHT + IMG_PH (인포그래픽) |
| 커뮤니케이션  | COLS + HIGHLIGHT + IMG_PH (흐름도) |
```

### 절대 하지 말 것
- ❌ 헬퍼 함수를 스크립트 내에 다시 정의하지 말 것
- ❌ RGBColor를 직접 하드코딩하지 말 것 → `C["primary"]` 사용
- ❌ 폰트명을 직접 쓰지 말 것 → `FONT` 상수 사용
- ❌ "맑은 고딕" 등 다른 폰트 사용 금지 → Pretendard만 사용
- ❌ `→` 또는 `←` 화살표 문자를 텍스트로 사용하지 말 것 → FLOW()나 구분선 활용
- ❌ TB(pg=pg) 사용 시 별도 PN() 호출 금지 → TB()가 자체 호출 (중복됨)
- ❌ `sz=Pt(N)` 금지 → `sz=N` (T()가 내부에서 Pt() 적용)
- ❌ RBOX(color=clr) 금지 → `RBOX(s, l, t, w, h, clr)` (위치 인수)
- ❌ 다크 배경 위 회색 텍스트 금지 → `C["white"]` 또는 밝은 색 사용
- ❌ 일반 장표에 `WB()` 직접 호출 금지 → Win Theme 뱃지는 **간지(section divider)에만** 표시. `slide_section_divider()`가 `win_theme_key` 파라미터로 자동 삽입
- ❌ EMU int와 float 혼합 산술 시 직접 전달 금지 → `CW * 0.5` 같은 float 결과는 `int()` 래핑 필요

### ★ 컴포넌트 API 파라미터 형식 (v3.8)

| 컴포넌트 | items 형식 | 비고 |
|---------|-----------|------|
| `FLOW()` | `[("제목", "설명"), ...]` | tuple만 |
| `STEP_ARROW()` | `[("번호", "제목", "설명"), ...]` | 3-tuple |
| `TIMELINE()` | `[("기간", "내용"), ...]` 또는 `[{"label": ..., "desc": ...}, ...]` | tuple/dict 모두 OK |
| `PYRAMID()` | `[("텍스트", color), ...]` | tuple |
| `COLS()` | `[{"title": ..., "body": [...]}, ...]` | dict |
| `GRID()` | `[{"title": ..., "body": ...}, ...]` | dict |
| `STAT_ROW()` | `[{"value": ..., "label": ...}, ...]` | dict |
| `KPIS()` | `[{"value": ..., "label": ..., "basis": ...}, ...]` | dict |
| `ORG()` | `pm={"name": ..., "role": ...}, dirs=[{"name": ..., "role": ...}]` | dict |
| `slide_next_step()` | `steps=[("STEP 1", "제목", "설명", color), ...]` | 4-tuple |
| `slide_exec_summary()` | 위치 인수: `(prs, title, one_liner, win_themes_dict, kpis, why_us_points)` | 순서형 |
| `GANTT_CHART()` | 위치 인수: `(s, categories, months, data, y=, colors=)` | 순서형 |

### ★ 디자인 품질 체크리스트 (생성 후 필수 확인)

1. **소제목/라벨 가시성**: 최소 16pt (`SZ["subtitle"]`), 상단 0.8" 이상 여백
2. **색상 대비**: 다크 배경 → 흰색/밝은색 텍스트 only, 파란색 박스 → 회색 폰트 금지
3. **경계 초과**: 모든 요소 y+h ≤ 7.0" (하단 마진 확보)
4. **HIGHLIGHT → 다음 요소 간격**: sub 있으면 최소 1.4" 후 시작
5. **PN() 중복**: TB(pg=pg)와 PN(s, pg) 동시 호출 금지 (slide_kit 자동 방지 적용됨)
6. **FLOW desc 높이**: 0.4" 이하 (이전 0.8"에서 수정됨)
7. **VStack 사용 시**: breathe(0.3) 이상 필수 (컴포넌트 간 겹침 방지)

### 기본 사용 패턴

```python
prs = new_presentation()
WIN = {"data": "...", "story": "...", "ugc": "..."}

# 표지
slide_cover(prs, "프로젝트명", "발주처명")

# 목차
slide_toc(prs, "목차", [("01", "HOOK", "설명"), ...], pg=2)

# 섹션 구분자
slide_section_divider(prs, "01", "사업이해", "부제", "스토리", "data", WIN)

# 일반 콘텐츠
s = new_slide(prs)
bg(s, C["white"])
TB(s, "Action Title — 인사이트 기반 제목", pg=3)
MT(s, ML, Inches(1.3), CW, Inches(3), ["항목1", "항목2"], bul=True)

# 저장
save_pptx(prs, "output/파일명.pptx")
```

**v3.1 업데이트**: Win Theme, Executive Summary, Next Step, Action Title 시스템 도입
- Win Theme: 제안서 전체에 반복되는 핵심 수주 전략 메시지
- Executive Summary: 의사결정권자용 1페이지 핵심 요약
- Next Step: 다음 단계 안내 / Call to Action
- Action Title: 인사이트 기반 슬라이드 제목 (Topic Title → Action Title)

## 역할 분리 (v5.1 아키텍처)

### slide_kit (콘텐츠 엔진 — 경량)
- PPTX 생성 기본 도구 (도형, 텍스트, 레이아웃, 차트)
- 테마 시스템 (register_theme + apply_theme)
- **역할**: RFP/싱크탱크 참조로 콘텐츠를 채우는 데 집중

### DesignAgent (디자인 + 톤 결정)
- 싱크탱크 + Gamma 병합 → MergedDesignBrief
- 컬러/폰트/레이아웃/이미지 스타일 결정
- **콘텐츠 톤 결정** → ContentToneBrief (감성 레벨, IP 깊이, 네이밍 스타일)
- slide_kit 테마 동적 등록

### ContentTone (콘텐츠 톤 시스템 — v5.1 신규)
- 싱크탱크 레퍼런스에서 감성 톤/IP 깊이/네이밍 패턴 자동 분석
- 산업별 기본 톤 규칙 내장 (game_event, marketing_pr, event, it_system, public, consulting)
- `tone_rules[]` → 생성 스크립트가 직접 참조하는 콘텐츠 작성 지침
- 어떤 RFP든 해당 산업에서 수주한 제안서의 톤을 자동 재현

### Gamma MCP (디자인 워싱 — 선택적)
- `get_themes()`: 프로젝트 주제별 테마 컨설팅
- `generate()`: 전체 콘텐츠를 Gamma에 전송하여 디자인된 프레젠테이션 생성
- **Gamma Public API v1.0**: `https://public-api.gamma.app/v1.0/generations`
- **API 인증**: `X-API-KEY` 헤더 + `sk-gamma-xxxxx` 형식 키
- **디자인 오버라이드**: Gamma/사용자 편집은 slide_kit 기본값보다 항상 우선

### GammaMCPBridge v7.1 (역파싱 + 브릿지 + API 폴링)
- **전체 파이프라인**: PPTX 텍스트 추출 → Gamma inputText 변환 → generate() → API 폴링 → PPTX 자동 다운로드
- **API 폴링**: `poll_and_download_gamma()` — 5초 간격 GET 폴링, `downloadLink`/`exportUrl` 획득 → HTTP 다운로드
- **3-way 다운로드 전략**: `direct` (응답에 exportUrl 포함) / `api_poll` (GAMMA_API_KEY 보유) / `browser` (수동 fallback)
- **사용자 편집 역파싱**: `extract_user_edits()` — 14개 카테고리 변경 감지 (지오메트리, 채우기, 타이포 등)
- **디자인 오버라이드**: `save_design_overrides()` / `load_design_overrides()` — 다음 버전 자동 반영
- **Cloudflare 대응**: `User-Agent: GammaBridge/1.0` 헤더 필수 (기본 Python urllib 차단됨)
- **연속 에러 보호**: `max_consecutive_errors=10` — 무한 재시도 방지

## 디렉토리 구조

```
├── main.py                 # CLI 엔트리포인트
├── config/
│   ├── prompts/            # Phase별 프롬프트 템플릿
│   └── design/             # 디자인 설정
│       └── design_style.py    # Modern 스타일 정의
├── src/
│   ├── parsers/            # 문서 파싱 (PDF, DOCX)
│   ├── agents/             # Claude 에이전트
│   │   ├── base_agent.py       # 에이전트 추상 클래스
│   │   ├── rfp_analyzer.py     # RFP 분석
│   │   ├── content_generator.py # 콘텐츠 생성
│   │   └── design_agent.py     # ★ 디자인 에이전트 (싱크탱크+Gamma 병합)
│   ├── generators/         # PPTX 생성
│   │   └── slide_kit.py       # ★ 콘텐츠 엔진 (v3.8)
│   ├── integrations/       # 외부 도구 연동
│   │   └── design_bridge.py   # ★ GammaMCPBridge v7.1 (API 폴링 자동 다운로드 + 역파싱 14카테고리)
│   ├── image_pipeline/     # 이미지 수급/삽입
│   │   ├── manager.py         # ★ 파이프라인 매니저 (brief 통합)
│   │   ├── inserter.py        # PPTX IMG_PH → 실제 이미지 교체
│   │   └── sources/
│   │       ├── web_search.py      # Pexels/Unsplash API
│   │       ├── ai_generator.py    # DALL-E/Stability AI
│   │       └── diagram_renderer.py # matplotlib 다이어그램
│   ├── orchestrators/      # 워크플로우 조율
│   └── schemas/            # Pydantic 스키마
│       ├── proposal_schema.py  # Impact-8 스키마
│       ├── rfp_schema.py       # RFP 분석 스키마
│       ├── design_schema.py    # ★ MergedDesignBrief, ThemeRecommendation
│       └── ip_research_schema.py # IP 리서치 스키마
├── think_tank/             # ★ 레퍼런스 DB + 디자인 브리프
│   ├── db.py                  # SQLite 레퍼런스 DB
│   ├── models.py              # DesignProfile, ContentPattern, ProgramTemplate
│   ├── design_brief.py        # DesignBriefBuilder (레퍼런스 → 브리프)
│   ├── retrieval.py           # ThinkTankRetrieval (유사 레퍼런스 검색)
│   └── ingestion/             # 레퍼런스 문서 인제스팅
├── templates/              # PPTX 템플릿
├── company_data/           # 회사 정보
├── input/                  # RFP 입력
├── output/                 # PPTX 출력
└── 제안요청서/             # RFP 입력 문서 (PDF)
```

## 사용법

```bash
# 의존성 설치
pip install -r requirements.txt

# .env 설정
cp .env.example .env
# ANTHROPIC_API_KEY 설정

# 제안서 생성 (기본: Impact-8 구조)
python main.py generate input/rfp.pdf -n "프로젝트명" -c "발주처"

# 프로젝트 유형 지정
python main.py generate input/rfp.pdf -n "프로젝트명" -c "발주처" -t marketing_pr

# RFP 분석만 수행
python main.py analyze input/rfp.pdf
```

## 제안서 구조: Impact-8 Framework

실제 수주 성공 제안서 분석을 기반으로 개선된 8-Phase 구조

```
┌─────────────────────────────────────────────────────────────┐
│  PHASE 0: HOOK (티저)                         3-10p (5%)   │
│  → 임팩트 있는 오프닝, 핵심 메시지, 비전                      │
├─────────────────────────────────────────────────────────────┤
│  PHASE 1: SUMMARY                             3-5p (5%)    │
│  → Executive Summary (의사결정자용 5분 요약)                 │
├─────────────────────────────────────────────────────────────┤
│  PHASE 2: INSIGHT                             8-15p (10%)  │
│  → 시장 환경 + 문제 정의 + 숨겨진 니즈                       │
├─────────────────────────────────────────────────────────────┤
│  PHASE 3: CONCEPT & STRATEGY                  8-15p (12%)  │
│  → 핵심 컨셉 + 차별화 전략 + 경쟁 우위                       │
├─────────────────────────────────────────────────────────────┤
│  PHASE 4: ACTION PLAN (★핵심)                 30-60p (40%) │
│  → 상세 실행 계획 + 콘텐츠 예시 + 채널별 전략                 │
├─────────────────────────────────────────────────────────────┤
│  PHASE 5: MANAGEMENT                          6-12p (10%)  │
│  → 조직 + 운영 + 품질관리 + 리포팅                          │
├─────────────────────────────────────────────────────────────┤
│  PHASE 6: WHY US                              8-15p (12%)  │
│  → 수행 역량 + 유사 실적 + 레퍼런스                          │
├─────────────────────────────────────────────────────────────┤
│  PHASE 7: INVESTMENT & ROI                    4-8p (6%)    │
│  → 투자 비용 + 정량적 효과 + ROI                            │
└─────────────────────────────────────────────────────────────┘
  총 70-140p (프로젝트 규모에 따라 조정)
```

## 프로젝트 유형별 가중치

| Phase | Marketing/PR | Event | IT/System | Public | Consulting |
|-------|-------------|-------|-----------|--------|------------|
| 0. HOOK | 8% | 6% | 3% | 3% | 5% |
| 1. SUMMARY | 5% | 5% | 8% | 8% | 8% |
| 2. INSIGHT | 12% | 8% | 12% | 15% | 15% |
| 3. CONCEPT | 12% | 10% | 10% | 10% | 12% |
| 4. ACTION | **40%** | **45%** | 35% | 30% | 30% |
| 5. MANAGEMENT | 8% | 10% | 12% | 12% | 10% |
| 6. WHY US | 10% | 10% | 12% | 15% | 12% |
| 7. INVESTMENT | 5% | 6% | 8% | 7% | 8% |

## v3.1 핵심 컴포넌트

### Win Theme (수주 전략 메시지)
제안서 전체에 반복되는 3대 핵심 수주 전략 메시지

```python
WIN_THEMES = {
    "data": "데이터 기반 타겟 마케팅",
    "community": "시민 참여형 브랜드 빌딩",
    "integration": "온-오프라인 통합 시너지",
}
```

- 각 섹션 구분자에 관련 Win Theme 표시
- 슬라이드 내에서 Win Theme 뱃지로 강조
- 일관된 메시지 반복으로 수주 전략 강화

### Action Title (인사이트 기반 제목)
Topic Title에서 Action Title로 전환

| Before (Topic Title) | After (Action Title) |
|---------------------|---------------------|
| 타겟 분석 | MZ세대 2030이 핵심, 하루 SNS 55분 사용 |
| 채널 전략 | 인스타그램 중심, 릴스로 도달률 3배 확보 |
| 예산 계획 | 월 3,000만원으로 팔로워 50만 달성 |

### Executive Summary
의사결정권자용 1페이지 핵심 요약

구성요소:
- 프로젝트 목표 (One Sentence Pitch)
- 3대 Win Theme
- 핵심 KPI (산출 근거 포함)
- Why Us 핵심 차별점

### Next Step (Call to Action)
다음 단계 안내 및 행동 촉구

```
┌─────────────────────────────────────────┐
│  NEXT STEP                              │
│                                         │
│  STEP 1: 제안 설명회 (00월 00일)         │
│  STEP 2: Q&A 및 추가 협의               │
│  STEP 3: 계약 체결                      │
│                                         │
│  Contact: [담당자 정보]                 │
└─────────────────────────────────────────┘
```

### KPI 산출 근거
모든 KPI에 산출 근거 필수 포함

```
목표: 팔로워 +30%
산출 근거: 인플루언서 협업 +10% + 릴스 확대 +12% + 이벤트 +8%
데이터 출처: 유사 프로젝트 평균 성장률 참고
```

**★★★ KPI 데이터 무결성 원칙**
- ❌ 절대 금지: 통계/시장 데이터/퍼센트/배수 등 정량 수치를 AI가 임의 생성
- ❌ 절대 금지: 실존하지 않는 보고서명/기관명을 출처로 기재 (예: "McKinsey XX Report 2026")
- ✅ 필수: WebSearch로 실제 데이터를 검색하여 검증된 수치만 사용
- ✅ 필수: 출처를 찾을 수 없으면 `[데이터 확인 필요]` 플레이스홀더 사용
- ✅ 허용: RFP에 명시된 수치 그대로 인용 (예산 2,000만원, 참여인원 150명 등)
- ✅ 허용: 자체 추정 시 "자체 추정", "유사 사례 기반 추정"으로 명시

### Placeholder 표준화
미완성 콘텐츠 표기 형식 통일: `[대괄호]`

```
✅ [발주처명], [프로젝트명], [담당자 연락처]
❌ OOO, XXX, ___
```

## 디자인 스타일: Modern

실제 수주 성공 제안서를 분석하여 추출한 디자인 시스템

### 컬러 팔레트
- Primary: `#002C5F` (다크 블루)
- Secondary: `#00AAD2` (스카이 블루)
- Teal: `#00A19C` (틸 - Win Theme 뱃지용)
- Accent: `#E63312` (레드)
- Dark BG: `#1A1A1A`
- Light BG: `#F5F5F5` (밝은 배경)

### 타이포그래피
- Font: Pretendard
- 티저 타이틀: 72pt Bold
- 섹션 타이틀: 48pt Bold
- 슬라이드 타이틀: 36pt SemiBold
- 본문: 18pt Regular

### 레이아웃
- 16:9 비율 (1920 x 1080)
- 여백: 상 80px, 하 60px, 좌우 100px
- 섹션 구분자: 다크 배경, 대형 숫자 아웃라인

## 핵심 컴포넌트

### 스키마
- `src/schemas/proposal_schema.py` - ProposalContent, PhaseContent, WinTheme, KPIWithBasis
- `src/schemas/rfp_schema.py` - RFPAnalysis
- `src/schemas/design_schema.py` - **MergedDesignBrief**, **ContentToneBrief**, ThemeRecommendation, ImageStyleGuide

### 에이전트
- `src/agents/rfp_analyzer.py` - RFP 분석
- `src/agents/content_generator.py` - 콘텐츠 생성
- `src/agents/design_agent.py` - **디자인 에이전트** (싱크탱크+Gamma 병합)

### 생성기
- `src/generators/slide_kit.py` - PPTX 콘텐츠 엔진 (v3.8, 40+ 컴포넌트)

### 디자인 파이프라인
- `src/integrations/design_bridge.py` - **GammaMCPBridge v7.1** (Gamma 연동, API 폴링 자동 다운로드, 역파싱 14카테고리)
  - `run_gamma_pipeline()` — 텍스트 추출 + 파라미터 빌드 원스텝
  - `handle_gamma_response()` — 3-way 다운로드 전략 결정 (direct/api_poll/browser)
  - `poll_and_download_gamma()` — Gamma API v1.0 폴링 + PPTX 자동 다운로드
  - `extract_user_edits()` — 원본 vs 수정 PPTX 14카테고리 변경 감지
- `src/image_pipeline/manager.py` - **ImagePipelineManager** (brief 기반 소스 자동 선택)
- `src/image_pipeline/inserter.py` - **ImageInserter** (IMG_PH → 이미지 교체)

### 싱크탱크 (레퍼런스 DB)
- `think_tank/retrieval.py` - ThinkTankRetrieval (유사 레퍼런스 검색, **콘텐츠 톤 패턴 검색**)
- `think_tank/design_brief.py` - DesignBriefBuilder (레퍼런스 → 디자인 브리프 + **콘텐츠 톤 분석**)
- `think_tank/models.py` - DesignProfile, ContentPattern, ProgramTemplate, **ContentToneProfile**

### 콘텐츠 가이드라인
- `config/prompts/content_guidelines.txt` - Action Title, Win Theme, KPI 산출 근거 작성 가이드

## 마케팅/PR 특화 기능

### 콘텐츠 예시 생성
- 실제 포스팅 예시 (비주얼 설명, 카피)
- 해시태그 전략
- 캠페인 상세 기획

### 채널별 전략
- Instagram: 피드, 스토리, 릴스
- YouTube: 롱폼, 숏폼, 커뮤니티
- Facebook, X, TikTok, Blog

### 캠페인 기획
- 캠페인 컨셉 및 목표
- 실행 계획
- 예상 성과

## Gamma API 기술 상세 (v1.0)

### API 엔드포인트
- **Base URL**: `https://public-api.gamma.app/v1.0/generations`
- **인증**: `X-API-KEY: sk-gamma-xxxxx` 헤더
- **v0.2 종료**: 2026년 1월 16일 일몰 (sunset), v1.0 필수

### 생성 플로우
```
1. POST /generations  (Gamma MCP generate() 호출)
   → {"generationId": "xxx", "status": "pending"}

2. GET /generations/{id}  (5초 간격 폴링)
   → {"status": "in_progress", ...}
   → {"status": "completed", "downloadLink": "https://...pptx"}

3. HTTP GET downloadLink  → PPTX 파일 다운로드
```

### 다운로드 필드 우선순위
```
downloadLink > exportUrl > pptxUrl > pdfUrl > files[0].url
```

### 필수 헤더 (Cloudflare 봇 감지 대응)
```python
headers = {
    "X-API-KEY": api_key,
    "User-Agent": "GammaBridge/1.0",  # Python-urllib 기본값은 차단됨
    "Accept": "application/json",
}
```

### 환경변수
| 변수 | 용도 | 필수 |
|------|------|------|
| `GAMMA_API_KEY` | Gamma Public API v1.0 인증 키 | STEP 6 사용 시 |
| `ANTHROPIC_API_KEY` | Claude API 키 | 항상 |
| `PEXELS_API_KEY` | Pexels 이미지 검색 API | STEP 5 이미지 삽입 시 |

### 파라미터 빌드 (`build_gamma_params`)
```python
params = {
    "inputText": extracted_text,     # PPTX에서 추출한 전체 텍스트
    "format": "presentation",
    "textMode": "preserve",          # 원본 구조 유지
    "exportAs": "pptx",              # downloadLink 활성화 필수
    "numCards": 70,                  # 슬라이드 수
    "themeId": "stratos",            # Gamma 테마 ID
    "imageOptions": {
        "source": "aiGenerated",
        "model": "imagen-4-pro",     # AI 이미지 모델
    },
    "textOptions": {"language": "ko"},
}
```

## 레퍼런스

- 실제 수주 성공 제안서 (200p+) — 구조 분석 레퍼런스
  - 구조: INTRO(13p) + CONCEPT(31p) + STRATEGY(14p) + ACTION PLAN(101p) + MANAGEMENT(16p) + CREDENTIALS(44p)
  - 핵심: ACTION PLAN이 전체의 46% 차지
  - 특징: 실제 콘텐츠 예시, AI 캠페인, 숏폼-롱폼 연계 전략
