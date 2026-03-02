# Proposal Agent — AI 입찰 제안서 자동 생성 에이전트

RFP(제안요청서) PDF를 입력하면 **40~80장 PPTX 입찰 제안서**를 자동 생성하는 AI 에이전트 시스템

## 핵심 특징

- **Impact-8 Framework**: 실제 수주 성공 제안서 분석 기반 8-Phase 구조
- **3-Layer 분리 아키텍처**: 콘텐츠(slide_kit) / 디자인(DesignAgent) / 톤(ContentTone)
- **Think Tank**: 수주 성공 레퍼런스 DB — 디자인/콘텐츠 패턴 자동 학습 및 검색
- **slide_kit v3.8**: 40+ 컴포넌트, 20 레이아웃, 동적 테마, VStack 자동 배치
- **Gamma MCP 연동**: 테마 컨설팅 + 디자인 워싱 + API 폴링 자동 다운로드
- **이미지 파이프라인**: Pexels/DALL-E/matplotlib 소스 → IMG_PH 자동 교체
- **콘텐츠 톤 시스템**: 산업별 감성 레벨, IP 깊이, 프로그램 네이밍 자동 결정

## 빠른 시작

### Claude Code 방식 (권장)

```bash
pip install -r requirements.txt
```

Claude Code에게 자연어로 요청:

```
"제안요청서 폴더에 있는 테스트 XX 폴더 내 파일을 분석한 후 제안서를 제작해줘"
```

### CLI 방식

```bash
pip install -r requirements.txt

# .env 설정
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env

# 제안서 생성
python main.py generate input/rfp.pdf -n "프로젝트명" -c "발주처명" -t event
```

## 파이프라인 (6-Step)

```
RFP (PDF)
    │
    ▼
STEP 1: RFP 분석 — PDF 파싱 + 프로젝트 유형 판별
    │
    ▼
STEP 2: 콘텐츠 기획 — Impact-8 Phase 구조 + Win Theme + Action Title + KPI
    │
    ▼
STEP 2.5: 디자인 에이전트
    ├─ 2.5a: Think Tank 조회 (수주 성공 레퍼런스 검색)
    ├─ 2.5b: Gamma 테마 추천 (선택적)
    ├─ 2.5c: 병합 → MergedDesignBrief
    ├─ 2.5d: 콘텐츠 톤 가이드 적용
    └─ 2.5e: 사용자 편집 역파싱 (오버라이드 반영)
    │
    ▼
STEP 3: 생성 스크립트 작성 — slide_kit 기반 Python 스크립트
    │
    ▼
STEP 4: PPTX 생성 — 40~80장 PPTX 출력
    │
    ▼
STEP 5: 이미지 삽입 (선택적) — IMG_PH → 실제 이미지 교체
    │
    ▼
STEP 6: Gamma 디자인 워싱 (선택적) — API 폴링 + PPTX 자동 다운로드
```

## Impact-8 Framework

| Phase | 이름 | 비중 | 설명 |
|-------|------|------|------|
| 0 | HOOK | 5% | 임팩트 있는 오프닝 |
| 1 | EXECUTIVE SUMMARY | 5% | 의사결정자용 요약 + Win Theme 정의 |
| 2 | INSIGHT | 12% | 시장 환경 + Pain Point |
| 3 | CONCEPT & STRATEGY | 12% | 핵심 컨셉 + 차별화 전략 |
| 4 | ACTION PLAN | **40%** | 상세 실행 계획 (핵심) |
| 5 | MANAGEMENT | 8% | 조직 + 운영 + 품질관리 |
| 6 | WHY US | 12% | 수행 역량 + 실적 |
| 7 | INVESTMENT & ROI | 6% | 비용 + 기대효과 |

## 프로젝트 유형별 자동 적응

| 유형 | Phase 4 비중 | 특화 콘텐츠 |
|------|-------------|-------------|
| 마케팅/PR | 40% | 채널별 전략, 콘텐츠 예시, 인플루언서 |
| 이벤트 | 45% | 공간 설계, 프로그램표, 참가자 여정 |
| IT/시스템 | 35% | 시스템 아키텍처, WBS, 간트 |
| 공공 | 30% | RFP 대응표, 정책 연계 |
| 컨설팅 | 30% | 전략 프레임워크, 벤치마킹 |

## 아키텍처

```
┌──────────────────────────────────────────────────────────────┐
│                    proposal-agent v5.1                        │
│                                                              │
│  ┌──────────┐  ┌───────────┐  ┌───────────┐  ┌──────────┐  │
│  │   RFP    │→│  Content  │→│  Design   │→│ slide_kit│→ PPTX
│  │ Analyzer │  │ Generator │  │   Agent   │  │   v3.8  │  │
│  └──────────┘  └───────────┘  └───────────┘  └──────────┘  │
│       ↑              ↑              ↑                        │
│  ┌──────────┐  ┌───────────┐  ┌───────────┐                │
│  │Think Tank│  │    IP     │  │   Gamma   │                │
│  │    DB    │  │Researcher │  │    MCP    │                │
│  └──────────┘  └───────────┘  └───────────┘                │
│                                     ↓                        │
│  ┌──────────┐               ┌──────────────┐               │
│  │  Image   │               │GammaMCPBridge│               │
│  │ Pipeline │               │    v7.1      │               │
│  └──────────┘               └──────────────┘               │
└──────────────────────────────────────────────────────────────┘
```

### 핵심 모듈

| 모듈 | 역할 |
|------|------|
| **slide_kit v3.8** | PPTX 콘텐츠 엔진 — 40+ 컴포넌트, 20 레이아웃, 동적 테마 |
| **DesignAgent** | 디자인 결정 — Think Tank(1순위) + Gamma(2순위) 병합 → MergedDesignBrief |
| **ContentTone** | 콘텐츠 톤 — 감성 레벨(1~5), IP 깊이, 네이밍 스타일, tone_rules |
| **Think Tank** | 레퍼런스 DB — 수주 성공 제안서 패턴 학습/검색 (SQLite) |
| **GammaMCPBridge v7.1** | Gamma 연동 — API 폴링 자동 다운로드 + 14카테고리 역파싱 |
| **ImagePipeline** | 이미지 수급 — Pexels/DALL-E/matplotlib → IMG_PH 교체 |

## 디렉토리 구조

```
├── main.py                          # CLI 엔트리포인트
├── config/
│   ├── prompts/                     # Phase별 프롬프트 템플릿
│   ├── design/                      # 디자인 시스템
│   ├── industry_profiles/           # 산업별 프로파일 (6종)
│   └── pipeline_config.yaml         # 파이프라인 ON/OFF 설정
├── src/
│   ├── agents/
│   │   ├── rfp_analyzer.py          # STEP 1: RFP 분석
│   │   ├── content_generator.py     # STEP 2: 콘텐츠 생성
│   │   ├── ip_researcher.py         # IP 딥 리서치
│   │   └── design_agent.py          # STEP 2.5: 디자인 에이전트
│   ├── generators/
│   │   └── slide_kit.py             # STEP 3-4: PPTX 엔진 (v3.8)
│   ├── integrations/
│   │   └── design_bridge.py         # STEP 6: GammaMCPBridge v7.1
│   ├── image_pipeline/              # STEP 5: 이미지 파이프라인
│   │   ├── manager.py               #   소스 자동 선택
│   │   ├── inserter.py              #   IMG_PH → 이미지 교체
│   │   └── sources/                 #   Pexels / DALL-E / matplotlib
│   ├── schemas/
│   │   ├── design_schema.py         #   MergedDesignBrief, ContentToneBrief
│   │   ├── proposal_schema.py       #   Impact-8 구조
│   │   └── rfp_schema.py            #   RFP 분석
│   └── pipeline/                    # 파이프라인 엔진 (ON/OFF 제어)
├── think_tank/                      # 레퍼런스 DB
│   ├── db.py                        #   SQLite 저장소
│   ├── models.py                    #   DesignProfile, ContentToneProfile
│   ├── design_brief.py              #   DesignBriefBuilder
│   ├── retrieval.py                 #   유사 레퍼런스 검색
│   └── ingestion/                   #   레퍼런스 인제스팅 (PDF/PPTX)
├── ref/                             # 레퍼런스 제안서 (git 제외)
├── input/                           # RFP 입력 문서 (git 제외)
├── output/                          # 생성된 PPTX (git 제외)
└── docs/                            # 가이드 문서
```

## 환경변수

| 변수 | 용도 | 필수 |
|------|------|------|
| `ANTHROPIC_API_KEY` | Claude API | 항상 |
| `GAMMA_API_KEY` | Gamma Public API v1.0 | STEP 6 사용 시 |
| `PEXELS_API_KEY` | Pexels 이미지 검색 | STEP 5 사용 시 |

## 기술 스택

| 카테고리 | 기술 |
|---------|------|
| AI | Claude (Claude Code / Anthropic API) |
| 문서 처리 | pypdf, pdfplumber, python-pptx |
| 데이터 | Pydantic v2, SQLite |
| 디자인 연동 | Gamma MCP (Public API v1.0) |
| 이미지 | Pexels API, DALL-E, matplotlib |
| CLI | Typer, Rich |

## 가이드 문서

- [설치 및 사용 가이드](docs/INSTALL_AND_USAGE.md)
- [에이전트 구축 방식 · 시스템 구조](docs/입찰제안서_에이전트_가이드.md)
- [상세 사용 가이드](docs/제안서_에이전트_사용_가이드.md)
- [기술 문서](docs/PROPOSAL_AGENT_GUIDE.md)
- [v4.0 구현 계획서](docs/4.0%20구현%20계획서.md)
- [v5.1 구현 계획서](docs/5.1%20구현%20계획서.md)

## 버전 히스토리

| 버전 | 주요 변경 |
|------|----------|
| v3.0 | slide_kit 초기 (기본 컴포넌트) |
| v3.5 | VStack, 네이티브 차트, 테마 시스템, 20 레이아웃 |
| v3.6 | Win Theme 전달 체인 + Action Title + C-E-I 설득 구조 |
| v3.8 | NIKKE AGF 검증, 겹침/공백/경계 규칙 강화 |
| v4.0 | 파이프라인 엔진, Think Tank, IP 리서치, 산업 프로파일 |
| v5.0 | DesignAgent, ContentTone, register_theme/apply_theme |
| v5.1 | ContentToneBrief 상세화, tone_rules 시스템 |
| v7.0 | GammaPipelineResult, run_gamma_pipeline 원스텝 |
| **v7.1** | API 폴링 자동 다운로드, Gamma API v1.0, Cloudflare 대응 |

## 라이선스

Private — 비공개 프로젝트
