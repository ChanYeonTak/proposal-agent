"""
게임 이벤트 산업 프로파일

레퍼런스 분석 결과 기반:
- PACK 구조 (Booth Design Pack, Event Pack, Interaction Pack, Stage Pack, Campaign Pack, Operation Pack)
- 프로그램당 2-3 슬라이드 딥다이브
- CHECK POINT 프레임워크: Analytics -> Solution -> Effect
- 시각 요소 필수: 3D 렌더링, 동선도, 평면도, 프로그램 플로우
"""

from .base_profile import IndustryProfile, PhaseGuide

GAME_EVENT_PROFILE = IndustryProfile(
    name="게임 이벤트/부스 운영",
    description="게임 IP 기반 오프라인 이벤트, 부스 운영, 게임쇼 참여 제안서",
    industry_type="game_event",

    phase_guides={
        0: PhaseGuide(
            phase_number=0,
            structure_type="teaser",
            depth_guide="IP의 세계관과 이벤트 비전을 연결하는 임팩트 오프닝. 대형 비주얼 중심.",
            required_elements=["컨셉 키비주얼", "핵심 슬로건", "IP 세계관 연결"],
            slides_per_item=2,
            visual_requirements=["풀블리드 이미지", "대형 타이포그래피"],
        ),
        2: PhaseGuide(
            phase_number=2,
            structure_type="analytics",
            depth_guide=(
                "IP 현황 데이터(MAU, 커뮤니티 규모) + 행사장 환경 분석 + "
                "타겟 관람객 프로파일. CHECK POINT: Analytics 파트."
            ),
            required_elements=[
                "IP 유저 데이터 (MAU, DAU, 커뮤니티 규모)",
                "행사장 환경 분석 (위치, 규모, 동선)",
                "타겟 관람객 프로파일 (연령, 관심사)",
                "경쟁 부스 벤치마킹",
            ],
            slides_per_item=2,
            visual_requirements=["데이터 차트", "관람객 퍼소나 카드", "벤치마크 비교표"],
            persuasion_framework="CHECK POINT - Analytics",
        ),
        3: PhaseGuide(
            phase_number=3,
            structure_type="concept",
            depth_guide=(
                "IP 세계관에서 파생된 부스 컨셉. 내러티브 아크로 관람객 경험 설계. "
                "3대 전략 축과 순환 구조 시각화."
            ),
            required_elements=[
                "핵심 컨셉 키워드 (IP 세계관 연결)",
                "내러티브 아크 (관람객 여정)",
                "3대 전략 축 (체험/콘텐츠/커뮤니티)",
                "컨셉 비주얼 무드보드",
            ],
            slides_per_item=2,
            visual_requirements=["컨셉 다이어그램", "내러티브 아크 플로우", "무드보드"],
        ),
        4: PhaseGuide(
            phase_number=4,
            structure_type="pack",
            depth_guide=(
                "PACK 단위로 분해. 각 PACK은 프로그램별 2-3 슬라이드 딥다이브. "
                "CHECK POINT: Solution -> Effect. "
                "프로그램마다 컨셉/메커니즘/보상/운영계획/기대효과 포함."
            ),
            required_elements=[
                "BOOTH DESIGN PACK (부스 외관, 내부 동선, 3D 렌더링)",
                "EVENT PACK (메인 이벤트, 프로그램별 딥다이브)",
                "INTERACTION PACK (체험형 콘텐츠, 미니게임, 포토존)",
                "STAGE PACK (스테이지 프로그램, 타임테이블)",
                "CAMPAIGN PACK (SNS 캠페인, 사전/현장/사후 마케팅)",
                "OPERATION PACK (운영 인력, 동선 관리, 비상 대응)",
            ],
            slides_per_item=3,
            visual_requirements=[
                "3D 부스 렌더링", "평면도/동선도", "프로그램 플로우차트",
                "타임테이블", "캠페인 비주얼 예시",
            ],
            persuasion_framework="CHECK POINT - Solution/Effect",
        ),
        5: PhaseGuide(
            phase_number=5,
            structure_type="operation",
            depth_guide=(
                "관람객 산출 근거, 동선 시뮬레이션, 인력 배치도, "
                "비상 대응 체계, 안전 관리 계획 포함."
            ),
            required_elements=[
                "관람객 규모 산출 (근거 포함)",
                "동선 시뮬레이션 (시간대별)",
                "인력 배치도 (역할별)",
                "비상 대응 매뉴얼",
                "품질 관리 체크리스트",
            ],
            slides_per_item=2,
            visual_requirements=["동선도", "인력 배치 다이어그램", "타임라인"],
        ),
    },

    industry_keywords=[
        "부스", "이벤트", "게임쇼", "IP", "세계관", "코스프레",
        "포토존", "미니게임", "굿즈", "팬미팅", "스테이지",
        "SNS 캠페인", "사전예약", "현장 이벤트", "관람객",
        "동선", "체험형", "인터랙션", "라이브", "스트리밍",
    ],

    default_persuasion="CHECK_POINT",

    action_plan_structure="pack",
    action_plan_items=[
        "BOOTH DESIGN PACK - 부스 외관/내부 설계, 3D 렌더링, 동선",
        "EVENT PACK - 메인 이벤트 프로그램, 미니게임, 체험 콘텐츠",
        "INTERACTION PACK - 참여형 콘텐츠, 포토존, AR/VR",
        "STAGE PACK - 스테이지 프로그램, MC/게스트, 타임테이블",
        "CAMPAIGN PACK - 사전/현장/사후 SNS 캠페인, 해시태그",
        "OPERATION PACK - 인력 배치, 동선 관리, 비상 대응",
    ],

    content_depth_guide="""
## 게임 이벤트 콘텐츠 깊이 지침

### 프로그램 딥다이브 (각 프로그램당 2-3 슬라이드)
1. **컨셉 슬라이드**: 프로그램명, 컨셉 설명, IP 연결점
2. **메커니즘 슬라이드**: 참여 방법, 보상 구조, 운영 플로우
3. **기대효과 슬라이드**: 예상 참여자 수, SNS 확산, KPI

### CHECK POINT 프레임워크
- Analytics: 데이터로 문제/기회 정의
- Solution: 구체적 프로그램/전략 제시
- Effect: 정량적 기대효과 (참여자 수, SNS 도달률 등)

### 필수 시각 요소
- 3D 부스 렌더링 (외관 + 내부)
- 평면도 + 관람객 동선도
- 프로그램 플로우차트
- 타임테이블 (시간대별)
- 캠페인 비주얼 예시
""",

    visual_guide="""
- 부스 디자인: 3D 렌더링 필수 (외관 2컷 + 내부 2컷)
- 동선도: 시간대별 관람객 흐름 시뮬레이션
- 프로그램: 각 프로그램별 비주얼 컨셉 이미지
- 캠페인: SNS 포스트 예시, 해시태그 비주얼
""",
)
