"""
마케팅/PR 산업 프로파일

기존 Impact-8 v3.6의 채널 구조를 유지하면서 강화.
- 채널별 전략 (Instagram, YouTube, TikTok 등)
- 콘텐츠 예시 포함
- 캠페인별 상세 기획
"""

from .base_profile import IndustryProfile, PhaseGuide

MARKETING_PR_PROFILE = IndustryProfile(
    name="마케팅/PR/소셜미디어",
    description="디지털 마케팅, PR, 소셜미디어 운영 제안서",
    industry_type="marketing_pr",

    phase_guides={
        2: PhaseGuide(
            phase_number=2,
            structure_type="analytics",
            depth_guide=(
                "시장 환경 + 타겟 분석 + 경쟁사 벤치마킹. "
                "소셜 리스닝 데이터, 트렌드 분석 포함."
            ),
            required_elements=[
                "시장 환경 분석 (규모, 성장률)",
                "타겟 오디언스 프로파일 (페르소나)",
                "경쟁사 소셜미디어 벤치마킹",
                "소셜 리스닝 인사이트",
            ],
            slides_per_item=2,
            visual_requirements=["데이터 차트", "페르소나 카드", "벤치마크 비교"],
        ),
        4: PhaseGuide(
            phase_number=4,
            structure_type="channel",
            depth_guide=(
                "채널별 전략 분해. 각 채널마다 콘텐츠 전략/포맷/빈도/KPI 포함. "
                "실제 포스팅 예시(비주얼+카피) 필수. 캠페인별 상세 기획."
            ),
            required_elements=[
                "Instagram 전략 (피드/스토리/릴스)",
                "YouTube 전략 (롱폼/숏폼/커뮤니티)",
                "TikTok/X/기타 채널 전략",
                "캠페인 기획 (최소 3개)",
                "콘텐츠 캘린더",
                "인플루언서 협업 전략",
            ],
            slides_per_item=2,
            visual_requirements=["포스팅 예시 비주얼", "콘텐츠 캘린더", "채널 플로우"],
            persuasion_framework="CEI",
        ),
    },

    industry_keywords=[
        "SNS", "소셜미디어", "인스타그램", "유튜브", "틱톡",
        "콘텐츠", "릴스", "숏폼", "롱폼", "인플루언서",
        "해시태그", "바이럴", "UGC", "캠페인", "브랜딩",
        "팔로워", "도달률", "참여율", "전환율",
    ],

    default_persuasion="CEI",

    action_plan_structure="channel",
    action_plan_items=[
        "Instagram 전략 - 피드/스토리/릴스 콘텐츠 + 포스팅 예시",
        "YouTube 전략 - 롱폼/숏폼/커뮤니티 + 콘텐츠 포맷",
        "TikTok/X 전략 - 숏폼 콘텐츠 + 트렌드 활용",
        "캠페인 기획 - 시즌별 캠페인 3-5개 상세",
        "인플루언서 협업 - 대상/방식/기대효과",
        "콘텐츠 캘린더 - 월별 콘텐츠 플랜",
    ],

    content_depth_guide="""
## 마케팅/PR 콘텐츠 깊이 지침

### 채널별 전략 (각 채널 2-3 슬라이드)
1. 채널 분석 + 전략 방향
2. 콘텐츠 포맷 + 빈도 + 실제 예시
3. KPI + 성과 측정 방법

### 캠페인 상세 기획 (각 캠페인 2-3 슬라이드)
1. 캠페인 컨셉 + 목표
2. 실행 계획 + 콘텐츠 예시
3. 예상 성과 + 산출 근거

### 필수 포함 요소
- 실제 포스팅 예시 (비주얼 설명 + 카피)
- 해시태그 전략
- 콘텐츠 캘린더
""",
)
