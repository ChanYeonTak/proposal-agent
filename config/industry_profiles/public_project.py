"""
공공 사업 산업 프로파일

공공 입찰 특화 - 법적 요구사항, 정책 정합성, 사회적 가치 강조.
"""

from .base_profile import IndustryProfile, PhaseGuide

PUBLIC_PROJECT_PROFILE = IndustryProfile(
    name="공공/정부 사업",
    description="공공기관 발주 사업, 정부 과제, 공공 입찰 제안서",
    industry_type="public",

    phase_guides={
        2: PhaseGuide(
            phase_number=2,
            structure_type="analytics",
            depth_guide="정책 환경 분석 + 법적 요구사항 + 선행 사례 분석 + 이해관계자 분석",
            required_elements=[
                "정책 환경 및 법적 근거",
                "선행 사례/유사 사업 분석",
                "이해관계자 분석",
                "지역 환경/현황 분석",
            ],
        ),
        4: PhaseGuide(
            phase_number=4,
            structure_type="section",
            depth_guide=(
                "과업 범위별 상세 추진 계획. 공공 평가 기준에 맞춰 "
                "기술성/경제성/사회적 가치 모두 포함."
            ),
            required_elements=[
                "과업별 추진 계획",
                "추진 일정 (WBS/Gantt)",
                "투입 인력 계획",
                "품질 관리 계획",
                "성과 지표 (KPI)",
                "사회적 가치 실현 방안",
            ],
            slides_per_item=2,
        ),
    },

    industry_keywords=[
        "과업", "사업비", "평가", "관리감독", "보고",
        "산출물", "납품", "검수", "유지보수", "보안",
        "정책", "법적 근거", "지역", "사회적 가치", "일자리",
    ],

    action_plan_structure="section",
    action_plan_items=[
        "과업 1: [RFP 과업 반영] - 상세 추진 계획",
        "과업 2: [RFP 과업 반영] - 상세 추진 계획",
        "품질 관리 - 검수/검증 체계",
        "보안 관리 - 보안 정책 준수",
        "성과 관리 - KPI + 모니터링",
    ],

    content_depth_guide="""
## 공공 사업 콘텐츠 깊이 지침

### 과업별 상세 (각 과업 2-3 슬라이드)
1. 과업 개요 + 추진 방법론
2. 상세 추진 계획 + 일정
3. 산출물 + 품질 기준

### 공공 필수 요소
- RFP 평가 기준과 직접 매칭
- 법적 근거/정책 정합성
- 사회적 가치 (일자리 창출, 지역 경제 등)
- 보안/개인정보보호 계획
""",
)
