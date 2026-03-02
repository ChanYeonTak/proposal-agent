"""
IT/시스템 산업 프로파일

모듈 구조 기반의 시스템 제안서.
"""

from .base_profile import IndustryProfile, PhaseGuide

IT_SYSTEM_PROFILE = IndustryProfile(
    name="IT/시스템 개발",
    description="IT 시스템 구축, 소프트웨어 개발, 플랫폼 구축 제안서",
    industry_type="it_system",

    phase_guides={
        2: PhaseGuide(
            phase_number=2,
            structure_type="analytics",
            depth_guide="현행 시스템 분석 + 문제점 도출 + To-Be 아키텍처 방향",
            required_elements=[
                "현행 시스템(As-Is) 분석",
                "문제점 및 개선 필요 사항",
                "기술 트렌드 분석",
                "벤치마킹",
            ],
        ),
        4: PhaseGuide(
            phase_number=4,
            structure_type="module",
            depth_guide=(
                "모듈 단위 분해. 각 모듈별 기능명세/아키텍처/화면설계 포함. "
                "To-Be 시스템 전체 아키텍처 + 모듈별 상세."
            ),
            required_elements=[
                "시스템 아키텍처 (To-Be)",
                "모듈별 기능 명세",
                "데이터 모델 설계",
                "화면 설계 (UI/UX)",
                "연동/인터페이스 설계",
                "개발 일정 (WBS)",
            ],
            slides_per_item=2,
            visual_requirements=["아키텍처 다이어그램", "ERD", "화면 와이어프레임", "WBS"],
        ),
    },

    industry_keywords=[
        "시스템", "플랫폼", "아키텍처", "모듈", "API",
        "데이터베이스", "클라우드", "보안", "인프라", "UI/UX",
        "마이그레이션", "연동", "테스트", "배포", "운영",
    ],

    action_plan_structure="module",
    action_plan_items=[
        "시스템 아키텍처 - To-Be 전체 설계",
        "핵심 모듈 개발 - 모듈별 기능 명세 + 화면 설계",
        "데이터 설계 - ERD + 데이터 흐름",
        "연동/인터페이스 - 외부 시스템 연동 설계",
        "테스트 전략 - 단위/통합/성능/보안 테스트",
        "이관/배포 - 마이그레이션 + 배포 계획",
    ],

    content_depth_guide="""
## IT/시스템 콘텐츠 깊이 지침

### 모듈별 상세 (각 모듈 2-3 슬라이드)
1. 모듈 개요 + 핵심 기능
2. 상세 기능 명세 + 화면 설계
3. 기술 스택 + 연동 방안

### 필수 시각 요소
- 시스템 아키텍처 다이어그램
- ERD (Entity Relationship Diagram)
- 화면 와이어프레임/목업
- WBS (Work Breakdown Structure)
""",
)
