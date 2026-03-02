"""
범용화된 ContentToneAnalyzer 검증 테스트

1. 기존 game_event 레퍼런스가 깨지지 않는지 확인
2. 가상의 IT/공공 텍스트로 비게임 산업 분석 품질 확인
"""
import sys
import io
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from think_tank.db import ThinkTankDB
from think_tank.ingestion.content_tone_analyzer import ContentToneAnalyzer
from think_tank.design_brief import DesignBriefBuilder


def test_game_event_regression():
    """기존 game_event 분석이 깨지지 않는지 확인"""
    print("=" * 70)
    print("[1] 기존 game_event 레퍼런스 회귀 테스트")
    print("=" * 70)

    db = ThinkTankDB()
    analyzer = ContentToneAnalyzer(industry="game_event")

    # BD2 제안서 (ID=3) 로드
    doc = db.get_reference(3)
    if not doc:
        print("  ⚠ ID=3 문서 없음 — 건너뜀")
        return True

    tone = analyzer.analyze(doc.full_text, doc.file_name)
    print(f"  BD2 제안서: 톤={tone.emotional_tone_level}/5 IP={tone.ip_depth_score:.2f}")
    print(f"    캐릭터: {tone.ip_lore_terms[:5]}")
    print(f"    커뮤니티: {tone.ip_community_terms[:5]}")
    print(f"    규칙: {len(tone.tone_rules)}건")

    checks = {
        "감성톤 >= 4": tone.emotional_tone_level >= 4,
        "IP 깊이 >= 0.5": tone.ip_depth_score >= 0.5,
        "톤 규칙 >= 5": len(tone.tone_rules) >= 5,
    }
    for name, ok in checks.items():
        print(f"  {'✅' if ok else '❌'} {name}")
    return all(checks.values())


def test_it_system():
    """IT/시스템 제안서 가상 텍스트 분석"""
    print("\n" + "=" * 70)
    print("[2] IT/시스템 산업 가상 텍스트 분석")
    print("=" * 70)

    it_text = """
    스마트시티 통합 플랫폼 구축 제안서

    본 제안은 MSA 기반 클라우드 네이티브 아키텍처를 적용하여
    기존 모놀리식 시스템을 마이크로서비스로 전환합니다.

    Phase 1: Discovery Sprint — 현황 진단 및 요구사항 분석
    Phase 2: Design Sprint — 서비스 설계 및 API 명세
    Phase 3: Build Sprint — Kubernetes 기반 구축 및 CI/CD 파이프라인
    Phase 4: Launch — 무중단 마이그레이션 및 모니터링

    핵심 기술 스택:
    - 컨테이너 오케스트레이션: Kubernetes + Docker
    - API Gateway: Kong + gRPC
    - 데이터 파이프라인: Apache Kafka + ETL
    - 보안: 제로트러스트 아키텍처 + ISMS 인증
    - AI/ML: 자연어처리 기반 민원 자동 분류

    성과 지표:
    - 시스템 응답시간 50% 단축 (3초 → 1.5초)
    - 장애 복구시간 MTTR 30분 이내
    - DevOps 배포 주기: 월 1회 → 일 1회

    AWS 기반 멀티 AZ 구성으로 99.95% 가용성을 보장하며,
    Terraform을 활용한 IaC로 인프라 표준화를 달성합니다.

    애자일 방법론 기반 스프린트 단위 개발로 2주마다 중간 성과를 확인하며,
    DevSecOps 체계를 구축하여 보안을 개발 프로세스에 내재화합니다.
    """

    analyzer = ContentToneAnalyzer(industry="it_system")
    tone = analyzer.analyze(it_text, "IT_시스템_가상_제안서")

    print(f"  감성톤: {tone.emotional_tone_level}/5")
    print(f"  프레이밍: {tone.narrative_framing.style}")
    print(f"  도메인 깊이: {tone.ip_depth_score:.2f}")
    print(f"  로어 용어: {tone.ip_lore_terms[:10]}")
    print(f"  커뮤니티: {tone.ip_community_terms[:10]}")
    print(f"  네이밍 스타일: {tone.program_naming_style}")
    print(f"  네이밍 예시: {tone.program_naming_examples[:5]}")
    print(f"  규칙: {len(tone.tone_rules)}건")
    for i, rule in enumerate(tone.tone_rules, 1):
        print(f"    {i}. {rule}")

    checks = {
        "감성톤 <= 3 (기술적 문서)": tone.emotional_tone_level <= 3,
        "도메인 깊이 > 0 (IT 용어 감지)": tone.ip_depth_score > 0,
        "프레이밍 data_driven 또는 hybrid": tone.narrative_framing.style in ("data_driven", "hybrid"),
        "톤 규칙 생성됨": len(tone.tone_rules) >= 3,
    }
    for name, ok in checks.items():
        print(f"  {'✅' if ok else '❌'} {name}")
    return all(checks.values())


def test_public_sector():
    """공공 부문 제안서 가상 텍스트 분석"""
    print("\n" + "=" * 70)
    print("[3] 공공 부문 산업 가상 텍스트 분석")
    print("=" * 70)

    public_text = """
    디지털정부 혁신 마스터플랜 수립 용역 제안서

    국정과제 연계: 디지털 플랫폼 정부 실현
    민관협력 거버넌스를 통한 시민 체감형 서비스 혁신

    사업 배경:
    정부는 디지털정부 전환을 가속화하기 위해 공공데이터 개방,
    마이데이터 활성화, 규제샌드박스 확대 등 전자정부 혁신을 추진 중입니다.

    주요 추진 과제:
    제1차 디지털 전환 전략 — 행정서비스 디지털화
    제2차 데이터 기반 행정 — 정책 의사결정 지원
    제3차 시민 참여 플랫폼 — 공론화 및 정보공개

    성과 목표:
    - 주민참여율 30% 향상
    - 민원 처리시간 40% 단축
    - 시민 만족도 85점 이상 달성

    관련 법령: 전자정부법, 데이터3법, 공공데이터 관리지침
    행정절차법에 따른 규정 준수를 전제로 합니다.

    주민 여러분과 함께 만들어가는 체계적이고 지속가능한 디지털정부,
    그 비전을 실현하기 위해 검증된 전문성과 축적된 노하우로 제안합니다.
    """

    analyzer = ContentToneAnalyzer(industry="public")
    tone = analyzer.analyze(public_text, "공공_디지털정부_가상_제안서")

    print(f"  감성톤: {tone.emotional_tone_level}/5")
    print(f"  프레이밍: {tone.narrative_framing.style}")
    print(f"  도메인 깊이: {tone.ip_depth_score:.2f}")
    print(f"  도메인 용어: {tone.ip_lore_terms[:10]}")
    print(f"  네이밍 스타일: {tone.program_naming_style}")
    print(f"  규칙: {len(tone.tone_rules)}건")
    for i, rule in enumerate(tone.tone_rules, 1):
        print(f"    {i}. {rule}")

    checks = {
        "감성톤 <= 3 (공공 문서)": tone.emotional_tone_level <= 3,
        "도메인 깊이 > 0 (정책 용어 감지)": tone.ip_depth_score > 0,
        "톤 규칙 생성됨": len(tone.tone_rules) >= 3,
    }
    for name, ok in checks.items():
        print(f"  {'✅' if ok else '❌'} {name}")
    return all(checks.values())


def test_design_brief_game():
    """DesignBriefBuilder — game_event 통합 테스트"""
    print("\n" + "=" * 70)
    print("[4] DesignBriefBuilder game_event 통합 (회귀 확인)")
    print("=" * 70)

    builder = DesignBriefBuilder()
    brief = builder.build(project_type="event", industry="game_event", target_slides=70)
    ct = brief.content_tone

    print(f"  감성톤: {ct.get('emotional_tone_level')}/5")
    print(f"  소스: {ct.get('source_analysis', '')[:60]}")
    print(f"  IP 로어: {ct.get('ip_lore_terms', [])[:5]}")
    print(f"  규칙: {len(ct.get('tone_rules', []))}건")

    checks = {
        "딥 분석 경로 사용": "ContentToneProfile 병합" in ct.get("source_analysis", ""),
        "감성톤 >= 4": ct.get("emotional_tone_level", 0) >= 4,
    }
    for name, ok in checks.items():
        print(f"  {'✅' if ok else '❌'} {name}")
    return all(checks.values())


def main():
    results = {
        "game_event 회귀": test_game_event_regression(),
        "IT/시스템 분석": test_it_system(),
        "공공 부문 분석": test_public_sector(),
        "DesignBrief 통합": test_design_brief_game(),
    }

    print("\n" + "=" * 70)
    print("최종 결과")
    print("=" * 70)
    all_pass = True
    for name, ok in results.items():
        print(f"  {'✅' if ok else '❌'} {name}")
        if not ok:
            all_pass = False
    print(f"\n  {'✅ 모든 테스트 통과!' if all_pass else '❌ 일부 실패'}")
    return all_pass


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
