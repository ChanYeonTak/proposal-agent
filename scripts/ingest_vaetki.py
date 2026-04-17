"""VAETKI Commerce 수주작을 싱크탱크에 인제스트.

수주 성공 레퍼런스로 표시하여 향후 유사 프로젝트에서 우선 참조되도록 함.
추가로 VAETKI 고유 디자인 특성(파스텔 그라디언트, 2색 브랜드 그라디언트)을
태그로 남겨 검색/추천 시 활용.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from think_tank.db import ThinkTankDB
from think_tank.ingestion.pptx_ingester import PPTXIngester
from think_tank.models import DocType, Industry


def main():
    pptx_path = Path(ROOT) / "input" / "2026 VAETKI" / \
        "[LAON] NC_AI_VAETKI_Commerce_오프라인_쇼케이스_운영대행_0414 (제출본).pptx"
    if not pptx_path.exists():
        print(f"[X] 파일 없음: {pptx_path}")
        return

    print(f"[INFO] 인제스트 시작: {pptx_path.name}")
    print(f"       크기: {pptx_path.stat().st_size / 1e6:.1f} MB")

    ingester = PPTXIngester()
    db = ThinkTankDB()

    # 중복 체크
    from think_tank.ingestion.dedup_checker import compute_file_hash
    file_hash = compute_file_hash(pptx_path)
    if db.exists(file_hash):
        print(f"[!] 이미 인제스트됨 (hash={file_hash[:12]}...)")
        existing = db.get_by_hash(file_hash)
        print(f"    기존 ID: {existing.id}, 제목: {existing.file_name}")
        return

    # 인제스트
    doc = ingester.ingest(
        file_path=pptx_path,
        doc_type=DocType.PROPOSAL,
        industry=Industry.MARKETING_PR,   # AI 테크 론칭 + 인플루언서 이벤트
        project_type="marketing_pr",
        won_bid=True,                       # ★ 실전 수주 성공
        tags=[
            "AI", "commerce", "VAETKI", "NC AI",
            "인플루언서", "오프라인 쇼케이스", "론칭",
            "파스텔", "그라디언트", "에디토리얼",
            "라이트 모드", "편집형", "프리미엄",
            "3-stop 그라디언트 배경",
            "2색 그라디언트 텍스트",
            "평행사변형 존 구분",
            "실전 수주 2026",
        ],
        notes=(
            "2026년 4월 NC AI VARCO Commerce 오프라인 쇼케이스 운영대행 "
            "제안서 실전 수주작. 46p, 파스텔 그라디언트 에디토리얼 라이트 "
            "스타일. 브랜드 2색 그라디언트(#6868F1→#DD6495)를 주요 "
            "헤드라인에 일관되게 적용. 섹션 디바이더는 여백 90% + 좌상단 "
            "검정 영문. TIME TABLE은 열별 파스텔 색 구분 셀. "
            "CREDENTIALS는 풀블리드 포토 + 하단 캡션. "
            "BUDGET은 심플 표 + 하단 검정 합계 바. "
            "대응 slide_kit 팔레트: 'vaetki_pastel'. "
            "대응 컴포넌트: bg_pastel_gradient, gradient_headline, "
            "PARALLELOGRAM_ZONE, slide_divider_light, PAGE_HEADER_LIGHT."
        ),
    )

    doc_id = db.save_reference(doc)
    print(f"\n[OK] 인제스트 완료")
    print(f"     ID: {doc_id}")
    print(f"     슬라이드: {doc.total_pages}")
    print(f"     섹션: {len(doc.sections)}")
    print(f"     수주 성공: {doc.won_bid}")
    print(f"     태그: {len(doc.tags)}개")

    # 디자인 프로파일 요약
    dp = doc.design_profile
    if dp.colors:
        print(f"\n  디자인 프로파일:")
        print(f"    상위 색상 5:")
        for c in dp.colors[:5]:
            print(f"      {c.hex} ({c.usage}, freq={c.frequency:.2f})")
    if dp.fonts:
        print(f"    상위 폰트 3:")
        for f in dp.fonts[:3]:
            print(f"      {f.name} {f.size_pt}pt ({f.usage})")

    # 통계
    stats = db.get_stats()
    print(f"\n  DB 전체 통계:")
    print(f"    총 레퍼런스: {stats.get('total', 0)}")
    print(f"    수주 성공: {stats.get('won_bid', 0)}")


if __name__ == "__main__":
    main()
