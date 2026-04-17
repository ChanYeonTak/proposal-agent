"""메커톤 (메이플스토리 월드) 제안서를 싱크탱크에 인제스트.

VAETKI와 함께 핵심 레퍼런스 2개 중 하나.
다크 에디토리얼 스타일 대표.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from think_tank.db import ThinkTankDB
from think_tank.ingestion.pptx_ingester import PPTXIngester
from think_tank.ingestion.dedup_checker import compute_file_hash
from think_tank.models import DocType, Industry


def main():
    pptx_path = Path(r"D:\2026\제안서\1. 2026 메이플스토리 월드 메커톤\1. 제안서\[LAON]2026_메이플스토리_월드_메커톤_제안서_0401.pptx")
    if not pptx_path.exists():
        print(f"[X] 없음: {pptx_path}")
        return

    print(f"[INFO] 인제스트: {pptx_path.name}")
    print(f"       크기: {pptx_path.stat().st_size / 1e6:.1f} MB")

    db = ThinkTankDB()
    file_hash = compute_file_hash(pptx_path)
    if db.exists(file_hash):
        existing = db.get_by_hash(file_hash)
        print(f"[!] 이미 인제스트됨 (id={existing.id})")
        return

    ingester = PPTXIngester()
    doc = ingester.ingest(
        file_path=pptx_path,
        doc_type=DocType.PROPOSAL,
        industry=Industry.GAME_EVENT,
        project_type="event",
        won_bid=True,
        tags=[
            "게임", "메이플스토리", "MSW", "해커톤", "메커톤",
            "넥슨", "NYPC", "청소년 창작",
            "다크 에디토리얼", "네이비 배경", "네온 사이언 악센트",
            "섹션 디바이더 여백", "3-pillar 포토카드",
            "원형 사진 타임라인", "쉐브론 연결자",
            "실전 수주 2026",
        ],
        notes=(
            "2026년 4월 넥슨 메이플스토리 월드 메커톤 운영 대행 제안서 "
            "실전 수주작. 65p, 다크 에디토리얼 스타일. "
            "딥 네이비 배경(#1C1F28) + 사이언 eyebrow(#66FFFF) + "
            "퍼플 악센트(#5F70FC). 섹션 디바이더는 여백 90% + 거대 영문. "
            "콘텐츠 워크호스는 PHOTO_CARD_TRIO 3-컬럼. STAT_ROW로 "
            "큰 수치(160명/48H/50+) 강조. CIRCULAR_PHOTO_FLOW로 "
            "이벤트 타임라인. DATA_TABLE_DARK로 상세 일정. "
            "대응 slide_kit 팔레트: 'editorial_dark'. "
            "대응 컴포넌트: slide_divider_hero, PAGE_HEADER, "
            "PHOTO_CARD_TRIO, STAT_ROW_HERO, CIRCULAR_PHOTO_FLOW."
        ),
    )
    doc_id = db.save_reference(doc)
    print(f"\n[OK] 완료 id={doc_id}, 슬라이드={doc.total_pages}, "
           f"섹션={len(doc.sections)}, 태그={len(doc.tags)}")

    stats = db.get_stats()
    print(f"  DB 총 {stats.get('total')}건, 수주 {stats.get('won_bid')}건")


if __name__ == "__main__":
    main()
