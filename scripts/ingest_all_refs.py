"""
/ref 폴더 전체 레퍼런스 인제스팅 스크립트

모든 PPTX 레퍼런스를 think_tank DB에 등록합니다.
중복 파일(SHA-256 해시 기준)은 자동 건너뜁니다.
"""

import sys
import time
import traceback
from pathlib import Path

# 프로젝트 루트 추가
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from think_tank.db import ThinkTankDB
from think_tank.models import DocType, Industry
from think_tank.ingestion.pptx_ingester import PPTXIngester
from think_tank.ingestion.dedup_checker import compute_file_hash

# ── 파일별 메타데이터 정의 ───────────────────────────────
FILES_META = [
    {
        "filename": "(GB)천년어게인_오프라인행사_운영계획안_250423_v12_14시버전.pptx",
        "doc_type": DocType.PLAN,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["천년어게인", "고블린게임즈", "오프라인행사", "운영계획안"],
        "notes": "천년어게인 오프라인 행사 운영계획안 (고블린게임즈). 2025년 4월.",
    },
    {
        "filename": "(HPN) SVWB_2025_Launch_Event_Proposal_250106.pptx",
        "doc_type": DocType.PROPOSAL,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["SVWB", "HPN", "론칭이벤트", "제안서"],
        "notes": "SVWB 2025 론칭 이벤트 제안서 (HPN). 2025년 1월.",
    },
    {
        "filename": "(HPN)마비노기 영웅전_15th Anniversary_운영매뉴얼(1224).pptx",
        "doc_type": DocType.MANUAL,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["마비노기영웅전", "HPN", "15주년", "운영매뉴얼", "Anniversary"],
        "notes": "마비노기 영웅전 15주년 기념 이벤트 운영매뉴얼 (HPN). 2024년 12월.",
    },
    {
        "filename": "(HPN)마비노기_원데이_클래스_시즌2_결과보고서_250205.pptx",
        "doc_type": DocType.REPORT,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["마비노기", "HPN", "원데이클래스", "시즌2", "결과보고서"],
        "notes": "마비노기 원데이 클래스 시즌2 결과보고서 (HPN). 2025년 2월.",
    },
    {
        "filename": "(HPN)마비노기_원데이_클래스_시즌2_운영계획안_250123.pptx",
        "doc_type": DocType.PLAN,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["마비노기", "HPN", "원데이클래스", "시즌2", "운영계획안"],
        "notes": "마비노기 원데이 클래스 시즌2 운영계획안 (HPN). 2025년 1월.",
    },
    {
        "filename": "250618_마비노기_21st_판타지파티_실행계획안_v12.pptx",
        "doc_type": DocType.PLAN,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["마비노기", "21주년", "판타지파티", "실행계획안"],
        "notes": "마비노기 21주년 판타지파티 실행계획안 v12. 2025년 6월. (1.7GB 대형 파일)",
    },
    {
        "filename": "가디스오더_사전 체험단 초청 이벤트_운영계획안_0806.pptx",
        "doc_type": DocType.PLAN,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["가디스오더", "사전체험단", "초청이벤트", "운영계획안"],
        "notes": "가디스오더 사전 체험단 초청 이벤트 운영계획안. 2025년 8월.",
    },
    {
        "filename": "브라운더스트2_2025_AGF_실행계획안_251201_전달용.pptx",
        "doc_type": DocType.PLAN,
        "industry": Industry.GAME_EVENT,
        "project_type": "event",
        "won_bid": True,
        "tags": ["브라운더스트2", "AGF", "실행계획안", "2025"],
        "notes": "브라운더스트2 2025 AGF 실행계획안 (BD2 제안서와 별도 문서). 2025년 12월.",
    },
]

# ── 인제스팅 실행 ─────────────────────────────────────────

def main():
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    ref_dir = ROOT / "ref"
    db = ThinkTankDB()
    ingester = PPTXIngester()

    print("=" * 70)
    print("Think Tank 레퍼런스 전체 인제스팅")
    print(f"ref 폴더: {ref_dir}")
    print(f"DB: {db.db_path}")
    print(f"처리 대상: {len(FILES_META)}개 파일")
    print("=" * 70)

    stats = {"success": 0, "skipped_dup": 0, "skipped_missing": 0, "failed": 0}

    for i, meta in enumerate(FILES_META, 1):
        filename = meta["filename"]
        file_path = ref_dir / filename
        print(f"\n[{i}/{len(FILES_META)}] {filename}")
        print(f"  유형: {meta['doc_type'].value} | 산업: {meta['industry'].value} | 수주: {meta['won_bid']}")

        # 파일 존재 확인
        if not file_path.exists():
            print(f"  ⚠ 파일 없음 — 건너뜀")
            stats["skipped_missing"] += 1
            continue

        size_mb = file_path.stat().st_size / (1024 * 1024)
        print(f"  크기: {size_mb:.1f} MB")

        # 해시 계산 + 중복 확인
        print(f"  해시 계산 중...", end=" ", flush=True)
        t0 = time.time()
        file_hash = compute_file_hash(file_path)
        hash_time = time.time() - t0
        print(f"완료 ({hash_time:.1f}s) — {file_hash[:16]}...")

        if db.exists(file_hash):
            print(f"  ⏭ 이미 DB에 존재 — 건너뜀")
            stats["skipped_dup"] += 1
            continue

        # 인제스팅 실행
        print(f"  📥 인제스팅 중...", flush=True)
        t0 = time.time()
        try:
            doc = ingester.ingest(
                file_path=file_path,
                doc_type=meta["doc_type"],
                industry=meta["industry"],
                project_type=meta["project_type"],
                won_bid=meta["won_bid"],
                tags=meta.get("tags", []),
                notes=meta.get("notes", ""),
            )

            ingest_time = time.time() - t0
            print(f"  ✅ 추출 완료 ({ingest_time:.1f}s)")
            print(f"     슬라이드: {doc.total_pages}p")
            print(f"     섹션: {len(doc.sections)}개")
            print(f"     색상: {len(doc.design_profile.colors)}개")
            print(f"     폰트: {len(doc.design_profile.fonts)}개")
            print(f"     콘텐츠 패턴: {len(doc.content_patterns)}개")
            print(f"     프로그램 템플릿: {len(doc.program_templates)}개")
            print(f"     텍스트 길이: {len(doc.full_text):,}자")

            # DB 저장
            doc_id = db.save_reference(doc)
            print(f"  💾 DB 저장: id={doc_id}")
            stats["success"] += 1

        except Exception as e:
            elapsed = time.time() - t0
            print(f"  ❌ 실패 ({elapsed:.1f}s): {e}")
            traceback.print_exc()
            stats["failed"] += 1

    # ── 결과 요약 ─────────────────────────────────────
    print("\n" + "=" * 70)
    print("인제스팅 완료")
    print(f"  ✅ 성공: {stats['success']}건")
    print(f"  ⏭ 중복 건너뜀: {stats['skipped_dup']}건")
    print(f"  ⚠ 파일 없음: {stats['skipped_missing']}건")
    print(f"  ❌ 실패: {stats['failed']}건")
    print("=" * 70)

    # 최종 DB 상태
    final_stats = db.get_stats()
    print(f"\nDB 최종 상태:")
    print(f"  전체: {final_stats['total']}건")
    print(f"  수주 성공: {final_stats['won_bid']}건")
    print(f"  유형별: {final_stats['by_type']}")
    print(f"  산업별: {final_stats['by_industry']}")


if __name__ == "__main__":
    main()
