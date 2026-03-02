"""
Think Tank SQLite 저장소 (v4.0)

레퍼런스 문서의 구조화된 데이터를 SQLite에 저장/검색합니다.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import List, Optional

from .models import (
    ReferenceDocument,
    DesignProfile,
    ContentPattern,
    ContentToneProfile,
    ProgramTemplate,
    SectionStructure,
    DocType,
    Industry,
    SearchResult,
)
from src.utils.logger import get_logger

logger = get_logger("think_tank_db")

# 기본 DB 경로
DEFAULT_DB_PATH = Path(__file__).parent / "references.db"


class ThinkTankDB:
    """
    Think Tank SQLite 저장소

    레퍼런스 문서를 저장하고 검색합니다.
    """

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = db_path or DEFAULT_DB_PATH
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        """DB 연결 반환"""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        return conn

    def _init_db(self):
        """DB 테이블 초기화"""
        conn = self._get_conn()
        try:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS references_doc (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT NOT NULL,
                    file_hash TEXT NOT NULL UNIQUE,
                    file_name TEXT NOT NULL,
                    file_size INTEGER DEFAULT 0,

                    doc_type TEXT DEFAULT 'other',
                    industry TEXT DEFAULT 'other',
                    project_type TEXT DEFAULT '',
                    won_bid INTEGER DEFAULT 0,

                    total_pages INTEGER DEFAULT 0,
                    sections_json TEXT DEFAULT '[]',
                    toc_json TEXT DEFAULT '[]',

                    design_profile_json TEXT DEFAULT '{}',
                    content_patterns_json TEXT DEFAULT '[]',
                    program_templates_json TEXT DEFAULT '[]',

                    full_text TEXT DEFAULT '',
                    summary TEXT DEFAULT '',

                    ingested_at TEXT NOT NULL,
                    tags_json TEXT DEFAULT '[]',
                    notes TEXT DEFAULT ''
                );

                CREATE INDEX IF NOT EXISTS idx_file_hash ON references_doc(file_hash);
                CREATE INDEX IF NOT EXISTS idx_doc_type ON references_doc(doc_type);
                CREATE INDEX IF NOT EXISTS idx_industry ON references_doc(industry);
                CREATE INDEX IF NOT EXISTS idx_project_type ON references_doc(project_type);
                CREATE INDEX IF NOT EXISTS idx_won_bid ON references_doc(won_bid);
            """)
            # content_tone_json 컬럼 마이그레이션 (v5.1)
            try:
                conn.execute(
                    "ALTER TABLE references_doc ADD COLUMN content_tone_json TEXT DEFAULT '{}'"
                )
                conn.commit()
                logger.info("DB 마이그레이션: content_tone_json 컬럼 추가")
            except sqlite3.OperationalError:
                pass  # 이미 존재
            conn.commit()
            logger.info(f"Think Tank DB 초기화: {self.db_path}")
        finally:
            conn.close()

    # ─── CRUD ──────────────────────────────────────────────────

    def exists(self, file_hash: str) -> bool:
        """파일 해시로 중복 확인"""
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT id FROM references_doc WHERE file_hash = ?",
                (file_hash,),
            ).fetchone()
            return row is not None
        finally:
            conn.close()

    def save_reference(self, doc: ReferenceDocument) -> int:
        """
        레퍼런스 문서 저장

        Returns:
            저장된 문서의 ID
        """
        conn = self._get_conn()
        try:
            cursor = conn.execute(
                """
                INSERT INTO references_doc (
                    file_path, file_hash, file_name, file_size,
                    doc_type, industry, project_type, won_bid,
                    total_pages, sections_json, toc_json,
                    design_profile_json, content_patterns_json, program_templates_json,
                    full_text, summary,
                    ingested_at, tags_json, notes,
                    content_tone_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    doc.file_path,
                    doc.file_hash,
                    doc.file_name,
                    doc.file_size,
                    doc.doc_type.value,
                    doc.industry.value,
                    doc.project_type,
                    1 if doc.won_bid else 0,
                    doc.total_pages,
                    json.dumps([s.model_dump() for s in doc.sections], ensure_ascii=False),
                    json.dumps(doc.table_of_contents, ensure_ascii=False),
                    doc.design_profile.model_dump_json(ensure_ascii=False) if doc.design_profile else "{}",
                    json.dumps([p.model_dump() for p in doc.content_patterns], ensure_ascii=False),
                    json.dumps([p.model_dump() for p in doc.program_templates], ensure_ascii=False),
                    doc.full_text,
                    doc.summary,
                    doc.ingested_at,
                    json.dumps(doc.tags, ensure_ascii=False),
                    doc.notes,
                    doc.content_tone.model_dump_json(ensure_ascii=False) if doc.content_tone else "{}",
                ),
            )
            conn.commit()
            doc_id = cursor.lastrowid
            logger.info(f"레퍼런스 저장: id={doc_id}, file={doc.file_name}")
            return doc_id
        finally:
            conn.close()

    def get_reference(self, doc_id: int) -> Optional[ReferenceDocument]:
        """ID로 레퍼런스 조회"""
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT * FROM references_doc WHERE id = ?",
                (doc_id,),
            ).fetchone()
            if row is None:
                return None
            return self._row_to_document(row)
        finally:
            conn.close()

    def get_by_hash(self, file_hash: str) -> Optional[ReferenceDocument]:
        """파일 해시로 조회"""
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT * FROM references_doc WHERE file_hash = ?",
                (file_hash,),
            ).fetchone()
            if row is None:
                return None
            return self._row_to_document(row)
        finally:
            conn.close()

    def search_by_type(
        self,
        doc_type: Optional[str] = None,
        industry: Optional[str] = None,
        project_type: Optional[str] = None,
        won_bid_only: bool = False,
        limit: int = 10,
    ) -> List[ReferenceDocument]:
        """유형별 검색"""
        conn = self._get_conn()
        try:
            conditions = []
            params = []

            if doc_type:
                conditions.append("doc_type = ?")
                params.append(doc_type)
            if industry:
                conditions.append("industry = ?")
                params.append(industry)
            if project_type:
                conditions.append("project_type = ?")
                params.append(project_type)
            if won_bid_only:
                conditions.append("won_bid = 1")

            where = " AND ".join(conditions) if conditions else "1=1"
            query = f"SELECT * FROM references_doc WHERE {where} ORDER BY ingested_at DESC LIMIT ?"
            params.append(limit)

            rows = conn.execute(query, params).fetchall()
            return [self._row_to_document(row) for row in rows]
        finally:
            conn.close()

    def list_all(self, limit: int = 50) -> List[ReferenceDocument]:
        """전체 목록 조회"""
        return self.search_by_type(limit=limit)

    def get_stats(self) -> dict:
        """DB 통계"""
        conn = self._get_conn()
        try:
            total = conn.execute("SELECT COUNT(*) FROM references_doc").fetchone()[0]
            won = conn.execute("SELECT COUNT(*) FROM references_doc WHERE won_bid = 1").fetchone()[0]
            types = conn.execute(
                "SELECT doc_type, COUNT(*) as cnt FROM references_doc GROUP BY doc_type"
            ).fetchall()
            industries = conn.execute(
                "SELECT industry, COUNT(*) as cnt FROM references_doc GROUP BY industry"
            ).fetchall()

            return {
                "total": total,
                "won_bid": won,
                "by_type": {row["doc_type"]: row["cnt"] for row in types},
                "by_industry": {row["industry"]: row["cnt"] for row in industries},
            }
        finally:
            conn.close()

    def update_content_tone(self, doc_id: int, tone: ContentToneProfile) -> bool:
        """content_tone 업데이트 (딥 분석 결과 저장)"""
        conn = self._get_conn()
        try:
            cursor = conn.execute(
                "UPDATE references_doc SET content_tone_json = ? WHERE id = ?",
                (tone.model_dump_json(ensure_ascii=False), doc_id),
            )
            conn.commit()
            if cursor.rowcount > 0:
                logger.info(f"ContentTone 업데이트: id={doc_id}")
            return cursor.rowcount > 0
        finally:
            conn.close()

    def delete_reference(self, doc_id: int) -> bool:
        """레퍼런스 삭제"""
        conn = self._get_conn()
        try:
            cursor = conn.execute("DELETE FROM references_doc WHERE id = ?", (doc_id,))
            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    # ─── 내부 헬퍼 ──────────────────────────────────────────────

    def _row_to_document(self, row: sqlite3.Row) -> ReferenceDocument:
        """DB 행 → ReferenceDocument 변환"""
        sections_raw = json.loads(row["sections_json"] or "[]")
        sections = [SectionStructure(**s) for s in sections_raw]

        design_raw = row["design_profile_json"] or "{}"
        design_profile = DesignProfile.model_validate_json(design_raw) if design_raw != "{}" else DesignProfile()

        patterns_raw = json.loads(row["content_patterns_json"] or "[]")
        content_patterns = [ContentPattern(**p) for p in patterns_raw]

        programs_raw = json.loads(row["program_templates_json"] or "[]")
        program_templates = [ProgramTemplate(**p) for p in programs_raw]

        tags = json.loads(row["tags_json"] or "[]")
        toc = json.loads(row["toc_json"] or "[]")

        # content_tone 로딩 (v5.1)
        content_tone_raw = row["content_tone_json"] if "content_tone_json" in row.keys() else "{}"
        content_tone = (
            ContentToneProfile.model_validate_json(content_tone_raw)
            if content_tone_raw and content_tone_raw != "{}"
            else ContentToneProfile()
        )

        return ReferenceDocument(
            id=row["id"],
            file_path=row["file_path"],
            file_hash=row["file_hash"],
            file_name=row["file_name"],
            file_size=row["file_size"],
            doc_type=DocType(row["doc_type"]),
            industry=Industry(row["industry"]),
            project_type=row["project_type"],
            won_bid=bool(row["won_bid"]),
            total_pages=row["total_pages"],
            sections=sections,
            table_of_contents=toc,
            design_profile=design_profile,
            content_patterns=content_patterns,
            program_templates=program_templates,
            full_text=row["full_text"],
            summary=row["summary"],
            ingested_at=row["ingested_at"],
            tags=tags,
            notes=row["notes"],
            content_tone=content_tone,
        )
