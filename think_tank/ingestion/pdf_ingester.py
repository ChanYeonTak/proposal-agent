"""
PDF 레퍼런스 입수

PDF 문서에서 텍스트, 섹션 구조, 페이지 수를 추출합니다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from ..models import (
    ReferenceDocument,
    SectionStructure,
    DocType,
    Industry,
)
from .dedup_checker import compute_file_hash
from src.utils.logger import get_logger

logger = get_logger("pdf_ingester")


class PDFIngester:
    """PDF 레퍼런스 입수"""

    def ingest(
        self,
        file_path: Path,
        doc_type: DocType = DocType.PROPOSAL,
        industry: Industry = Industry.OTHER,
        project_type: str = "",
        won_bid: bool = False,
        tags: Optional[List[str]] = None,
        notes: str = "",
    ) -> ReferenceDocument:
        """
        PDF 파일을 분석하여 ReferenceDocument 생성

        Args:
            file_path: PDF 파일 경로
            doc_type: 문서 유형
            industry: 산업 분류
            project_type: 프로젝트 유형
            won_bid: 수주 성공 여부
            tags: 태그
            notes: 메모

        Returns:
            ReferenceDocument: 추출된 레퍼런스 데이터
        """
        file_hash = compute_file_hash(file_path)

        # 텍스트 추출 (pypdf)
        text, total_pages = self._extract_text(file_path)

        # 섹션 구조 분석
        sections = self._analyze_sections(text)

        doc = ReferenceDocument(
            file_path=str(file_path.absolute()),
            file_hash=file_hash,
            file_name=file_path.name,
            file_size=file_path.stat().st_size,
            doc_type=doc_type,
            industry=industry,
            project_type=project_type,
            won_bid=won_bid,
            total_pages=total_pages,
            sections=sections,
            full_text=text[:100000],  # 최대 100K 문자
            tags=tags or [],
            notes=notes,
        )

        logger.info(f"PDF 입수 완료: {file_path.name} ({total_pages}p, {len(text)}자)")
        return doc

    def _extract_text(self, file_path: Path) -> tuple[str, int]:
        """PDF에서 텍스트 + 페이지 수 추출"""
        try:
            import pypdf
            reader = pypdf.PdfReader(str(file_path))
            total_pages = len(reader.pages)

            texts = []
            for page in reader.pages:
                page_text = page.extract_text() or ""
                texts.append(page_text)

            full_text = "\n\n".join(texts)

            # 텍스트가 거의 없으면 이미지 기반 PDF일 수 있음
            if len(full_text.strip()) < 100 and total_pages > 5:
                logger.warning(
                    f"텍스트가 거의 없는 PDF ({len(full_text.strip())}자) — "
                    f"이미지 기반 PDF일 수 있습니다"
                )

            return full_text, total_pages

        except Exception as e:
            logger.error(f"PDF 텍스트 추출 실패: {e}")
            # 페이지 수만이라도 추출 시도
            try:
                import pypdf
                reader = pypdf.PdfReader(str(file_path))
                return "", len(reader.pages)
            except Exception:
                return "", 0

    def _analyze_sections(self, text: str) -> List[SectionStructure]:
        """텍스트에서 섹션 구조 분석 (간단한 휴리스틱)"""
        if not text.strip():
            return []

        sections = []
        lines = text.split("\n")
        current_section = None
        current_lines = 0

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            # 간단한 섹션 감지 (숫자로 시작하는 큰 제목)
            is_section = (
                len(stripped) < 50
                and (
                    stripped.startswith(("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."))
                    or stripped.startswith(("I.", "II.", "III.", "IV.", "V."))
                    or stripped.startswith(("Phase", "PHASE", "PART", "Part", "Chapter"))
                    or stripped.isupper() and len(stripped) > 3
                )
            )

            if is_section:
                if current_section:
                    sections.append(SectionStructure(
                        name=current_section,
                        slide_count=max(1, current_lines // 20),  # 대략적 추정
                    ))
                current_section = stripped
                current_lines = 0
            else:
                current_lines += 1

        # 마지막 섹션
        if current_section:
            sections.append(SectionStructure(
                name=current_section,
                slide_count=max(1, current_lines // 20),
            ))

        return sections
