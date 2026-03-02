"""
PPTX 검증 모듈 (v1.0)

병합·생성된 PPTX의 OOXML 구조를 검증합니다.

주요 컴포넌트:
    - PptxMergeValidator: Gamma PPTX 병합 후 구조적 무결성 검증
"""

from .pptx_merge_validator import PptxMergeValidator, ValidationResult, ValidationIssue

__all__ = ["PptxMergeValidator", "ValidationResult", "ValidationIssue"]
