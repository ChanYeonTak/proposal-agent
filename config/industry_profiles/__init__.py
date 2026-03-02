"""
산업별 프로파일 (v4.0)

각 산업/프로젝트 유형에 맞는 콘텐츠 구조, 깊이, 키워드를 정의합니다.
"""

from .base_profile import IndustryProfile, get_industry_profile, list_profiles

__all__ = ["IndustryProfile", "get_industry_profile", "list_profiles"]
