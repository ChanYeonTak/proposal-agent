"""
IP 딥 리서치 스키마 (v4.0)

RFP 밖의 실제 데이터를 수집하고 구조화합니다.
브랜드 현황, 커뮤니티 동향, 경쟁사, 트렌드 등.
"""

from __future__ import annotations

from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class ConfidenceLevel(str, Enum):
    """데이터 신뢰도"""
    VERIFIED = "verified"           # 공식 출처 확인
    ESTIMATED = "estimated"         # 추정 (근거 있음)
    AI_GENERATED = "ai_generated"   # AI 생성 (검증 필요)


class DataPoint(BaseModel):
    """데이터 포인트 (출처 추적)"""
    value: str = ""
    source: str = ""                        # 출처 (URL, 공식 발표 등)
    confidence: ConfidenceLevel = ConfidenceLevel.ESTIMATED
    date: str = ""                          # 데이터 기준일


# ═══════════════════════════════════════════════════════════════
# 브랜드 데이터
# ═══════════════════════════════════════════════════════════════

class BrandData(BaseModel):
    """브랜드/IP 공식 데이터"""
    brand_name: str = ""
    company: str = ""                       # 제작사/퍼블리셔
    genre: str = ""                         # 장르/분류
    release_date: str = ""                  # 출시일
    platforms: List[str] = Field(default_factory=list)

    # 유저 규모
    mau: DataPoint = Field(default_factory=DataPoint)
    dau: DataPoint = Field(default_factory=DataPoint)
    total_downloads: DataPoint = Field(default_factory=DataPoint)
    app_store_ranking: DataPoint = Field(default_factory=DataPoint)

    # 소셜 미디어
    social_media: Dict[str, DataPoint] = Field(default_factory=dict)
    # {"instagram": DataPoint(value="50만 팔로워"), "youtube": DataPoint(...)}

    # 매출/시장
    revenue: DataPoint = Field(default_factory=DataPoint)
    market_share: DataPoint = Field(default_factory=DataPoint)

    # IP 특성
    key_characters: List[str] = Field(default_factory=list)
    ip_strengths: List[str] = Field(default_factory=list)
    brand_keywords: List[str] = Field(default_factory=list)

    # 최근 동향
    recent_updates: List[str] = Field(default_factory=list)
    upcoming_events: List[str] = Field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# 커뮤니티 인사이트
# ═══════════════════════════════════════════════════════════════

class CommunityInsight(BaseModel):
    """커뮤니티/유저 동향"""
    platform: str = ""                      # 커뮤니티 플랫폼 (디시, 레딧, X 등)
    active_users: DataPoint = Field(default_factory=DataPoint)

    # 감성 분석
    overall_sentiment: str = ""             # positive/neutral/negative
    sentiment_details: str = ""             # 감성 상세

    # 트렌드
    viral_topics: List[str] = Field(default_factory=list)
    popular_memes: List[str] = Field(default_factory=list)
    fan_content_trends: List[str] = Field(default_factory=list)

    # 유저 니즈
    user_demands: List[str] = Field(default_factory=list)
    pain_points: List[str] = Field(default_factory=list)
    wishlist: List[str] = Field(default_factory=list)

    # 인기 캐릭터/콘텐츠
    popular_characters: List[str] = Field(default_factory=list)
    popular_content_types: List[str] = Field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# 경쟁사 프로파일
# ═══════════════════════════════════════════════════════════════

class CompetitorEvent(BaseModel):
    """경쟁사 이벤트/행사"""
    name: str = ""
    date: str = ""
    scale: str = ""                         # 규모 (참여자 수, 부스 크기 등)
    highlights: List[str] = Field(default_factory=list)
    differentiators: List[str] = Field(default_factory=list)


class CompetitorProfile(BaseModel):
    """경쟁사 프로파일"""
    name: str = ""
    brand: str = ""
    recent_events: List[CompetitorEvent] = Field(default_factory=list)
    strengths: List[str] = Field(default_factory=list)
    weaknesses: List[str] = Field(default_factory=list)
    market_position: str = ""


# ═══════════════════════════════════════════════════════════════
# 협력사 후보
# ═══════════════════════════════════════════════════════════════

class CollaboratorCandidate(BaseModel):
    """잠재 협력사/인플루언서"""
    name: str = ""
    category: str = ""                      # cosplayer, influencer, artist, streamer, etc.
    platform: str = ""
    followers: DataPoint = Field(default_factory=DataPoint)
    relevance: str = ""                     # 관련성 설명
    past_collaborations: List[str] = Field(default_factory=list)
    estimated_cost: str = ""


# ═══════════════════════════════════════════════════════════════
# 트렌드
# ═══════════════════════════════════════════════════════════════

class IndustryTrend(BaseModel):
    """산업 트렌드"""
    trend_name: str = ""
    description: str = ""
    relevance_to_project: str = ""
    examples: List[str] = Field(default_factory=list)
    source: str = ""


# ═══════════════════════════════════════════════════════════════
# 통합 결과
# ═══════════════════════════════════════════════════════════════

class IPResearchResult(BaseModel):
    """
    IP 딥 리서치 통합 결과

    RFP 밖의 실제 데이터를 종합한 결과물입니다.
    content_generator가 이 데이터를 활용하여 더 구체적인 콘텐츠를 생성합니다.
    """
    # 대상 IP/브랜드
    target_brand: str = ""
    research_scope: str = ""                # 리서치 범위 설명

    # 핵심 데이터
    brand_data: BrandData = Field(default_factory=BrandData)
    community_insights: List[CommunityInsight] = Field(default_factory=list)
    competitor_profiles: List[CompetitorProfile] = Field(default_factory=list)
    collaborator_candidates: List[CollaboratorCandidate] = Field(default_factory=list)
    industry_trends: List[IndustryTrend] = Field(default_factory=list)

    # 전략적 인사이트 (AI 종합)
    strategic_insights: List[str] = Field(default_factory=list)
    differentiation_opportunities: List[str] = Field(default_factory=list)
    risk_factors: List[str] = Field(default_factory=list)

    # 데이터 품질
    data_quality_score: float = 0.0         # 0~1 (검증된 데이터 비율)
    verified_data_count: int = 0
    estimated_data_count: int = 0
    ai_generated_count: int = 0

    # 메타
    research_timestamp: str = ""
    search_queries_used: List[str] = Field(default_factory=list)
