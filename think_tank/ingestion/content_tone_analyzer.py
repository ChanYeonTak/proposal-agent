"""
콘텐츠 톤 딥 분석기 (v2.0 — 범용)

레퍼런스 문서의 텍스트를 심도 있게 분석하여
방향성, 어휘, 감각적 표현, 도메인 깊이감, 네이밍 패턴을 추출합니다.

v2.0: 게임 이벤트 전용 → 전 산업 범용 분석으로 확장
- INDUSTRY_VOCAB: 산업별 전문 어휘 사전
- DOMAIN_DEPTH_MARKERS: 산업별 도메인 깊이 지표
- PROGRAM_NAME_PATTERNS: 산업별 프로그램 네이밍 패턴
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List, Optional, Tuple

from ..models import ContentToneProfile, NarrativeFraming
from src.utils.logger import get_logger

logger = get_logger("content_tone_analyzer")


# ═══════════════════════════════════════════════════════════════
# 분석용 키워드 사전
# ═══════════════════════════════════════════════════════════════

# ── 범용 감성 표현 사전 (전 산업 공통 + 산업 특화) ──
EMOTIONAL_WORDS = {
    # 흥분/기대 (범용 + 엔터/게임)
    "excitement": [
        "폭발", "극대화", "압도", "몰입", "열광", "환호", "도파민",
        "짜릿", "전율", "소름", "심장", "터지", "미친", "대박",
        "최고", "레전드", "역대급", "신화", "쾌거", "기적",
        "혁신", "획기적", "파격", "돌파구", "비약", "약진",
    ],
    # 감동/서사 (범용)
    "narrative": [
        "이야기", "스토리", "여정", "모험", "전설", "서사", "운명",
        "시작", "개막", "탄생", "기다림", "꿈", "희망", "판타지",
        "동화", "마법", "세계관", "로어", "에픽",
        "비전", "미래", "도약", "전환점", "새 시대", "패러다임",
    ],
    # 감각/묘사 (범용)
    "sensory": [
        "지지직", "화면", "백색광", "맹렬", "엄습", "포근", "낡은",
        "불협화음", "노이즈", "불안감", "황폐", "훈훈", "어둠",
        "빛", "눈부신", "찬란", "몽환", "퇴폐", "아름다움",
        "생생한", "직관적", "선명한", "감각적",
    ],
    # 커뮤니티/소속감 (범용)
    "community": [
        "팬", "유저", "커뮤니티", "팬덤", "우리", "함께", "만남",
        "교감", "소통", "참여", "축하", "축제", "파티", "기념",
        "시민", "주민", "파트너", "동반자", "협력", "상생",
    ],
    # 신뢰/전문성 (공공/컨설팅/IT 특화)
    "authority": [
        "체계적", "검증된", "축적된", "선도적", "독보적",
        "일관된", "안정적", "최적화", "효율적", "지속가능",
        "전문성", "노하우", "레퍼런스", "트랙레코드",
    ],
}

# ── 산업별 전문 어휘 사전 ──
INDUSTRY_VOCAB = {
    "game_event": {
        "mechanic": [
            "가챠", "gacha", "뽑기", "리세", "리세마라", "파밍", "레벨업",
            "티어", "랭킹", "PVP", "PVE", "레이드", "던전", "퀘스트",
            "스킬", "버프", "너프", "밸런스", "메타", "빌드",
            "쿨타임", "시너지", "콤보", "크리티컬",
        ],
        "fan_language": [
            "갓겜", "쁘더", "쁘린이", "커피콩", "최애", "덕질",
            "입덕", "탈덕", "성덕", "굿즈", "MD", "컬렉터",
            "코스프레", "코스어", "2차 창작", "팬아트", "동인",
            "짤", "밈", "바이럴", "떡상", "역주행",
        ],
        "event_ops": [
            "부스", "스테이지", "무대", "LED", "트러스", "리깅",
            "동선", "시공", "세팅", "철거", "리허설", "콘솔",
            "인포데스크", "포토존", "체험존", "게임존", "굿즈존",
            "타임테이블", "럭키드로우", "경품", "스탬프 랠리",
        ],
    },
    "marketing_pr": {
        "digital": [
            "SEO", "SEM", "CPC", "CPM", "CTR", "CVR", "ROAS",
            "리타겟팅", "퍼포먼스", "그로스해킹", "퍼널", "전환율",
            "A/B 테스트", "어트리뷰션", "LTV", "CAC",
        ],
        "branding": [
            "브랜드 아이덴티티", "톤앤매너", "퍼소나", "포지셔닝",
            "브랜드 에쿼티", "USP", "tagline", "슬로건",
            "브랜딩", "리브랜딩", "BI", "CI",
        ],
        "content": [
            "콘텐츠 마케팅", "네이티브 광고", "인플루언서",
            "UGC", "숏폼", "릴스", "틱톡", "유튜브",
            "바이럴", "밈", "해시태그", "SNS", "KOL",
        ],
        "metrics": [
            "MAU", "DAU", "PV", "UV", "팔로워",
            "인게이지먼트", "도달률", "노출수", "impression",
        ],
    },
    "event": {
        "operations": [
            "동선", "시공", "철거", "세팅", "리허설",
            "스테이지", "무대", "LED", "트러스", "리깅",
            "케이터링", "이벤트 MC", "사회자",
        ],
        "experience": [
            "포토존", "체험존", "전시존", "네트워킹",
            "인터랙티브", "스탬프 랠리", "럭키드로우",
            "팝업", "팝업스토어", "인스타그래머블",
        ],
    },
    "it_system": {
        "architecture": [
            "MSA", "API", "REST", "GraphQL", "gRPC",
            "클라우드", "SaaS", "PaaS", "IaaS", "온프레미스",
            "컨테이너", "쿠버네티스", "도커", "서버리스",
            "마이크로서비스", "모놀리식", "아키텍처",
        ],
        "methodology": [
            "애자일", "스프린트", "스크럼", "칸반", "DevOps",
            "CI/CD", "TDD", "리팩토링", "코드리뷰",
            "요구사항", "설계", "테스트", "배포", "운영",
        ],
        "data_ai": [
            "AI", "ML", "딥러닝", "자연어처리", "NLP",
            "빅데이터", "데이터레이크", "ETL", "BI",
            "대시보드", "분석", "예측", "모델링",
        ],
        "security": [
            "제로트러스트", "ISMS", "ISO27001", "취약점",
            "침해사고", "보안관제", "접근통제", "암호화",
        ],
    },
    "public": {
        "policy": [
            "민관협력", "PPP", "거버넌스", "규제샌드박스",
            "행정혁신", "디지털정부", "전자정부",
            "국정과제", "정책목표", "성과지표",
        ],
        "citizen": [
            "주민참여", "시민체감", "공론화", "정보공개",
            "민원", "공공서비스", "복지", "접근성",
        ],
        "regulation": [
            "관련 법령", "조례", "규정", "지침", "가이드라인",
            "행정절차", "입찰공고", "계약조건",
        ],
    },
    "consulting": {
        "framework": [
            "SWOT", "PESTLE", "밸류체인", "BPR", "RPA",
            "DX", "디지털전환", "이해관계자", "stakeholder",
            "로드맵", "마일스톤", "KPI", "OKR",
        ],
        "analysis": [
            "Gap 분석", "벤치마크", "As-Is", "To-Be",
            "진단", "현황분석", "개선방안", "최적화",
            "TCO", "ROI", "BEP", "IRR",
        ],
        "delivery": [
            "워크숍", "인터뷰", "FGI", "서베이",
            "산출물", "중간보고", "최종보고", "이행계획",
        ],
    },
    "finance": {
        "product": [
            "펀드", "ETF", "보험", "대출", "금리",
            "포트폴리오", "자산배분", "리밸런싱",
        ],
        "regulation": [
            "금융위", "금감원", "바젤", "IFRS", "AML",
            "KYC", "준법감시", "리스크관리",
        ],
    },
    "education": {
        "pedagogy": [
            "커리큘럼", "학습설계", "LMS", "EdTech",
            "플립드러닝", "PBL", "블렌디드", "마이크로러닝",
        ],
        "assessment": [
            "성취도", "역량평가", "루브릭", "포트폴리오",
            "수행평가", "형성평가", "총괄평가",
        ],
    },
    "healthcare": {
        "clinical": [
            "EMR", "EHR", "PACS", "CDSS", "원격의료",
            "디지털헬스", "바이오마커", "임상시험",
        ],
        "regulation": [
            "GMP", "GCP", "식약처", "FDA", "CE마킹",
            "의료기기", "인허가", "규제과학",
        ],
    },
}

# 하위호환: 기존 코드에서 GAME_INDUSTRY_VOCAB를 참조하는 경우 대비
GAME_INDUSTRY_VOCAB = INDUSTRY_VOCAB.get("game_event", {})

# ── 산업별 도메인 깊이 지표 ──
DOMAIN_DEPTH_MARKERS = {
    "game_event": {
        "proper_nouns": [
            "빌헬미나", "올리비에", "모르페아", "에클립스", "벤타나",
            "단아", "사냐", "네반", "에린", "밀레시안",
            "글루피", "베일라", "아스칼론",
        ],
        "domain_terms": [
            "세계관", "로어", "에린", "판타지", "왕국",
            "동화", "잔혹 동화", "퇴폐", "운명", "마법",
            "문파", "무림", "무공", "강호",
        ],
        "community_slang": [
            "쁘더", "쁘린이", "커피콩", "도파민", "갓겜",
            "역주행", "떡상", "인게임", "컷신",
        ],
    },
    "marketing_pr": {
        "proper_nouns": [],  # RFP/브랜드별 동적 추출
        "domain_terms": [
            "퍼포먼스 마케팅", "브랜드 에쿼티", "그로스해킹",
            "옴니채널", "터치포인트", "고객여정", "페르소나",
        ],
        "community_slang": [
            "인스타그래머블", "갓생", "핫플", "인생샷",
            "바이럴", "밈", "챌린지", "OOTD",
        ],
    },
    "it_system": {
        "proper_nouns": [
            "AWS", "Azure", "GCP", "Kubernetes", "Docker",
            "Terraform", "Jenkins", "GitHub", "Jira",
        ],
        "domain_terms": [
            "마이크로서비스", "컨테이너", "오케스트레이션",
            "서버리스", "이벤트드리븐", "CQRS", "사가패턴",
            "제로트러스트", "DevSecOps",
        ],
        "community_slang": [],
    },
    "public": {
        "proper_nouns": [],  # 발주처/사업명별 동적 추출
        "domain_terms": [
            "거버넌스", "규제샌드박스", "디지털정부",
            "전자정부", "국정과제", "민관협력",
            "데이터3법", "공공데이터", "마이데이터",
        ],
        "community_slang": [],
    },
    "consulting": {
        "proper_nouns": [],
        "domain_terms": [
            "밸류체인", "벤치마크", "디지털전환",
            "BPR", "RPA", "프로세스 마이닝",
            "체인지 매니지먼트", "거버넌스",
        ],
        "community_slang": [],
    },
    "event": {
        "proper_nouns": [],
        "domain_terms": [
            "인스타그래머블", "네트워킹", "인터랙티브",
            "이머시브", "팝업", "컨벤션", "MICE",
        ],
        "community_slang": [],
    },
    "finance": {
        "proper_nouns": [],
        "domain_terms": [
            "포트폴리오", "리밸런싱", "알파", "베타",
            "AML", "KYC", "RegTech", "핀테크",
        ],
        "community_slang": [],
    },
    "education": {
        "proper_nouns": [],
        "domain_terms": [
            "플립드러닝", "PBL", "EdTech", "LMS",
            "마이크로러닝", "역량기반교육",
        ],
        "community_slang": [],
    },
    "healthcare": {
        "proper_nouns": [],
        "domain_terms": [
            "디지털헬스", "원격의료", "CDSS",
            "바이오마커", "규제과학", "RWE",
        ],
        "community_slang": [],
    },
}

# 하위호환
IP_DEPTH_MARKERS = DOMAIN_DEPTH_MARKERS.get("game_event", {})

# ── 산업별 프로그램 네이밍 패턴 ──
PROGRAM_NAME_PATTERNS_BY_INDUSTRY = {
    "game_event": [
        r"(?:프로즌|겨울왕국|오로라|글루피|에린|판타지)\s*(?:스테이지|스토어|존|데이트|퀸|도서관|파티|캠핑)",
        r"[A-Z][A-Za-z]+\s+(?:존|스토어|스테이지|스팟|데스크|프로그램|이벤트)",
        r"(?:모닥불정령|에린 연대기|뒤틀린 동화|도파민 러쉬|대유쾌마운틴)",
        r"(?:Match|Battle|Challenge|Quest|Rush|Festival|Party)\s+[Oo]f\s+\w+",
    ],
    "it_system": [
        r"Phase\s+\d+[:\s]+[A-Za-z\s]{3,25}",
        r"Sprint\s+\d+",
        r"[A-Za-z]+\s+모듈",
        r"[가-힣]+\s+(?:시스템|플랫폼|모듈|엔진|서버|DB)",
    ],
    "public": [
        r"[가-힣]+\s*(?:사업|프로젝트|프로그램|위원회|협의회)",
        r"제\d+[차기]\s+[가-힣]+",
    ],
    "consulting": [
        r"(?:워크숍|워크샵)\s+\d+",
        r"[A-Za-z]+\s+(?:Assessment|Workshop|Sprint)",
        r"(?:진단|분석|설계|이행)\s+단계",
    ],
    "marketing_pr": [
        r"[A-Za-z]+\s+(?:캠페인|챌린지|이벤트|프로모션)",
        r"시즌\s*\d+",
        r"[가-힣]+\s+(?:프로젝트|캠페인|이벤트)",
    ],
    "event": [
        r"[A-Za-z]+\s+(?:존|스테이지|스토어|스팟|프로그램)",
        r"[가-힣]+\s+(?:존|스테이지|체험|이벤트|프로그램)",
    ],
}

# 하위호환 (기존 PROGRAM_NAME_PATTERNS를 참조하는 코드 대비)
PROGRAM_NAME_PATTERNS = PROGRAM_NAME_PATTERNS_BY_INDUSTRY.get("game_event", [])

# ── 공식/전문 산업군 (감성 점수 보정 대상) ──
FORMAL_INDUSTRIES = {
    "public", "it_system", "consulting", "finance", "healthcare", "education",
}

# ── 감성어처럼 보이지만 공식 문서에서는 표준 어휘인 단어들 ──
# 이 단어들은 FORMAL_INDUSTRIES에서 감성 가중치를 70% 할인
FORMAL_CONTEXT_WORDS = {
    # excitement 카테고리이지만 정책/기술 문서의 표준 어휘
    "혁신", "획기적", "파격", "돌파구", "비약", "약진",
    # narrative 카테고리이지만 전략 문서의 표준 어휘
    "비전", "미래", "도약", "전환점", "새 시대", "패러다임",
    # community 카테고리이지만 공공/컨설팅의 표준 어휘
    "시민", "주민", "참여", "협력", "상생",
}

# 전략 표현 패턴 (범용)
STRATEGY_PHRASES = [
    # 영문 전략 키워드
    r"[A-Z][A-Z\s]{3,30}",
    # "~하는 전략" 패턴
    r"[가-힣]+(?:을|를)\s+(?:극대화|폭발|강화|확대|확보|선점|혁신|개선|달성)(?:하는|시키는)\s+전략",
    # "~의 핵심" 패턴
    r"[가-힣]+(?:의|만의)\s+(?:핵심|강점|매력|차별화|경쟁력|우위|전문성)",
    # "~ 기반 ~" 패턴 (IT/컨설팅)
    r"[가-힣A-Za-z]+\s+기반\s+[가-힣]+(?:전략|체계|시스템|플랫폼)",
]


class ContentToneAnalyzer:
    """
    레퍼런스 텍스트 딥 분석기 (v2.0 범용)

    단순 키워드 카운팅이 아니라,
    실제 문장/구절 단위에서 스타일 패턴을 추출합니다.

    산업(industry) 파라미터에 따라 해당 산업의 전문 어휘 사전,
    도메인 깊이 지표, 프로그램 네이밍 패턴을 자동 선택합니다.
    """

    def __init__(self, industry: str = ""):
        """
        Args:
            industry: 산업 분류 (game_event, marketing_pr, it_system, public, consulting 등)
                     빈 문자열이면 전체 산업 사전을 병합하여 사용
        """
        self.industry = industry

        # 산업별 어휘 사전 선택
        if industry and industry in INDUSTRY_VOCAB:
            self._vocab = INDUSTRY_VOCAB[industry]
        else:
            # 전 산업 병합 (산업 미지정 시)
            self._vocab = {}
            for ind_vocab in INDUSTRY_VOCAB.values():
                for cat, words in ind_vocab.items():
                    if cat not in self._vocab:
                        self._vocab[cat] = []
                    self._vocab[cat].extend(words)

        # 산업별 도메인 깊이 지표 선택
        if industry and industry in DOMAIN_DEPTH_MARKERS:
            self._depth_markers = DOMAIN_DEPTH_MARKERS[industry]
        else:
            # game_event 기본 (하위호환) + 추가로 공통 항목
            self._depth_markers = DOMAIN_DEPTH_MARKERS.get("game_event", {
                "proper_nouns": [], "domain_terms": [], "community_slang": [],
            })

        # 산업별 프로그램 네이밍 패턴
        if industry and industry in PROGRAM_NAME_PATTERNS_BY_INDUSTRY:
            self._naming_patterns = PROGRAM_NAME_PATTERNS_BY_INDUSTRY[industry]
        else:
            self._naming_patterns = PROGRAM_NAME_PATTERNS

    def analyze(self, full_text: str, file_name: str = "") -> ContentToneProfile:
        """
        전체 텍스트를 심도 있게 분석하여 ContentToneProfile 생성

        Args:
            full_text: 문서 전체 텍스트
            file_name: 파일명 (로깅용)

        Returns:
            ContentToneProfile: 분석 결과
        """
        if not full_text or len(full_text) < 100:
            logger.warning(f"텍스트 부족: {file_name} ({len(full_text or '')}자)")
            return ContentToneProfile()

        # 1. 어휘 분석
        vocab = self._analyze_vocabulary(full_text)

        # 2. 감성 표현 분석
        emotional = self._analyze_emotional_patterns(full_text)

        # 3. IP 깊이 분석
        ip_depth = self._analyze_ip_depth(full_text)

        # 4. 프로그램 네이밍 분석
        naming = self._analyze_program_naming(full_text)

        # 5. 내러티브/프레이밍 분석
        framing = self._analyze_narrative_framing(full_text)

        # 6. 텍스트 밀도/스타일 분석
        density = self._analyze_text_density(full_text)

        # 7. Win Theme 스타일 분석
        win_theme = self._analyze_win_theme_style(full_text)

        # 8. 종합 톤 규칙 생성
        tone_rules = self._generate_tone_rules(
            vocab, emotional, ip_depth, naming, framing, density, win_theme
        )

        # 9. 핵심 표현 사전 생성 (실제 예시 기반)
        vocabulary_examples = self._extract_vocabulary_examples(full_text)

        profile = ContentToneProfile(
            emotional_tone_level=emotional["tone_level"],
            narrative_framing=NarrativeFraming(
                style=framing["style"],
                core_metaphor=framing["core_metaphor"],
                entry_hook=framing["entry_hook"],
                recurring_motif=framing["recurring_motif"],
                description=framing["description"],
            ),
            ip_depth_score=ip_depth["score"],
            ip_character_count=ip_depth["character_count"],
            ip_lore_terms=ip_depth["lore_terms"],
            ip_community_terms=ip_depth["community_terms"],
            program_naming_style=naming["style"],
            program_naming_examples=naming["examples"],
            win_theme_style=win_theme["style"],
            win_theme_examples=win_theme["examples"],
            text_density_style=density["style"],
            image_slide_ratio=density.get("image_ratio", 0.0),
            text_only_ratio=density.get("text_only_ratio", 0.0),
            tone_rules=tone_rules,
            source_analysis=(
                f"딥 분석 완료: {file_name} ({len(full_text):,}자). "
                f"감성어 {emotional['total_count']}건, "
                f"IP용어 {ip_depth['total_terms']}건, "
                f"프로그램명 {len(naming['examples'])}건, "
                f"어휘풍부도 {vocab['richness']:.2f}"
            ),
        )

        logger.info(
            f"딥 분석: {file_name} | "
            f"tone={emotional['tone_level']}/5, "
            f"ip={ip_depth['score']:.2f}, "
            f"rules={len(tone_rules)}"
        )
        return profile

    # ═══════════════════════════════════════════════════════════
    # 1. 어휘 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_vocabulary(self, text: str) -> Dict:
        """어휘 풍부도, 전문성, 감각적 표현 비중 분석"""
        # 한글 단어 추출
        words = re.findall(r'[가-힣]{2,}', text)
        total_words = len(words)

        if total_words == 0:
            return {"richness": 0, "total_words": 0, "unique_ratio": 0}

        unique_words = len(set(words))
        richness = unique_words / total_words  # 어휘 다양성 지수

        # 전문 어휘 비율 (산업별 사전 활용)
        domain_terms = 0
        for cat_words in self._vocab.values():
            for w in cat_words:
                domain_terms += text.lower().count(w.lower())

        emotional_count = 0
        for cat_words in EMOTIONAL_WORDS.values():
            for w in cat_words:
                emotional_count += text.count(w)

        return {
            "richness": richness,
            "total_words": total_words,
            "unique_ratio": unique_words / total_words,
            "domain_term_density": domain_terms / max(total_words, 1),
            "emotional_density": emotional_count / max(total_words, 1),
        }

    # ═══════════════════════════════════════════════════════════
    # 2. 감성 표현 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_emotional_patterns(self, text: str) -> Dict:
        """감성적 표현 패턴 분석 → 톤 레벨 (1~5) 결정

        v2.1: 산업별 감성 보정
        - authority 카테고리는 전문성 지표이므로 감성 밀도에서 제외
        - FORMAL_INDUSTRIES에서는 FORMAL_CONTEXT_WORDS 가중치를 70% 할인
        - authority 밀도가 높으면 톤 레벨 하향 (전문적 톤 = 감성 억제)
        """
        category_counts = {}
        total_count = 0
        found_phrases = []

        for cat, words in EMOTIONAL_WORDS.items():
            count = 0
            for w in words:
                matches = text.count(w)
                count += matches
                if matches > 0:
                    found_phrases.append((w, cat, matches))
            category_counts[cat] = count
            total_count += count

        # 감성 문장 추출 (실제 예시)
        emotional_sentences = self._extract_emotional_sentences(text)

        # ── 톤 레벨 결정 (v2.1 산업별 보정) ──
        text_len = len(text)

        # 1) authority 카테고리는 감성이 아닌 전문성 → 감성 밀도에서 제외
        authority_count = category_counts.get("authority", 0)
        emotional_count = total_count - authority_count

        # 2) FORMAL_INDUSTRIES에서는 공식 어휘(FORMAL_CONTEXT_WORDS) 가중치 할인
        formal_discount = 0
        is_formal = self.industry in FORMAL_INDUSTRIES
        if is_formal:
            for w in FORMAL_CONTEXT_WORDS:
                formal_discount += text.count(w)
            # 공식 어휘의 70%를 감성 카운트에서 차감 (30%만 감성으로 인정)
            emotional_count = max(emotional_count - formal_discount * 0.7, 0)

        density = emotional_count / max(text_len / 1000, 1)  # 1000자당 유효 감성어 빈도

        if density > 8:
            tone_level = 5  # 풀 내러티브
        elif density > 5:
            tone_level = 4  # 감성+전문
        elif density > 3:
            tone_level = 3  # 중립
        elif density > 1:
            tone_level = 2  # 간결 전문가
        else:
            tone_level = 1  # 사무적/기능적

        # 3) authority 댐프너: 전문 용어가 감성어보다 많으면 톤 하향
        #    (전문적 문서에 감성어가 산재된 경우 → 감성적 의도가 아님)
        if authority_count >= 3 and authority_count >= emotional_count * 0.5:
            if tone_level > 2:
                tone_level -= 1

        # 4) 스토리텔링 요소가 있으면 레벨 상향
        storytelling_markers = [
            "어느 ", "그리고 ", "그 순간", "그때", "시작되",
            "이야기", "스토리", "여정", "모험",
        ]
        story_count = sum(text.count(m) for m in storytelling_markers)
        if story_count > 5 and tone_level < 5:
            tone_level = min(5, tone_level + 1)

        return {
            "tone_level": tone_level,
            "total_count": total_count,
            "category_counts": category_counts,
            "density": round(density, 2),
            "authority_count": authority_count,
            "formal_discount": formal_discount,
            "effective_emotional": round(emotional_count, 1),
            "top_phrases": sorted(found_phrases, key=lambda x: -x[2])[:20],
            "emotional_sentences": emotional_sentences[:10],
        }

    def _extract_emotional_sentences(self, text: str) -> List[str]:
        """감성적 문장 추출 (실제 예시)"""
        sentences = re.split(r'[.!?\n]', text)
        emotional = []

        trigger_words = [
            "폭발", "극대화", "몰입", "열광", "도파민", "전율",
            "이야기", "스토리", "판타지", "동화", "마법",
            "기다림", "꿈", "축하", "축제",
        ]

        for sent in sentences:
            sent = sent.strip()
            if 10 < len(sent) < 100:
                if any(w in sent for w in trigger_words):
                    emotional.append(sent)

        return emotional

    # ═══════════════════════════════════════════════════════════
    # 3. IP 깊이 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_ip_depth(self, text: str) -> Dict:
        """도메인 전문성 깊이 분석 (게임 IP / IT 아키텍처 / 정책 전문성 등)"""
        text_lower = text.lower()

        # 고유명사 검출 (게임 캐릭터 / IT 제품 / 기관명 등)
        found_characters = []
        for name in self._depth_markers.get("proper_nouns", []):
            if name in text:
                found_characters.append(name)

        # 도메인 전문 용어 검출 (세계관 / 아키텍처 / 정책 등)
        found_worldview = []
        for term in self._depth_markers.get("domain_terms", []):
            count = text.count(term)
            if count > 0:
                found_worldview.append((term, count))

        # 커뮤니티/은어 검출
        found_slang = []
        for term in self._depth_markers.get("community_slang", []):
            count = text_lower.count(term.lower())
            if count > 0:
                found_slang.append((term, count))

        # 산업 어휘 중 커뮤니티성 용어 검출 (팬 언어 / 업계 은어)
        found_community = []
        community_cats = ["fan_language", "community_slang"]
        for cat in community_cats:
            for term in self._vocab.get(cat, []):
                count = text_lower.count(term.lower())
                if count > 0:
                    found_community.append((term, count))

        total_terms = (
            len(found_characters) +
            sum(c for _, c in found_worldview) +
            sum(c for _, c in found_slang) +
            sum(c for _, c in found_community)
        )

        # IP 깊이 점수 (0~1)
        char_score = min(len(found_characters) / 5, 1.0) * 0.3
        world_score = min(sum(c for _, c in found_worldview) / 20, 1.0) * 0.3
        slang_score = min(sum(c for _, c in found_slang) / 10, 1.0) * 0.2
        community_score = min(sum(c for _, c in found_community) / 10, 1.0) * 0.2

        score = char_score + world_score + slang_score + community_score

        # 실제 용어 리스트 (중복 제거, 빈도순)
        lore_terms = (
            found_characters +
            [t for t, _ in sorted(found_worldview, key=lambda x: -x[1])]
        )
        community_terms = (
            [t for t, _ in sorted(found_slang, key=lambda x: -x[1])] +
            [t for t, _ in sorted(found_community, key=lambda x: -x[1])]
        )

        return {
            "score": round(score, 2),
            "character_count": len(found_characters),
            "characters": found_characters,
            "worldview_terms": found_worldview,
            "slang_terms": found_slang,
            "community_vocab": found_community,
            "total_terms": total_terms,
            "lore_terms": lore_terms[:20],
            "community_terms": community_terms[:20],
        }

    # ═══════════════════════════════════════════════════════════
    # 4. 프로그램 네이밍 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_program_naming(self, text: str) -> Dict:
        """프로그램/존/스테이지 네이밍 패턴 추출"""
        examples = []

        # 전처리: 줄바꿈을 공백으로 치환한 클린 텍스트 준비
        clean_text = re.sub(r'\n+', ' ', text)
        clean_text = re.sub(r'\s{2,}', ' ', clean_text)

        # 패턴 기반 추출 (산업별 패턴 사용, 클린 텍스트에서)
        for pattern in self._naming_patterns:
            matches = re.findall(pattern, clean_text)
            for m in matches:
                m = m.strip()
                if self._is_valid_program_name(m) and m not in examples:
                    examples.append(m)

        # 추가: "~ 존", "~ 스테이지", "~ 스토어" 등 한글 추출 (클린 텍스트)
        zone_pattern = r'([가-힣A-Za-z\s]{2,20}(?:존|스테이지|스토어|스팟|데스크|데이트|캠핑존|교역소|교환소|포토존|게임존|미니게임))'
        for m in re.findall(zone_pattern, clean_text):
            m = m.strip()
            if self._is_valid_program_name(m) and m not in examples:
                examples.append(m)

        # 원본 텍스트에서도 한글 프로그램명 추출 (줄 단위로 깔끔한 것만)
        for line in text.split('\n'):
            line = line.strip()
            # "XX 존", "XX 스테이지" 등 깔끔한 한 줄 프로그램명
            if (5 <= len(line) <= 35 and
                not '\n' in line and
                re.match(r'^[가-힣A-Za-z\s]+(?:존|스테이지|스토어|파티|캠핑|이벤트|프로그램)$', line) and
                line not in examples):
                if self._is_valid_program_name(line):
                    examples.append(line)

        # IP 세계관 네이밍 vs 기능적 네이밍 분류
        ip_names = []
        functional_names = []
        branded_names = []

        for name in examples:
            has_ip = any(
                term in name for term in
                self._depth_markers.get("proper_nouns", []) +
                self._depth_markers.get("domain_terms", []) +
                ["프로즌", "글루피", "에린", "오로라", "모닥불", "아스칼론"]
            )
            has_english = bool(re.search(r'[A-Z]', name))

            if has_ip:
                ip_names.append(name)
            elif has_english:
                branded_names.append(name)
            else:
                functional_names.append(name)

        # 스타일 결정
        if len(ip_names) > len(functional_names):
            style = "ip_narrative"
        elif len(branded_names) > len(functional_names):
            style = "branded"
        elif ip_names or branded_names:
            style = "hybrid"
        else:
            style = "functional"

        return {
            "style": style,
            "examples": examples[:20],
            "ip_names": ip_names,
            "branded_names": branded_names,
            "functional_names": functional_names,
        }

    @staticmethod
    def _is_valid_program_name(name: str) -> bool:
        """프로그램명 유효성 검사 — 노이즈 필터링"""
        if not name or len(name) < 3 or len(name) > 35:
            return False
        # 줄바꿈 포함 → 노이즈
        if '\n' in name or '\r' in name:
            return False
        # 너무 많은 공백 (3개 이상이면 문장 파편)
        if name.count(' ') > 3:
            return False
        # 한글 조사로 시작하면 문장 파편
        if re.match(r'^[은는이가을를의에서로와과에게]', name):
            return False
        # 한글 조사로 끝나면 문장 파편
        if re.search(r'[은는을를의에과와]$', name):
            return False
        # 너무 일반적인 단어/문장 파편 필터링
        generic = [
            "체험존", "전시존", "판매존", "관람석", "운영 구역",
            "일 간의", "개최", "지역의", "공식", "발적",
            "속 세계관", "한국은", "서울에서", "게임을",
            "에게 특별한", "신규 유저", "기존 유저", "팬들에게",
            "설치 안내", "추억을 선사", "선사할",
            "통한", "위한", "따른", "대한",
        ]
        if any(g in name for g in generic):
            return False
        # 15자 넘으면서 조사 포함 → 문장일 가능성 높음
        if len(name) > 15:
            particle_count = len(re.findall(r'[은는이가을를의에서로와과]', name))
            if particle_count >= 3:
                return False
        return True

    # ═══════════════════════════════════════════════════════════
    # 5. 내러티브/프레이밍 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_narrative_framing(self, text: str) -> Dict:
        """내러티브 스타일, 메타포, 진입 후크 분석"""
        # 세계관 기반 키워드
        worldview_markers = [
            "세계관", "로어", "캐릭터", "스토리", "동화", "판타지",
            "왕국", "전설", "모험", "여정", "마법", "운명",
        ]
        # 데이터 기반 키워드
        data_markers = [
            "MAU", "DAU", "PV", "UV", "%", "만원", "억원",
            "전년 대비", "성장률", "달성", "목표", "KPI", "ROI",
        ]
        # 감성 기반 키워드
        emotion_markers = [
            "기다림", "꿈", "희망", "감동", "열정", "사랑",
            "추억", "기억", "축하", "축제", "함께", "우리",
        ]

        wv_count = sum(text.count(w) for w in worldview_markers)
        data_count = sum(text.count(w) for w in data_markers)
        emo_count = sum(text.count(w) for w in emotion_markers)

        total = wv_count + data_count + emo_count + 1
        if wv_count / total > 0.45:
            style = "worldview_based"
        elif data_count / total > 0.45:
            style = "data_driven"
        elif emo_count / total > 0.4:
            style = "emotion_led"
        else:
            style = "hybrid"

        # 메타포 추출
        metaphor_patterns = [
            r'「([^」]+)」',  # 일본식 큰따옴표
            r'\'([^\']+)\'',  # 작은따옴표
            r'『([^』]+)』',  # 큰괄호
        ]
        metaphors = []
        for p in metaphor_patterns:
            for m in re.findall(p, text):
                if 3 <= len(m) <= 30:
                    metaphors.append(m)

        core_metaphor = metaphors[0] if metaphors else ""

        # 진입 후크 패턴 추출
        hook_patterns = self._extract_hook_patterns(text)
        entry_hook = hook_patterns[0] if hook_patterns else ""

        # 반복 모티프
        recurring = self._find_recurring_motifs(text)

        return {
            "style": style,
            "core_metaphor": core_metaphor,
            "entry_hook": entry_hook,
            "recurring_motif": ", ".join(recurring[:3]),
            "description": (
                f"프레이밍: {style} "
                f"(세계관={wv_count}, 데이터={data_count}, 감성={emo_count}). "
                f"메타포: {len(metaphors)}건, 후크: {len(hook_patterns)}건"
            ),
            "metaphors": metaphors[:10],
            "hooks": hook_patterns[:5],
        }

    def _extract_hook_patterns(self, text: str) -> List[str]:
        """진입 후크 패턴 추출 (질문형, 선언형, 서사형)"""
        hooks = []
        lines = text.split('\n')

        for line in lines:
            line = line.strip()
            if not line or len(line) < 5 or len(line) > 80:
                continue

            # 질문형 후크
            if line.endswith('?') or line.endswith('?'):
                hooks.append(f"[질문형] {line}")
            # 선언형 후크 (짧은 강렬한 문장)
            elif line.endswith('!') and len(line) < 30:
                hooks.append(f"[선언형] {line}")
            # 서사형 후크 (스토리텔링 시작)
            elif any(line.startswith(s) for s in ["어느 ", "그때 ", "만약 "]):
                hooks.append(f"[서사형] {line}")
            # 영문 전략 키워드
            elif re.match(r'^[A-Z][A-Z\s]{5,}$', line):
                hooks.append(f"[전략키워드] {line}")

        return hooks

    def _find_recurring_motifs(self, text: str) -> List[str]:
        """반복되는 모티프 키워드 (3회 이상 등장하는 감성 단어)"""
        emotional_all = []
        for cat_words in EMOTIONAL_WORDS.values():
            emotional_all.extend(cat_words)

        motifs = []
        for w in emotional_all:
            count = text.count(w)
            if count >= 3:
                motifs.append((w, count))

        motifs.sort(key=lambda x: -x[1])
        return [w for w, _ in motifs[:5]]

    # ═══════════════════════════════════════════════════════════
    # 6. 텍스트 밀도/스타일 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_text_density(self, text: str) -> Dict:
        """텍스트 밀도, 문장 길이, 비주얼 비중 분석"""
        lines = [l.strip() for l in text.split('\n') if l.strip()]

        if not lines:
            return {"style": "balanced", "avg_line_len": 0}

        line_lengths = [len(l) for l in lines]
        avg_len = sum(line_lengths) / len(line_lengths)

        # 짧은 라인 (비주얼 레이블, 제목 등) 비율
        short_lines = sum(1 for l in line_lengths if l < 15)
        short_ratio = short_lines / len(lines)

        # 긴 라인 (본문, 설명) 비율
        long_lines = sum(1 for l in line_lengths if l > 60)
        long_ratio = long_lines / len(lines)

        # 이미지 관련 키워드 (IMG, 사진, 이미지, 비주얼 등) → 비주얼 비중 추정
        visual_markers = ["이미지", "사진", "비주얼", "일러스트", "조감도", "평면도", "렌더링"]
        visual_count = sum(text.count(m) for m in visual_markers)
        image_ratio = min(visual_count / max(len(lines) / 10, 1), 1.0)

        if avg_len < 15 and short_ratio > 0.7:
            style = "minimal"  # 비주얼 중심, 텍스트 최소
        elif avg_len < 25 and short_ratio > 0.5:
            style = "balanced"
        elif long_ratio > 0.3:
            style = "rich"  # 텍스트 풍부
        else:
            style = "balanced"

        return {
            "style": style,
            "avg_line_len": round(avg_len, 1),
            "short_ratio": round(short_ratio, 2),
            "long_ratio": round(long_ratio, 2),
            "image_ratio": round(image_ratio, 2),
            "text_only_ratio": round(long_ratio * 0.5, 2),
        }

    # ═══════════════════════════════════════════════════════════
    # 7. Win Theme 스타일 분석
    # ═══════════════════════════════════════════════════════════

    def _analyze_win_theme_style(self, text: str) -> Dict:
        """Win Theme 스타일 및 전략적 표현 패턴 분석"""
        # 섹션 헤더 블랙리스트 (Win Theme이 아닌 일반 섹션명)
        section_headers = {
            "MISSION", "MARKET ANALYSIS", "CORE TARGET", "MARKET TREND",
            "KEY STRATEGY", "ACTION PLAN", "MANAGEMENT", "BUDGET",
            "SCHEDULE", "TIMELINE", "OUR REFERENCE", "WHY US",
            "EXECUTIVE SUMMARY", "INSIGHT", "CONCEPT", "CLOSING",
            "THANK YOU", "APPENDIX", "TABLE OF CONTENTS", "OVERVIEW",
            "INTRODUCTION", "SUMMARY", "CONCLUSION", "INDEX",
            "PROPOSAL", "REPORT", "KPI", "ROI", "SNS",
        }

        # 전략 키워드 추출
        strategy_phrases = []
        for pattern in STRATEGY_PHRASES:
            for m in re.findall(pattern, text):
                m = m.strip()
                if 4 <= len(m) <= 50:
                    # 줄바꿈 포함 → 노이즈
                    if '\n' in m or '\r' in m:
                        continue
                    # 섹션 헤더 필터링
                    upper = m.upper().strip()
                    if upper in section_headers:
                        continue
                    # 1~2글자 단어만으로 이루어진 것 필터 (약어 반복)
                    if all(len(w) <= 2 for w in m.split()):
                        continue
                    # 의미 있는 전략 표현인지 확인 (최소 2단어 또는 한글 포함)
                    has_korean = bool(re.search(r'[가-힣]', m))
                    word_count = len(m.split())
                    if has_korean or word_count >= 2:
                        strategy_phrases.append(m)

        # IP 세계관 기반 전략 표현
        ip_strategy = [
            p for p in strategy_phrases
            if any(t in p for t in ["세계관", "캐릭터", "팬", "IP", "도파민", "몰입"])
        ]
        # 데이터/키워드 기반 전략 표현
        data_strategy = [
            p for p in strategy_phrases
            if any(t in p.upper() for t in ["MAU", "KPI", "%", "달성", "성장"])
        ]
        # 감성 후크 전략
        emotion_strategy = [
            p for p in strategy_phrases
            if any(t in p for t in ["기다림", "꿈", "축제", "함께", "만남", "경험"])
        ]

        if len(ip_strategy) > len(data_strategy):
            style = "ip_worldview"
        elif len(data_strategy) > len(emotion_strategy):
            style = "keyword_functional"
        elif emotion_strategy:
            style = "emotional_hook"
        else:
            style = "keyword_functional"

        # 상위 전략 표현 정제
        unique_phrases = list(dict.fromkeys(strategy_phrases))

        return {
            "style": style,
            "examples": unique_phrases[:10],
            "ip_count": len(ip_strategy),
            "data_count": len(data_strategy),
            "emotion_count": len(emotion_strategy),
        }

    # ═══════════════════════════════════════════════════════════
    # 8. 어휘 예시 추출
    # ═══════════════════════════════════════════════════════════

    def _extract_vocabulary_examples(self, text: str) -> Dict[str, List[str]]:
        """실제 문서에서 사용된 핵심 어휘 예시 추출"""
        examples = {
            "hook_phrases": [],      # 후크 표현
            "emotional_phrases": [], # 감성적 표현
            "ip_integration": [],    # IP 통합 표현
            "program_names": [],     # 프로그램 네이밍
            "strategy_keywords": [], # 전략 키워드
        }

        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line or len(line) < 5:
                continue

            # 후크 표현 (짧고 강렬)
            if (len(line) < 40 and
                any(w in line for w in ["!", "?", "…"]) and
                len(line) > 8):
                examples["hook_phrases"].append(line)

            # 도메인 고유명사 통합 표현
            if any(name in line for name in self._depth_markers.get("proper_nouns", [])):
                if len(line) < 80:
                    examples["ip_integration"].append(line)

        # 중복 제거 + 상위 선택
        for key in examples:
            examples[key] = list(dict.fromkeys(examples[key]))[:15]

        return examples

    # ═══════════════════════════════════════════════════════════
    # 9. 종합 톤 규칙 생성
    # ═══════════════════════════════════════════════════════════

    def _generate_tone_rules(
        self, vocab, emotional, ip_depth, naming, framing, density, win_theme
    ) -> List[str]:
        """분석 결과를 기반으로 구체적이고 실행 가능한 톤 규칙 생성"""
        rules = []

        # ── IP 깊이 기반 규칙 ──
        if ip_depth["score"] >= 0.5:
            rules.append("게임 IP의 세계관과 캐릭터를 제안서 전체에 자연스럽게 녹여라")
            if ip_depth["characters"]:
                chars = ", ".join(ip_depth["characters"][:5])
                rules.append(f"IP 캐릭터({chars})의 고유 성격과 대사를 프로그램 기획에 반영하라")
            if ip_depth["community_terms"]:
                terms = ", ".join(t for t, _ in ip_depth.get("slang_terms", [])[:5])
                if terms:
                    rules.append(f"팬 커뮤니티 슬랭({terms})을 자연스럽게 사용하여 '인사이더' 느낌을 줘라")
            if ip_depth.get("worldview_terms"):
                rules.append("IP 세계관 용어를 섹션 제목과 프로그램명에 직접 활용하라")
        elif ip_depth["score"] >= 0.2:
            rules.append("브랜드 고유 용어를 핵심 포인트에 전략적으로 배치하라")

        # ── 감성 톤 기반 규칙 ──
        tone = emotional["tone_level"]
        if tone >= 5:
            rules.append("기능 나열이 아닌 감성적 스토리텔링을 전면에 배치하라")
            rules.append("'어느 날~', '그 순간~' 같은 서사적 진입부를 활용하라")
            if emotional.get("emotional_sentences"):
                ex = emotional["emotional_sentences"][0]
                rules.append(f"실제 레퍼런스 감성 표현 참고: \"{ex}\"")
        elif tone >= 4:
            rules.append("데이터와 감성을 7:3으로 배합하되, 도입부는 감성으로 시작하라")
            rules.append("Action Title에 감성적 후크를 포함하라 (사실 전달 < 감정 유발)")
        elif tone >= 3:
            rules.append("전문적 톤을 유지하되 핵심 메시지에 감성 코드를 삽입하라")
        else:
            rules.append("간결하고 객관적인 톤을 유지하라. 데이터와 근거 중심으로 작성하라")

        # ── 프로그램 네이밍 규칙 ──
        if naming["style"] == "ip_narrative":
            rules.append("프로그램명에 IP 세계관 용어를 삽입하라 (기능 설명형 네이밍 지양)")
            if naming["ip_names"]:
                ex = ", ".join(naming["ip_names"][:3])
                rules.append(f"레퍼런스 네이밍 참고: {ex}")
        elif naming["style"] == "branded":
            rules.append("프로그램명은 영문 브랜드명 + 한글 설명 조합으로 구성하라")
            if naming["branded_names"]:
                ex = ", ".join(naming["branded_names"][:3])
                rules.append(f"레퍼런스 네이밍 참고: {ex}")
        elif naming["style"] == "hybrid":
            rules.append("IP 이름 + 기능 설명을 혼합한 네이밍 전략을 사용하라")

        # ── 내러티브 프레이밍 규칙 ──
        if framing["style"] == "worldview_based":
            rules.append("IP/브랜드 세계관을 제안서 구조의 뼈대로 사용하라")
            if framing["core_metaphor"]:
                rules.append(f"핵심 메타포 '{framing['core_metaphor']}'를 관통 모티프로 활용하라")
        elif framing["style"] == "emotion_led":
            rules.append("감정 곡선을 설계하라: 도입(기대) → 전개(공감) → 절정(감동) → 마무리(행동)")
        elif framing["style"] == "data_driven":
            rules.append("데이터 인사이트를 서두에 배치하고, 해석으로 전환하라")

        # ── 텍스트 밀도 규칙 ──
        if density["style"] == "minimal":
            rules.append("텍스트를 최소화하고 이미지/비주얼 비중을 70% 이상 유지하라")
            rules.append("한 슬라이드에 핵심 메시지 1개만 담아라 (Show, Don't Tell)")
        elif density["style"] == "rich":
            rules.append("상세한 설명을 포함하되, 계층적 타이포그래피로 스캔 가능하게 구성하라")
        else:
            rules.append("비주얼과 텍스트를 균형 있게 배합하라 (50:50 목표)")

        # ── Win Theme 스타일 규칙 ──
        if win_theme["style"] == "ip_worldview":
            rules.append("Win Theme을 IP 세계관의 핵심 가치와 연결하여 제시하라")
        elif win_theme["style"] == "emotional_hook":
            rules.append("Win Theme을 감정적 공감 포인트로 포장하라")

        # ── 후크 패턴 규칙 ──
        hooks = framing.get("hooks", [])
        if hooks:
            hook_types = [h.split("]")[0].replace("[", "") for h in hooks]
            most_common_hook = Counter(hook_types).most_common(1)
            if most_common_hook:
                hook_type = most_common_hook[0][0]
                rules.append(f"섹션 도입부에 '{hook_type}' 후크를 우선 사용하라")

        return rules


# ═══════════════════════════════════════════════════════════════
# 편의 함수
# ═══════════════════════════════════════════════════════════════

def deep_analyze_document(
    full_text: str,
    file_name: str = "",
    industry: str = "",
) -> ContentToneProfile:
    """단일 문서 딥 분석 편의 함수"""
    analyzer = ContentToneAnalyzer(industry=industry)
    return analyzer.analyze(full_text, file_name)
