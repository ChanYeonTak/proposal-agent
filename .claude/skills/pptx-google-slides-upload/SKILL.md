---
name: pptx-google-slides-upload
description: |
  PPTX 파일을 Google Slides에 업로드하는 전체 워크플로우. 폰트 치환(Pretendard →
  Noto Sans KR), 시각적 무손실 이미지 최적화(1920px 다운샘플), OAuth Resumable 업로드,
  Apps Script 폰트 일괄 적용까지 5단계. 100MB 초과 PPTX를 Google Slides 변환 한계
  아래로 줄이면서 한글 렌더링을 보존한다. NIKKE 3.5주년 운영매뉴얼(158MB → 91MB)
  실전 검증 완료.
---

# PPTX → Google Slides 업로드 워크플로우

실전 검증: **158MB PPTX → 91MB → Google Slides 업로드 → Noto Sans KR 일괄 적용** 성공 (2026-04-15, NIKKE 3.5주년 운영매뉴얼)

## 배경 — 왜 이 과정이 필요한가

Google Slides는 PPTX를 직접 import할 수 있지만 3가지 문제가 발생한다:

1. **Pretendard 폰트**: Google Slides 기본 폰트에 없고, "더 많은 글꼴"로 등록해도 **변환 시점에 Arial로 영구 고정**됨 (나중에 추가해도 소용없음)
2. **용량 한계**: PPTX → Slides 자동 변환은 **100MB 이하만 허용**. 이미지가 많은 마케팅 제안서는 쉽게 150~250MB에 도달
3. **서비스 계정 금지**: 2025년부터 Google이 **서비스 계정의 개인 Drive 파일 소유를 차단**. OAuth 사용자 위임 필수

## 파이프라인 전체 구조

```
원본 PPTX
  ↓  1. pptx_font_replacer.py        → Pretendard 계열 → Noto Sans KR
  ↓  2. pptx_inspector.py (선택)      → 이미지 분포 분석
  ↓  3. pptx_optimizer.py             → 1920px 다운샘플 + PNG/JPEG 재압축
  ↓  4. google_slides_bridge.py       → OAuth Resumable 업로드 + Slides 변환
  ↓  5. Apps Script (apps_script.js)  → 잔여 폰트 일괄 적용
Google Slides (편집 가능)
```

## 사전 준비 (최초 1회)

### A. OAuth 클라이언트 발급
1. https://console.cloud.google.com/apis/credentials → **OAuth 클라이언트 ID**
2. 동의 화면 구성 필요 시: User Type **외부** → 테스트 사용자에 본인 Gmail 추가 (필수)
3. 애플리케이션 유형 **데스크톱 앱** → JSON 다운로드
4. `config/google_oauth_client.json`에 저장

### B. 패키지
```bash
pip install google-api-python-client google-auth google-auth-oauthlib Pillow
```

### C. Drive 폴더
- Google Drive에 업로드 대상 폴더 생성
- URL에서 폴더 ID 복사: `.../folders/<FOLDER_ID>`

### D. gitignore 확인
```
config/google_oauth_client.json
config/google_service_account.json
config/google_token.pickle
```

## 실행 순서

### STEP 1 — 폰트 치환 (Pretendard → Noto Sans KR)

```bash
python src/integrations/pptx_font_replacer.py \
  "input/프로젝트/원본.pptx" \
  "input/프로젝트/원본_noto.pptx"
```

**동작**: PPTX 내부 XML에서 `typeface="Pretendard*"` 를 `typeface="Noto Sans KR"` 로 일괄 치환 (latin/ea/cs 모두).

**⚠️ 알려진 버그**: 현재 치환기는 정확한 `"Pretendard"` 만 매칭 → `Pretendard ExtraBold` 같은 웨이트 베리언트는 놓침 (slide5.xml 기준 2/120개 잔여). Apps Script STEP 5에서 처리됨. 근본 수정하려면 regex를 `typeface="Pretendard[^"]*"` 로 변경.

### STEP 2 — (선택) 분석 리포트

```bash
python src/integrations/pptx_inspector.py "input/프로젝트/원본_noto.pptx"
```

출력: 이미지 포맷 분포, 2400px 초과 개수, 임베드 폰트/중복 이미지, 100MB 목표 예상 달성 여부.

### STEP 3 — 시각적 무손실 최적화

```bash
python src/integrations/pptx_optimizer.py \
  "input/프로젝트/원본_noto.pptx" \
  "input/프로젝트/원본_noto_optimized.pptx"
```

**파라미터 기본값** (LosslessPPTXOptimizer):
- `max_image_dim=1920` — Google Slides 최대 렌더링 해상도
- `jpeg_quality=90` — 시각적 무손실 범위
- `strip_fonts=True` — Google은 임베드 폰트 무시하므로 제거
- `zip_level=9` — 최대 압축

**레이아웃 불변 보장**:
- 슬라이드 XML은 한 글자도 수정 안 함
- 이미지 표시 크기(`cx`, `cy`)는 유지, 픽셀 해상도만 축소
- 과대 해상도만 Lanczos 다운샘플

**실전 결과 (NIKKE)**: 158.2 MB → 91.0 MB (42.5% 감소, 이미지 44개 다운샘플)

### STEP 4 — OAuth 업로드 + Slides 변환

```bash
python src/integrations/google_slides_bridge.py \
  "input/프로젝트/원본_noto_optimized.pptx" \
  --title "프로젝트명 표시 제목" \
  --folder "<Drive_FOLDER_ID>"
```

**첫 실행 시**: 브라우저가 열리고 Google 로그인 요구
→ 본인 Gmail 선택
→ "확인되지 않은 앱" 경고 → **고급 → 이동(안전하지 않음)**
→ 권한 허용
→ 토큰이 `config/google_token.pickle`에 저장

**이후 실행**: 브라우저 없이 바로 진행 (토큰 만료 시 자동 갱신)

**동작**:
- 8MB 청크 Resumable 업로드 (5TB까지 지원)
- `mimeType: application/vnd.google-apps.presentation` 지정 → Slides 포맷 자동 변환
- 파일 소유자: 본인 Gmail (폴더 공유 불필요)

**출력**:
```
완료: https://docs.google.com/presentation/d/<FILE_ID>/edit?usp=drivesdk
파일 ID: <FILE_ID>
```

### STEP 5 — Apps Script로 폰트 최종 일괄 적용 ⭐

**이 단계는 필수**. STEP 1 치환기가 놓친 웨이트 베리언트(`Pretendard ExtraBold` 등)와 Google 변환기가 Arial로 고정시킨 잔여분을 일괄 교체한다.

1. https://script.google.com/ → **새 프로젝트**
2. `apps_script.js` 내용 붙여넣기
3. `FILE_ID` 상수를 STEP 4 출력의 파일 ID로 변경
4. **저장 → 실행**
5. 첫 실행 시 권한 요청 → 본인 Gmail로 승인 (동일하게 "확인되지 않은 앱" → 고급 → 이동)
6. 실행 로그에 `완료: 총 XXX개 텍스트 요소 폰트 변경` 확인
7. Slides 파일 새로고침(F5) → 전체 Noto Sans KR 반영

## 파일 구조

```
src/integrations/
├── pptx_font_replacer.py      # STEP 1: 폰트 XML 치환
├── pptx_inspector.py           # STEP 2: 분석 리포트
├── pptx_optimizer.py           # STEP 3: 시각적 무손실 최적화
└── google_slides_bridge.py     # STEP 4: OAuth Resumable 업로드

.claude/skills/pptx-google-slides-upload/
├── SKILL.md                    # 본 문서
└── apps_script.js              # STEP 5: Apps Script 템플릿

config/
├── google_oauth_client.json    # OAuth 클라이언트 시크릿 (gitignore)
├── google_token.pickle         # OAuth 사용자 토큰 (gitignore, 자동생성)
└── google_service_account.json # 참고용 (사용 안 함, gitignore)
```

## 트러블슈팅

| 증상 | 원인 | 해결 |
|------|------|------|
| `access_denied` 403 | 테스트 사용자 미등록 | Cloud Console → OAuth 동의 화면 → 테스트 사용자에 본인 Gmail 추가 |
| `storageQuotaExceeded` | 서비스 계정 사용 | OAuth 클라이언트로 전환 (서비스 계정은 2025+ 개인 Drive 소유 불가) |
| 변환 후 Arial로 표시 | Google 변환기가 폰트명 미인식 | Apps Script STEP 5 실행 |
| 100MB 초과로 변환 실패 | 이미지 과대 | `max_image_dim=1600` 으로 더 공격적 다운샘플 |
| Apps Script 에러 "권한 없음" | 파일 소유자 불일치 | FILE_ID가 본인 소유 파일인지 확인 |
| 그룹 내 텍스트가 안 바뀜 | Apps Script 재귀 누락 | apps_script.js는 GROUP 재귀 처리 포함됨 확인 |

## 재사용 체크리스트

새 프로젝트에 이 워크플로우 적용 시:

- [ ] 원본 PPTX가 `Pretendard` 폰트 사용 중인가? (다른 폰트면 replacer의 mapping 수정)
- [ ] 원본 용량이 100MB 초과인가? (아니면 STEP 3 생략 가능)
- [ ] Drive 대상 폴더 ID 확보
- [ ] `google_token.pickle` 존재 → STEP 4 브라우저 없이 실행됨
- [ ] STEP 4 완료 후 FILE_ID 복사 → Apps Script FILE_ID에 반영
- [ ] Apps Script 실행 후 Slides 새로고침으로 반영 확인

## 현실적 한계 (기록용)

- **Layer A 비트 무손실 불가능**: Google은 import 시 모든 이미지를 무조건 재인코딩 (JPEG 양자화). 이를 우회할 Google 제공 경로 없음.
- **Layer B 시각적 무손실만 달성**: 1920px 이상의 과대 해상도는 어차피 렌더링에 기여하지 않으므로 다운샘플해도 사용자 경험 동일.
- **OOXML 직접 import 경로 없음**: XML을 Slides API로 직접 밀어넣는 경로 없음 — Drive API 바이너리 업로드가 유일한 입구.
- **그라디언트/그림자/라운드박스 일부**: Google 렌더러가 OOXML 파라미터 일부를 무시. slide_kit의 modern 스타일 중 일부는 변환 후 납작해짐 — 근본 해결 불가, 디자인 단계에서 선택 필요.
