# 핸드오프: PPTX → Google Slides 최적화 업로드 시스템

**브랜치**: `claude/pptx-xml-google-slides-V6iqF`
**관련 문서**: `docs/PPTX_GOOGLE_SLIDES_UPLOAD_PLAN.md` (상세 설계)
**대상 파일**: `input/2026 NIKKE/*.pptx`

---

## 로컬 Claude Code가 이어서 할 일

이 문서 하나만 읽고 바로 작업 이어받으세요.

### 확정된 요구사항 (변경 금지)

1. **편집 가능** — Google Slides에서 텍스트/도형 네이티브 편집 가능해야 함
2. **전체 업로드** — 파일 크기 제한 없이 Resumable 청크 업로드
3. **이미지 시각적 무손실 (Layer B)** — 원본과 보는 사람이 구분 불가
4. **레이아웃 보존** — 좌표/크기/회전/z-order/마진 절대 변경 금지

### 핵심 결정사항

- **폰트**: `FONT = "Pretendard"` 그대로 유지. 사용자가 구글 슬라이드 "더 많은 글꼴"에서 Pretendard를 계정에 한 번 추가. PPTX 임베드 폰트는 strip (구글이 무시함).
- **이미지**: PNG는 oxipng 무손실 재압축, JPEG는 원본 유지. 긴 변 2400px 초과 시에만 Lanczos 다운샘플 (구글 렌더링 최대 1920px의 1.25x 여유).
- **레이아웃**: 슬라이드 내 이미지 도형의 `cx/cy` 치수는 건드리지 않음. 픽셀 해상도만 줄임.
- **업로드**: Drive API Resumable, 8MB 청크, 진행률 콜백.
- **변환**: mimeType `application/vnd.google-apps.presentation` 으로 create → 자동 Slides 변환.

### Google 플랫폼 제약 (이미 검토됨)

| 항목 | 값 |
|------|---|
| Drive 단순 업로드 | 5MB |
| Drive Resumable | 5TB |
| PPTX→Slides 자동 변환 한계 | ~100MB |
| Slides 최대 렌더링 해상도 | 1920×1080 |
| 구글이 import 시 이미지 재인코딩 | 예 (피할 수 없음) |

→ 목표 최적화 후 크기: **100MB 이하** (여유롭게 50MB 이하 권장)

---

## 구현 파일 3종 (신규 작성)

### 1. `src/integrations/pptx_inspector.py`

PPTX 분석 리포트 생성. 파일은 건드리지 않고 내부 구조만 분석.

**분석 항목**:
- 파일 전체 크기, 슬라이드 수
- `ppt/media/` 내 이미지 목록 (파일명, 포맷, 해상도, 바이트 크기)
- 중복 이미지 해시 검사 (MD5)
- `ppt/fonts/` 임베드 폰트 존재 여부와 크기
- `ppt/embeddings/` Excel 링크 등
- EMF/WMF 벡터 존재 여부 (Google Slides에서 깨짐)
- 각 이미지가 슬라이드에서 실제로 차지하는 `cx/cy` 값 → 과잉 해상도 계산
- 예상 최적화 후 크기 시뮬레이션 (파라미터별)

**출력**: JSON + 사람 읽기용 텍스트 리포트

```python
# 사용 예
from src.integrations.pptx_inspector import PPTXInspector
report = PPTXInspector("input/2026 NIKKE/원본.pptx").inspect()
print(report.summary())
report.save_json("inspect_report.json")
```

### 2. `src/integrations/pptx_optimizer.py`

`LosslessPPTXOptimizer` 클래스. 상세 구현은 `docs/PPTX_GOOGLE_SLIDES_UPLOAD_PLAN.md`의 코드 블록 참조.

**핵심 동작**:
```python
class LosslessPPTXOptimizer:
    def __init__(self,
                 max_image_dim: int = 2400,      # 과잉 해상도 상한
                 use_oxipng: bool = True,         # PNG 무손실 재압축
                 strip_fonts: bool = True,        # 임베드 폰트 제거
                 strip_metadata: bool = True,     # EXIF/색프로파일 정리
                 dedupe_media: bool = False,      # 1차: 비활성. 2차에서 rels 재작성 포함
                 zip_level: int = 9):
        ...

    def optimize(self, src: Path, dst: Path) -> dict:
        """ZIP을 순회하며 media/fonts만 재작성. XML은 한 글자도 수정 안 함.
        통계 딕셔너리 반환."""
```

**절대 하지 말 것**:
- `ppt/slides/*.xml` 내부 수정 (레이아웃 보존 위반)
- `ppt/slideLayouts/*`, `ppt/theme/*` 수정
- JPEG 무조건 재인코딩 (품질 손실)
- 이미지 리사이즈 시 도형의 `cx/cy` 변경

**필수 동작**:
- 이미지 포맷별 분기 (PNG / JPEG / 나머지는 원본 유지)
- PNG: `Pillow.save(format="PNG", optimize=True, compress_level=9)` → 성공 시 `oxipng --opt max --strip safe`로 한 번 더
- JPEG: 리사이즈 필요한 경우만 `quality=95, subsampling=0, progressive=True`로 재저장. 아니면 원본 유지
- 크기 증가 시 원본으로 롤백
- 폰트: `ppt/fonts/` 및 `ppt/embeddings/fonts/` 전체 skip

### 3. `src/integrations/google_slides_bridge.py`

`GoogleSlidesUploader` 클래스. 상세는 계획서 참조.

```python
class GoogleSlidesUploader:
    def __init__(self, credentials_json: Path):
        # service_account.Credentials.from_service_account_file
        # scopes=["https://www.googleapis.com/auth/drive.file"]
        ...

    def upload_and_convert(self,
                           pptx_path: Path,
                           title: str,
                           folder_id: str | None = None,
                           on_progress=None) -> dict:
        """MediaFileUpload(resumable=True, chunksize=8*1024*1024)
        mimeType="application/vnd.google-apps.presentation" 로 create
        next_chunk() 루프 → 진행률 콜백
        webViewLink 반환"""
        ...
```

**OAuth 사용자 플로우 대안**도 추가로 제공하면 좋음 (`google_auth_oauthlib.flow.InstalledAppFlow` 사용).

---

## 실행 스크립트

### `scripts/pptx_inspect.py` (신규)

```python
"""PPTX 파일을 분석해 리포트 출력."""
import sys
from pathlib import Path
from src.integrations.pptx_inspector import PPTXInspector

if __name__ == "__main__":
    target = Path(sys.argv[1])
    report = PPTXInspector(target).inspect()
    print(report.summary())
    report.save_json(target.with_suffix(".inspect.json"))
```

### `scripts/pptx_optimize_and_upload.py` (신규)

```python
"""최적화 + 업로드 원스텝."""
import sys
from pathlib import Path
from src.integrations.pptx_optimizer import LosslessPPTXOptimizer
from src.integrations.google_slides_bridge import GoogleSlidesUploader

def main():
    src = Path(sys.argv[1])
    dst = src.with_name(src.stem + "_optimized.pptx")

    # 1) 최적화
    opt = LosslessPPTXOptimizer(
        max_image_dim=2400,
        use_oxipng=True,
        strip_fonts=True,
        strip_metadata=True,
    )
    stats = opt.optimize(src, dst)
    print(f"원본 {stats['original_mb']:.1f}MB → "
          f"최적화 {stats['optimized_mb']:.1f}MB "
          f"({(1 - stats['ratio']) * 100:.0f}% 감소)")

    # 2) 업로드
    uploader = GoogleSlidesUploader(
        Path("config/google_service_account.json"),
    )
    result = uploader.upload_and_convert(
        dst,
        title=src.stem,
        folder_id=None,  # 필요시 Drive 폴더 ID
        on_progress=lambda p: print(f"  업로드 {p * 100:.0f}%"),
    )
    print(f"완료: {result['webViewLink']}")

if __name__ == "__main__":
    main()
```

---

## 의존성 추가 (`requirements.txt`에 append)

```
google-api-python-client>=2.100.0
google-auth>=2.23.0
google-auth-oauthlib>=1.1.0
Pillow>=10.0.0
```

**바이너리 (선택)**:
- `oxipng` — PNG 무손실 재압축. 없으면 Pillow 폴백. Windows는 `cargo install oxipng` 또는 GitHub 릴리즈 바이너리

---

## 사용자 준비물 (Google Cloud)

1. Google Cloud Console → 새 프로젝트
2. Google Drive API 활성화
3. **서비스 계정** 생성 → JSON 키 다운로드 → `config/google_service_account.json` 에 배치
4. Google Drive에서 업로드 폴더를 서비스 계정 이메일(`xxx@xxx.iam.gserviceaccount.com`)에 **편집자 권한으로 공유**
5. (선택) 폴더 ID를 URL에서 복사해두기

**OAuth 대안**: 서비스 계정 대신 OAuth 사용자 플로우로 하면 개인 Drive에 바로 올라감. `google_slides_bridge.py`에 `from_oauth_credentials()` 클래스메서드 추가하면 됨.

---

## 로컬 작업 순서

```bash
# 1. 최신 받기
cd D:\code\proposal-agent
git pull origin claude/pptx-xml-google-slides-V6iqF

# 2. 의존성 설치
pip install -r requirements.txt

# 3. Claude Code에게 이 문서 읽고 구현 시작시키기:
#    "docs/HANDOFF_GOOGLE_SLIDES_UPLOAD.md 읽고
#     pptx_inspector.py, pptx_optimizer.py, google_slides_bridge.py
#     3개 파일과 2개 스크립트를 구현해줘"

# 4. 구현 끝나면 분석부터
python scripts/pptx_inspect.py "input/2026 NIKKE/원본.pptx"

# 5. 리포트 보고 파라미터 확정 후 업로드
python scripts/pptx_optimize_and_upload.py "input/2026 NIKKE/원본.pptx"

# 6. 구글 슬라이드에서 "더 많은 글꼴" → Pretendard 추가 (최초 1회)
```

---

## 체크리스트

### 구현 단계
- [ ] `src/integrations/pptx_inspector.py` 구현
- [ ] `src/integrations/pptx_optimizer.py` 구현 (`LosslessPPTXOptimizer`)
- [ ] `src/integrations/google_slides_bridge.py` 구현 (`GoogleSlidesUploader`)
- [ ] `scripts/pptx_inspect.py` 작성
- [ ] `scripts/pptx_optimize_and_upload.py` 작성
- [ ] `requirements.txt` 의존성 추가
- [ ] `.gitignore` 에 `config/google_service_account.json` 추가

### 사용자 준비 단계
- [ ] Google Cloud 서비스 계정 JSON 발급 → `config/` 에 배치
- [ ] Google Drive 폴더 공유 (서비스 계정 이메일에 편집자)
- [ ] 구글 슬라이드에 Pretendard 추가 (계정당 1회)
- [ ] (선택) oxipng 설치

### 검증 단계
- [ ] `pptx_inspect.py`로 원본 분석 리포트 확인
- [ ] 이미지 중복/과잉 해상도/EMF 존재 여부 파악
- [ ] 파라미터 확정 (특히 `max_image_dim`, `dedupe_media`)
- [ ] 최적화 실행 → 결과 크기가 100MB 이하인지 확인
- [ ] 최적화 PPTX를 로컬 PowerPoint로 열어 레이아웃 보존 확인
- [ ] 업로드 실행
- [ ] 구글 슬라이드에서 열어 폰트/이미지/레이아웃 검증

---

## 주의사항 (레이아웃 보존 위반 방지)

1. **`ppt/slides/*.xml` 절대 수정 금지** — 도형 좌표/크기는 여기 저장됨
2. **이미지 치환 시 바이트만 교체**, 파일명(Part Name)과 확장자 유지
3. **PNG→JPEG 같은 포맷 변환 금지** — rels 참조가 깨짐
4. **`ppt/slideLayouts/`, `ppt/slideMasters/`, `ppt/theme/` 수정 금지**
5. **최적화 후 python-pptx로 열어 `prs.slides[n].shapes` 순회가 정상인지 smoke test**

---

## 상세 설계 참조

더 자세한 배경과 대안 분석은 `docs/PPTX_GOOGLE_SLIDES_UPLOAD_PLAN.md` 참조:
- 왜 XML 변환이 답이 아닌가
- 무손실 Layer A vs Layer B 차이
- 3가지 경로 비교 (Strict Lossless / Display-Adaptive / 외부 호스팅)
- 폰트 전략 3가지
- 풀 사이즈 `LosslessPPTXOptimizer` 코드 블록
- `GoogleSlidesUploader` 코드 블록
