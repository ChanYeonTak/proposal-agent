# PPTX → Google Slides 최적화 업로드 시스템 계획서

**브랜치**: `claude/pptx-xml-google-slides-V6iqF`
**목표**: slide_kit으로 생성된 마케팅 제안서 PPTX를 Google Slides에 **편집 가능 + 레이아웃 보존 + 이미지 시각적 무손실** 상태로 업로드

---

## 배경 — PPTX를 Google Slides에 올릴 때 깨지는 이유

### 왜 XML 변환은 답이 아닌가
- `.pptx`는 이미 XML (OOXML, ZIP 아카이브 내부에 `ppt/slides/slide1.xml` 등 존재)
- Google Slides는 OOXML XML을 직접 import 하는 경로 없음
- Import 경로는 오직 `.pptx` 바이너리(ZIP) 포맷만 지원

### 실제 호환성 문제
1. **폰트 (Pretendard)**: Google Workspace 기본 폰트에 없어 Arial 등으로 대체 → 줄바꿈/너비 어긋남
2. **그라디언트 / 그림자 / 라운드 박스**: OOXML 파라미터 일부가 Google Slides 렌더러에서 무시됨
3. **차트**: 네이티브 차트는 import되지만 Google Slides에서 **편집 불가 객체**로 고정
4. **이미지 재인코딩**: Google은 import 시 이미지를 자체 포맷으로 재인코딩 (주로 JPEG, 최대 1920×1080)
5. **파일 용량**: Drive 단순 업로드는 5MB, PPTX→Slides 자동 변환 한계 ~100MB

---

## 요구사항 확정 (사용자 합의)

| # | 요구사항 | 구현 전략 |
|---|---------|----------|
| 1 | **편집 가능** | 텍스트/도형 네이티브 유지. Pretendard는 Google Fonts "더 많은 글꼴" 경로. 차트는 분석 후 결정 |
| 2 | **전체 업로드** | Drive API Resumable Upload (8MB 청크, 재개 지원, 진행률 표시) |
| 3 | **이미지 시각적 무손실** | PNG 무손실 재압축(oxipng) + JPEG 원본 유지 + 과잉 해상도만 2400px로 다운샘플 |
| 4 | **레이아웃 그대로** | 좌표/크기/회전/z-order/마진 절대 변경 금지. 이미지 교체는 동일 치수로만 |

### "무손실"의 정의 — Layer B (시각적 무손실)
- **Layer A (비트 무손실)**: 픽셀값 1개도 안 바뀜 — 현실적으로 불가능 (구글이 import 시 재인코딩)
- **Layer B (시각적 무손실)**: 보는 사람이 원본과 구분 불가 — 본 프로젝트 채택
  - 구글 슬라이드 최대 렌더링 해상도 1920×1080 기준으로 2x 여유(2400px)까지만 유지
  - 2400px 초과분은 Lanczos 다운샘플 (화면에 절대 보이지 않는 픽셀)
  - 구글이 어차피 재인코딩하므로 여기서 손실 발생 여부는 사용자 경험에 영향 없음

### 레이아웃 보존이 특히 중요한 지점
- 이미지 리사이즈 시 **슬라이드 내 표시 크기(`cx`, `cy`)는 건드리지 않음** — 픽셀 해상도만 줄이고 도형 박스는 그대로
- 폰트 strip 후에도 텍스트박스 위치/크기 불변 (Google Fonts Pretendard가 같은 metric으로 렌더링)
- ZIP 재압축 시 **XML 내용은 한 글자도 수정 안 함** — 오직 `ppt/media/`와 `ppt/fonts/` 폴더만 건드림
- Dedupe 적용 시에도 rels 재작성으로 **참조만 병합**, 슬라이드 상의 배치는 불변

---

## 폰트 전략: Pretendard를 Google Slides에서 그대로 쓰기

### 사실 관계
- Google Slides는 **임의 웹폰트 @font-face 로딩 불가** (CSS 주입 훅 없음)
- PPTX의 폰트 임베딩은 Google Slides가 **무시**함
- 사용 가능한 폰트 풀: 기본 내장 + Google Fonts 등재 폰트 + Workspace 관리자 업로드 커스텀 폰트

### 권장 경로: Google Fonts의 Pretendard (1순위)
Pretendard는 Google Fonts에 등재되어 있음.

1. 구글 슬라이드 → 폰트 드롭다운 → **"더 많은 글꼴"**
2. 검색창에 `Pretendard` 입력 → 체크 → 확인
3. 이후 이 계정의 모든 프레젠테이션에서 Pretendard 사용 가능
4. PPTX import 시 폰트명 `Pretendard`가 **그대로 매칭**되어 렌더링
5. **PPTX 쪽에서는 아무것도 바꿀 필요 없음** — `slide_kit`의 `FONT = "Pretendard"` 그대로 유지

### 대안
- **Google Workspace 커스텀 폰트 업로드**: 조직 관리자가 Admin Console에서 Pretendard `.ttf/.otf` 업로드 → 조직 전체 배포
- **Noto Sans KR 폴백**: Pretendard가 Inter + Noto Sans KR 기반이라 거의 구분 안 감 (최종 폴백용)

---

## 용량 최적화 전략

### 분석 결과 (일반)

| 요소 | 일반 비중 | 최적화 후 |
|------|---------|----------|
| 이미지 (Pexels/DALL-E 원본) | 70~85% | 50~70% 감소 |
| 임베드 폰트 (Pretendard 하위셋) | 5~10% | 100% 제거 (구글이 무시) |
| XML / 도형 정의 | 5~10% | 변화 없음 |
| 차트/테마 리소스 | 2~5% | 변화 없음 |

→ **이미지만 제대로 다루면 80MB → 15~30MB로 줄어듦**

### LosslessPPTXOptimizer 설계

```python
# src/integrations/pptx_optimizer.py
import subprocess, zipfile, shutil
from pathlib import Path
from io import BytesIO
from PIL import Image

class LosslessPPTXOptimizer:
    """시각적 무손실 최적화. 레이아웃/도형/텍스트 불변."""

    def __init__(self,
                 max_image_dim: int = 2400,   # 구글 렌더링 1920px의 1.25x 여유
                 use_oxipng: bool = True,
                 strip_fonts: bool = True,
                 strip_metadata: bool = True,
                 dedupe_media: bool = True,
                 zip_level: int = 9):
        self.max_image_dim = max_image_dim
        self.use_oxipng = use_oxipng and shutil.which("oxipng") is not None
        self.strip_fonts = strip_fonts
        self.strip_metadata = strip_metadata
        self.dedupe = dedupe_media
        self.zip_level = zip_level

    def optimize(self, src: Path, dst: Path) -> dict:
        """PPTX ZIP을 재작성. 통계 반환."""
        stats = {
            "original_mb": src.stat().st_size / 1e6,
            "images_processed": 0,
            "png_recompressed": 0,
            "jpeg_preserved": 0,
            "fonts_stripped": 0,
            "bytes_saved": 0,
        }

        with zipfile.ZipFile(src, "r") as zin, \
             zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED,
                             compresslevel=self.zip_level) as zout:

            for entry in zin.infolist():
                name = entry.filename
                data = zin.read(name)
                original_size = len(data)

                # 1) 임베드 폰트 제거 (lossless — 구글이 무시)
                if self.strip_fonts and (
                    name.startswith("ppt/fonts/") or
                    name.startswith("ppt/embeddings/fonts/")
                ):
                    stats["fonts_stripped"] += 1
                    stats["bytes_saved"] += original_size
                    continue

                # 2) 이미지 무손실 처리
                if name.startswith("ppt/media/"):
                    new_data = self._process_image(data, name)
                    if new_data is not None and len(new_data) < original_size:
                        stats["bytes_saved"] += original_size - len(new_data)
                        data = new_data
                        stats["images_processed"] += 1

                zout.writestr(entry, data,
                              compress_type=zipfile.ZIP_DEFLATED,
                              compresslevel=self.zip_level)

        stats["optimized_mb"] = dst.stat().st_size / 1e6
        stats["ratio"] = stats["optimized_mb"] / stats["original_mb"]
        return stats

    def _process_image(self, data: bytes, name: str) -> bytes | None:
        lower = name.lower()

        try:
            img = Image.open(BytesIO(data))
        except Exception:
            return None

        # 과잉 해상도만 다운샘플 (시각적 무손실)
        w, h = img.size
        if max(w, h) > self.max_image_dim:
            ratio = self.max_image_dim / max(w, h)
            img = img.resize(
                (int(w * ratio), int(h * ratio)),
                Image.LANCZOS,
            )
        elif lower.endswith((".jpg", ".jpeg")):
            # JPEG 원본 크기면 건드리지 않음 (재인코딩 손실 방지)
            return None

        buf = BytesIO()
        if lower.endswith(".png"):
            # PNG 무손실 재압축
            img.save(buf, format="PNG", optimize=True, compress_level=9)
            out = buf.getvalue()
            if self.use_oxipng:
                out = self._oxipng_lossless(out) or out
            return out
        elif lower.endswith((".jpg", ".jpeg")):
            # 리사이즈된 경우만 JPEG 재저장 (quality 95 = 시각적 무손실)
            img.convert("RGB").save(buf, format="JPEG",
                                     quality=95, optimize=True,
                                     progressive=True, subsampling=0)
            return buf.getvalue()
        return None

    def _oxipng_lossless(self, png_bytes: bytes) -> bytes | None:
        try:
            result = subprocess.run(
                ["oxipng", "--opt", "max", "--strip", "safe", "-"],
                input=png_bytes, capture_output=True, timeout=60,
            )
            if result.returncode == 0 and len(result.stdout) < len(png_bytes):
                return result.stdout
        except Exception:
            pass
        return None
```

### Dedupe (2차 작업)
동일 해시 이미지가 여러 번 임베드되는 경우 `ppt/slides/_rels/*.xml.rels`의 `Target` 참조를 재작성해 하나로 병합. 1차 구현 후 분석 결과 필요 시 추가.

---

## Resumable 청크 업로드 설계

### Google Drive API 한계
| 제한 | 값 |
|------|---|
| 단순 업로드 | 5MB |
| Resumable 업로드 | **5TB** |
| PPTX→Slides 자동 변환 | **~100MB** |

### GoogleSlidesUploader 설계

```python
# src/integrations/google_slides_bridge.py
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

CHUNK_SIZE = 8 * 1024 * 1024  # 8MB

class GoogleSlidesUploader:
    def __init__(self, credentials_json: Path):
        creds = service_account.Credentials.from_service_account_file(
            str(credentials_json),
            scopes=["https://www.googleapis.com/auth/drive.file"],
        )
        self.drive = build("drive", "v3", credentials=creds)

    def upload_and_convert(self,
                           pptx_path: Path,
                           title: str,
                           folder_id: str | None = None,
                           on_progress=None) -> dict:
        """Resumable 업로드 + Slides 변환 트리거. webViewLink 반환."""
        metadata = {
            "name": title,
            # ★ 이 mimeType이 Slides 변환 트리거
            "mimeType": "application/vnd.google-apps.presentation",
        }
        if folder_id:
            metadata["parents"] = [folder_id]

        media = MediaFileUpload(
            str(pptx_path),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            chunksize=CHUNK_SIZE,
            resumable=True,  # ★ 핵심
        )

        request = self.drive.files().create(
            body=metadata, media_body=media,
            fields="id,name,webViewLink,size",
        )

        response = None
        while response is None:
            status, response = request.next_chunk()
            if status and on_progress:
                on_progress(status.progress())  # 0.0 ~ 1.0

        return response
```

### 원스텝 파이프라인

```python
# output/테스트 XX/upload_to_google.py
from pathlib import Path
from src.integrations.pptx_optimizer import LosslessPPTXOptimizer
from src.integrations.google_slides_bridge import GoogleSlidesUploader

src = Path("output/NIKKE_AGF_2025/NIKKE_AGF_2025.pptx")
optimized = src.with_name(src.stem + "_optimized.pptx")

# 1) 시각적 무손실 최적화
opt = LosslessPPTXOptimizer(
    max_image_dim=2400,
    use_oxipng=True,
    strip_fonts=True,
    strip_metadata=True,
)
stats = opt.optimize(src, optimized)
print(f"원본: {stats['original_mb']:.1f}MB → 최적화: {stats['optimized_mb']:.1f}MB "
      f"({(1-stats['ratio'])*100:.0f}% 감소)")

# 2) Resumable 업로드 + Slides 변환
uploader = GoogleSlidesUploader(Path("config/google_service_account.json"))
result = uploader.upload_and_convert(
    optimized,
    title="NIKKE AGF 2025 제안서",
    folder_id="YOUR_DRIVE_FOLDER_ID",
    on_progress=lambda p: print(f"  업로드 중... {p*100:.0f}%"),
)
print(f"완료: {result['webViewLink']}")
```

---

## PPTX 분석 (실행 직전 필수)

원본 파일 수령 시 `PPTXInspector`로 아래 항목을 자동 리포트:

1. **파일 크기 + 슬라이드 수** — 압축 목표치 계산
2. **이미지 인벤토리** (`ppt/media/` 내부)
   - 개수, 포맷별 분포 (PNG / JPEG / WebP / EMF)
   - 각 이미지 해상도 및 용량
   - 슬라이드 내 실제 표시 크기 대비 원본 해상도 비율 (과잉 픽셀 계산)
   - 중복 이미지 해시 검사
3. **폰트 임베딩 상태** (`ppt/fonts/`) — Pretendard 서브셋 용량
4. **차트/임베드 오브젝트** (`ppt/embeddings/`) — Excel 링크 등
5. **XML 블로트** — 미사용 슬라이드 레이아웃, 테마 리소스
6. **ZIP 압축 수준** — 재압축 여유
7. **EMF/WMF 벡터 존재 여부** — Google Slides에서 깨지는 주범

분석 리포트 기반으로 다음 파라미터 최종 확정:

```python
LosslessPPTXOptimizer(
    max_image_dim=???,        # 분석 결과 기반 (2400 기본, 이미지 분포 보고 조정)
    use_oxipng=True,          # PNG 무손실 재압축
    strip_fonts=True,         # Pretendard 임베드 제거 (Google Fonts 경로로 대체)
    strip_metadata=True,      # EXIF/색프로파일 정리
    dedupe_media=???,         # 중복 이미지 있으면 True
    convert_emf_to_png=???,   # EMF/WMF 있으면 True (렌더링 호환)
    zip_level=9,
)
```

---

## 실행 순서

```
[대기] 원본 PPTX 파일 수령
  ↓
[1] PPTXInspector 실행 → 분석 리포트 출력
  ↓
[2] 리포트 기반으로 LosslessPPTXOptimizer 파라미터 확정
  ↓
[3] 최적화 실행 → {filename}_optimized.pptx 생성
  ↓
[4] 예상 용량 확인 (목표: 100MB 이하)
  ↓
[5] GoogleSlidesUploader로 Resumable 업로드
  ↓
[6] Drive API가 Slides 포맷으로 자동 변환
  ↓
[7] webViewLink 반환 → 사용자에게 전달
  ↓
[8] (사용자) 구글 슬라이드에서 "더 많은 글꼴" → Pretendard 추가 (계정당 1회)
```

---

## 사전 준비물 (사용자 액션)

1. **Google Cloud 프로젝트** 생성 → Drive API 활성화
2. **서비스 계정** 생성 → JSON 키 다운로드 → `config/google_service_account.json`
   - 또는 OAuth 사용자 인증 플로우 (선택 가능)
3. Google Drive에서 업로드할 폴더를 서비스 계정 이메일에 **공유** (Editor 권한)
4. 의존성 설치:
   ```bash
   pip install google-api-python-client google-auth Pillow
   # oxipng 바이너리 (선택, 있으면 PNG 압축 품질 향상)
   # macOS: brew install oxipng
   # Ubuntu: cargo install oxipng
   ```
5. 구글 슬라이드 계정에서 "더 많은 글꼴" → Pretendard 추가 (최초 1회)

---

## 산출물 파일 구조 (구현 예정)

```
src/integrations/
├── pptx_inspector.py          # 신규: 원본 PPTX 분석 리포트 생성
├── pptx_optimizer.py          # 신규: LosslessPPTXOptimizer (시각적 무손실)
└── google_slides_bridge.py    # 신규: GoogleSlidesUploader (Resumable)

config/
└── google_service_account.json  # 사용자 준비 (gitignore)

docs/
└── PPTX_GOOGLE_SLIDES_UPLOAD_PLAN.md  # 본 문서
```

---

## 대기 사항

- [ ] **원본 PPTX 파일 수령** — 받으면 즉시 분석 → 파라미터 확정 → 구현 → 테스트 → 업로드까지 한 번에 진행
- [ ] Google Cloud 서비스 계정 JSON 준비 상태 확인
- [ ] 업로드 대상 Drive 폴더 ID (있으면 바로 연결, 없으면 루트)
