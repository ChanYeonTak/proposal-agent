"""Google Drive → Slides 업로드 브릿지.

- Resumable 청크 업로드 (8MB) — 5TB까지 지원
- 업로드 시 mimeType을 google-apps.presentation 으로 지정 → Slides 변환 트리거
- 진행률 콜백 지원
- OAuth 사용자 인증 (개인 Drive 소유권 유지)
"""
from __future__ import annotations

import pickle
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

CHUNK_SIZE = 8 * 1024 * 1024  # 8MB
PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
SLIDES_MIME = "application/vnd.google-apps.presentation"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]


@dataclass
class UploadResult:
    file_id: str
    name: str
    web_view_link: str
    size_bytes: int | None = None


class GoogleSlidesUploader:
    """OAuth 사용자 인증 기반. 첫 실행 시 브라우저로 동의 → token.pickle 저장."""

    def __init__(
        self,
        client_secret_json: Path,
        token_cache: Path = Path("config/google_token.pickle"),
    ):
        creds: UserCredentials | None = None
        if token_cache.exists():
            with token_cache.open("rb") as f:
                creds = pickle.load(f)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(client_secret_json), SCOPES
                )
                creds = flow.run_local_server(port=0)
            token_cache.parent.mkdir(parents=True, exist_ok=True)
            with token_cache.open("wb") as f:
                pickle.dump(creds, f)

        self.drive = build("drive", "v3", credentials=creds)

    def upload(
        self,
        pptx_path: Path,
        title: str,
        folder_id: str | None = None,
        convert_to_slides: bool = True,
        on_progress: Callable[[float], None] | None = None,
    ) -> UploadResult:
        """PPTX를 Drive에 업로드. convert_to_slides=True면 Slides로 변환."""
        metadata: dict = {
            "name": title,
            "mimeType": SLIDES_MIME if convert_to_slides else PPTX_MIME,
        }
        if folder_id:
            metadata["parents"] = [folder_id]

        media = MediaFileUpload(
            str(pptx_path),
            mimetype=PPTX_MIME,
            chunksize=CHUNK_SIZE,
            resumable=True,
        )

        request = self.drive.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,size",
            supportsAllDrives=True,
        )

        response = None
        while response is None:
            status, response = request.next_chunk()
            if status and on_progress:
                on_progress(status.progress())

        if on_progress:
            on_progress(1.0)

        return UploadResult(
            file_id=response["id"],
            name=response.get("name", title),
            web_view_link=response.get("webViewLink", ""),
            size_bytes=int(response["size"]) if response.get("size") else None,
        )


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser()
    p.add_argument("pptx", type=Path)
    p.add_argument("--title", required=True)
    p.add_argument(
        "--creds",
        type=Path,
        default=Path("config/google_oauth_client.json"),
    )
    p.add_argument("--folder", default=None, help="Drive folder ID (optional)")
    p.add_argument(
        "--no-convert",
        action="store_true",
        help="Store as .pptx without converting to Slides",
    )
    args = p.parse_args()

    def progress(pct: float) -> None:
        bar = "#" * int(pct * 30)
        print(f"\r  [{bar:<30}] {pct*100:5.1f}%", end="", flush=True)

    up = GoogleSlidesUploader(args.creds)
    result = up.upload(
        args.pptx,
        title=args.title,
        folder_id=args.folder,
        convert_to_slides=not args.no_convert,
        on_progress=progress,
    )
    print()
    print(f"완료: {result.web_view_link}")
    print(f"파일 ID: {result.file_id}")
