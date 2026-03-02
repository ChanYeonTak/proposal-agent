"""
Gamma REST API 클라이언트 (v4.0)

Gamma (https://gamma.app) API를 통해 PPTX를 업로드하고
디자인이 적용된 결과물을 받아옵니다.

Note: Gamma API가 공식 제공되면 구현을 완성합니다.
현재는 인터페이스 정의 + 수동 업로드 안내.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from ..utils.logger import get_logger

logger = get_logger("gamma_api")


class GammaAPI:
    """
    Gamma REST API 클라이언트

    현재 상태: Gamma 공식 API 출시 대기 중.
    API가 제공되면 upload_pptx / get_result 구현 예정.
    """

    def __init__(self):
        self.api_key = os.getenv("GAMMA_API_KEY", "")
        self.base_url = os.getenv("GAMMA_API_URL", "https://api.gamma.app/v1")

    def is_configured(self) -> bool:
        """API 키 설정 여부"""
        return bool(self.api_key)

    async def upload_pptx(self, pptx_path: Path) -> "DesignBridgeResult":
        """
        PPTX 업로드 → Gamma 디자인 적용

        Args:
            pptx_path: PPTX 파일 경로

        Returns:
            DesignBridgeResult
        """
        from .design_bridge import DesignBridgeResult

        if not self.is_configured():
            return DesignBridgeResult(
                success=False,
                provider="gamma",
                error="GAMMA_API_KEY 미설정",
            )

        # TODO: Gamma API 공식 출시 후 구현
        # 현재는 placeholder
        logger.info(f"Gamma API 업로드 예정: {pptx_path}")

        return DesignBridgeResult(
            success=False,
            provider="gamma",
            error=(
                "Gamma API 연동이 준비 중입니다.\n"
                "https://gamma.app 에서 수동으로 업로드해주세요."
            ),
        )

    async def get_result(self, job_id: str) -> Optional[dict]:
        """
        디자인 적용 결과 조회

        Args:
            job_id: 업로드 작업 ID

        Returns:
            결과 딕셔너리 또는 None
        """
        # TODO: 구현 예정
        return None
