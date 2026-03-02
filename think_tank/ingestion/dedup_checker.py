"""
중복 체크 (SHA-256 해시)

이미 학습된 파일의 재투입을 방지합니다.
"""

import hashlib
from pathlib import Path

from src.utils.logger import get_logger

logger = get_logger("dedup_checker")

# 64KB 단위로 읽기
CHUNK_SIZE = 65536


def compute_file_hash(file_path: Path) -> str:
    """
    파일의 SHA-256 해시 계산

    Args:
        file_path: 파일 경로

    Returns:
        SHA-256 해시 (hex)
    """
    sha256 = hashlib.sha256()

    with open(file_path, "rb") as f:
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                break
            sha256.update(chunk)

    return sha256.hexdigest()
