from __future__ import annotations

import re
from pathlib import Path

INVALID_WINDOWS_CHARS = r'<>:"/\\|?*'


def sanitize_job_description(description: str) -> str:
    cleaned = description.strip()
    if not cleaned:
        return ""
    pattern = f"[{re.escape(INVALID_WINDOWS_CHARS)}]"
    cleaned = re.sub(pattern, " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def detect_next_sequence(client_path: str, abbr: str) -> int:
    path = Path(client_path)
    if not path.exists():
        return 1

    pattern = re.compile(
        rf"^U-{re.escape(abbr)}-(\d+)(?:\s+.*)?$",
        flags=re.IGNORECASE,
    )

    sequences: list[int] = []
    for child in path.iterdir():
        if not child.is_dir():
            continue
        match = pattern.match(child.name)
        if match:
            sequences.append(int(match.group(1)))

    return (max(sequences) + 1) if sequences else 1


def build_order_folder_name(
    prefix: str,
    abbr: str,
    sequence: int,
    description: str,
) -> str:
    base = f"{prefix}-{abbr}-{sequence}"
    sanitized = sanitize_job_description(description)
    if sanitized:
        return f"{base} {sanitized}"
    return base


def build_workbook_filename(
    prefix: str,
    abbr: str,
    sequence: int,
    extension: str = ".xls",
) -> str:
    ext = extension if extension.startswith(".") else f".{extension}"
    return f"{prefix}-{abbr}-{sequence}{ext}"


def build_folder_job_number(abbr: str, sequence: int) -> str:
    return f"{abbr}-{sequence}"
