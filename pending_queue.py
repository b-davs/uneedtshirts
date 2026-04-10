"""Pending bizactivity sync queue.

When the bizactivity workbook is locked (Dan has it open in Excel),
write_job_to_bizactivity cannot land its write. Instead of dropping the
event, we serialize the payload to a JSON file and drain it later when
the lock clears.

Dedup policy: last-write-wins keyed by job_number. If two saves for the
same job happen while bizactivity is open, only the most recent payload
is retained.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Any, Optional


_QUEUE_FILENAME = "pending_syncs.json"
_queue_lock = Lock()


def _queue_path() -> Path:
    from storage import get_app_data_dir
    return get_app_data_dir() / _QUEUE_FILENAME


def _load(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return [entry for entry in data if isinstance(entry, dict)]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def _save(path: Path, entries: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(entries, f, default=str, indent=2)
    tmp_path.replace(path)


def enqueue(
    values: dict[str, Any],
    *,
    queue_path: Optional[Path] = None,
    logger: Optional[logging.Logger] = None,
) -> None:
    """Append a sync payload to the queue, replacing any prior entry
    for the same job_number (last write wins)."""
    job_number = values.get("job_number")
    if not job_number or not str(job_number).strip():
        return
    job_number = str(job_number).strip()

    path = queue_path or _queue_path()
    with _queue_lock:
        entries = _load(path)
        entries = [
            e for e in entries
            if str(e.get("values", {}).get("job_number", "")).strip() != job_number
        ]
        entries.append({
            "enqueued_at": datetime.now().isoformat(timespec="seconds"),
            "values": values,
        })
        _save(path, entries)

    if logger:
        logger.info(
            "Pending queue: enqueued job %s (depth=%d)",
            job_number, len(entries),
        )


def peek(*, queue_path: Optional[Path] = None) -> list[dict[str, Any]]:
    """Return a snapshot of queue entries without modifying the queue."""
    path = queue_path or _queue_path()
    with _queue_lock:
        return _load(path)


def size(*, queue_path: Optional[Path] = None) -> int:
    return len(peek(queue_path=queue_path))


def clear(*, queue_path: Optional[Path] = None) -> None:
    path = queue_path or _queue_path()
    with _queue_lock:
        if path.exists():
            path.unlink()


def drain(
    writer: Any,
    bizactivity_path: str,
    *,
    queue_path: Optional[Path] = None,
    logger: Optional[logging.Logger] = None,
) -> dict[str, int]:
    """Drain all queued payloads through `writer(bizactivity_path, values, logger=...)`.

    `writer` should return an object with a `.success` bool attribute
    (BizactivityResult-shaped). Successful entries are removed; failed
    entries remain for the next drain attempt.
    """
    path = queue_path or _queue_path()
    counts = {"drained": 0, "failed": 0, "remaining": 0}

    # Atomically claim the current batch: load entries AND clear the file.
    # Any events enqueued after this point will land in a fresh file and
    # be merged back in at the end if there are failed entries to persist.
    with _queue_lock:
        entries = _load(path)
        if entries and path.exists():
            path.unlink()

    if not entries:
        return counts

    remaining: list[dict[str, Any]] = []
    for entry in entries:
        values = entry.get("values") or {}
        try:
            result = writer(bizactivity_path, values, logger=logger)
            if getattr(result, "success", False):
                counts["drained"] += 1
            else:
                counts["failed"] += 1
                remaining.append(entry)
        except Exception:
            if logger:
                logger.exception(
                    "Pending queue: drain failed for job %s",
                    values.get("job_number", "?"),
                )
            counts["failed"] += 1
            remaining.append(entry)

    with _queue_lock:
        # Merge any entries added mid-drain (fresh file state) with our
        # failed-to-drain entries. Dedup by job_number, last-write-wins.
        current = _load(path)
        merged_by_job: dict[str, dict[str, Any]] = {}
        for e in remaining + current:
            jn = str(e.get("values", {}).get("job_number", "")).strip()
            if jn:
                merged_by_job[jn] = e
        final_entries = list(merged_by_job.values())
        if final_entries:
            _save(path, final_entries)
        elif path.exists():
            path.unlink()
        counts["remaining"] = len(final_entries)

    if logger and (counts["drained"] or counts["failed"]):
        logger.info(
            "Pending queue: drained=%d failed=%d remaining=%d",
            counts["drained"], counts["failed"], counts["remaining"],
        )

    return counts
