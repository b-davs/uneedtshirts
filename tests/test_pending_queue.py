from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

import pending_queue


@dataclass
class FakeResult:
    success: bool


class TestEnqueue:
    def test_enqueue_creates_file_and_entry(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue(
            {"job_number": "A-1", "client": "A"},
            queue_path=queue_file,
        )
        entries = pending_queue.peek(queue_path=queue_file)
        assert len(entries) == 1
        assert entries[0]["values"]["job_number"] == "A-1"
        assert "enqueued_at" in entries[0]

    def test_enqueue_skips_blank_job_number(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue({"job_number": "  "}, queue_path=queue_file)
        pending_queue.enqueue({"job_number": None}, queue_path=queue_file)
        assert pending_queue.size(queue_path=queue_file) == 0

    def test_last_write_wins_dedup_by_job_number(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue(
            {"job_number": "A-1", "deposit": 100},
            queue_path=queue_file,
        )
        pending_queue.enqueue(
            {"job_number": "A-1", "deposit": 250},
            queue_path=queue_file,
        )
        pending_queue.enqueue(
            {"job_number": "B-2", "deposit": 50},
            queue_path=queue_file,
        )
        entries = pending_queue.peek(queue_path=queue_file)
        assert len(entries) == 2
        by_job = {e["values"]["job_number"]: e for e in entries}
        assert by_job["A-1"]["values"]["deposit"] == 250
        assert by_job["B-2"]["values"]["deposit"] == 50


class TestClear:
    def test_clear_removes_file(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue({"job_number": "X-1"}, queue_path=queue_file)
        pending_queue.clear(queue_path=queue_file)
        assert not queue_file.exists()
        assert pending_queue.size(queue_path=queue_file) == 0


class TestDrain:
    def test_drain_empty_queue_noop(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"

        def writer(
            path: str, values: dict[str, Any],
            source_path: Any = None, logger: Any = None,
        ) -> FakeResult:
            return FakeResult(success=True)

        counts = pending_queue.drain(writer, "/fake/biz.xlsx", queue_path=queue_file)
        assert counts == {"drained": 0, "failed": 0, "remaining": 0}

    def test_drain_flushes_successful_entries(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue({"job_number": "A-1"}, queue_path=queue_file)
        pending_queue.enqueue({"job_number": "B-2"}, queue_path=queue_file)
        calls: list[str] = []

        def writer(
            path: str, values: dict[str, Any],
            source_path: Any = None, logger: Any = None,
        ) -> FakeResult:
            calls.append(values["job_number"])
            return FakeResult(success=True)

        counts = pending_queue.drain(writer, "/fake/biz.xlsx", queue_path=queue_file)
        assert counts["drained"] == 2
        assert counts["failed"] == 0
        assert counts["remaining"] == 0
        assert set(calls) == {"A-1", "B-2"}
        assert not queue_file.exists()

    def test_drain_keeps_failed_entries(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue({"job_number": "OK-1"}, queue_path=queue_file)
        pending_queue.enqueue({"job_number": "FAIL-1"}, queue_path=queue_file)

        def writer(
            path: str, values: dict[str, Any],
            source_path: Any = None, logger: Any = None,
        ) -> FakeResult:
            return FakeResult(success=values["job_number"] != "FAIL-1")

        counts = pending_queue.drain(writer, "/fake/biz.xlsx", queue_path=queue_file)
        assert counts["drained"] == 1
        assert counts["failed"] == 1
        assert counts["remaining"] == 1
        remaining = pending_queue.peek(queue_path=queue_file)
        assert len(remaining) == 1
        assert remaining[0]["values"]["job_number"] == "FAIL-1"

    def test_drain_forwards_source_path_to_writer(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue(
            {"job_number": "HL-1"},
            source_path=r"D:\Jobs\U-HL-1.xls",
            queue_path=queue_file,
        )
        captured: list[Any] = []

        def writer(
            path: str, values: dict[str, Any],
            source_path: Any = None, logger: Any = None,
        ) -> FakeResult:
            captured.append(source_path)
            return FakeResult(success=True)

        counts = pending_queue.drain(writer, "/fake/biz.xlsx", queue_path=queue_file)
        assert counts["drained"] == 1
        assert captured == [r"D:\Jobs\U-HL-1.xls"]

    def test_drain_handles_writer_exception(self, tmp_path: Path) -> None:
        queue_file = tmp_path / "q.json"
        pending_queue.enqueue({"job_number": "BOOM-1"}, queue_path=queue_file)

        def writer(
            path: str, values: dict[str, Any],
            source_path: Any = None, logger: Any = None,
        ) -> FakeResult:
            raise RuntimeError("simulated COM failure")

        counts = pending_queue.drain(writer, "/fake/biz.xlsx", queue_path=queue_file)
        assert counts["drained"] == 0
        assert counts["failed"] == 1
        assert counts["remaining"] == 1
