"""Bizactivity file watcher — monitors Whole Job Docs for changes and syncs to bizactivity.

Runs as a standalone background process on Dan's Windows 11 machine.
Uses watchdog to detect file saves, debounces rapid events, then reads
the Map sheet and writes to the bizactivity workbook via COM.

Usage:
    python watcher.py              # Run in foreground (Ctrl+C to stop)
    pythonw watcher.py             # Run hidden (no console window)

Reads config from config.json in the same directory as the script.
Requires: clients_root and bizactivity_path to be set in config.
"""

from __future__ import annotations

import logging
import re
import sys
import time
from pathlib import Path
from threading import Lock, Timer
from typing import Optional

from watchdog.events import FileSystemEvent, FileSystemEventHandler  # type: ignore[import-not-found]
from watchdog.observers import Observer  # type: ignore[import-not-found]

from bizactivity import read_map_sheet, write_job_to_bizactivity
from config import ConfigError, load_runtime_config
from logging_setup import setup_logging

# Seconds to wait after last file event before processing.
# Gives Excel time to finish writing and release the file lock.
DEBOUNCE_SECONDS = 5.0

# Pattern for order workbook filenames: U-ABBR-SEQ.xls(m|x)
_WORKBOOK_PATTERN = re.compile(r"^U-.+\.(xls|xlsm|xlsx)$", re.IGNORECASE)

# File extensions we care about
_EXCEL_EXTENSIONS = {".xls", ".xlsm", ".xlsx"}

# Ignore Excel temp/lock files
_IGNORE_PREFIXES = ("~$", "~")


class _DebouncedHandler(FileSystemEventHandler):
    """Watches for Excel file modifications and debounces rapid events.

    When a file is saved, Excel can trigger multiple events in quick
    succession (temp file create, write, rename). This handler collects
    events per file and only processes after DEBOUNCE_SECONDS of quiet.
    """

    def __init__(self, bizactivity_path: str, logger: logging.Logger) -> None:
        super().__init__()
        self._bizactivity_path = bizactivity_path
        self._logger = logger
        self._timers: dict[str, Timer] = {}
        self._lock = Lock()

    def on_modified(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        self._schedule(event.src_path)

    def on_created(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        self._schedule(event.src_path)

    def _schedule(self, file_path: str) -> None:
        path = Path(file_path)

        # Skip non-Excel files
        if path.suffix.lower() not in _EXCEL_EXTENSIONS:
            return

        # Skip Excel temp/lock files
        if any(path.name.startswith(p) for p in _IGNORE_PREFIXES):
            return

        # Only process files matching U-ABBR-SEQ pattern
        if not _WORKBOOK_PATTERN.match(path.name):
            return

        # Must be inside a U-* order folder
        if not path.parent.name.startswith("U-"):
            return

        normalized = str(path.resolve())

        with self._lock:
            existing = self._timers.get(normalized)
            if existing is not None:
                existing.cancel()

            timer = Timer(DEBOUNCE_SECONDS, self._process, args=[normalized])
            timer.daemon = True
            self._timers[normalized] = timer
            timer.start()

    def _process(self, file_path: str) -> None:
        with self._lock:
            self._timers.pop(file_path, None)

        self._logger.info("Detected change: %s", file_path)

        try:
            values = read_map_sheet(file_path, logger=self._logger)
            if values is None:
                self._logger.info("Skipped %s (no valid Map data)", file_path)
                return

            result = write_job_to_bizactivity(
                self._bizactivity_path, values, logger=self._logger
            )
            if result.success:
                self._logger.info(
                    "Synced job %s -> row %d (month %d, action=%s)",
                    values.get("job_number", "?"),
                    result.target_row or 0,
                    result.month or 0,
                    result.action,
                )
            else:
                self._logger.warning(
                    "Sync failed for %s: %s",
                    values.get("job_number", "?"),
                    result.error_message,
                )
        except Exception:
            self._logger.exception("Error processing %s", file_path)


def run_watcher(
    clients_root: str,
    bizactivity_path: str,
    logger: logging.Logger,
) -> None:
    """Start the file watcher. Blocks until interrupted."""
    root_path = Path(clients_root)
    if not root_path.exists():
        logger.error("clients_root does not exist: %s", clients_root)
        return

    biz_path = Path(bizactivity_path)
    if not biz_path.exists():
        logger.error("bizactivity file does not exist: %s", bizactivity_path)
        return

    handler = _DebouncedHandler(bizactivity_path, logger)
    observer = Observer()
    observer.schedule(handler, str(root_path), recursive=True)
    observer.start()

    logger.info("Watcher started. Monitoring: %s", clients_root)
    logger.info("Syncing to: %s", bizactivity_path)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("Watcher stopping...")
    finally:
        observer.stop()
        observer.join()
        logger.info("Watcher stopped.")


def main() -> None:
    logger = setup_logging()
    logger.info("Bizactivity watcher starting...")

    try:
        config, _ = load_runtime_config()
    except ConfigError as exc:
        logger.error("Configuration error: %s", exc)
        print(f"Configuration error: {exc}", file=sys.stderr)
        sys.exit(1)

    if not config.bizactivity_path:
        logger.error("bizactivity_path is not set in config.json")
        print("Error: bizactivity_path is not set in config.json", file=sys.stderr)
        sys.exit(1)

    run_watcher(
        config.root_paths.clients_root,
        config.bizactivity_path,
        logger,
    )


if __name__ == "__main__":
    main()
