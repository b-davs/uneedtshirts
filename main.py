from __future__ import annotations

import logging
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from threading import Thread
from tkinter import messagebox

from config import ConfigError, load_runtime_config, runtime_base_dir
from logging_setup import setup_logging
from models import AppConfig
from storage import ensure_schema, seed_clients_from_csv_if_needed
from ui_main import MainWindow
from updater import check_for_update_async

WATCHER_EXE_NAME = "BizactivityWatcher.exe"


def _start_watcher_if_needed(base_dir: Path, logger: logging.Logger) -> None:
    """Start BizactivityWatcher.exe if not already running. Windows only."""
    if sys.platform != "win32":
        return

    watcher_path = base_dir / WATCHER_EXE_NAME
    if not watcher_path.exists():
        logger.info("Watcher exe not found at %s, skipping", watcher_path)
        return

    # Check if already running
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {WATCHER_EXE_NAME}", "/NH"],
            capture_output=True, text=True, timeout=5,
        )
        if WATCHER_EXE_NAME.lower() in result.stdout.lower():
            logger.info("Watcher already running, skipping start")
            return
    except Exception:
        logger.exception("Failed to check watcher process status")

    # Start detached — survives launcher close
    try:
        CREATE_NO_WINDOW = 0x08000000
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            [str(watcher_path)],
            creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
            close_fds=True,
        )
        logger.info("Started watcher: %s", watcher_path)
    except Exception:
        logger.exception("Failed to start watcher")


def _sync_bizactivity_async(config: AppConfig, logger: logging.Logger) -> None:
    """Run bizactivity sync in a background thread on launch."""
    if not config.bizactivity_path:
        return

    def _sync() -> None:
        from bizactivity import sync_all_to_bizactivity

        try:
            counts = sync_all_to_bizactivity(
                config.root_paths.clients_root,
                config.bizactivity_path,
                logger=logger,
            )
            logger.info("Bizactivity launch sync: %s", counts)
        except Exception:
            logger.exception("Bizactivity launch sync failed")

    Thread(target=_sync, daemon=True).start()


def main() -> None:
    logger = setup_logging()
    try:
        config, config_path = load_runtime_config()
    except ConfigError as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Configuration Error", str(exc))
        return

    ensure_schema()
    csv_path = runtime_base_dir() / "clients.csv"
    try:
        report = seed_clients_from_csv_if_needed(config, csv_path)
        if report.created or report.updated or report.warnings:
            logger.info("Client seed report: %s", report.summary())
            for warning in report.warnings[:20]:
                logger.warning(warning)
    except Exception:
        logger.exception("Client bootstrap failed. App will continue to launch.")

    root = tk.Tk()
    MainWindow(root, config, config_path, logger)
    check_for_update_async(root, runtime_base_dir(), logger)
    _sync_bizactivity_async(config, logger)
    _start_watcher_if_needed(runtime_base_dir(), logger)
    root.mainloop()


if __name__ == "__main__":
    main()
