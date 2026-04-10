from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import tempfile
import tkinter as tk
import zipfile
from pathlib import Path
from threading import Thread
from tkinter import messagebox
from typing import Optional
from urllib.request import Request, urlopen

REPO = "b-davs/uneedtshirts"
RELEASES_URL = f"https://api.github.com/repos/{REPO}/releases/latest"
PROTECTED_FILES = {"config.json"}
WATCHER_EXE_NAME = "BizactivityWatcher.exe"


def _current_version(base_dir: Path) -> str:
    version_file = base_dir / "version.txt"
    if version_file.exists():
        return version_file.read_text(encoding="utf-8").strip()
    return "0.0.0"


def _fetch_latest_release() -> Optional[dict]:
    req = Request(RELEASES_URL, headers={"User-Agent": "NewOrderLauncher-Updater"})
    try:
        with urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())
    except Exception:
        return None


def _find_zip_url(release: dict) -> Optional[str]:
    for asset in release.get("assets", []):
        if asset.get("name") == "NewOrderLauncher.zip":
            return asset.get("browser_download_url")
    return None


def _stop_watcher(logger: logging.Logger) -> None:
    """Kill BizactivityWatcher.exe if running, so the exe can be overwritten."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", WATCHER_EXE_NAME],
            capture_output=True, timeout=10,
        )
        logger.info("Stopped %s before update", WATCHER_EXE_NAME)
    except Exception:
        # Not running, or taskkill not available — either way, continue
        pass


def _download_and_apply(zip_url: str, version: str, base_dir: Path, logger: logging.Logger) -> bool:
    try:
        _stop_watcher(logger)
        req = Request(zip_url, headers={"User-Agent": "NewOrderLauncher-Updater"})
        with tempfile.TemporaryDirectory() as tmp:
            zip_path = Path(tmp) / "update.zip"
            extract_dir = Path(tmp) / "contents"

            with urlopen(req, timeout=120) as resp:
                zip_path.write_bytes(resp.read())

            with zipfile.ZipFile(zip_path) as zf:
                zf.extractall(extract_dir)

            for item in extract_dir.iterdir():
                if item.is_file() and item.name not in PROTECTED_FILES:
                    dest = base_dir / item.name
                    dest.write_bytes(item.read_bytes())
                    logger.info("Updated %s", item.name)
                elif item.name in PROTECTED_FILES:
                    logger.info("Skipped %s (protected)", item.name)

            version_file = base_dir / "version.txt"
            version_file.write_text(version, encoding="utf-8")

        return True
    except Exception:
        logger.exception("Update failed")
        return False


def _restart_app() -> None:
    exe = sys.executable
    if getattr(sys, "frozen", False):
        subprocess.Popen([exe])
    else:
        subprocess.Popen([exe] + sys.argv)
    sys.exit(0)


def check_for_update_async(
    root: tk.Tk,
    base_dir: Path,
    logger: logging.Logger,
) -> None:
    """Checks for updates in a background thread. Shows a dialog on the main
    thread if a new version is available."""

    def _check() -> None:
        current = _current_version(base_dir)
        release = _fetch_latest_release()
        if release is None:
            return

        latest = release.get("tag_name", "").lstrip("v")
        if not latest or latest == current:
            return

        zip_url = _find_zip_url(release)
        if not zip_url:
            return

        root.after(0, lambda: _prompt_update(root, base_dir, logger, current, latest, zip_url))

    Thread(target=_check, daemon=True).start()


def _prompt_update(
    root: tk.Tk,
    base_dir: Path,
    logger: logging.Logger,
    current: str,
    latest: str,
    zip_url: str,
) -> None:
    answer = messagebox.askyesno(
        "Update Available",
        f"Version {latest} is available (you have {current}).\n\nUpdate now?",
        parent=root,
    )
    if not answer:
        return

    logger.info("Updating from %s to %s", current, latest)
    success = _download_and_apply(zip_url, latest, base_dir, logger)

    if success:
        restart = messagebox.askyesno(
            "Update Complete",
            f"Updated to version {latest}.\n\nRestart now?",
            parent=root,
        )
        if restart:
            _restart_app()
    else:
        messagebox.showerror(
            "Update Failed",
            "Something went wrong. The app will continue with the current version.\nCheck the log for details.",
            parent=root,
        )
