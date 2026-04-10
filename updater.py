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


def _fetch_latest_release(logger: Optional[logging.Logger] = None) -> Optional[dict]:
    req = Request(RELEASES_URL, headers={"User-Agent": "NewOrderLauncher-Updater"})
    try:
        with urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())
    except Exception:
        if logger:
            logger.exception("Failed to fetch latest release from %s", RELEASES_URL)
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


def _download_and_extract(zip_url: str, logger: logging.Logger) -> Optional[Path]:
    """Download the release zip and extract to a persistent temp directory.
    Returns the extraction directory path, or None on failure."""
    try:
        req = Request(zip_url, headers={"User-Agent": "NewOrderLauncher-Updater"})
        staging = Path(tempfile.mkdtemp(prefix="nol_update_"))

        zip_path = staging / "update.zip"
        with urlopen(req, timeout=120) as resp:
            zip_path.write_bytes(resp.read())

        extract_dir = staging / "contents"
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(extract_dir)

        zip_path.unlink()
        logger.info("Downloaded and extracted update to %s", extract_dir)
        return extract_dir
    except Exception:
        logger.exception("Download/extract failed")
        return None


def _write_apply_script(extract_dir: Path, base_dir: Path, version: str, logger: logging.Logger) -> Optional[Path]:
    """Write a batch script that waits for the app to exit, copies new files,
    writes the version, and relaunches the app."""
    exe_name = Path(sys.executable).name if getattr(sys, "frozen", False) else "python.exe"
    exe_path = base_dir / exe_name

    # Build protected-file checks for the :copy_one subroutine
    skip_checks = "\n".join(
        f'if /I "%~nx1"=="{f}" (echo   Skipped %~nx1 ^(protected^) & exit /b)'
        for f in PROTECTED_FILES
    )

    script = f"""@echo off
echo Waiting for app to close...
:wait
tasklist /FI "IMAGENAME eq {exe_name}" /NH 2>nul | find /I "{exe_name}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)

echo Applying update...
for %%F in ("{extract_dir}\\*") do call :copy_one "%%F"

echo {version}> "{base_dir}\\version.txt"

echo Cleaning up...
rmdir /S /Q "{extract_dir.parent}"

echo Restarting...
start "" "{exe_path}"
exit

:copy_one
{skip_checks}
copy /Y "%~1" "{base_dir}" >nul
echo   Updated %~nx1
exit /b
"""
    try:
        script_path = extract_dir.parent / "apply_update.bat"
        script_path.write_text(script, encoding="utf-8")
        logger.info("Wrote apply script to %s", script_path)
        return script_path
    except Exception:
        logger.exception("Failed to write apply script")
        return None


def check_for_update_async(
    root: tk.Tk,
    base_dir: Path,
    logger: logging.Logger,
) -> None:
    """Checks for updates in a background thread. Shows a dialog on the main
    thread if a new version is available."""

    def _check() -> None:
        try:
            current = _current_version(base_dir)
            logger.info("Update check: current version = %s", current)

            release = _fetch_latest_release(logger)
            if release is None:
                logger.warning("Update check: failed to fetch latest release from GitHub")
                return

            latest = release.get("tag_name", "").lstrip("v")
            logger.info("Update check: latest release = %s", latest)

            if not latest:
                logger.warning("Update check: release has no tag_name")
                return
            if latest == current:
                logger.info("Update check: already up to date")
                return

            zip_url = _find_zip_url(release)
            if not zip_url:
                logger.warning("Update check: no NewOrderLauncher.zip asset found in release")
                return

            logger.info("Update check: update available %s -> %s", current, latest)
            root.after(0, lambda: _prompt_update(root, base_dir, logger, current, latest, zip_url))
        except Exception:
            logger.exception("Update check: unexpected error")

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

    _stop_watcher(logger)

    extract_dir = _download_and_extract(zip_url, logger)
    if not extract_dir:
        messagebox.showerror(
            "Update Failed",
            "Could not download the update.\nCheck the log for details.",
            parent=root,
        )
        return

    script_path = _write_apply_script(extract_dir, base_dir, latest, logger)
    if not script_path:
        messagebox.showerror(
            "Update Failed",
            "Could not prepare the update.\nCheck the log for details.",
            parent=root,
        )
        return

    logger.info("Launching apply script and exiting")
    subprocess.Popen(
        ["cmd", "/c", str(script_path)],
        creationflags=0x08000000,  # CREATE_NO_WINDOW
    )
    root.destroy()
    sys.exit(0)
