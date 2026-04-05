from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from config import ConfigError, load_runtime_config, runtime_base_dir
from logging_setup import setup_logging
from storage import ensure_schema, seed_clients_from_csv_if_needed
from ui_main import MainWindow
from updater import check_for_update_async


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
    root.mainloop()


if __name__ == "__main__":
    main()
