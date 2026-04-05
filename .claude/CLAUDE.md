# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

New Order Launcher — a Windows 11 Tkinter desktop app for UNeedTShirts that creates order folders and `.xls` workbooks from templates. Python 3.11+, no web stack.

## Commands

```bash
# Run the app
python main.py

# Run all tests
python -m pytest

# Run a single test file
python -m pytest tests/test_storage_clients.py

# Run a single test
python -m pytest tests/test_storage_clients.py::test_function_name -v

# Build one-file Windows exe (PowerShell)
./build_exe.ps1
```

## Architecture

**Data flow:** `config.json` → `config.py` parses into `AppConfig` → `main.py` bootstraps SQLite schema + CSV seed → `ui_main.py` launches Tkinter UI → user creates order → `order_service.py` orchestrates folder/workbook creation → `excel_writer.py` writes via COM.

**Module responsibilities (fixed layout — do not add/rename modules):**
- `models.py` — All dataclasses (`AppConfig`, `ClientRecord`, `OrderRequest`, `OrderResult`, etc.)
- `config.py` — JSON config loading, legacy client seed parsing. Falls back from `config.json` to `config.example.json`
- `storage.py` — SQLite operations (client CRUD, order ID generation, CSV import, address parsing, abbreviation generation). DB at `%LOCALAPPDATA%/UneedTShirtsNewOrder/state.db`
- `sequence.py` — Pure functions for folder naming, sequence detection (scans filesystem), description sanitization
- `order_service.py` — Orchestrates order creation: resolves client → detects sequence → creates folder → copies template → writes Excel → records event
- `excel_writer.py` — Windows Excel COM automation via `win32com.client.DispatchEx`. Best-effort: failures don't block folder creation
- `ui_main.py`, `ui_new_client.py`, `ui_manage_clients.py` — Tkinter UI layer
- `logging_setup.py` — Rotating file logger to `%LOCALAPPDATA%/UneedTShirtsNewOrder/logs/`

**Key patterns:**
- Client registry lives in SQLite (authoritative), not config. Config `clients` array is legacy seed only
- First-run seeds from `clients.csv` (upsert by case-insensitive name), falls back to config legacy clients
- Order sequence is per-client folder: regex scan of `U-{ABBR}-(\d+)` directories, next = max+1
- All storage/order functions accept optional `db_path` parameter for test isolation (temp SQLite files)
- Excel COM write is injectable — `order_service.create_order` accepts `excel_write_func` kwarg for testing

## Governance

- `PLAN.md` is the authoritative spec. Any behavior change must update `PLAN.md` and `DECISIONS.md`
- `AGENTS.md` has the full rules for this project (non-goals, quality gates, definition of done)
- Type hints required on all new code
- Tests required for naming, sequence, config, CSV import, client registry, and order flow
