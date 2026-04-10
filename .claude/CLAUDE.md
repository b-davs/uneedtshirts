# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

New Order Launcher — a Windows 11 Tkinter desktop app for UNeedTShirts that creates order folders and `.xls` workbooks from templates. Also includes a bizactivity integration that syncs job data from Whole Job Docs Map sheets into a single Business Activity workbook (Job Reports sheet), and a standalone file watcher that keeps bizactivity up to date in real time. Python 3.11+, no web stack.

## Commands

```bash
# Run all tests (logic layer — works on Mac)
python -m pytest

# Run a single test file
python -m pytest tests/test_storage_clients.py

# Run a single test
python -m pytest tests/test_storage_clients.py::test_function_name -v

# Tag and release (builds Windows exe via GitHub Actions)
git tag v1.x.x && git push origin v1.x.x
```

See `.claude/rules/release-workflow.md` for the full release and delivery process.

## Architecture

**Data flow:** `config.json` → `config.py` parses into `AppConfig` → `main.py` bootstraps SQLite schema + CSV seed → `ui_main.py` launches Tkinter UI → `updater.py` checks GitHub for new version in background thread → `bizactivity.sync_all_to_bizactivity` runs background sync → user creates order → `order_service.py` orchestrates folder/workbook creation → `excel_writer.py` writes Map sheet via COM → `bizactivity.write_job_to_bizactivity` writes initial row to bizactivity via COM.

**Bizactivity data flow:** Individual Whole Job Docs each have a Map sheet with one job's data. The Business Activity workbook (`bizactivity.xlsx`) has a Job Reports sheet with 12 monthly sections (70 rows each). Data flows Map → Job Reports via three triggers: (1) initial row on order creation, (2) full sync on launcher startup, (3) real-time sync via `watcher.py` file watcher. Column mapping and month assignment logic live in `bizactivity.py`. See `excel-mapping/docs/map_to_jobreports_integration_spec.md` for the full spec.

**Module responsibilities:**
- `models.py` — All dataclasses (`AppConfig`, `ClientRecord`, `OrderRequest`, `OrderResult`, etc.)
- `config.py` — JSON config loading, legacy client seed parsing. Falls back from `config.json` to `config.example.json`
- `storage.py` — SQLite operations (client CRUD, order ID generation, CSV import, address parsing, abbreviation generation). DB at `%LOCALAPPDATA%/UneedTShirtsNewOrder/state.db`
- `sequence.py` — Pure functions for folder naming, sequence detection (scans filesystem), description sanitization
- `order_service.py` — Orchestrates order creation: resolves client → detects sequence → creates folder → copies template → writes Excel → writes bizactivity initial row → records event
- `excel_writer.py` — Windows Excel COM automation via `win32com.client.DispatchEx`. Best-effort: failures don't block folder creation
- `ui_main.py`, `ui_new_client.py`, `ui_manage_clients.py` — Tkinter UI layer
- `bizactivity.py` — Core logic for syncing job data to the Business Activity workbook. Reads Whole Job Docs Map sheets and writes/updates job rows in the Job Reports sheet via Excel COM. Handles month assignment (job_start_date > create_date), row matching by job number across all 12 sections, insert/update/move operations. Column mapping defined as constants matching the integration spec.
- `watcher.py` — Standalone file watcher (separate exe). Uses `watchdog` to monitor `clients_root` for Whole Job Docs changes in real time. Debounces rapid file events (5s), then reads the changed Map sheet and syncs to bizactivity. Built as `BizactivityWatcher.exe`. Auto-started by the launcher on startup, auto-killed before updates, auto-restarted after. Dan never interacts with it directly.
- `updater.py` — Auto-update check on launch: queries GitHub releases API in background thread, prompts user to download and restart if newer version exists
- `logging_setup.py` — Rotating file logger to `%LOCALAPPDATA%/UneedTShirtsNewOrder/logs/`

**Key patterns:**
- Client registry lives in SQLite (authoritative), not config. Config `clients` array is legacy seed only
- First-run seeds from `clients.csv` (upsert by case-insensitive name), falls back to config legacy clients
- Order sequence is per-client folder: regex scan of `U-{ABBR}-(\d+)` directories, next = max+1
- All storage/order functions accept optional `db_path` parameter for test isolation (temp SQLite files)
- Excel COM write is injectable — `order_service.create_order` accepts `excel_write_func` kwarg for testing

## External References

- **Integration spec:** The full column mapping and Job Reports sheet structure is documented in `/Users/B/Documents/the_matrix/excel-mapping/docs/map_to_jobreports_integration_spec.md` (outside this repo). This is the authoritative reference for Map → Job Reports field mapping, month section layout, merged cells, formula columns to avoid, and paid marker system.

## Governance

- `PLAN.md` is the authoritative spec. Any behavior change must update `PLAN.md`
- Type hints required on all new code
- Tests required for naming, sequence, config, CSV import, client registry, order flow, and bizactivity logic
