# New Order Launcher Plan

## Summary
Build and maintain a Windows 11 Tkinter launcher for creating new orders. Client records are authoritative in SQLite, seeded from `clients.csv` when empty, and managed through in-app UI. The app creates correctly named order folders/workbooks and writes mapped data into Excel via COM.

## Functional Requirements
- Load app settings from `config.json` (fallback `config.example.json`).
- Keep app settings in config, but keep client registry in SQLite.
- Auto-seed clients from `clients.csv` on first run if DB has zero clients.
- Fallback seed from legacy `clients` list in config when CSV is missing.
- Main form fields:
  - `Client` required dropdown
  - `Job Description` optional
  - `Due Date` optional
- Main actions:
  - Create order
  - Open folder/workbook
  - Retry Excel write
  - Manage clients
- Manage Clients screen:
  - Add, Edit, Archive, Import CSV, Show Archived

## Naming and Sequence Rules
- Root clients directory: `D:/A Client Sites & Images`
- Folder naming: `U-{ABBR}-{SEQ}` + optional ` {Job Description}`
- Workbook naming: `U-{ABBR}-{SEQ}.xls`
- Sequence detection: scan client folder names with regex `^U-{ABBR}-(\d+)(?:\s+.*)?$`, next is max+1.

## Config Schema
- Config remains source of:
  - root paths
  - naming rules
  - behavior flags
  - template list/default
  - Excel mapping
- Legacy `clients` in config is optional seed input only.

## SQLite Schema
- `clients`
  - id, name, abbr, folder_path
  - contact_person, phone, email
  - street_address, city_state_zip
  - status (`active`/`archived`), created_at, updated_at
- `internal_order_ids`
  - existing order ID tracking fields
  - `client_id` link for registry integration
- `app_meta`
  - key/value metadata (CSV fingerprint and import timestamp)

## CSV Bootstrap Rules
- Read `clients.csv` from app directory (`utf-8-sig`).
- Expected headers:
  - `Client name`, `Contact person`, `phone number`, `client address`, `email`
- Merge policy for manual/seed import:
  - Upsert by case-insensitive client name
  - Existing abbreviations preserved unless blank
  - Existing folder path preserved unless blank
- New abbreviations are deterministic and unique (`^[A-Z0-9]+$`).

## Address Parsing Rules
- Input: raw `client address` from CSV.
- Output:
  - `street_address`
  - `city_state_zip`
- Parse from right side for city/state/zip.
- Keep suite/unit markers in street (`#106`, `Ste`, `Bay`, etc.).
- If ambiguous, keep full address as street and leave `city_state_zip` blank.

## Excel COM Mapping
- Sheet: `Map`
- Existing writes:
  - `A2`: client_name
  - `B2`: job_number (`ABBR-N`)
  - `C2`: job_description
  - `E2`: due_date
- New writes:
  - `AB2`: contact_person
  - `AC2`: phone
  - `AD2`: email
  - `AE2`: street_address
  - `AF2`: city_state_zip
- If `skip_non_empty_cells=true`, do not overwrite existing values.

## Error Handling and Logging
- Runtime data:
  - `%LOCALAPPDATA%/UneedTShirtsNewOrder/state.db`
  - `%LOCALAPPDATA%/UneedTShirtsNewOrder/logs/`
- On Excel write failure:
  - keep created folder/workbook
  - show actionable UI message
  - keep retry available

## Packaging and Deployment
- Build two one-file exes: `NewOrderLauncher.exe` (main app) and `BizactivityWatcher.exe` (file watcher).
- GitHub Actions builds both via PyInstaller on `windows-latest` and bundles into `NewOrderLauncher.zip`.
- `build_exe.ps1` is the legacy local build script (also builds both exes).
- Optional desktop shortcut with `create_shortcut.ps1`.

## Test Plan
- Unit tests for:
  - config loading and legacy-client parsing
  - address parsing and abbreviation generation
  - CSV seed/import merge behavior
  - archive filtering
  - order creation using DB clients
  - excel payload includes new contact/address fields
  - bizactivity month section row calculations
  - bizactivity month assignment logic (priority: job_start_date > create_date)
  - bizactivity row matching (insert, update, move)
  - bizactivity column mapping consistency (Map fields → Job Reports columns)
  - bizactivity companion state read/write/reset helpers
  - bizactivity empty-row detection tightened for companion columns
  - bizactivity lock detection via Python try-open
  - pending-sync queue enqueue/drain/dedup behavior
  - watcher file pattern matching and ignore prefixes

## Manual Acceptance Checklist
- First launch seeds clients from `clients.csv`.
- Manage Clients supports add/edit/archive/import and persistence.
- Archived clients hidden from main dropdown.
- Created order folder/workbook names follow convention.
- Workbook `Map` writes include `A2/B2/C2/E2/AB2/AC2/AD2/AE2/AF2`.
- On order creation, initial row written to bizactivity (correct month section, correct columns).
- On launcher startup, background sync reads all Whole Job Docs and updates bizactivity.
- Watcher auto-starts on launcher launch, detects file saves, syncs to bizactivity in real time.
- Watcher auto-stops before updates and auto-restarts after.

## Business Activity (Bizactivity) Integration
- Single `bizactivity.xlsx` workbook with a `Job Reports` sheet acts as master ledger.
- 12 monthly sections stacked vertically, 70 data rows each, totals row per section.
- Column mapping from Whole Job Docs Map sheet to Job Reports is defined in `bizactivity.py`.
- **Mode 1 (order creation):** After folder/workbook creation, write initial row (client, job #, description, create date) to the correct month section. Best-effort — failure does not block order creation.
- **Mode 2 (launch sync):** Background thread on app startup scans all Whole Job Docs under `clients_root`, reads each Map sheet, and updates/inserts rows in bizactivity. Catches financial data Dan fills in after order creation. Skipped with a warning if bizactivity is currently open in Excel (file watcher will pick up any later changes).
- Month assignment: `job_start_date` > `create_date` > current month. If month changes, job row is moved between sections.
- Row matching: scan all 12 sections by job number (column D). Update in place, move if month changed, or insert in first empty row.
- Empty-row detection: a row is considered empty only if columns B/C/D **and** all companion columns are empty, preventing reuse of a row with stale state.
- **Companion columns** (H, Q, S, U, W, Y, AA, AD, AF, AH) hold user-owned state that the watcher must not clobber: H is a client-paid marker, Q/S/U/W/Y/AA/AD/AF/AH are vendor-paid "P" dropdowns that drive CF on their left-adjacent vendor-cost cells. On a month-boundary move, companion cell values AND manual fill colors are snapshotted from the old row and written to the new row; the old row is cleared (mapped + companion) so it can be reused safely.
- **Sheet protection:** Dan manually unlocks the companion column ranges once and protects the Job Reports sheet with password `password` (see `docs/bizactivity-protection-setup.md`). The watcher unprotects → writes → reprotects around each sync operation. This is accident-prevention only, not a security boundary.
- **Clickable Job # cells:** The job_number cell (column D) is written as an Excel `HYPERLINK` formula pointing to the originating Whole Job Docs workbook, so clicking it opens the source file directly. The source path is threaded through every write path (order creation, launch sync, watcher live sync) and is persisted on pending-queue entries so queued writes still land with the link when the queue drains.
- **Pending-sync queue:** If `bizactivity.xls` is open in Excel when a sync fires, the watcher detects the file lock (via a Python `open(path, "r+b")` probe — version-agnostic, works with `.xls` and `.xlsx`) and serializes the payload to `%LOCALAPPDATA%/UneedTShirtsNewOrder/pending_syncs.json`. Dedup is last-write-wins keyed by `job_number`. A background drain loop in the watcher checks every 30s and flushes the queue as soon as the workbook is unlocked.
- Config: `bizactivity_path` in `config.json` points to the workbook on Dan's machine.
- Uses Excel COM (`win32com.client`), consistent with the existing `excel_writer.py` approach.
- **File watcher (`watcher.py`):** Standalone background process using `watchdog` to monitor `clients_root` for Whole Job Docs file saves. Debounces events (5s delay), reads changed Map sheet, and syncs to bizactivity. Also runs the pending-sync queue drain loop on a 30s interval. Built as separate `BizactivityWatcher.exe`. Auto-managed by the launcher: auto-started on launch (if not already running), auto-killed before updates, auto-restarted after. Dan never interacts with it directly.

## Assumptions and Decisions
- SQLite is authoritative for clients.
- CSV import is upsert-by-name.
- All 5 new map fields are active.
- Address parsing favors safe fallback over aggressive inference.
- `PLAN.md` is authoritative.
