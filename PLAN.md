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
- Build one-file exe with `build_exe.ps1`.
- Optional desktop shortcut with `create_shortcut.ps1`.

## Test Plan
- Unit tests for:
  - config loading and legacy-client parsing
  - address parsing and abbreviation generation
  - CSV seed/import merge behavior
  - archive filtering
  - order creation using DB clients
  - excel payload includes new contact/address fields

## Manual Acceptance Checklist
- First launch seeds clients from `clients.csv`.
- Manage Clients supports add/edit/archive/import and persistence.
- Archived clients hidden from main dropdown.
- Created order folder/workbook names follow convention.
- Workbook `Map` writes include `A2/B2/C2/E2/AB2/AC2/AD2/AE2/AF2`.

## Assumptions and Decisions
- SQLite is authoritative for clients.
- CSV import is upsert-by-name.
- All 5 new map fields are active.
- Address parsing favors safe fallback over aggressive inference.
- `PLAN.md` is authoritative.
