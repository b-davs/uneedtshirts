# New Order Launcher

Windows-focused Tkinter app for creating order folders and `.xls` workbooks from a template.

## Core Behavior
- Client registry is stored in local SQLite (`%LOCALAPPDATA%/UneedTShirtsNewOrder/state.db`).
- Clients are seeded from `clients.csv` on first run when the registry is empty.
- Order sequence is per client folder (`U-ABBR-N ...`) and never resets.
- Workbook naming is `U-ABBR-N.xls`.
- Excel COM writes mapped fields into `Map` sheet including contact/address fields (`A2/B2/C2/E2/AB2:AF2`).

## Client Management
- `+ New Client...` quick add flow from main screen.
- `Manage Clients...` dialog with:
  - Add
  - Edit
  - Archive
  - Show archived filter
  - Import CSV
- Editing name/abbreviation is allowed with warning if the client already has orders.

## CSV Bootstrap
Expected headers in `clients.csv` (first row):
- `Client name`
- `Contact person`
- `phone number`
- `client address`
- `email`

Address parsing behavior:
- Attempts to split address into `street_address` and `city_state_zip`.
- If ambiguous, keeps full address in street and leaves `city_state_zip` blank.

## Setup
1. Install Python 3.11+ on Windows.
2. Install dependencies:
- `python -m pip install -r requirements.txt`
3. Copy `config.example.json` to `config.json` and verify paths.
4. Run:
- `python main.py`

## Build One-File EXE
Use PowerShell:
- `./build_exe.ps1`

Output:
- `dist/NewOrderLauncher.exe`

## Create Desktop Shortcut
After building:
- `./create_shortcut.ps1`

## Tests
Run:
- `python -m pytest`

## Troubleshooting
- If Excel write fails, folder/workbook remain and `Retry Excel Write` stays available.
- Ensure Excel is installed on the target Windows machine for COM automation.
- Verify template path exists and workbook is not locked by another process.
