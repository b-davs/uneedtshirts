from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from models import BizactivityResult

# ---------------------------------------------------------------------------
# Sheet protection
# ---------------------------------------------------------------------------
# Accident-prevention only — not a security boundary. Dan manually unlocks
# companion column ranges once, then protects the sheet with this password.
# The watcher unprotects/reprotects around each write.
_SHEET_PASSWORD = "password"

# Excel Interior.Pattern constants
_XL_PATTERN_NONE = -4142  # xlPatternNone (no fill)

# ---------------------------------------------------------------------------
# Month section layout constants (from integration spec)
# ---------------------------------------------------------------------------
_SECTION_STRIDE = 76
_FIRST_DATA_ROW_BASE = 13  # January first data row
_DATA_ROWS_PER_SECTION = 70


def _first_data_row(month: int) -> int:
    """Return the first data row for a 1-based month number."""
    return _FIRST_DATA_ROW_BASE + (month - 1) * _SECTION_STRIDE


def _last_data_row(month: int) -> int:
    """Return the last data row for a 1-based month number."""
    return _first_data_row(month) + _DATA_ROWS_PER_SECTION - 1


# ---------------------------------------------------------------------------
# Column mapping: internal field key -> Job Reports column letter
# ---------------------------------------------------------------------------
FIELD_TO_JR_COL: dict[str, str] = {
    "client": "B",
    "job_description": "C",
    "job_number": "D",
    "job_start_date": "E",
    "tax_or_exempt": "F",
    "gross_sales_before_tax": "G",
    "total_sale_incl_tax": "I",
    "deposit": "J",
    "dep_how_paid": "K",
    "balance_due": "L",
    "job_deliver_date": "M",
    "bal_how_paid": "N",
    "profit": "O",
    "all_day_shirts": "P",
    "sanmar": "R",
    "ss": "T",
    "tko": "V",
    "florida_dtf": "X",
    "embroid": "Z",
    "square_fee": "AB",
    "screen": "AC",
    "shipping": "AE",
    "other": "AG",
    "description_of_other": "AJ",
    "exempt_part_of_sale": "BT",
    "taxable_amount_of_sale": "BU",
    "sales_tax": "BV",
    "create_date": "BW",
}

# Map sheet column letter -> internal field key (for reading Whole Job Docs)
MAP_COL_TO_FIELD: dict[str, str] = {
    "A": "client",
    "B": "job_number",
    "C": "job_description",
    "D": "create_date",
    "E": "job_start_date",
    "G": "tax_or_exempt",
    "H": "gross_sales_before_tax",
    "I": "total_sale_incl_tax",
    "J": "deposit",
    "K": "dep_how_paid",
    "L": "balance_due",
    "M": "job_deliver_date",
    "N": "bal_how_paid",
    "R": "profit",
    "S": "all_day_shirts",
    "T": "sanmar",
    "U": "ss",
    "V": "tko",
    "W": "florida_dtf",
    "X": "embroid",
    "Y": "square_fee",
    "Z": "screen",
    "AA": "shipping",
    "AB": "other",
    "AC": "description_of_other",
    "O": "exempt_part_of_sale",
    "P": "taxable_amount_of_sale",
    "Q": "sales_tax",
}

# Job Reports columns with formulas — never overwrite
_JR_FORMULA_COLS = {"AY", "AZ", "BR"}

# Field key whose value should be rendered as a HYPERLINK formula pointing
# back to the originating Whole Job Docs workbook when a source path is known.
_HYPERLINK_FIELD = "job_number"


def _build_hyperlink_formula(target_path: str, display: str) -> str:
    """Build an Excel HYPERLINK formula string.

    Escapes embedded double quotes per Excel string-literal convention
    (doubled quotes). Works for .xls, .xlsm, and .xlsx — HYPERLINK has
    been a built-in since Excel 97.
    """
    safe_path = str(target_path).replace('"', '""')
    safe_display = str(display).replace('"', '""')
    return f'=HYPERLINK("{safe_path}","{safe_display}")'

# Companion columns: user-owned cells that carry state the watcher must
# not clobber. H is Dan's client-paid marker (adjacent to G gross_sales).
# Q/S/U/W/Y/AA/AD/AF/AH are vendor-paid markers ("P" dropdown) that drive
# CF on their left-adjacent vendor cost cells. Cell values AND manual fill
# colors on these columns must travel with the job on month moves.
COMPANION_COLS: list[str] = [
    "H", "Q", "S", "U", "W", "Y", "AA", "AD", "AF", "AH",
]


# ---------------------------------------------------------------------------
# Month assignment
# ---------------------------------------------------------------------------
def determine_month(values: dict[str, Any]) -> int:
    """Determine target month from job dates.

    Priority: job_start_date > create_date > current month.
    """
    for key in ("job_start_date", "create_date"):
        raw = values.get(key)
        if raw is None:
            continue
        if isinstance(raw, datetime):
            return raw.month
        if isinstance(raw, str) and raw.strip():
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
                try:
                    return datetime.strptime(raw.strip(), fmt).month
                except ValueError:
                    continue
    return datetime.now().month


# ---------------------------------------------------------------------------
# COM helper: column letter -> Range reference for a row
# ---------------------------------------------------------------------------
def _cell_ref(col_letter: str, row: int) -> str:
    """Build a cell reference like 'B13' from column letter and row number."""
    return f"{col_letter}{row}"


# ---------------------------------------------------------------------------
# Row operations (COM sheet object)
# ---------------------------------------------------------------------------
def _find_job_row(sheet: Any, job_number: str) -> Optional[tuple[int, int]]:
    """Scan all 12 sections for a row where column D matches job_number.

    Returns (row_number, month) or None.
    """
    for month in range(1, 13):
        first = _first_data_row(month)
        last = _last_data_row(month)
        for row in range(first, last + 1):
            cell_val = sheet.Range(f"D{row}").Value
            if cell_val is not None and str(cell_val).strip() == job_number.strip():
                return row, month
    return None


def _find_first_empty_row(sheet: Any, month: int) -> Optional[int]:
    """Find the first empty data row in a month section.

    A row is empty if B/C/D are all empty AND all companion columns are
    also empty. The companion check prevents a freshly-cleared row from
    being reused while a stale "P" selection still sits in one of the
    vendor-paid markers.
    """
    first = _first_data_row(month)
    last = _last_data_row(month)
    for row in range(first, last + 1):
        b = sheet.Range(f"B{row}").Value
        c = sheet.Range(f"C{row}").Value
        d = sheet.Range(f"D{row}").Value
        if b not in (None, "") or c not in (None, "") or d not in (None, ""):
            continue
        companion_dirty = False
        for col in COMPANION_COLS:
            val = sheet.Range(f"{col}{row}").Value
            if val not in (None, ""):
                companion_dirty = True
                break
        if not companion_dirty:
            return row
    return None


def _clear_row(sheet: Any, row: int) -> None:
    """Clear all mapped columns in a data row."""
    for col_letter in FIELD_TO_JR_COL.values():
        sheet.Range(_cell_ref(col_letter, row)).Value = None


def _read_companion_state(sheet: Any, row: int) -> dict[str, dict[str, Any]]:
    """Snapshot companion cell values + fill state for a row.

    Reads `Interior.Pattern` and `Interior.Color` so we can reproduce
    Dan's manual fill exactly on the destination row after a move.
    `Interior.Color` on companion cells reflects only direct manual fill
    (no CF on the companion cell itself), so this read is unambiguous.
    """
    state: dict[str, dict[str, Any]] = {}
    for col in COMPANION_COLS:
        cell = sheet.Range(f"{col}{row}")
        interior = cell.Interior
        state[col] = {
            "value": cell.Value,
            "pattern": interior.Pattern,
            "color": interior.Color,
        }
    return state


def _write_companion_state(
    sheet: Any, row: int, state: dict[str, dict[str, Any]]
) -> None:
    """Apply a companion state snapshot to `row`."""
    for col, s in state.items():
        cell = sheet.Range(f"{col}{row}")
        value = s.get("value")
        if value is not None:
            cell.Value = value
        pattern = s.get("pattern", _XL_PATTERN_NONE)
        cell.Interior.Pattern = pattern
        if pattern != _XL_PATTERN_NONE:
            color = s.get("color")
            if color is not None:
                cell.Interior.Color = color


def _reset_companion_state(sheet: Any, row: int) -> None:
    """Clear companion cell values and fill on a row so it can be
    reused as an empty slot for a future job."""
    for col in COMPANION_COLS:
        cell = sheet.Range(f"{col}{row}")
        cell.Value = None
        cell.Interior.Pattern = _XL_PATTERN_NONE


def _write_row(
    sheet: Any,
    row: int,
    values: dict[str, Any],
    source_path: Optional[str] = None,
) -> list[str]:
    """Write field values into a Job Reports data row. Returns list of cells written.

    When `source_path` is provided, the job_number cell is written as a
    HYPERLINK formula pointing to the originating Whole Job Docs workbook,
    so Dan can click it to open the file directly from Job Reports.
    """
    written: list[str] = []
    for field_key, value in values.items():
        col_letter = FIELD_TO_JR_COL.get(field_key)
        if col_letter is None:
            continue
        if col_letter in _JR_FORMULA_COLS:
            continue
        if value is None:
            continue
        ref = _cell_ref(col_letter, row)
        if field_key == _HYPERLINK_FIELD and source_path:
            formula = _build_hyperlink_formula(source_path, str(value))
            sheet.Range(ref).Formula = formula
        else:
            sheet.Range(ref).Value = value
        written.append(ref)
    return written


# ---------------------------------------------------------------------------
# Lock detection — version-agnostic, avoids depending on ~$ lock files
# ---------------------------------------------------------------------------
def is_bizactivity_locked(path: str | Path) -> bool:
    """Return True if the bizactivity workbook is locked for exclusive
    write by another process (typically Dan's Excel having it open).

    Uses a read/write file open as the probe — if Excel holds an
    exclusive lock, Windows raises PermissionError. No bytes are written.
    Works on both `.xls` and `.xlsx` regardless of Excel version.
    """
    p = Path(path)
    if not p.exists():
        return False
    try:
        with open(p, "r+b"):
            return False
    except (PermissionError, OSError):
        return True


# ---------------------------------------------------------------------------
# Sheet protection
# ---------------------------------------------------------------------------
def _unprotect_sheet(sheet: Any, logger: logging.Logger | None = None) -> None:
    """Best-effort unprotect. Silent no-op if the sheet isn't protected."""
    try:
        sheet.Unprotect(Password=_SHEET_PASSWORD)
    except Exception:
        if logger:
            logger.debug("Unprotect no-op (sheet may not be protected)")


def _protect_sheet(sheet: Any, logger: logging.Logger | None = None) -> None:
    """Best-effort reprotect. Logs but does not raise on failure."""
    try:
        sheet.Protect(Password=_SHEET_PASSWORD)
    except Exception:
        if logger:
            logger.warning("Failed to reprotect Job Reports sheet")


# ---------------------------------------------------------------------------
# COM lifecycle helpers
# ---------------------------------------------------------------------------
def _open_excel() -> Any:
    """Create a new hidden Excel COM instance."""
    import win32com.client as win32  # type: ignore[import-not-found]

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    return excel


def _open_workbook(excel: Any, path: str, read_only: bool = False) -> Any:
    """Open a workbook via COM."""
    return excel.Workbooks.Open(
        str(Path(path).resolve()),
        UpdateLinks=0,
        ReadOnly=read_only,
        CorruptLoad=1,
    )


def _close_workbook(workbook: Any, save: bool, logger: logging.Logger | None = None) -> None:
    """Close a workbook, optionally saving."""
    try:
        workbook.Close(SaveChanges=save)
    except Exception:
        if logger:
            logger.exception("Failed to close workbook")


def _quit_excel(excel: Any, logger: logging.Logger | None = None) -> None:
    """Quit the Excel COM instance."""
    try:
        excel.Quit()
    except Exception:
        if logger:
            logger.exception("Failed to quit Excel instance")


# ---------------------------------------------------------------------------
# Public API: write a single job to bizactivity
# ---------------------------------------------------------------------------
def write_job_to_bizactivity(
    bizactivity_path: str,
    values: dict[str, Any],
    *,
    source_path: Optional[str] = None,
    logger: logging.Logger | None = None,
    allow_queue: bool = True,
) -> BizactivityResult:
    """Write or update a single job row in the bizactivity workbook.

    Args:
        bizactivity_path: Path to bizactivity.xlsx (or .xls).
        values: Dict of field keys (matching FIELD_TO_JR_COL keys) to values.
            Must include 'job_number'. Should include date fields for month assignment.
        source_path: Absolute path to the originating Whole Job Docs workbook.
            When provided, the job_number cell is rendered as a HYPERLINK
            formula so clicking it opens the source file. When omitted, the
            job_number is written as a plain string.
        logger: Optional logger.
        allow_queue: When True (default), if the workbook is currently locked
            (Dan has it open), serialize the payload to the pending-sync queue
            and return a "queued" result. When False (used by the drain loop
            itself), lock collisions bubble up as skipped results to avoid
            re-queueing the same payload in a loop.

    Returns:
        BizactivityResult with success status and action taken.
    """
    job_number = values.get("job_number")
    if not job_number or not str(job_number).strip():
        return BizactivityResult(
            success=False,
            action="skipped",
            error_message="No job_number provided",
        )

    job_number = str(job_number).strip()
    target_month = determine_month(values)

    if not bizactivity_path or not str(bizactivity_path).strip():
        return BizactivityResult(
            success=False,
            action="skipped",
            error_message="bizactivity_path is empty",
        )

    biz_path = Path(bizactivity_path)
    if not biz_path.exists() or not biz_path.is_file():
        return BizactivityResult(
            success=False,
            action="skipped",
            error_message=f"Bizactivity file not found: {bizactivity_path}",
        )

    if is_bizactivity_locked(biz_path):
        if allow_queue:
            from pending_queue import enqueue
            enqueue(values, source_path=source_path, logger=logger)
            return BizactivityResult(
                success=True, action="queued", month=target_month,
            )
        return BizactivityResult(
            success=False, action="skipped", month=target_month,
            error_message="Bizactivity workbook is locked (open in Excel)",
        )

    excel = None
    workbook = None
    sheet = None
    try:
        excel = _open_excel()
        workbook = _open_workbook(excel, str(biz_path))
        sheet = workbook.Worksheets("Job Reports")

        _unprotect_sheet(sheet, logger=logger)

        existing = _find_job_row(sheet, job_number)

        if existing is not None:
            existing_row, existing_month = existing
            if existing_month == target_month:
                _write_row(sheet, existing_row, values, source_path=source_path)
                _protect_sheet(sheet, logger=logger)
                workbook.Save()
                if logger:
                    logger.info(
                        "Bizactivity: updated job %s in row %d (month %d)",
                        job_number, existing_row, target_month,
                    )
                return BizactivityResult(
                    success=True, action="updated",
                    target_row=existing_row, month=target_month,
                )
            else:
                # Month changed — carry companion state across the move
                companion_state = _read_companion_state(sheet, existing_row)
                _clear_row(sheet, existing_row)
                _reset_companion_state(sheet, existing_row)
                new_row = _find_first_empty_row(sheet, target_month)
                if new_row is None:
                    _protect_sheet(sheet, logger=logger)
                    workbook.Save()
                    return BizactivityResult(
                        success=False, action="skipped", month=target_month,
                        error_message=f"Month {target_month} section is full (70 rows)",
                    )
                _write_row(sheet, new_row, values, source_path=source_path)
                _write_companion_state(sheet, new_row, companion_state)
                _protect_sheet(sheet, logger=logger)
                workbook.Save()
                if logger:
                    logger.info(
                        "Bizactivity: moved job %s from row %d (month %d) to row %d (month %d)",
                        job_number, existing_row, existing_month, new_row, target_month,
                    )
                return BizactivityResult(
                    success=True, action="moved",
                    target_row=new_row, month=target_month,
                )
        else:
            new_row = _find_first_empty_row(sheet, target_month)
            if new_row is None:
                _protect_sheet(sheet, logger=logger)
                return BizactivityResult(
                    success=False, action="skipped", month=target_month,
                    error_message=f"Month {target_month} section is full (70 rows)",
                )
            _write_row(sheet, new_row, values, source_path=source_path)
            _protect_sheet(sheet, logger=logger)
            workbook.Save()
            if logger:
                logger.info(
                    "Bizactivity: inserted job %s in row %d (month %d)",
                    job_number, new_row, target_month,
                )
            return BizactivityResult(
                success=True, action="inserted",
                target_row=new_row, month=target_month,
            )
    except Exception as exc:
        if sheet is not None:
            _protect_sheet(sheet, logger=logger)
        if logger:
            logger.exception("Bizactivity write failed for job %s", job_number)
        return BizactivityResult(
            success=False, action="skipped", error_message=str(exc),
        )
    finally:
        if workbook is not None:
            _close_workbook(workbook, save=True, logger=logger)
        if excel is not None:
            _quit_excel(excel, logger=logger)


# ---------------------------------------------------------------------------
# Map sheet reader (for Mode 2 sync)
# ---------------------------------------------------------------------------
def read_map_sheet(
    workbook_path: str,
    *,
    excel: Any = None,
    logger: logging.Logger | None = None,
) -> Optional[dict[str, Any]]:
    """Read the Map sheet (row 2) from a Whole Job Docs workbook.

    If an existing Excel COM instance is passed via `excel`, it will be reused
    (no new instance created or quit). Otherwise a fresh instance is used.

    Returns a dict of field keys -> values, or None if the file can't be read.
    """
    created_excel = False
    workbook = None
    try:
        if excel is None:
            excel = _open_excel()
            created_excel = True

        workbook = _open_workbook(excel, workbook_path, read_only=True)

        # Check if Map sheet exists
        sheet_names = [workbook.Worksheets(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
        if "Map" not in sheet_names:
            return None

        sheet = workbook.Worksheets("Map")
        values: dict[str, Any] = {}
        for col_letter, field_key in MAP_COL_TO_FIELD.items():
            cell_val = sheet.Range(f"{col_letter}2").Value
            if cell_val is not None and cell_val != "":
                values[field_key] = cell_val

        return values if values.get("job_number") else None
    except Exception:
        if logger:
            logger.exception("Failed to read Map sheet from %s", workbook_path)
        return None
    finally:
        if workbook is not None:
            _close_workbook(workbook, save=False, logger=logger)
        if created_excel and excel is not None:
            _quit_excel(excel, logger=logger)


# ---------------------------------------------------------------------------
# Batch sync (Mode 2): scan all workbooks and update bizactivity
# ---------------------------------------------------------------------------
def _find_workbooks(clients_root: str) -> list[Path]:
    """Find all Whole Job Docs workbooks under clients_root.

    Looks for .xls/.xlsm/.xlsx files in order folders matching U-*-* pattern.
    """
    root = Path(clients_root)
    if not root.exists():
        return []

    workbooks: list[Path] = []
    for client_dir in root.iterdir():
        if not client_dir.is_dir():
            continue
        for order_dir in client_dir.iterdir():
            if not order_dir.is_dir() or not order_dir.name.startswith("U-"):
                continue
            for f in order_dir.iterdir():
                if (
                    f.is_file()
                    and f.name.startswith("U-")
                    and f.suffix.lower() in (".xls", ".xlsm", ".xlsx")
                ):
                    workbooks.append(f)
    return workbooks


def sync_all_to_bizactivity(
    clients_root: str,
    bizactivity_path: str,
    *,
    logger: logging.Logger | None = None,
) -> dict[str, int]:
    """Scan all Whole Job Docs under clients_root and sync to bizactivity.

    Uses a single Excel COM instance for reading all Map sheets, then a
    separate instance for writing to bizactivity.

    Returns counts: {"synced": N, "skipped": N, "errors": N}.
    """
    from models import SyncReport

    report = SyncReport()

    if not bizactivity_path or not Path(bizactivity_path).exists():
        if logger:
            logger.warning("Bizactivity sync skipped: file not found at %s", bizactivity_path)
        return {"synced": 0, "skipped": 0, "errors": 0}

    if is_bizactivity_locked(bizactivity_path):
        if logger:
            logger.warning(
                "Bizactivity sync skipped: workbook is locked (open in Excel). "
                "Watcher will catch any file changes once it closes."
            )
        return {"synced": 0, "skipped": 0, "errors": 0}

    workbooks = _find_workbooks(clients_root)
    if logger:
        logger.info("Bizactivity sync: found %d workbooks to scan", len(workbooks))

    if not workbooks:
        return {"synced": 0, "skipped": 0, "errors": 0}

    # Phase 1: Read all Map sheets using one shared Excel instance.
    # Each entry carries (source_path, values) so Phase 2 can render the
    # job_number cell as a HYPERLINK back to the originating workbook.
    all_values: list[tuple[str, dict[str, Any]]] = []
    read_excel = None
    try:
        read_excel = _open_excel()
        for wb_path in workbooks:
            try:
                values = read_map_sheet(str(wb_path), excel=read_excel, logger=logger)
                if values is not None:
                    all_values.append((str(wb_path), values))
                else:
                    report.skipped += 1
            except Exception as exc:
                report.errors += 1
                if logger:
                    logger.exception("Bizactivity sync read error for %s", wb_path)
                report.details.append(f"{wb_path.name}: {exc}")
    finally:
        if read_excel is not None:
            _quit_excel(read_excel, logger=logger)

    if not all_values:
        if logger:
            logger.info("Bizactivity sync: no valid Map sheets found")
        return {"synced": report.synced, "skipped": report.skipped, "errors": report.errors}

    # Phase 2: Write all jobs to bizactivity
    for src_path, values in all_values:
        try:
            result = write_job_to_bizactivity(
                bizactivity_path, values, source_path=src_path, logger=logger
            )
            if result.success:
                report.synced += 1
            else:
                report.skipped += 1
                if result.error_message:
                    report.details.append(
                        f"{values.get('job_number', '?')}: {result.error_message}"
                    )
        except Exception as exc:
            report.errors += 1
            if logger:
                logger.exception("Bizactivity sync write error")
            report.details.append(f"{values.get('job_number', '?')}: {exc}")

    if logger:
        logger.info("Bizactivity sync complete: %s", report.summary())

    return {"synced": report.synced, "skipped": report.skipped, "errors": report.errors}
