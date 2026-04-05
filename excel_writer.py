from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from models import WriteResult


def _parse_mapping(mapping: dict[str, Any]) -> tuple[str, dict[str, str]]:
    if "fields" in mapping:
        sheet_name = str(mapping.get("sheet_name", "Map"))
        fields = dict(mapping["fields"])
        return sheet_name, fields
    return "Map", dict(mapping)


def write_header_block(
    workbook_path: str,
    mapping: dict,
    values: dict,
    skip_non_empty: bool,
    logger: logging.Logger | None = None,
) -> WriteResult:
    sheet_name, fields = _parse_mapping(mapping)
    written_cells: list[str] = []
    skipped_cells: list[str] = []

    excel = None
    workbook = None
    created_excel = False

    try:
        import win32com.client as win32  # type: ignore[import-not-found]

        excel = win32.DispatchEx("Excel.Application")
        created_excel = True
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable

        workbook = excel.Workbooks.Open(
            str(Path(workbook_path).resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            CorruptLoad=1,
        )
        sheet = workbook.Worksheets(sheet_name)

        for field_key, cell_ref in fields.items():
            if field_key not in values:
                continue
            value = values[field_key]
            if value is None:
                continue
            if isinstance(value, str) and value.strip() == "":
                continue

            cell = sheet.Range(cell_ref)
            existing = cell.Value
            if skip_non_empty and existing not in (None, ""):
                skipped_cells.append(cell_ref)
                continue

            cell.Value = value
            written_cells.append(cell_ref)

        workbook.Save()
        return WriteResult(
            success=True,
            written_cells=written_cells,
            skipped_cells=skipped_cells,
            error_message=None,
        )
    except Exception as exc:  # pragma: no cover - exercised via mocking
        if logger:
            logger.exception("Excel COM write failed for workbook %s", workbook_path)
        return WriteResult(
            success=False,
            written_cells=written_cells,
            skipped_cells=skipped_cells,
            error_message=str(exc),
        )
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                if logger:
                    logger.exception("Failed to close workbook %s", workbook_path)
        if excel is not None and created_excel:
            try:
                excel.Quit()
            except Exception:
                if logger:
                    logger.exception("Failed to quit Excel instance")
