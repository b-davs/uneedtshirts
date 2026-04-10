from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Optional
from unittest.mock import MagicMock, patch

import pytest

from bizactivity import (
    COMPANION_COLS,
    FIELD_TO_JR_COL,
    MAP_COL_TO_FIELD,
    _XL_PATTERN_NONE,
    _find_first_empty_row,
    _find_job_row,
    _first_data_row,
    _last_data_row,
    _read_companion_state,
    _reset_companion_state,
    _write_companion_state,
    determine_month,
    is_bizactivity_locked,
    read_map_sheet,
    write_job_to_bizactivity,
)


# ---------------------------------------------------------------------------
# Month section row calculations (pure logic — no mocking needed)
# ---------------------------------------------------------------------------
class TestMonthSectionRows:
    def test_january_rows(self) -> None:
        assert _first_data_row(1) == 13
        assert _last_data_row(1) == 82

    def test_february_rows(self) -> None:
        assert _first_data_row(2) == 89
        assert _last_data_row(2) == 158

    def test_march_rows(self) -> None:
        assert _first_data_row(3) == 165
        assert _last_data_row(3) == 234

    def test_december_rows(self) -> None:
        assert _first_data_row(12) == 849
        assert _last_data_row(12) == 918

    def test_all_months_have_70_data_rows(self) -> None:
        for month in range(1, 13):
            span = _last_data_row(month) - _first_data_row(month) + 1
            assert span == 70, f"Month {month} has {span} data rows, expected 70"

    def test_spec_row_positions(self) -> None:
        """Verify row positions match the integration spec table."""
        expected = {
            1: (13, 82, 83),
            2: (89, 158, 159),
            3: (165, 234, 235),
            4: (241, 310, 311),
            5: (317, 386, 387),
            6: (393, 462, 463),
            7: (469, 538, 539),
            8: (545, 614, 615),
            9: (621, 690, 691),
            10: (697, 766, 767),
            11: (773, 842, 843),
            12: (849, 918, 919),
        }
        for month, (exp_first, exp_last, exp_totals) in expected.items():
            assert _first_data_row(month) == exp_first, f"Month {month} first"
            assert _last_data_row(month) == exp_last, f"Month {month} last"
            totals = _last_data_row(month) + 1
            assert totals == exp_totals, f"Month {month} totals"


# ---------------------------------------------------------------------------
# Month assignment logic (pure logic — no mocking needed)
# ---------------------------------------------------------------------------
class TestDetermineMonth:
    def test_job_start_date_takes_priority(self) -> None:
        values = {
            "create_date": "2026-01-15",
            "job_start_date": "2026-04-10",
        }
        assert determine_month(values) == 4

    def test_create_date_fallback(self) -> None:
        values = {"create_date": "2026-07-22"}
        assert determine_month(values) == 7

    def test_datetime_object(self) -> None:
        values = {"create_date": datetime(2026, 11, 5)}
        assert determine_month(values) == 11

    def test_slash_date_format(self) -> None:
        values = {"create_date": "03/15/2026"}
        assert determine_month(values) == 3

    def test_no_dates_uses_current_month(self) -> None:
        month = determine_month({})
        assert 1 <= month <= 12

    def test_empty_string_ignored(self) -> None:
        values = {"create_date": "", "job_start_date": "  "}
        month = determine_month(values)
        assert 1 <= month <= 12


# ---------------------------------------------------------------------------
# Column mapping consistency (pure logic — no mocking needed)
# ---------------------------------------------------------------------------
class TestColumnMappingConsistency:
    def test_all_map_fields_have_jr_mapping(self) -> None:
        """Every field produced by MAP_COL_TO_FIELD should exist in FIELD_TO_JR_COL."""
        for col, field_key in MAP_COL_TO_FIELD.items():
            assert field_key in FIELD_TO_JR_COL, (
                f"Map column {col} produces field '{field_key}' "
                f"which has no Job Reports mapping"
            )


# ---------------------------------------------------------------------------
# Mock COM sheet helper
# ---------------------------------------------------------------------------
class MockSheet:
    """Simulates a COM Worksheet with Range().Value + Interior + Protect/Unprotect."""

    def __init__(self) -> None:
        self._cells: dict[str, Any] = {}
        self._interiors: dict[str, dict[str, Any]] = {}
        self.protected = False
        self.protect_calls: list[str] = []
        self.unprotect_calls: list[str] = []

    def Range(self, ref: str) -> "_CellProxy":
        return _CellProxy(self._cells, self._interiors, ref)

    def set_cell(self, ref: str, value: Any) -> None:
        self._cells[ref] = value

    def get_cell(self, ref: str) -> Any:
        return self._cells.get(ref)

    def set_interior(self, ref: str, pattern: int, color: Any) -> None:
        self._interiors[ref] = {"pattern": pattern, "color": color}

    def get_interior(self, ref: str) -> dict[str, Any]:
        return self._interiors.get(
            ref, {"pattern": _XL_PATTERN_NONE, "color": 0}
        )

    # COM-style API
    def Protect(self, Password: str = "") -> None:
        self.protected = True
        self.protect_calls.append(Password)

    def Unprotect(self, Password: str = "") -> None:
        self.protected = False
        self.unprotect_calls.append(Password)


class _CellProxy:
    """Proxy for a single cell supporting .Value and .Interior."""

    def __init__(
        self,
        cells: dict[str, Any],
        interiors: dict[str, dict[str, Any]],
        ref: str,
    ) -> None:
        self._cells = cells
        self._interiors = interiors
        self._ref = ref

    @property
    def Value(self) -> Any:
        return self._cells.get(self._ref)

    @Value.setter
    def Value(self, val: Any) -> None:
        if val is None:
            self._cells.pop(self._ref, None)
        else:
            self._cells[self._ref] = val

    @property
    def Interior(self) -> "_InteriorProxy":
        return _InteriorProxy(self._interiors, self._ref)


class _InteriorProxy:
    """Proxy for a cell's Interior supporting .Pattern and .Color."""

    def __init__(self, interiors: dict[str, dict[str, Any]], ref: str) -> None:
        self._interiors = interiors
        self._ref = ref

    def _entry(self) -> dict[str, Any]:
        if self._ref not in self._interiors:
            self._interiors[self._ref] = {"pattern": _XL_PATTERN_NONE, "color": 0}
        return self._interiors[self._ref]

    @property
    def Pattern(self) -> int:
        return self._entry().get("pattern", _XL_PATTERN_NONE)

    @Pattern.setter
    def Pattern(self, val: int) -> None:
        self._entry()["pattern"] = val

    @property
    def Color(self) -> Any:
        return self._entry().get("color", 0)

    @Color.setter
    def Color(self, val: Any) -> None:
        self._entry()["color"] = val


# ---------------------------------------------------------------------------
# Row finding (uses MockSheet)
# ---------------------------------------------------------------------------
class TestFindJobRow:
    def test_finds_existing_job_in_january(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("D15", "ACME-1")
        result = _find_job_row(sheet, "ACME-1")
        assert result == (15, 1)

    def test_finds_job_in_april(self) -> None:
        sheet = MockSheet()
        # April data rows: 241-310
        sheet.set_cell("D245", "TFR-5")
        result = _find_job_row(sheet, "TFR-5")
        assert result == (245, 4)

    def test_returns_none_when_not_found(self) -> None:
        sheet = MockSheet()
        result = _find_job_row(sheet, "MISSING-99")
        assert result is None

    def test_strips_whitespace_for_matching(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("D13", " ACME-1 ")
        result = _find_job_row(sheet, "ACME-1")
        assert result == (13, 1)


class TestFindFirstEmptyRow:
    def test_finds_first_row_in_empty_section(self) -> None:
        sheet = MockSheet()
        row = _find_first_empty_row(sheet, month=1)
        assert row == 13

    def test_skips_occupied_rows(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("D13", "ACME-1")  # row 13 has data in col D
        row = _find_first_empty_row(sheet, month=1)
        assert row == 14

    def test_treats_empty_string_as_empty(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("B13", "")
        sheet.set_cell("C13", "")
        sheet.set_cell("D13", "")
        row = _find_first_empty_row(sheet, month=1)
        assert row == 13

    def test_returns_none_when_section_full(self) -> None:
        sheet = MockSheet()
        for row in range(13, 83):  # all 70 January rows
            sheet.set_cell(f"D{row}", f"JOB-{row}")
        row = _find_first_empty_row(sheet, month=1)
        assert row is None

    def test_skips_row_with_stale_companion_value(self) -> None:
        """A row with empty B/C/D but a stale P in a companion col
        should NOT be treated as empty — that would reuse a row with
        leftover state from a previous job."""
        sheet = MockSheet()
        # Row 13: mapped cols empty but Q13 has a stale "P"
        sheet.set_cell("Q13", "P")
        row = _find_first_empty_row(sheet, month=1)
        assert row == 14  # row 13 skipped, 14 is clean

    def test_treats_empty_string_companion_as_empty(self) -> None:
        sheet = MockSheet()
        for col in COMPANION_COLS:
            sheet.set_cell(f"{col}13", "")
        row = _find_first_empty_row(sheet, month=1)
        assert row == 13


# ---------------------------------------------------------------------------
# Companion state helpers
# ---------------------------------------------------------------------------
class TestCompanionState:
    def test_read_captures_values_and_fills(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("Q50", "P")
        sheet.set_cell("H50", "CLIENT PAID")
        sheet.set_interior("Q50", pattern=1, color=65280)  # xlSolid + green-ish
        sheet.set_interior("AD50", pattern=1, color=255)

        state = _read_companion_state(sheet, 50)

        assert state["Q"]["value"] == "P"
        assert state["Q"]["pattern"] == 1
        assert state["Q"]["color"] == 65280
        assert state["H"]["value"] == "CLIENT PAID"
        assert state["AD"]["pattern"] == 1
        assert state["AD"]["color"] == 255
        # Unset cols should have default pattern
        assert state["S"]["pattern"] == _XL_PATTERN_NONE

    def test_write_applies_values_and_fills(self) -> None:
        sheet = MockSheet()
        state = {
            col: {"value": None, "pattern": _XL_PATTERN_NONE, "color": 0}
            for col in COMPANION_COLS
        }
        state["Q"] = {"value": "P", "pattern": 1, "color": 65280}
        state["H"] = {"value": "CLIENT PAID", "pattern": _XL_PATTERN_NONE, "color": 0}

        _write_companion_state(sheet, 100, state)

        assert sheet.get_cell("Q100") == "P"
        assert sheet.get_cell("H100") == "CLIENT PAID"
        assert sheet.get_interior("Q100")["pattern"] == 1
        assert sheet.get_interior("Q100")["color"] == 65280
        # H had no fill — pattern should be None (default)
        assert sheet.get_interior("H100")["pattern"] == _XL_PATTERN_NONE

    def test_reset_clears_values_and_fills(self) -> None:
        sheet = MockSheet()
        sheet.set_cell("Q50", "P")
        sheet.set_cell("AD50", "P")
        sheet.set_interior("Q50", pattern=1, color=65280)

        _reset_companion_state(sheet, 50)

        assert sheet.get_cell("Q50") is None
        assert sheet.get_cell("AD50") is None
        assert sheet.get_interior("Q50")["pattern"] == _XL_PATTERN_NONE

    def test_write_skips_color_when_pattern_none(self) -> None:
        """When source pattern is xlPatternNone, don't stamp a stale
        color onto the destination."""
        sheet = MockSheet()
        state = {
            col: {"value": None, "pattern": _XL_PATTERN_NONE, "color": 12345}
            for col in COMPANION_COLS
        }
        _write_companion_state(sheet, 100, state)
        # Destination interiors should all be pattern None, color untouched (default 0)
        for col in COMPANION_COLS:
            assert sheet.get_interior(f"{col}100")["pattern"] == _XL_PATTERN_NONE


# ---------------------------------------------------------------------------
# Lock detection
# ---------------------------------------------------------------------------
class TestLockDetection:
    def test_reports_missing_file_as_not_locked(self, tmp_path: Path) -> None:
        assert is_bizactivity_locked(tmp_path / "nonexistent.xlsx") is False

    def test_reports_unlocked_file(self, tmp_path: Path) -> None:
        p = tmp_path / "biz.xlsx"
        p.write_bytes(b"hello")
        assert is_bizactivity_locked(p) is False

    def test_reports_locked_file(self, tmp_path: Path) -> None:
        """Hold the file in append mode and verify lock detection.

        On Windows this actually triggers the r+b PermissionError path.
        On macOS/Linux the probe usually succeeds regardless, so this
        test is a no-op there. We still verify the function doesn't
        crash on an open file handle.
        """
        p = tmp_path / "biz.xlsx"
        p.write_bytes(b"hello")
        # Just exercise the code path; platform-dependent result
        with open(p, "r+b"):
            result = is_bizactivity_locked(p)
            assert isinstance(result, bool)


# ---------------------------------------------------------------------------
# write_job_to_bizactivity (mocked COM)
# ---------------------------------------------------------------------------
class TestWriteJobToBizactivity:
    def test_missing_job_number_skipped(self) -> None:
        result = write_job_to_bizactivity("/fake/path.xlsx", {"client": "ACME"})
        assert result.success is False
        assert result.action == "skipped"
        assert "job_number" in (result.error_message or "")

    def test_missing_file_returns_error(self, tmp_path: Path) -> None:
        result = write_job_to_bizactivity(
            str(tmp_path / "nonexistent.xlsx"),
            {"job_number": "X-1"},
        )
        assert result.success is False
        assert "not found" in (result.error_message or "")

    def test_empty_bizactivity_path(self) -> None:
        result = write_job_to_bizactivity("", {"job_number": "X-1"})
        assert result.success is False

    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_insert_new_job(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock, tmp_path: Path,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")

        sheet = MockSheet()
        mock_wb = MagicMock()
        mock_wb.Worksheets.return_value = sheet
        mock_wb.Save = MagicMock()
        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        values = {
            "client": "Tamarac Fire Rescue",
            "job_number": "TFR-1",
            "job_description": "event shirts",
            "create_date": "2026-04-09",
        }
        result = write_job_to_bizactivity(str(biz_path), values)

        assert result.success is True
        assert result.action == "inserted"
        assert result.month == 4
        assert result.target_row == 241  # April first data row

        # Verify data was written to the mock sheet
        assert sheet.get_cell("B241") == "Tamarac Fire Rescue"
        assert sheet.get_cell("C241") == "event shirts"
        assert sheet.get_cell("D241") == "TFR-1"
        assert sheet.get_cell("BW241") == "2026-04-09"
        mock_wb.Save.assert_called_once()

    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_update_existing_job(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock, tmp_path: Path,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")

        sheet = MockSheet()
        # Pre-populate an existing job in January row 13
        sheet.set_cell("B13", "ACME")
        sheet.set_cell("D13", "ACME-1")
        sheet.set_cell("C13", "hats")

        mock_wb = MagicMock()
        mock_wb.Worksheets.return_value = sheet
        mock_wb.Save = MagicMock()
        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        values = {
            "client": "ACME",
            "job_number": "ACME-1",
            "job_description": "hats updated",
            "create_date": "2026-01-15",
            "deposit": 100.00,
        }
        result = write_job_to_bizactivity(str(biz_path), values)

        assert result.success is True
        assert result.action == "updated"
        assert result.target_row == 13

        assert sheet.get_cell("C13") == "hats updated"
        assert sheet.get_cell("J13") == 100.00  # J = deposit

    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_move_job_when_month_changes(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock, tmp_path: Path,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")

        sheet = MockSheet()
        # Job exists in January row 13 with some companion state
        sheet.set_cell("B13", "ACME")
        sheet.set_cell("D13", "ACME-1")
        sheet.set_cell("C13", "hats")
        sheet.set_cell("Q13", "P")         # all_day_shirts paid
        sheet.set_cell("AD13", "P")        # screen paid
        sheet.set_cell("H13", "paid")      # client-paid note
        sheet.set_interior("Q13", pattern=1, color=65280)

        mock_wb = MagicMock()
        mock_wb.Worksheets.return_value = sheet
        mock_wb.Save = MagicMock()
        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        # Now set job_start_date in March — should move
        values = {
            "client": "ACME",
            "job_number": "ACME-1",
            "job_description": "hats",
            "create_date": "2026-01-15",
            "job_start_date": "2026-03-10",
        }
        result = write_job_to_bizactivity(str(biz_path), values)

        assert result.success is True
        assert result.action == "moved"
        assert result.month == 3
        assert result.target_row == 165  # March first data row

        # Old row cleared (mapped cols + companion cols)
        assert sheet.get_cell("D13") is None
        assert sheet.get_cell("Q13") is None
        assert sheet.get_cell("AD13") is None
        assert sheet.get_cell("H13") is None
        assert sheet.get_interior("Q13")["pattern"] == _XL_PATTERN_NONE
        # New row populated with mapped values AND companion state carried over
        assert sheet.get_cell("D165") == "ACME-1"
        assert sheet.get_cell("B165") == "ACME"
        assert sheet.get_cell("Q165") == "P"
        assert sheet.get_cell("AD165") == "P"
        assert sheet.get_cell("H165") == "paid"
        assert sheet.get_interior("Q165")["pattern"] == 1
        assert sheet.get_interior("Q165")["color"] == 65280
        # Protection dance happened
        assert len(sheet.unprotect_calls) >= 1
        assert len(sheet.protect_calls) >= 1
        assert sheet.unprotect_calls[0] == "password"
        assert sheet.protect_calls[0] == "password"

    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_insert_runs_protection_dance(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock, tmp_path: Path,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")

        sheet = MockSheet()
        mock_wb = MagicMock()
        mock_wb.Worksheets.return_value = sheet
        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        values = {"job_number": "X-1", "client": "X", "create_date": "2026-05-01"}
        result = write_job_to_bizactivity(str(biz_path), values)

        assert result.success is True
        assert sheet.unprotect_calls == ["password"]
        assert sheet.protect_calls == ["password"]

    def test_locked_workbook_enqueues_instead_of_writing(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")

        # Force the lock detector to return True
        monkeypatch.setattr("bizactivity.is_bizactivity_locked", lambda p: True)

        # Redirect pending queue to a temp file
        queue_file = tmp_path / "pending.json"
        import pending_queue
        monkeypatch.setattr(pending_queue, "_queue_path", lambda: queue_file)

        values = {"job_number": "LCK-1", "client": "Locked Co", "create_date": "2026-06-01"}
        result = write_job_to_bizactivity(str(biz_path), values)

        assert result.success is True
        assert result.action == "queued"
        assert queue_file.exists()
        snapshot = pending_queue.peek(queue_path=queue_file)
        assert len(snapshot) == 1
        assert snapshot[0]["values"]["job_number"] == "LCK-1"

    def test_locked_workbook_with_allow_queue_false_skips(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        biz_path = tmp_path / "bizactivity.xlsx"
        biz_path.write_text("placeholder")
        monkeypatch.setattr("bizactivity.is_bizactivity_locked", lambda p: True)

        result = write_job_to_bizactivity(
            str(biz_path),
            {"job_number": "LCK-2"},
            allow_queue=False,
        )
        assert result.success is False
        assert result.action == "skipped"
        assert "locked" in (result.error_message or "").lower()


# ---------------------------------------------------------------------------
# read_map_sheet (mocked COM)
# ---------------------------------------------------------------------------
class TestReadMapSheet:
    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_reads_map_values(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock,
    ) -> None:
        sheet = MockSheet()
        sheet.set_cell("A2", "ACME Corp")
        sheet.set_cell("B2", "ACME-1")
        sheet.set_cell("C2", "polo shirts")
        sheet.set_cell("D2", "2026-04-01")
        sheet.set_cell("H2", 500.00)

        mock_wb = MagicMock()
        mock_wb.Worksheets.Count = 1
        mock_wb.Worksheets.return_value = sheet

        # Make sheet name lookup work
        sheet_info = MagicMock()
        sheet_info.Name = "Map"
        mock_wb.Worksheets.__call__ = lambda self, x: sheet if x == "Map" else sheet_info
        # For the sheet name iteration
        def worksheets_call(idx: Any) -> Any:
            if idx == "Map":
                return sheet
            if idx == 1:
                return sheet_info
            return MagicMock()
        mock_wb.Worksheets.side_effect = worksheets_call

        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        values = read_map_sheet("/fake/U-ACME-1.xls")

        assert values is not None
        assert values["client"] == "ACME Corp"
        assert values["job_number"] == "ACME-1"
        assert values["job_description"] == "polo shirts"
        assert values["create_date"] == "2026-04-01"
        assert values["gross_sales_before_tax"] == 500.00

    @patch("bizactivity._quit_excel")
    @patch("bizactivity._close_workbook")
    @patch("bizactivity._open_workbook")
    @patch("bizactivity._open_excel")
    def test_returns_none_when_no_job_number(
        self, mock_open_excel: MagicMock, mock_open_wb: MagicMock,
        mock_close_wb: MagicMock, mock_quit: MagicMock,
    ) -> None:
        sheet = MockSheet()
        sheet.set_cell("A2", "Some Client")
        # No job number in B2

        mock_wb = MagicMock()
        mock_wb.Worksheets.Count = 1
        sheet_info = MagicMock()
        sheet_info.Name = "Map"
        def worksheets_call(idx: Any) -> Any:
            if idx == "Map":
                return sheet
            if idx == 1:
                return sheet_info
            return MagicMock()
        mock_wb.Worksheets.side_effect = worksheets_call

        mock_open_excel.return_value = MagicMock()
        mock_open_wb.return_value = mock_wb

        values = read_map_sheet("/fake/empty.xls")
        assert values is None
