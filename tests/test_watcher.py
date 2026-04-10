from __future__ import annotations

import re
import sys

import pytest

# watcher.py imports watchdog which is Windows-only; test the constants directly
try:
    from watcher import _WORKBOOK_PATTERN, _IGNORE_PREFIXES
except ImportError:
    # Replicate the constants so pattern tests run on Mac
    _WORKBOOK_PATTERN = re.compile(r"^U-.+\.(xls|xlsm|xlsx)$", re.IGNORECASE)
    _IGNORE_PREFIXES = ("~$", "~")


class TestWorkbookPattern:
    def test_matches_xls(self) -> None:
        assert _WORKBOOK_PATTERN.match("U-ACME-1.xls")

    def test_matches_xlsm(self) -> None:
        assert _WORKBOOK_PATTERN.match("U-TFR-3.xlsm")

    def test_matches_xlsx(self) -> None:
        assert _WORKBOOK_PATTERN.match("U-CINTAS-12.xlsx")

    def test_matches_with_description(self) -> None:
        # The file itself won't have description in name (workbook pattern is U-ABBR-SEQ.ext)
        # but the regex should still match U- prefix files
        assert _WORKBOOK_PATTERN.match("U-TFR-1.xls")

    def test_rejects_non_u_prefix(self) -> None:
        assert not _WORKBOOK_PATTERN.match("template.xls")

    def test_rejects_non_excel(self) -> None:
        assert not _WORKBOOK_PATTERN.match("U-ACME-1.pdf")

    def test_case_insensitive(self) -> None:
        assert _WORKBOOK_PATTERN.match("U-ACME-1.XLS")
        assert _WORKBOOK_PATTERN.match("U-ACME-1.Xlsm")


class TestIgnorePrefixes:
    def test_ignores_excel_lock_file(self) -> None:
        assert any("~$U-ACME-1.xls".startswith(p) for p in _IGNORE_PREFIXES)

    def test_ignores_tilde_temp(self) -> None:
        assert any("~temp.xls".startswith(p) for p in _IGNORE_PREFIXES)

    def test_does_not_ignore_normal_file(self) -> None:
        assert not any("U-ACME-1.xls".startswith(p) for p in _IGNORE_PREFIXES)
