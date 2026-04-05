from sequence import (
    build_folder_job_number,
    build_order_folder_name,
    build_workbook_filename,
    sanitize_job_description,
)


def test_sanitize_job_description_removes_windows_invalid_chars() -> None:
    raw = ' event: shirts / spring? "2026" '
    assert sanitize_job_description(raw) == "event shirts spring 2026"


def test_build_order_folder_name_with_description() -> None:
    assert (
        build_order_folder_name("U", "TFR", 7, "event shirts")
        == "U-TFR-7 event shirts"
    )


def test_build_order_folder_name_without_description() -> None:
    assert build_order_folder_name("U", "TFR", 7, "   ") == "U-TFR-7"


def test_build_workbook_filename() -> None:
    assert build_workbook_filename("U", "TFR", 7, ".xls") == "U-TFR-7.xls"


def test_build_folder_job_number() -> None:
    assert build_folder_job_number("TFR", 7) == "TFR-7"
