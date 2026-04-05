from pathlib import Path

from sequence import detect_next_sequence


def _mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def test_detect_next_sequence_ignores_malformed_names(tmp_path: Path) -> None:
    _mkdir(tmp_path / "U-TFR-1 event shirts")
    _mkdir(tmp_path / "U-TFR-2")
    _mkdir(tmp_path / "U-TFR-two")
    _mkdir(tmp_path / "U-TFR-")
    _mkdir(tmp_path / "random")

    assert detect_next_sequence(str(tmp_path), "TFR") == 3


def test_detect_next_sequence_handles_gaps(tmp_path: Path) -> None:
    _mkdir(tmp_path / "U-CINTAS-1")
    _mkdir(tmp_path / "U-CINTAS-4 late")

    assert detect_next_sequence(str(tmp_path), "CINTAS") == 5


def test_detect_next_sequence_returns_one_when_none(tmp_path: Path) -> None:
    assert detect_next_sequence(str(tmp_path), "ABC") == 1
