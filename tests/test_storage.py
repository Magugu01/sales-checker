from pathlib import Path

from storage import history_files, month_path, save_monthly_workbook


def test_monthly_history_prevents_duplicates_and_requires_update(
    tmp_path: Path,
) -> None:
    first = b"first workbook"
    second = b"updated workbook"

    assert save_monthly_workbook(first, 2026, 10, tmp_path) == "saved"
    assert save_monthly_workbook(first, 2026, 10, tmp_path) == "duplicate"
    assert save_monthly_workbook(second, 2026, 10, tmp_path) == "rejected"
    assert month_path(tmp_path, 2026, 10).read_bytes() == first

    assert (
        save_monthly_workbook(second, 2026, 10, tmp_path, replace_existing=True)
        == "saved"
    )
    assert month_path(tmp_path, 2026, 10).read_bytes() == second
    assert [path.name for path in history_files(tmp_path)] == ["2026-10.xlsx"]
