from datetime import date
from io import BytesIO

import pytest
from openpyxl import Workbook

from analyzer import (
    analyze_period,
    infer_month_from_filename,
    is_sunday_operator,
    parse_workbook,
)


def sample_workbook() -> bytes:
    workbook = Workbook()
    consolidated = workbook.active
    consolidated.title = "Consolidated"
    consolidated.append(["Code", "Name", "Branch", "Variances"])
    consolidated.append([1, "POMONA", 100, 0])
    consolidated.append([2, "MBARE", 200, 0])
    consolidated.append([3, "OTHER", 300, 0])

    day_one = workbook.create_sheet("Day 1")
    day_one.append(["Code", "Name", "NAV Revenue", "Branch", "Variances"])
    day_one.append([1, "POMONA", -100, 100, 0.10])
    day_one.append([2, "MBARE", -200, 200, -0.10])
    day_one.append([3, "OTHER", -300, 0, 0.1001])

    sunday = workbook.create_sheet("Day 4")
    sunday.append(["Code", "Name", "NAV Revenue", "Branch", "Variances"])
    sunday.append([1, "POMONA", -120, 120, 0])
    sunday.append([2, "MBARE", -240, None, 0])
    sunday.append([3, "OTHER", -350, 350, 0])

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def test_extracts_branch_roster_daily_sales_and_variance() -> None:
    parsed = parse_workbook(sample_workbook())

    assert len(parsed.branches) == 3
    pomona = next(branch for branch in parsed.branches if branch.name == "POMONA")
    assert parsed.daily_entries[1][pomona.key].sales == 100
    assert parsed.daily_entries[1][pomona.key].variance == 0.10


def test_missing_sales_sunday_rules_and_variance_tolerance() -> None:
    parsed = parse_workbook(sample_workbook())
    records, issues, coverage = analyze_period(
        parsed,
        date(2026, 1, 1),
        date(2026, 1, 4),
        variance_tolerance=0.10,
        today=date(2026, 1, 4),
    )

    day_one_issues = issues[issues["Date"] == date(2026, 1, 1)]
    assert "OTHER" in day_one_issues.loc[
        day_one_issues["Issue"] == "Variance exceeds tolerance", "Branch"
    ].tolist()
    assert "POMONA" not in day_one_issues["Branch"].tolist()
    sunday_records = records[records["Date"] == date(2026, 1, 4)]
    assert set(sunday_records["Branch"]) == {"POMONA", "MBARE"}
    assert sunday_records["Sales recorded"].sum() == 1
    sunday_coverage = coverage[coverage["Date"] == date(2026, 1, 4)].iloc[0]
    assert sunday_coverage["Expected branches"] == 2
    assert sunday_coverage["Missing sales"] == 1


def test_sunday_branch_name_aliases() -> None:
    assert is_sunday_operator("THE VET-CENTRE")
    assert is_sunday_operator("FCC MSASA")
    assert not is_sunday_operator("OTHER")


def test_filename_month_detection_does_not_match_summary_substrings() -> None:
    assert infer_month_from_filename("SALES SUMMARY OCTOBER.xlsx") == 10
    assert infer_month_from_filename("daily workbook.xlsx") is None


def test_future_end_date_needs_confirmation() -> None:
    parsed = parse_workbook(sample_workbook())
    with pytest.raises(ValueError, match="future"):
        analyze_period(
            parsed,
            date(2026, 1, 1),
            date(2026, 1, 5),
            today=date(2026, 1, 4),
        )


def test_excluded_branches_are_removed_from_analysis_and_coverage() -> None:
    parsed = parse_workbook(sample_workbook())
    records, issues, coverage = analyze_period(
        parsed,
        date(2026, 1, 1),
        date(2026, 1, 1),
        excluded_branches={"OTHER"},
        today=date(2026, 1, 1),
    )

    assert set(records["Branch"]) == {"POMONA", "MBARE"}
    assert "OTHER" not in set(issues["Branch"])
    assert coverage.iloc[0]["Expected branches"] == 2


def test_rejects_non_excel_upload_with_clear_message() -> None:
    with pytest.raises(ValueError, match="valid .xlsx"):
        parse_workbook(b"not an Excel workbook")
