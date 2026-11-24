from datetime import date
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from reporting import build_excel_report, build_pdf_report


def test_reports_are_readable_and_escape_workbook_text() -> None:
    records = pd.DataFrame(
        [
            {
                "Date": date(2026, 10, 1),
                "Day": 1,
                "Branch": "<POMONA>",
                "Branch code": "1",
                "Sales": 100.0,
                "Variance": 0.2,
                "Worksheet available": True,
                "Sales recorded": True,
                "Robust variance score": 2.0,
            }
        ]
    )
    issues = pd.DataFrame(
        [
            {
                "Date": date(2026, 10, 1),
                "Day": 1,
                "Branch": "<POMONA>",
                "Branch code": "1",
                "Sales": 100.0,
                "Variance": 0.2,
                "Issue": "Variance exceeds tolerance",
                "Details": "Review variance",
            }
        ]
    )
    coverage = pd.DataFrame(
        [
            {
                "Date": date(2026, 10, 1),
                "Expected branches": 1,
                "Recorded sales": 1,
                "Missing sales": 0,
                "Coverage %": 100.0,
            }
        ]
    )

    excel_bytes = build_excel_report(
        records, issues, coverage, date(2026, 10, 1), date(2026, 10, 1),
        "=monthly.xlsx", 0.10
    )
    workbook = load_workbook(BytesIO(excel_bytes), data_only=True)
    assert workbook.sheetnames == [
        "Summary",
        "Daily coverage",
        "Sales gaps",
        "Variance alerts",
        "Input errors",
    ]
    assert workbook["Summary"]["B2"].value == "'=monthly.xlsx"
    workbook.close()

def test_pdf_report_is_a_valid_downloadable_document() -> None:
    records = pd.DataFrame(
        [
            {
                "Date": date(2026, 10, 1),
                "Day": 1,
                "Branch": "POMONA",
                "Branch code": "1",
                "Sales": 100.0,
                "Variance": 0.0,
                "Worksheet available": True,
                "Sales recorded": True,
                "Robust variance score": 0.0,
            }
        ]
    )
    issues = pd.DataFrame(columns=["Date", "Branch", "Variance", "Issue"])
    coverage = pd.DataFrame(
        [
            {
                "Date": date(2026, 10, 1),
                "Expected branches": 1,
                "Recorded sales": 1,
                "Missing sales": 0,
                "Coverage %": 100.0,
            }
        ]
    )

    pdf_report = build_pdf_report(
        records,
        issues,
        coverage,
        date(2026, 10, 1),
        date(2026, 10, 1),
        "October.xlsx",
        0.10,
    )

    assert pdf_report.startswith(b"%PDF-")
    assert len(pdf_report) > 1000
