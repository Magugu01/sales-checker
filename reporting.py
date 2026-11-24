from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Any
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    LongTable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    TableStyle,
)


def _safe_excel_value(value: Any) -> Any:
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    if pd.isna(value):
        return None
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime()
    return value


def _append_table(workbook: Workbook, title: str, frame: pd.DataFrame) -> None:
    sheet = workbook.create_sheet(title=title)
    if frame.empty:
        sheet.append(["No records"])
        return
    sheet.append([str(column) for column in frame.columns])
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.alignment = Alignment(wrap_text=True)
    for row in frame.itertuples(index=False, name=None):
        sheet.append([_safe_excel_value(value) for value in row])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column_cells in sheet.columns:
        values = [str(cell.value or "") for cell in list(column_cells)[:100]]
        width = min(max(max((len(value) for value in values), default=10) + 2, 12), 48)
        sheet.column_dimensions[column_cells[0].column_letter].width = width


def build_excel_report(
    records: pd.DataFrame,
    issues: pd.DataFrame,
    coverage: pd.DataFrame,
    start_date: date,
    end_date: date,
    source_name: str,
    variance_tolerance: float,
) -> bytes:
    missing = (
        issues[issues["Issue"].isin(["Missing sales", "Invalid sales value"])]
        if not issues.empty
        else issues
    )
    variance_alerts = (
        issues[issues["Issue"] == "Variance exceeds tolerance"]
        if not issues.empty
        else issues
    )
    input_errors = (
        issues[issues["Issue"].isin(["Invalid sales value", "Invalid variance value"])]
        if not issues.empty
        else issues
    )

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Sales assurance report"])
    summary.append(["Workbook", _safe_excel_value(source_name)])
    summary.append(["Analysis period", f"{start_date.isoformat()} to {end_date.isoformat()}"])
    summary.append(["Variance tolerance", variance_tolerance])
    summary.append(["Expected branch-days", len(records)])
    summary.append(["Sales recorded", int(records["Sales recorded"].sum()) if not records.empty else 0])
    summary.append(["Missing sales entries", len(missing)])
    summary.append(["Variance alerts", len(variance_alerts)])
    summary.append(["Input errors", len(input_errors)])
    summary["A1"].font = Font(bold=True, size=16, color="17324D")
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 60
    _append_table(workbook, "Daily coverage", coverage)
    _append_table(workbook, "Sales gaps", missing)
    _append_table(workbook, "Variance alerts", variance_alerts)
    _append_table(workbook, "Input errors", input_errors)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_pdf_report(
    records: pd.DataFrame,
    issues: pd.DataFrame,
    coverage: pd.DataFrame,
    start_date: date,
    end_date: date,
    source_name: str,
    variance_tolerance: float,
) -> bytes:
    missing = (
        issues[issues["Issue"].isin(["Missing sales", "Invalid sales value"])]
        if not issues.empty
        else issues
    )
    variance_alerts = (
        issues[issues["Issue"] == "Variance exceeds tolerance"]
        if not issues.empty
        else issues
    )
    recorded = int(records["Sales recorded"].sum()) if not records.empty else 0
    expected = len(records)
    completeness = recorded / expected * 100 if expected else 0.0
    sales_total = records["Sales"].sum(min_count=1) if expected else 0
    sales_total = 0 if pd.isna(sales_total) else float(sales_total)

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=13 * mm,
        leftMargin=13 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        title="Sales assurance report",
        author="IT-19 Sales Analysis Model",
    )
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ReportSubtitle",
            parent=styles["Normal"],
            textColor=colors.HexColor("#5e7184"),
            fontSize=9,
            leading=13,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportCell",
            parent=styles["Normal"],
            fontSize=7,
            leading=9,
            alignment=TA_LEFT,
            wordWrap="CJK",
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportHeading",
            parent=styles["Heading2"],
            textColor=colors.HexColor("#17324d"),
            spaceBefore=8,
            spaceAfter=6,
        )
    )

    def value_text(value: Any) -> str:
        if pd.isna(value):
            return ""
        if hasattr(value, "strftime"):
            return escape(value.strftime("%Y-%m-%d"))
        if isinstance(value, float):
            return escape(f"{value:,.2f}")
        return escape(str(value))

    def make_table(
        frame: pd.DataFrame,
        column_widths: list[float],
        max_rows: int | None = None,
    ) -> LongTable:
        selected = frame if max_rows is None else frame.head(max_rows)
        if selected.empty:
            selected = pd.DataFrame({"Status": ["No items to review."]})
        header = [Paragraph(escape(str(column)), styles["ReportCell"]) for column in selected.columns]
        body = [
            [Paragraph(value_text(value), styles["ReportCell"]) for value in row]
            for row in selected.itertuples(index=False, name=None)
        ]
        table = LongTable(
            [header, *body],
            colWidths=column_widths,
            repeatRows=1,
            hAlign="LEFT",
        )
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17324d")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d5dee7")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6f9")]),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        return table

    def add_page_number(canvas: Any, doc: Any) -> None:
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#5e7184"))
        canvas.drawRightString(
            landscape(A4)[0] - 13 * mm,
            8 * mm,
            f"Page {doc.page}",
        )
        canvas.restoreState()

    story = [
        Paragraph("Sales assurance report", styles["Title"]),
        Paragraph(
            f"Source: {escape(source_name)} · Period: {start_date.isoformat()} "
            f"to {end_date.isoformat()} · Variance tolerance: {variance_tolerance:,.2f}",
            styles["ReportSubtitle"],
        ),
        Spacer(1, 8 * mm),
    ]
    summary = [
        ["SALES COVERAGE", "ENTRIES RECORDED", "MISSING SALES", "VARIANCE ALERTS", "RECORDED SALES"],
        [
            f"{completeness:.1f}%",
            f"{recorded:,} / {expected:,}",
            f"{len(missing):,}",
            f"{len(variance_alerts):,}",
            f"{sales_total:,.2f}",
        ],
    ]
    summary_table = LongTable(summary, colWidths=[51 * mm] * 5, hAlign="LEFT")
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eaf1f7")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#53677a")),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#17324d")),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d5dee7")),
                ("PADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.extend(
        [
            summary_table,
            Paragraph("Daily coverage", styles["ReportHeading"]),
            make_table(coverage, [30 * mm, 53 * mm, 48 * mm, 45 * mm, 45 * mm]),
            Paragraph("Missing or invalid sales", styles["ReportHeading"]),
            make_table(
                missing.reindex(
                    columns=["Date", "Branch", "Branch code", "Issue", "Details"]
                ),
                [25 * mm, 43 * mm, 22 * mm, 46 * mm, 120 * mm],
            ),
            Paragraph("Variance alerts", styles["ReportHeading"]),
            make_table(
                variance_alerts.reindex(
                    columns=[
                        "Date",
                        "Branch",
                        "Branch code",
                        "Sales",
                        "Variance",
                        "Issue",
                        "Robust variance score",
                        "Details",
                    ]
                ),
                [
                    23 * mm,
                    40 * mm,
                    18 * mm,
                    24 * mm,
                    24 * mm,
                    39 * mm,
                    28 * mm,
                    62 * mm,
                ],
            ),
        ]
    )
    if not issues.empty:
        input_errors = issues[
            issues["Issue"].isin(["Invalid sales value", "Invalid variance value"])
        ]
        if not input_errors.empty:
            story.extend(
                [
                    PageBreak(),
                    Paragraph("Invalid spreadsheet values", styles["ReportHeading"]),
                    make_table(
                        input_errors.reindex(
                            columns=[
                                "Date",
                                "Branch",
                                "Branch code",
                                "Issue",
                                "Details",
                            ]
                        ),
                        [25 * mm, 43 * mm, 22 * mm, 46 * mm, 120 * mm],
                    ),
                ]
            )
    document.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buffer.getvalue()
