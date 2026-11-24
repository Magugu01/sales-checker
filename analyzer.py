from __future__ import annotations

import calendar
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, timedelta
from io import BytesIO
from math import isclose
from statistics import median
from typing import Any
from zipfile import BadZipFile

import pandas as pd
from openpyxl import load_workbook


SUNDAY_BRANCHES = {
    "POMONA",
    "MBARE",
    "FCC MSASA",
    "VET-CENTRE",
    "MADOKERO",
    "NORTON",
}

NAME_HEADERS = {"name", "branchname", "storename", "outletname"}
CODE_HEADERS = {"code", "branchcode", "storecode", "locationcode"}
SALES_HEADERS = {"branch", "sales", "branchsales", "salesamount", "branchamount"}
VARIANCE_HEADERS = {"variance", "variances"}
DAY_SHEET_PATTERN = re.compile(r"\bday\s*0*(\d{1,2})\b", re.IGNORECASE)


@dataclass(frozen=True)
class Branch:
    key: str
    name: str
    code: str | None


@dataclass(frozen=True)
class SalesEntry:
    sales_raw: Any
    sales: float | None
    variance_raw: Any
    variance: float | None


@dataclass
class ParsedWorkbook:
    branches: list[Branch]
    daily_entries: dict[int, dict[str, SalesEntry]]
    available_days: set[int]
    warnings: list[str]
    has_variance_column: bool


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def _normalize_branch_name(value: Any) -> str:
    normalized = _clean_text(value)
    if normalized.startswith("the") and normalized[3:] in {
        _clean_text(branch) for branch in SUNDAY_BRANCHES
    }:
        normalized = normalized[3:]
    return normalized


def _normalize_code(value: Any) -> str | None:
    if value is None or not str(value).strip():
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip().casefold()


def _branch_key(code: Any, name: Any) -> str:
    normalized_code = _normalize_code(code)
    if normalized_code:
        return f"code:{normalized_code}"
    normalized_name = _normalize_branch_name(name)
    return f"name:{normalized_name}" if normalized_name else ""


def _to_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = text.replace(",", "")
    text = re.sub(r"[$£€\s]", "", text)
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def _header_columns(sheet: Any) -> tuple[int, dict[str, int]] | None:
    for row_number in range(1, min(sheet.max_row, 40) + 1):
        headers = {
            _clean_text(sheet.cell(row_number, column).value): column
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(row_number, column).value is not None
        }
        name_column = next(
            (headers[name] for name in NAME_HEADERS if name in headers), None
        )
        sales_column = next(
            (headers[name] for name in SALES_HEADERS if name in headers), None
        )
        if name_column is None or sales_column is None:
            continue
        columns = {"name": name_column, "sales": sales_column}
        for key, aliases in (
            ("code", CODE_HEADERS),
            ("variance", VARIANCE_HEADERS),
        ):
            column = next(
                (headers[name] for name in aliases if name in headers), None
            )
            if column is not None:
                columns[key] = column
        return row_number, columns
    return None


def _sheet_day(sheet_name: str) -> int | None:
    match = DAY_SHEET_PATTERN.search(sheet_name)
    return int(match.group(1)) if match else None


def _rows_after_header(sheet: Any, header_row: int, columns: dict[str, int]):
    for row_number in range(header_row + 1, sheet.max_row + 1):
        name = sheet.cell(row_number, columns["name"]).value
        code = (
            sheet.cell(row_number, columns["code"]).value
            if "code" in columns
            else None
        )
        if name is None and code is None:
            continue
        if str(name or "").strip().casefold() in {"total", "grand total"}:
            continue
        yield row_number, code, name


def parse_workbook(workbook_bytes: bytes) -> ParsedWorkbook:
    try:
        workbook = load_workbook(
            BytesIO(workbook_bytes), read_only=False, data_only=True
        )
    except (BadZipFile, OSError, ValueError, TypeError) as error:
        raise ValueError("Could not read this workbook. Upload a valid .xlsx file.") from error

    daily_sheets = [
        (sheet, _sheet_day(sheet.title))
        for sheet in workbook.worksheets
        if _sheet_day(sheet.title) is not None
    ]
    if not daily_sheets:
        workbook.close()
        raise ValueError(
            "No daily worksheets named like 'Day 1' or 'Day 2' were found. "
            "This workbook needs its daily branch sales tabs to check missing entries."
        )

    warnings: list[str] = []
    available_days = {day for _, day in daily_sheets if day is not None}
    branch_by_key: dict[str, Branch] = {}
    key_by_name: dict[str, str] = {}
    sheet_layouts: dict[int, tuple[Any, int, dict[str, int]]] = {}
    has_variance_column = False

    consolidated = next(
        (
            sheet
            for sheet in workbook.worksheets
            if _clean_text(sheet.title) == "consolidated"
        ),
        None,
    )
    if consolidated is not None:
        layout = _header_columns(consolidated)
        if layout is not None:
            header_row, columns = layout
            for _, code, name in _rows_after_header(
                consolidated, header_row, columns
            ):
                if name is None or not str(name).strip():
                    continue
                key = _branch_key(code, name)
                if not key:
                    continue
                branch_by_key[key] = Branch(
                    key, str(name).strip(), _normalize_code(code)
                )
                key_by_name[_normalize_branch_name(name)] = key
    has_consolidated_roster = bool(branch_by_key)

    for sheet, day in daily_sheets:
        if day is None:
            continue
        layout = _header_columns(sheet)
        if layout is None:
            warnings.append(
                f"'{sheet.title}' has no recognizable Name and Branch sales headers."
            )
            continue
        header_row, columns = layout
        if day in sheet_layouts:
            warnings.append(f"More than one daily worksheet was found for day {day}.")
            continue
        sheet_layouts[day] = (sheet, header_row, columns)
        has_variance_column = has_variance_column or "variance" in columns
        for _, code, name in _rows_after_header(sheet, header_row, columns):
            if name is None or not str(name).strip():
                continue
            key = _branch_key(code, name)
            if key not in branch_by_key and _normalize_branch_name(name) in key_by_name:
                key = key_by_name[_normalize_branch_name(name)]
            if key not in branch_by_key:
                branch_by_key[key] = Branch(
                    key, str(name).strip(), _normalize_code(code)
                )
                key_by_name[_normalize_branch_name(name)] = key
                if has_consolidated_roster:
                    warnings.append(
                        f"'{sheet.title}' includes branch '{str(name).strip()}' "
                        "not listed in Consolidated; it was included in the roster."
                    )

    if not branch_by_key:
        workbook.close()
        raise ValueError(
            "No branch rows were found. Expected a Name column and a Branch sales "
            "column in the consolidated sheet or daily worksheets."
        )

    daily_entries: dict[int, dict[str, SalesEntry]] = {}
    for day, (sheet, header_row, columns) in sheet_layouts.items():
        entries: dict[str, SalesEntry] = {}
        for row_number, code, name in _rows_after_header(
            sheet, header_row, columns
        ):
            if name is None or not str(name).strip():
                continue
            key = _branch_key(code, name)
            if key not in branch_by_key and _normalize_branch_name(name) in key_by_name:
                key = key_by_name[_normalize_branch_name(name)]
            if key not in branch_by_key:
                warnings.append(
                    f"'{sheet.title}' row {row_number} contains an unmapped branch."
                )
                continue
            if key in entries:
                warnings.append(
                    f"'{sheet.title}' repeats branch '{name}'; the first row was used."
                )
                continue
            sales_raw = sheet.cell(row_number, columns["sales"]).value
            variance_raw = (
                sheet.cell(row_number, columns["variance"]).value
                if "variance" in columns
                else None
            )
            entries[key] = SalesEntry(
                sales_raw=sales_raw,
                sales=_to_number(sales_raw),
                variance_raw=variance_raw,
                variance=_to_number(variance_raw),
            )
        daily_entries[day] = entries

    workbook.close()
    return ParsedWorkbook(
        branches=sorted(
            branch_by_key.values(), key=lambda branch: branch.name.casefold()
        ),
        daily_entries=daily_entries,
        available_days=available_days,
        warnings=list(dict.fromkeys(warnings)),
        has_variance_column=has_variance_column,
    )


def is_sunday_operator(branch_name: str) -> bool:
    allowed = {_clean_text(branch) for branch in SUNDAY_BRANCHES}
    return _normalize_branch_name(branch_name) in allowed


def analyze_period(
    parsed: ParsedWorkbook,
    start_date: date,
    end_date: date,
    variance_tolerance: float = 0.10,
    confirm_future_dates: bool = False,
    today: date | None = None,
    excluded_branches: set[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if start_date > end_date:
        raise ValueError("The start date must be on or before the end date.")
    if variance_tolerance < 0:
        raise ValueError("Variance tolerance cannot be negative.")
    current_date = today or date.today()
    if end_date > current_date and not confirm_future_dates:
        raise ValueError(
            "The selected end date is in the future. Confirm that these dates should be included."
        )

    records: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    excluded = {name.casefold() for name in (excluded_branches or set())}
    for offset in range((end_date - start_date).days + 1):
        current_day = start_date + timedelta(days=offset)
        day_number = current_day.day
        if current_day.weekday() == 6:
            expected_branches = [
                branch
                for branch in parsed.branches
                if is_sunday_operator(branch.name)
            ]
        else:
            expected_branches = parsed.branches
        expected_branches = [
            branch for branch in expected_branches if branch.name.casefold() not in excluded
        ]
        day_entries = parsed.daily_entries.get(day_number, {})
        sheet_available = day_number in parsed.available_days

        for branch in expected_branches:
            entry = day_entries.get(branch.key)
            raw_sales = entry.sales_raw if entry is not None else None
            sales = entry.sales if entry is not None else None
            raw_variance = entry.variance_raw if entry is not None else None
            variance = entry.variance if entry is not None else None
            record = {
                "Date": current_day,
                "Day": day_number,
                "Branch": branch.name,
                "Branch code": branch.code or "",
                "Sales": sales,
                "Variance": variance,
                "Worksheet available": sheet_available,
                "Sales recorded": sales is not None,
            }
            records.append(record)

            if sales is None:
                if raw_sales is not None and str(raw_sales).strip():
                    issues.append(
                        {
                            **record,
                            "Issue": "Invalid sales value",
                            "Details": (
                                f"Branch sales value {raw_sales!r} is not numeric."
                            ),
                        }
                    )
                else:
                    reason = (
                        f"No '{day_number}' daily worksheet was found."
                        if not sheet_available
                        else "No numeric value was found in the Branch sales column."
                    )
                    issues.append(
                        {**record, "Issue": "Missing sales", "Details": reason}
                    )
                continue

            if (
                raw_variance is not None
                and str(raw_variance).strip()
                and variance is None
            ):
                issues.append(
                    {
                        **record,
                        "Issue": "Invalid variance value",
                        "Details": (
                            f"Variance value {raw_variance!r} is not numeric."
                        ),
                    }
                )
            elif variance is not None and abs(variance) > variance_tolerance and not isclose(
                abs(variance), variance_tolerance, rel_tol=0.0, abs_tol=1e-9
            ):
                issues.append(
                    {
                        **record,
                        "Issue": "Variance exceeds tolerance",
                        "Details": (
                            f"Absolute variance {abs(variance):,.2f} exceeds the "
                            f"configured tolerance of {variance_tolerance:,.2f}."
                        ),
                    }
                )

    records_frame = pd.DataFrame.from_records(records)
    if not records_frame.empty:
        valid_variances = records_frame["Variance"].dropna().astype(float)
        if valid_variances.empty:
            center = 0.0
            scale = variance_tolerance or 1e-9
        else:
            center = float(valid_variances.median())
            mad = float((valid_variances - center).abs().median())
            scale = max(1.4826 * mad, variance_tolerance, 1e-9)
        records_frame["Robust variance score"] = records_frame["Variance"].apply(
            lambda value: (
                round(abs(float(value) - center) / scale, 2)
                if pd.notna(value)
                else None
            )
        )
        score_by_key = {
            (row["Date"], row["Branch"]): row["Robust variance score"]
            for row in records_frame.to_dict(orient="records")
        }
        for issue in issues:
            issue["Robust variance score"] = score_by_key.get(
                (issue["Date"], issue["Branch"])
            )

    issues_frame = pd.DataFrame.from_records(issues)
    if issues_frame.empty:
        issues_frame = pd.DataFrame(
            columns=[
                "Date",
                "Day",
                "Branch",
                "Branch code",
                "Sales",
                "Variance",
                "Issue",
                "Robust variance score",
                "Details",
            ]
        )
    return records_frame, issues_frame, _daily_coverage(records_frame)


def _daily_coverage(records: pd.DataFrame) -> pd.DataFrame:
    if records.empty:
        return pd.DataFrame(
            columns=[
                "Date",
                "Expected branches",
                "Recorded sales",
                "Missing sales",
                "Coverage %",
            ]
        )
    coverage = records.groupby("Date", as_index=False).agg(
        **{
            "Expected branches": ("Branch", "count"),
            "Recorded sales": ("Sales recorded", "sum"),
        }
    )
    coverage["Missing sales"] = coverage["Expected branches"] - coverage["Recorded sales"]
    coverage["Coverage %"] = (
        coverage["Recorded sales"] / coverage["Expected branches"] * 100
    ).round(1)
    return coverage


def infer_month_from_filename(filename: str) -> int | None:
    tokens = re.findall(r"[a-z]+", filename.casefold())
    for month in range(1, 13):
        if (
            calendar.month_name[month].casefold() in tokens
            or calendar.month_abbr[month].casefold() in tokens
        ):
            return month
    return None


def infer_year_from_filename(filename: str) -> int | None:
    match = re.search(r"\b(20\d{2})\b", filename)
    return int(match.group(1)) if match else None


def format_month(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"
