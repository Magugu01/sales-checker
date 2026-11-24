from __future__ import annotations

import calendar
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from analyzer import (
    ParsedWorkbook,
    analyze_period,
    infer_month_from_filename,
    infer_year_from_filename,
    is_sunday_operator,
    parse_workbook,
)
from reporting import build_excel_report, build_pdf_report
from storage import (
    HISTORY_PATTERN,
    file_digest,
    history_files,
    month_path,
    save_monthly_workbook,
)


APP_DIRECTORY = Path(__file__).resolve().parent
HISTORY_DIRECTORY = APP_DIRECTORY / "history"
LOGO_PATH = APP_DIRECTORY / "assets" / "farm-city-centre-logo.bmp"
PAGES = [
    "Executive Dashboard",
    "Missing Sales Detector",
    "Variance & Trends",
    "Data Management",
    "Report Centre",
]
FOREST_GREEN = "#176B3A"
WARM_GOLD = "#D9A441"
SOFT_WHITE = "#F5F7F3"
CHARCOAL = "#26332D"
ALERT_RED = "#C0392B"
INFO_BLUE = "#2878B5"
CHART_COLORS = [FOREST_GREEN, WARM_GOLD, INFO_BLUE, ALERT_RED, CHARCOAL]
NAVIGATION_GROUPS = [
    ("OVERVIEW", ("Executive Dashboard",)),
    ("SALES INTELLIGENCE", ("Missing Sales Detector", "Variance & Trends")),
    ("DATA", ("Data Management",)),
    ("REPORTING", ("Report Centre",)),
]
PAGE_ICONS = {
    "Executive Dashboard": ":material/dashboard:",
    "Missing Sales Detector": ":material/search:",
    "Variance & Trends": ":material/trending_up:",
    "Data Management": ":material/folder_open:",
    "Report Centre": ":material/description:",
}
PLOT_CONFIG = {
    "displaylogo": False,
    "toImageButtonOptions": {
        "format": "png",
        "width": 1920,
        "height": 1080,
        "scale": 2,
    },
}


def _set_current_workbook(content: bytes, filename: str) -> None:
    with st.spinner("Validating and loading the workbook..."):
        parsed = parse_workbook(content)
    for key in tuple(st.session_state):
        if key.startswith(("start_", "end_")):
            st.session_state.pop(key, None)
    st.session_state["workbook_bytes"] = content
    st.session_state["workbook_hash"] = file_digest(content)
    st.session_state["workbook_name"] = filename
    st.session_state["workbook_loaded_at"] = datetime.now().astimezone().strftime(
        "%d %b %Y, %H:%M"
    )
    st.session_state["parsed_workbook"] = parsed
    st.session_state["analysis_result"] = None
    st.session_state["excluded_branches"] = []
    st.session_state["confirm_future_dates"] = False

    st.session_state["period_month"] = (
        infer_month_from_filename(filename) or date.today().month
    )
    st.session_state["period_year"] = (
        infer_year_from_filename(filename) or date.today().year
    )


def _clear_workbook() -> None:
    for key in (
        "workbook_bytes",
        "workbook_hash",
        "workbook_name",
        "workbook_loaded_at",
        "parsed_workbook",
        "analysis_result",
        "excluded_branches",
        "confirm_future_dates",
        "analyze_full_month_requested",
    ):
        st.session_state.pop(key, None)
    st.session_state["upload_widget_version"] = (
        int(st.session_state.get("upload_widget_version", 0)) + 1
    )


def _request_full_month_analysis() -> None:
    year = int(st.session_state["period_year"])
    month = int(st.session_state["period_month"])
    month_key = f"{year}_{month}"
    st.session_state[f"start_{month_key}"] = date(year, month, 1)
    st.session_state[f"end_{month_key}"] = date(
        year, month, calendar.monthrange(year, month)[1]
    )
    st.session_state["analyze_full_month_requested"] = True


def _welcome_header(active_page: str) -> None:
    page_descriptions = {
        "Executive Dashboard": "Monitor performance. Identify opportunities. Make smarter business decisions.",
        "Missing Sales Detector": "Find missing branch sales and follow up on reporting gaps.",
        "Variance & Trends": "Review variance exceptions and compare sales trends.",
        "Data Management": "Manage workbook imports, monthly history, and branch filters.",
        "Report Centre": "Review and export sales assurance reports.",
    }
    st.markdown(
        f"""
        <section class="welcome-hero">
          <p class="welcome-eyebrow">BRANCH PERFORMANCE INTELLIGENCE</p>
          <h1 class="welcome-title">
            <span>{active_page}</span>
          </h1>
          <p class="welcome-subtitle">{page_descriptions[active_page]}</p>
        </section>
        """,
        unsafe_allow_html=True,
    )
    if active_page != "Executive Dashboard":
        if st.button(
            "← Back to Executive Dashboard",
            key="return_to_executive_dashboard",
        ):
            st.session_state["active_page"] = "Executive Dashboard"
            st.rerun()


def _show_kpi_card(
    container: object,
    label: str,
    value: str,
    detail: str,
    accent: str = FOREST_GREEN,
) -> None:
    container.markdown(
        f"""
        <div class="kpi-card" style="--kpi-accent:{accent}">
          <p class="kpi-label">{label}</p>
          <p class="kpi-value">{value}</p>
          <p class="kpi-detail">{detail}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _style_chart(figure: object, height: int = 360) -> object:
    dark_mode = st.session_state.get("dark_mode", False)
    figure.update_layout(
        template="plotly_dark" if dark_mode else "plotly_white",
        height=height,
        margin=dict(l=12, r=16, t=30, b=12),
        font=dict(
            family="Segoe UI, Arial, sans-serif",
            color="#E6EEE8" if dark_mode else CHARCOAL,
        ),
        hoverlabel=dict(bgcolor=CHARCOAL, font_color="white"),
    )
    return figure


def _show_issue_table(
    issues: pd.DataFrame, issue_names: set[str], title: str | None = None
) -> None:
    subset = issues[issues["Issue"].isin(issue_names)] if not issues.empty else issues
    if title:
        st.subheader(title)
    if subset.empty:
        st.success("No items to review for this category.")
    else:
        st.dataframe(subset, hide_index=True, width="stretch")


def _selected_issues(issues: pd.DataFrame, issue_names: set[str]) -> pd.DataFrame:
    return issues[issues["Issue"].isin(issue_names)] if not issues.empty else issues


def _show_welcome_page() -> None:
    st.subheader("Your business intelligence workspace")
    st.write(
        "A clear monthly view of sales performance, reporting activity, and "
        "branch-level follow-up."
    )
    card_columns = st.columns(4)
    _show_kpi_card(
        card_columns[0], "TOTAL SALES", "—", "Selected reporting period"
    )
    _show_kpi_card(
        card_columns[1], "ACTIVE BRANCHES", "—", "Branches reporting", INFO_BLUE
    )
    _show_kpi_card(
        card_columns[2], "AVERAGE SALES", "—", "Per active branch", WARM_GOLD
    )
    _show_kpi_card(
        card_columns[3], "MISSING REPORTS", "—", "Branches to follow up", ALERT_RED
    )
    st.markdown(
        """
        <div class="welcome-callout">
          <strong>Start with your monthly workbook</strong>
          <span>Use the sidebar upload to validate branch reporting and unlock the sales dashboard.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="welcome-features">
          <div><span class="feature-mark">01</span><strong>See the whole picture</strong><p>Track sales, active branches, reporting completeness and exceptions.</p></div>
          <div><span class="feature-mark">02</span><strong>Focus on what matters</strong><p>Compare branch performance and follow up on missing reports.</p></div>
          <div><span class="feature-mark">03</span><strong>Share with confidence</strong><p>Export a clear Excel or PDF report for your team.</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        "for further assistance and help conduct IT-19 2026/27 Mr Magugu."
    )


def _show_executive_dashboard(
    records: pd.DataFrame,
    issues: pd.DataFrame,
    coverage: pd.DataFrame,
) -> None:
    missing = _selected_issues(
        issues, {"Missing sales", "Invalid sales value"}
    )
    variance_alerts = _selected_issues(
        issues, {"Variance exceeds tolerance"}
    )
    recorded = int(records["Sales recorded"].sum()) if not records.empty else 0
    expected = len(records)
    sales_total = records["Sales"].sum(min_count=1) if expected else 0
    sales_total = 0 if pd.isna(sales_total) else float(sales_total)
    days_reviewed = int(coverage["Date"].nunique()) if not coverage.empty else 0
    active_branches = (
        int(records.loc[records["Sales recorded"], "Branch"].nunique())
        if expected
        else 0
    )
    branches_with_gaps = (
        int(missing["Branch"].nunique()) if not missing.empty else 0
    )
    average_sales = sales_total / active_branches if active_branches else 0.0

    metrics = st.columns(4)
    _show_kpi_card(
        metrics[0],
        "TOTAL SALES",
        f"{sales_total:,.2f}",
        "Selected reporting period",
    )
    _show_kpi_card(
        metrics[1],
        "ACTIVE BRANCHES",
        f"{active_branches:,} / {records['Branch'].nunique() if expected else 0:,}",
        "Branches with at least one report",
        INFO_BLUE,
    )
    _show_kpi_card(
        metrics[2],
        "AVERAGE SALES",
        f"{average_sales:,.2f}",
        "Per active branch",
        WARM_GOLD,
    )
    _show_kpi_card(
        metrics[3],
        "MISSING REPORTS",
        f"{branches_with_gaps:,}",
        f"{len(missing):,} missing branch-days",
        ALERT_RED,
    )
    reporting_rate = recorded / expected * 100 if expected else 0.0
    st.caption(
        f"Reporting completeness: **{reporting_rate:.1f}%** · "
        f"Variance alerts: **{len(variance_alerts):,}** · "
        f"Sales records: **{recorded:,} / {expected:,}**"
    )

    left, right = st.columns(2)
    if coverage.empty:
        st.info("No branch-days are available for this reporting period.")
        return

    with left:
        st.subheader("Daily sales and reporting coverage")
        daily_sales = (
            records.groupby("Date", as_index=False)
            .agg(Sales=("Sales", "sum"))
            .merge(coverage[["Date", "Coverage %"]], on="Date", how="left")
        )
        sales_chart = px.bar(
            daily_sales,
            x="Date",
            y="Sales",
            color_discrete_sequence=[CHART_COLORS[0]],
            labels={"Sales": "Recorded sales"},
        )
        sales_chart.add_scatter(
            x=daily_sales["Date"],
            y=daily_sales["Coverage %"],
            name="Sales coverage %",
            mode="lines+markers",
            yaxis="y2",
            line=dict(color=CHART_COLORS[3], width=3),
        )
        sales_chart.update_layout(
            yaxis=dict(title="Recorded sales"),
            yaxis2=dict(
                title="Coverage %",
                overlaying="y",
                side="right",
                range=[0, 100],
                ticksuffix="%",
            ),
            legend=dict(orientation="h", y=1.12),
            barmode="group",
        )
        st.plotly_chart(
            _style_chart(sales_chart, 420),
            width="stretch",
            config=PLOT_CONFIG,
        )

    with right:
        st.subheader("Branch sales performance")
        branch_sales = (
            records.groupby("Branch", as_index=False)
            .agg(
                Sales=("Sales", "sum"),
                Days=("Date", "nunique"),
                Recorded=("Sales recorded", "sum"),
            )
            .sort_values("Sales", ascending=False)
            .head(15)
        )
        performance = px.bar(
            branch_sales.sort_values("Sales"),
            x="Sales",
            y="Branch",
            orientation="h",
            color="Sales",
            color_continuous_scale=["#e2efe6", FOREST_GREEN],
            hover_data=["Days", "Recorded"],
        )
        performance.update_layout(coloraxis_showscale=False, yaxis_title="")
        st.plotly_chart(
            _style_chart(performance, 420),
            width="stretch",
            config=PLOT_CONFIG,
        )

    left, right = st.columns(2)
    with left:
        st.subheader("Reporting coverage by day")
        coverage_chart = px.area(
            coverage,
            x="Date",
            y="Coverage %",
            markers=True,
            range_y=[0, 100],
            color_discrete_sequence=[CHART_COLORS[1]],
        )
        coverage_chart.update_layout(yaxis_title="Branches reporting (%)")
        st.plotly_chart(
            _style_chart(coverage_chart, 340),
            width="stretch",
            config=PLOT_CONFIG,
        )
    with right:
        st.subheader("Review queue")
        queue = pd.DataFrame(
            {
                "Review category": ["Missing sales", "Variance alerts"],
                "Items": [len(missing), len(variance_alerts)],
            }
        )
        queue_chart = px.bar(
            queue,
            x="Review category",
            y="Items",
            color="Review category",
            color_discrete_sequence=[CHART_COLORS[2], CHART_COLORS[3]],
            text="Items",
        )
        queue_chart.update_layout(showlegend=False, xaxis_title="")
        st.plotly_chart(
            _style_chart(queue_chart, 340),
            width="stretch",
            config=PLOT_CONFIG,
        )
    st.caption(
        f"{days_reviewed} reporting days analyzed. Sunday expectations include only "
        "the six configured Sunday branches."
    )


def _show_missing_detector(
    records: pd.DataFrame, issues: pd.DataFrame, coverage: pd.DataFrame
) -> None:
    missing = _selected_issues(issues, {"Missing sales", "Invalid sales value"})
    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("Missing-sales calendar")
        if records.empty:
            st.info("No branch-days are expected in the selected period.")
        else:
            calendar_data = records.pivot(
                index="Branch", columns="Day", values="Sales recorded"
            ).astype(float)
            figure = px.imshow(
                1 - calendar_data,
                aspect="auto",
                color_continuous_scale=[FOREST_GREEN, ALERT_RED],
                range_color=(0, 1),
                labels={"x": "Day of month", "y": "Branch", "color": "Missing"},
            )
            figure.update_layout(
                coloraxis_showscale=False,
                height=min(900, max(380, 19 * len(calendar_data.index))),
            )
            st.plotly_chart(
                _style_chart(figure),
                width="stretch",
                config=PLOT_CONFIG,
            )
            st.caption(
                "Red cells are missing sales; green cells are recorded sales. "
                "Non-operating Sundays are excluded."
            )
    with right:
        st.subheader("Branches needing follow-up")
        if missing.empty:
            st.success("All expected branches have sales entries.")
        else:
            by_branch = (
                missing.groupby("Branch", as_index=False)
                .size()
                .rename(columns={"size": "Missing days"})
                .sort_values("Missing days")
            )
            figure = px.bar(
                by_branch,
                x="Missing days",
                y="Branch",
                orientation="h",
                color="Missing days",
                color_continuous_scale=[WARM_GOLD, ALERT_RED],
            )
            figure.update_layout(coloraxis_showscale=False, yaxis_title="")
            st.plotly_chart(
                _style_chart(figure, max(360, min(800, 22 * len(by_branch)))),
                width="stretch",
                config=PLOT_CONFIG,
            )
        st.metric(
            "Expected branch-days",
            f"{int(coverage['Expected branches'].sum()) if not coverage.empty else 0:,}",
        )
        st.metric("Sales gaps", f"{len(missing):,}")

    st.divider()
    _show_issue_table(
        issues,
        {"Missing sales", "Invalid sales value"},
        "Sales entries to investigate",
    )


def _show_variance_analysis(
    records: pd.DataFrame,
    issues: pd.DataFrame,
    tolerance: float,
    start_date: date,
    end_date: date,
    excluded_branches: list[str],
) -> None:
    st.caption(
        f"Variances above +/-{tolerance:,.2f} are flagged. The robust score adds "
        "median/MAD context; confirm spreadsheet source values before correcting them."
    )
    alerts = _selected_issues(issues, {"Variance exceeds tolerance"})
    valid_variances = records.dropna(subset=["Variance"]).copy()
    metrics = st.columns(4)
    metrics[0].metric("Variance alerts", f"{len(alerts):,}")
    metrics[1].metric(
        "Positive alerts",
        f"{int((alerts['Variance'] > tolerance).sum()) if not alerts.empty else 0:,}",
    )
    metrics[2].metric(
        "Negative alerts",
        f"{int((alerts['Variance'] < -tolerance).sum()) if not alerts.empty else 0:,}",
    )
    median_variance = (
        float(valid_variances["Variance"].median())
        if not valid_variances.empty
        else 0.0
    )
    metrics[3].metric("Median variance", f"{median_variance:,.2f}")

    left, right = st.columns(2)
    with left:
        st.subheader("Variance trend by day")
        if valid_variances.empty:
            st.info("No numeric variance values were found.")
        else:
            daily = (
                valid_variances.groupby("Date", as_index=False)
                .agg(
                    Median=("Variance", "median"),
                    Maximum=("Variance", "max"),
                    Minimum=("Variance", "min"),
                )
            )
            figure = px.line(
                daily,
                x="Date",
                y=["Maximum", "Median", "Minimum"],
                markers=True,
                color_discrete_sequence=CHART_COLORS[:3],
            )
            figure.add_hline(
                y=tolerance, line_dash="dash", line_color=CHART_COLORS[3]
            )
            figure.add_hline(
                y=-tolerance, line_dash="dash", line_color=CHART_COLORS[3]
            )
            st.plotly_chart(
                _style_chart(figure, 400),
                width="stretch",
                config=PLOT_CONFIG,
            )
    with right:
        st.subheader("Variance distribution")
        if valid_variances.empty:
            st.info("No numeric variance values were found.")
        else:
            figure = px.histogram(
                valid_variances,
                x="Variance",
                nbins=36,
                color_discrete_sequence=[CHART_COLORS[1]],
            )
            figure.add_vline(
                x=tolerance, line_dash="dash", line_color=CHART_COLORS[3]
            )
            figure.add_vline(
                x=-tolerance, line_dash="dash", line_color=CHART_COLORS[3]
            )
            st.plotly_chart(
                _style_chart(figure, 400),
                width="stretch",
                config=PLOT_CONFIG,
            )

    st.subheader("Branch sales trend")
    branches = sorted(records["Branch"].unique(), key=str.casefold)
    if branches:
        selected_branch = st.selectbox("Choose a branch", branches)
        branch_data = records[records["Branch"] == selected_branch]
        figure = px.line(
            branch_data,
            x="Date",
            y="Sales",
            markers=True,
            color_discrete_sequence=[CHART_COLORS[0]],
        )
        figure.update_layout(yaxis_title="Recorded sales")
        st.plotly_chart(
            _style_chart(figure, 350),
            width="stretch",
            config=PLOT_CONFIG,
        )
    _show_issue_table(
        issues,
        {"Variance exceeds tolerance", "Invalid variance value"},
        "Variance and input alerts",
    )
    _show_saved_period_comparison(
        records, start_date, end_date, excluded_branches
    )


def _show_saved_period_comparison(
    current_records: pd.DataFrame,
    start_date: date,
    end_date: date,
    excluded_branches: list[str],
) -> None:
    current_month = f"{start_date.year:04d}-{start_date.month:02d}"
    comparison_files = [
        path for path in history_files(HISTORY_DIRECTORY) if path.stem != current_month
    ]
    st.subheader("Compare with a saved reporting period")
    if not comparison_files:
        st.info(
            "Save another month in Data Management to compare branch sales across "
            "periods."
        )
        return

    comparison_path = st.selectbox(
        "Comparison workbook",
        options=comparison_files,
        format_func=lambda path: path.stem,
        key="comparison_workbook",
    )
    match = HISTORY_PATTERN.fullmatch(comparison_path.name)
    if match is None:
        st.error("The selected history file has an invalid month name.")
        return

    comparison_year = int(match.group(1))
    comparison_month = int(match.group(2))
    month_days = calendar.monthrange(comparison_year, comparison_month)[1]
    comparison_start = date(
        comparison_year, comparison_month, min(start_date.day, month_days)
    )
    comparison_end = date(
        comparison_year, comparison_month, min(end_date.day, month_days)
    )
    try:
        comparison_parsed = parse_workbook(comparison_path.read_bytes())
        comparison_records, _, _ = analyze_period(
            comparison_parsed,
            comparison_start,
            comparison_end,
            variance_tolerance=0.0,
            confirm_future_dates=True,
            excluded_branches=set(excluded_branches),
        )
    except (OSError, ValueError) as error:
        st.error(f"Could not analyze the selected comparison workbook: {error}")
        return

    def total_by_branch(frame: pd.DataFrame, output_name: str) -> pd.DataFrame:
        if frame.empty:
            return pd.DataFrame(columns=["Branch key", output_name])
        totals = frame.groupby("Branch", as_index=False)["Sales"].sum()
        totals["Branch key"] = totals["Branch"].str.strip().str.casefold()
        return totals.rename(columns={"Sales": output_name}).drop(columns="Branch")

    current_totals = total_by_branch(current_records, "Current period sales")
    prior_totals = total_by_branch(
        comparison_records, "Comparison period sales"
    )
    comparison = current_totals.merge(prior_totals, on="Branch key", how="outer")
    comparison = comparison.fillna(
        {"Current period sales": 0.0, "Comparison period sales": 0.0}
    )
    names = current_records.groupby("Branch")["Branch"].first()
    prior_names = comparison_records.groupby("Branch")["Branch"].first()
    display_names = {
        branch.strip().casefold(): branch
        for branch in list(names.index) + list(prior_names.index)
    }
    comparison["Branch"] = comparison["Branch key"].map(display_names).fillna(
        comparison["Branch key"]
    )
    comparison["Change"] = (
        comparison["Current period sales"] - comparison["Comparison period sales"]
    )
    comparison["Change %"] = comparison.apply(
        lambda row: (
            row["Change"] / row["Comparison period sales"] * 100
            if row["Comparison period sales"] != 0
            else None
        ),
        axis=1,
    )
    comparison = comparison.sort_values(
        "Current period sales", ascending=False
    ).reset_index(drop=True)
    current_total = float(comparison["Current period sales"].sum())
    prior_total = float(comparison["Comparison period sales"].sum())
    change_total = current_total - prior_total
    metrics = st.columns(4)
    metrics[0].metric("Selected-period sales", f"{current_total:,.2f}")
    metrics[1].metric(
        f"{calendar.month_abbr[comparison_month]} comparison sales",
        f"{prior_total:,.2f}",
    )
    metrics[2].metric("Sales change", f"{change_total:+,.2f}")
    metrics[3].metric("Branches compared", f"{len(comparison):,}")
    st.caption(
        f"Comparing {start_date.strftime('%d %b')}–{end_date.strftime('%d %b %Y')} "
        f"with {comparison_start.strftime('%d %b')}–"
        f"{comparison_end.strftime('%d %b %Y')}. Periods use the same day numbers; "
        "the comparison period is shortened when that month has fewer days."
    )

    chart_data = comparison.head(20).melt(
        id_vars="Branch",
        value_vars=["Current period sales", "Comparison period sales"],
        var_name="Period",
        value_name="Sales",
    )
    figure = px.bar(
        chart_data,
        x="Branch",
        y="Sales",
        color="Period",
        barmode="group",
        color_discrete_sequence=[FOREST_GREEN, WARM_GOLD],
    )
    figure.update_layout(xaxis_title="", yaxis_title="Branch sales")
    st.plotly_chart(
        _style_chart(figure, 430),
        width="stretch",
        config=PLOT_CONFIG,
    )
    st.dataframe(
        comparison[
            [
                "Branch",
                "Current period sales",
                "Comparison period sales",
                "Change",
                "Change %",
            ]
        ],
        hide_index=True,
        width="stretch",
    )


def _show_data_management(
    parsed: ParsedWorkbook,
    workbook_name: str,
    workbook_bytes: bytes,
    selected_year: int,
    selected_month: int,
    excluded_branches: list[str],
    start_date: date,
    end_date: date,
) -> None:
    st.subheader("Workbook management")
    st.write(
        "Import an `.xlsx` workbook using the sidebar, review the detected branch "
        "roster and daily worksheets, then explicitly save it to monthly history."
    )
    columns = st.columns(4)
    columns[0].metric("Workbook", workbook_name)
    columns[1].metric("Branches detected", f"{len(parsed.branches):,}")
    columns[2].metric("Daily worksheets", f"{len(parsed.available_days):,}")
    columns[3].metric("Excluded branches", f"{len(excluded_branches):,}")
    st.caption(
        f"Loaded at {st.session_state.get('workbook_loaded_at', 'this session')} · "
        f"{len(workbook_bytes) / 1024:,.1f} KB · Excel workbook (.xlsx)"
    )
    st.caption(
        "The consolidated roster is used where available; extra branches found on "
        "daily worksheets are retained and noted."
    )

    target_history = month_path(HISTORY_DIRECTORY, selected_year, selected_month)
    existing_differs = (
        target_history.exists()
        and file_digest(target_history.read_bytes()) != file_digest(workbook_bytes)
    )
    replace_existing = False
    if existing_differs:
        st.warning(
            f"A different workbook is already saved for {selected_year:04d}-"
            f"{selected_month:02d}. Choose whether to replace it."
        )
        replace_existing = st.radio(
            "Monthly history action",
            ["Keep existing workbook", "Replace with this workbook"],
            index=0,
        ) == "Replace with this workbook"
    if target_history.exists() and not existing_differs:
        st.success("This exact workbook is already saved for this month.")
    if st.button(
        "Save or update monthly history",
        type="primary",
        disabled=target_history.exists() and not existing_differs,
    ):
        try:
            save_result = save_monthly_workbook(
                workbook_bytes,
                selected_year,
                selected_month,
                HISTORY_DIRECTORY,
                replace_existing=replace_existing,
            )
        except OSError as error:
            st.error(f"Could not save this month's workbook: {error}")
        else:
            if save_result == "saved":
                st.success("Monthly workbook history saved.")
                st.rerun()
            elif save_result == "duplicate":
                st.info("This identical workbook is already in history.")
            else:
                st.info("The existing workbook was kept; no update was saved.")

    with st.expander("Workbook validation and import notes", expanded=True):
        expected_days = set(range(start_date.day, end_date.day + 1))
        missing_days = sorted(expected_days - parsed.available_days)
        unusable_days = sorted(expected_days - set(parsed.daily_entries))
        st.write(
            f"Worksheets found for this period: "
            f"{len(expected_days) - len(missing_days)} / {len(expected_days)}"
        )
        if missing_days:
            st.error(
                "Missing daily worksheet(s): "
                + ", ".join(f"Day {day}" for day in missing_days)
            )
        elif unusable_days:
            st.error(
                "Daily worksheet(s) are present but do not have usable branch-sales "
                "headers: "
                + ", ".join(f"Day {day}" for day in unusable_days)
            )
        else:
            st.success("All daily worksheets for the selected period are available.")
        if not parsed.has_variance_column:
            st.warning("No daily Variances column was detected.")
        if parsed.warnings:
            for warning in parsed.warnings:
                st.warning(warning)
        else:
            st.success("No workbook structure notes.")

    roster = pd.DataFrame(
        [
            {
                "Branch code": branch.code or "",
                "Branch": branch.name,
                "Sunday reporting": (
                    "Yes" if is_sunday_operator(branch.name) else "No"
                ),
                "Analysis status": (
                    "Excluded" if branch.name in excluded_branches else "Included"
                ),
            }
            for branch in parsed.branches
        ]
    )
    st.subheader("Detected branch roster")
    st.dataframe(roster, hide_index=True, width="stretch")
    history = history_files(HISTORY_DIRECTORY)
    st.subheader("Saved monthly history")
    if history:
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Month": path.stem,
                        "File size (KB)": round(path.stat().st_size / 1024, 1),
                    }
                    for path in history
                ]
            ),
            hide_index=True,
            width="stretch",
        )
    else:
        st.info("No monthly workbooks have been saved yet.")


def _show_report_centre(
    records: pd.DataFrame,
    issues: pd.DataFrame,
    coverage: pd.DataFrame,
    result: dict[str, object],
) -> None:
    missing = _selected_issues(issues, {"Missing sales", "Invalid sales value"})
    variance_alerts = _selected_issues(issues, {"Variance exceeds tolerance"})
    columns = st.columns(4)
    columns[0].metric("Days in report", f"{coverage['Date'].nunique() if not coverage.empty else 0:,}")
    columns[1].metric("Expected entries", f"{len(records):,}")
    columns[2].metric("Missing sales", f"{len(missing):,}")
    columns[3].metric("Variance alerts", f"{len(variance_alerts):,}")
    st.write(
        "Reports use the selected dates, excluded branch list, and variance "
        "tolerance. Excel workbooks include a summary, daily coverage, sales gaps, "
        "variance alerts, and invalid input values."
    )
    with st.spinner("Preparing high-resolution reports..."):
        excel_report = build_excel_report(
            records,
            issues,
            coverage,
            result["start_date"],
            result["end_date"],
            result["source_name"],
            result["tolerance"],
        )
        pdf_report = build_pdf_report(
            records,
            issues,
            coverage,
            result["start_date"],
            result["end_date"],
            result["source_name"],
            result["tolerance"],
        )
    period = (
        f"{result['start_date'].isoformat()}_to_{result['end_date'].isoformat()}"
    )
    download_columns = st.columns(2)
    download_columns[0].download_button(
        "Download Excel report",
        data=excel_report,
        file_name=f"sales-analysis-{period}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    download_columns[1].download_button(
        "Download PDF report",
        data=pdf_report,
        file_name=f"sala-analysis-{period}.pdf",
        mime="application/pdf",
        width="stretch",
    )
    st.subheader("Report preview")
    st.dataframe(coverage, hide_index=True, width="stretch")
    with st.expander("Preview flagged sales entries"):
        _show_issue_table(
            issues, {"Missing sales", "Invalid sales value"}
        )
    with st.expander("Preview variance alerts"):
        _show_issue_table(issues, {"Variance exceeds tolerance"})


def _navigate_to(page: str) -> None:
    st.session_state["active_page"] = page


def _render_navigation_group(
    title: str, pages: tuple[str, ...], active_page: str
) -> None:
    st.markdown(
        f'<p class="nav-section-label">{title}</p>',
        unsafe_allow_html=True,
    )
    for page in pages:
        key = page.casefold().replace(" ", "_").replace("&", "and")
        st.button(
            page,
            key=f"navigate_{key}",
            icon=PAGE_ICONS[page],
            type="primary" if page == active_page else "secondary",
            width="stretch",
            on_click=_navigate_to,
            args=(page,),
        )


def _render_sidebar_navigation() -> str:
    if "active_page" not in st.session_state:
        st.session_state["active_page"] = PAGES[0]
    allowed_pages = {page for _, pages in NAVIGATION_GROUPS for page in pages}
    if st.session_state["active_page"] not in allowed_pages:
        st.session_state["active_page"] = PAGES[0]
    for title, pages in NAVIGATION_GROUPS:
        _render_navigation_group(
            title, pages, st.session_state["active_page"]
        )
    st.divider()
    return st.session_state["active_page"]


st.set_page_config(
    page_title="Farm & City Centre | Sales Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
if "active_page" not in st.session_state:
    st.session_state["active_page"] = PAGES[0]
if "dark_mode" not in st.session_state:
    st.session_state["dark_mode"] = False
selected_page = st.session_state["active_page"]
st.markdown(
    """
    <style>
    :root {
        --farm-forest: #176B3A;
        --farm-gold: #D9A441;
        --farm-paper: #F5F7F3;
        --farm-charcoal: #26332D;
        --farm-red: #C0392B;
        --farm-blue: #2878B5;
    }
    .stApp { background: var(--farm-paper); color: var(--farm-charcoal); }
    [data-testid="stHeader"] { background: rgba(245, 247, 243, .94); }
    [data-testid="stSidebar"] {
        background: var(--farm-charcoal);
        border-right: 1px solid rgba(217, 164, 65, .28);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
        color: var(--farm-paper);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        margin-bottom: .45rem;
    }
    [data-testid="stSidebar"] [data-testid="stImage"] {
        display: inline-flex; background: white; border-radius: 9px;
        padding: 5px 9px; margin: 4px 0 2px;
    }
    .sidebar-brand {
        font-size: .73rem !important; font-weight: 750 !important;
        letter-spacing: .13em; color: #dce5dc !important;
        margin: 0 0 1.15rem !important;
    }
    .sidebar-brand span { color: var(--farm-gold); margin-left: .4rem; }
    .nav-section-label {
        color: #dce5dc !important; font-size: .75rem !important;
        font-weight: 750 !important; letter-spacing: .12em;
        margin: .95rem 0 .35rem !important;
    }
    [data-testid="stSidebar"] button[kind="secondary"],
    [data-testid="stSidebar"] button[kind="primary"] {
        display: flex; text-align: left; white-space: nowrap;
    }
    [data-testid="stSidebar"] button[kind] > div {
        width: 100%; justify-content: flex-start !important;
    }
    [data-testid="stSidebar"] button[kind] > div > span {
        width: 100%; flex: 1 1 auto; gap: .65rem;
        justify-content: flex-start !important;
    }
    [data-testid="stSidebar"] button[kind] > div > span > span {
        width: 1.25rem; min-width: 1.25rem; flex: 0 0 1.25rem;
    }
    [data-testid="stSidebar"] button[kind] [data-testid="stIconMaterial"] {
        width: 1.25rem; min-width: 1.25rem; flex: 0 0 1.25rem;
        text-align: center;
    }
    [data-testid="stSidebar"] button[kind] [data-testid="stMarkdownContainer"] {
        flex: 1 1 auto; text-align: left;
    }
    [data-testid="stSidebar"] button[kind] p {
        margin: 0; text-align: left; white-space: nowrap;
    }
    [data-testid="stSidebar"] button[kind="secondary"] {
        background: transparent; color: #f5f7f3;
        border: 1px solid transparent; border-radius: 8px;
        justify-content: flex-start; min-height: 2.55rem;
    }
    [data-testid="stSidebar"] button[kind="secondary"]:hover {
        background: rgba(255,255,255,.09); border-color: rgba(217,164,65,.4);
        color: white;
    }
    [data-testid="stSidebar"] button[kind="primary"],
    button[kind="primary"] {
        background: var(--farm-forest); color: white;
        border: 1px solid var(--farm-gold); border-radius: 8px;
        justify-content: flex-start; min-height: 2.55rem;
    }
    [data-testid="stSidebar"] button[kind="primary"]:hover,
    button[kind="primary"]:hover { background: #12552f; color: white; }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] section {
        background: #33423a; border: 1px dashed #829285;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] section * {
        color: var(--farm-paper);
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid*="File"] {
        color: var(--farm-paper) !important;
        background-color: #33423a;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] button {
        background: var(--farm-forest); color: white;
        border-color: var(--farm-gold);
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] button svg {
        color: white; fill: currentColor;
    }
    [data-testid="stMainBlockContainer"] { padding-top: 2rem; }
    /* Streamlit 1.58 keeps the main menu visible in viewer toolbar mode. */
    [data-testid="stToolbar"] button[aria-label="Main menu"],
    [data-testid="stDeployButton"] { display: none !important; }
    [data-testid="stAlert"][kind="info"] {
        background: #eaf1e9; border-left: 4px solid var(--farm-forest);
        color: var(--farm-charcoal);
    }
    [data-testid="stMetric"] {
        background: #fff; border: 1px solid #e0e6dd; padding: 17px 18px;
        border-radius: 12px; min-height: 112px;
        box-shadow: 0 4px 14px rgba(38,51,45,.055);
    }
    [data-testid="stMetricLabel"] { color: #53645a; }
    [data-testid="stMetricValue"] { color: var(--farm-forest); }
    [data-testid="stMetricDelta"] { color: var(--farm-charcoal); }
    .readiness-grid {
        display: grid; grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: .65rem;
    }
    .readiness-metric {
        min-width: 0; min-height: 112px; padding: 14px;
        display: flex; flex-direction: column; justify-content: center;
        gap: .45rem; background: white; border: 1px solid #e0e6dd;
        border-radius: 12px; box-shadow: 0 4px 14px rgba(38,51,45,.055);
    }
    .readiness-label {
        color: #53645a; font-size: .73rem; font-weight: 650;
    }
    .readiness-value {
        color: var(--farm-forest); font-size: clamp(1rem, 1.45vw, 1.35rem);
        line-height: 1.15; font-weight: 750; overflow-wrap: anywhere;
    }
    [data-testid="stDataFrame"], [data-testid="stTable"] {
        border: 1px solid #e0e6dd; border-radius: 10px;
        box-shadow: 0 3px 12px rgba(38,51,45,.04);
    }
    div[data-testid="stPlotlyChart"] {
        background: white; border: 1px solid #e0e6dd;
        border-radius: 12px; padding: 5px;
        box-shadow: 0 4px 14px rgba(38,51,45,.045);
    }
    [data-testid="stAlert"] { border-radius: 9px; }
    .welcome-hero {
        min-height: 188px; padding: 24px 30px;
        display: flex; flex-direction: column; justify-content: center;
        background: white; border: 1px solid #e0e6dd;
        border-left: 5px solid var(--farm-forest);
        border-radius: 12px; box-shadow: 0 5px 18px rgba(38,51,45,.055);
    }
    .welcome-eyebrow {
        color: var(--farm-forest); font-size: .7rem; font-weight: 750;
        letter-spacing: .17em; margin: 0 0 8px;
    }
    .welcome-title {
        font-family: Segoe UI, Arial, sans-serif;
        font-size: clamp(1.8rem, 3.05vw, 2.8rem); font-weight: 760;
        line-height: 1.08; letter-spacing: -.04em; margin: 0;
        color: var(--farm-forest);
    }
    .welcome-title-second { color: var(--farm-charcoal); }
    .welcome-amp { color: var(--farm-gold); }
    .welcome-subtitle { color: #59685f; margin: 9px 0 5px; font-size: .91rem; }
    .kpi-card {
        min-height: 119px; padding: 15px 17px 13px;
        background: white; border: 1px solid #e0e6dd;
        border-left: 4px solid var(--kpi-accent);
        border-radius: 10px; box-shadow: 0 4px 14px rgba(38,51,45,.055);
    }
    .kpi-label {
        margin: 0 0 .45rem !important; color: #647168;
        font-size: .68rem !important; font-weight: 750 !important;
        letter-spacing: .12em;
    }
    .kpi-value {
        margin: 0 !important; color: var(--farm-charcoal);
        font-size: clamp(1.2rem, 1.8vw, 1.75rem) !important;
        line-height: 1.15; font-weight: 750 !important;
        overflow-wrap: anywhere;
    }
    .kpi-detail { margin: .4rem 0 0 !important; color: #758078; font-size: .77rem !important; }
    .welcome-callout {
        display: flex; flex-direction: column; gap: .3rem;
        background: #eaf1e9; border: 1px solid #d5e2d4;
        border-left: 4px solid var(--farm-forest); border-radius: 10px;
        padding: 1rem 1.15rem; color: var(--farm-charcoal);
    }
    .welcome-callout strong { color: var(--farm-forest); }
    .welcome-features {
        display: grid; grid-template-columns: repeat(3,minmax(0,1fr));
        gap: .8rem; margin-top: 1rem;
    }
    .welcome-features > div {
        display: flex; flex-direction: column; gap: .4rem;
        min-height: 124px; background: white;
        border: 1px solid #e0e6dd; border-radius: 10px;
        padding: 1rem; box-shadow: 0 3px 12px rgba(38,51,45,.04);
    }
    .welcome-features strong { color: var(--farm-forest); }
    .welcome-features p { color: #59685f; font-size: .84rem; margin: 0; }
    .feature-mark { color: var(--farm-gold); font-weight: 800; font-size: .75rem; }
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { scroll-behavior: auto !important; }
    }
    @media (max-width: 760px) {
      .welcome-features { grid-template-columns: 1fr; }
      .welcome-hero { padding: 17px; }
    }
    @media (max-width: 900px) {
      .kpi-card { padding: 12px; }
      .kpi-detail { font-size: .7rem !important; }
    }
    @media (max-width: 1100px) {
      .readiness-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
if st.session_state["dark_mode"]:
    st.markdown(
        """
        <style>
        .stApp, [data-testid="stMain"] { background: #151d18; color: #e6eee8; }
        [data-testid="stMainBlockContainer"] h1,
        [data-testid="stMainBlockContainer"] h2,
        [data-testid="stMainBlockContainer"] h3,
        [data-testid="stMainBlockContainer"] p,
        [data-testid="stMainBlockContainer"] label { color: #e6eee8; }
        [data-testid="stMetric"], .kpi-card, .readiness-metric, .welcome-hero,
        .welcome-features > div {
          background: #202b24; border-color: #3a4b3e; color: #e6eee8;
        }
        [data-testid="stMetricLabel"], .kpi-label, .kpi-detail,
        .readiness-label, .welcome-subtitle, .welcome-features p { color: #c4d0c6; }
        [data-testid="stMetricValue"], .kpi-value { color: #e6eee8; }
        .welcome-callout { background: #23352a; border-color: #42604a; color: #e6eee8; }
        [data-testid="stDataFrame"], [data-testid="stTable"],
        div[data-testid="stPlotlyChart"] { border-color: #3a4b3e; }
        div[data-testid="stPlotlyChart"] { background: #202b24; }
        </style>
        """,
        unsafe_allow_html=True,
    )
_welcome_header(selected_page)
st.divider()

HISTORY_DIRECTORY.mkdir(parents=True, exist_ok=True)
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=142)
    st.markdown(
        '<p class="sidebar-brand">FARM &amp; CITY CENTRE <span>IT-19</span></p>',
        unsafe_allow_html=True,
    )
    st.toggle("Dark mode", key="dark_mode")
    st.divider()
    st.markdown(
        '<p class="nav-section-label">WORKBOOK</p>',
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Import a monthly Excel workbook",
        type=["xlsx"],
        key=f"workbook_upload_{st.session_state.get('upload_widget_version', 0)}",
    )
    st.caption("Upload an .xlsx workbook with daily sheets named Day 1, Day 2, etc.")
    if uploaded_file is not None:
        uploaded_bytes = uploaded_file.getvalue()
        if file_digest(uploaded_bytes) != st.session_state.get("workbook_hash"):
            try:
                _set_current_workbook(uploaded_bytes, uploaded_file.name)
            except ValueError as error:
                st.error(str(error))
                for key in (
                    "workbook_bytes",
                    "workbook_hash",
                    "workbook_name",
                    "workbook_loaded_at",
                    "parsed_workbook",
                    "analysis_result",
                ):
                    st.session_state.pop(key, None)
    if "workbook_name" in st.session_state:
        workbook_bytes_sidebar = st.session_state["workbook_bytes"]
        workbook_size = len(workbook_bytes_sidebar) / 1024
        loaded_month = int(st.session_state["period_month"])
        loaded_year = int(st.session_state["period_year"])
        st.caption(
            f"{st.session_state['workbook_name']} · "
            f"{calendar.month_name[loaded_month]} {loaded_year} · "
            f"{len(st.session_state['parsed_workbook'].branches)} branches"
        )
        st.caption(
            f"{workbook_size:,.1f} KB · Loaded "
            f"{st.session_state.get('workbook_loaded_at', 'this session')}"
        )

    history = history_files(HISTORY_DIRECTORY)
    if history:
        history_choice = st.selectbox(
            "Saved monthly workbook",
            options=history,
            format_func=lambda path: path.stem,
        )
        if st.button("Load saved workbook", width="stretch"):
            match = HISTORY_PATTERN.fullmatch(history_choice.name)
            if match is None:
                st.error("The selected history file has an invalid month name.")
            else:
                try:
                    with st.spinner("Loading saved workbook..."):
                        _set_current_workbook(
                            history_choice.read_bytes(), history_choice.name
                        )
                except (OSError, ValueError) as error:
                    st.error(f"Could not load the saved workbook: {error}")
                else:
                    st.session_state["period_year"] = int(match.group(1))
                    st.session_state["period_month"] = int(match.group(2))
                    st.rerun()
    if "workbook_name" in st.session_state or uploaded_file is not None:
        st.button(
            "Clear workbook",
            width="stretch",
            on_click=_clear_workbook,
        )

    st.divider()
    selected_page = _render_sidebar_navigation()

parsed = st.session_state.get("parsed_workbook")
if parsed is None:
    _show_welcome_page()
    if selected_page == "Data Management":
        st.subheader("Getting started")
        st.write(
            "The importer reads the branch sales value from the **Branch** column "
            "and checks **Variances** where present. Only Pomona, Mbare, FCC Msasa, "
            "Vet-Centre, Madokero, and Norton are expected to report on Sundays."
        )
    st.stop()

workbook_bytes = st.session_state["workbook_bytes"]
workbook_name = st.session_state["workbook_name"]
selected_month = st.selectbox(
    "Reporting month",
    options=list(range(1, 13)),
    format_func=lambda month: calendar.month_name[month],
    key="period_month",
)
selected_year = st.number_input(
    "Reporting year",
    min_value=2000,
    max_value=2100,
    step=1,
    key="period_year",
)
month_start = date(int(selected_year), int(selected_month), 1)
month_end = date(
    int(selected_year),
    int(selected_month),
    calendar.monthrange(int(selected_year), int(selected_month))[1],
)
period_label = f"{calendar.month_name[int(selected_month)]} {int(selected_year)}"
daily_sheets_found = len(parsed.daily_entries)
readiness_columns = st.columns([4, 1.35], vertical_alignment="center")
readiness_columns[0].markdown(
    f"""
    <div class="readiness-grid">
      <div class="readiness-metric"><span class="readiness-label">Reporting period</span><strong class="readiness-value">{period_label}</strong></div>
      <div class="readiness-metric"><span class="readiness-label">Branches detected</span><strong class="readiness-value">{len(parsed.branches):,}</strong></div>
      <div class="readiness-metric"><span class="readiness-label">Daily sheets found</span><strong class="readiness-value">{daily_sheets_found:,}</strong></div>
      <div class="readiness-metric"><span class="readiness-label">Variances column</span><strong class="readiness-value">{"Available" if parsed.has_variance_column else "Not found"}</strong></div>
    </div>
    """,
    unsafe_allow_html=True,
)
readiness_columns[1].button(
    "Analyze full month",
    type="primary",
    width="stretch",
    on_click=_request_full_month_analysis,
)
st.caption(
    f"**{workbook_name}** · {period_label} · {len(parsed.branches)} branches. "
    "Set dates and branch exclusions below, then click **Analyze selected dates**."
)

date_columns = st.columns(4)
with date_columns[0]:
    start_date = st.date_input(
        "Analyze from",
        value=month_start,
        min_value=month_start,
        max_value=month_end,
        key=f"start_{selected_year}_{selected_month}",
    )
with date_columns[1]:
    end_date = st.date_input(
        "Analyze through",
        value=month_end,
        min_value=month_start,
        max_value=month_end,
        key=f"end_{selected_year}_{selected_month}",
    )
with date_columns[2]:
    variance_tolerance = st.number_input(
        "Variance tolerance",
        min_value=0.0,
        value=0.10,
        step=0.01,
        format="%.2f",
        key="variance_tolerance",
        help="Absolute differences up to this amount are treated as normal.",
    )
with date_columns[3]:
    branch_names = sorted(
        [branch.name for branch in parsed.branches], key=str.casefold
    )
    excluded_branches = st.multiselect(
        "Exclude branches",
        options=branch_names,
        key="excluded_branches",
        help="Selected branches are omitted from completeness, trend, and report calculations.",
    )

confirm_future_dates = False
if end_date > date.today():
    st.warning(
        "The selected end date is in the future. Days without entries may "
        "not have been reported yet."
    )
    confirm_future_dates = st.checkbox(
        "I confirm that future dates should be included",
        key="confirm_future_dates",
    )
analysis_buttons = st.columns([1, 3])
with analysis_buttons[0]:
    run_selected_dates = st.button(
        "Analyze selected dates",
        type="primary",
        width="stretch",
    )
run_full_month = st.session_state.pop("analyze_full_month_requested", False)
run_analysis = run_selected_dates or run_full_month

config = {
    "start_date": start_date,
    "end_date": end_date,
    "tolerance": float(variance_tolerance),
    "excluded_branches": tuple(sorted(excluded_branches, key=str.casefold)),
    "workbook_hash": file_digest(workbook_bytes),
}
if run_analysis:
    if start_date > end_date:
        st.error("The start date must be on or before the end date.")
    else:
        try:
            with st.spinner("Analyzing sales entries and building dashboard..."):
                records, issues, coverage = analyze_period(
                    parsed,
                    start_date,
                    end_date,
                    variance_tolerance=float(variance_tolerance),
                    confirm_future_dates=confirm_future_dates,
                    excluded_branches=set(excluded_branches),
                )
        except ValueError as error:
            st.error(str(error))
        else:
            st.session_state["analysis_result"] = {
                "records": records,
                "issues": issues,
                "coverage": coverage,
                "start_date": start_date,
                "end_date": end_date,
                "tolerance": float(variance_tolerance),
                "source_name": workbook_name,
                "analyzed_at": datetime.now().astimezone().strftime(
                    "%d %b %Y, %H:%M %Z"
                ),
                "config": config,
            }

result = st.session_state.get("analysis_result")
if result is not None:
    st.caption(f"Last analysed: {result.get('analyzed_at', 'Not recorded')}")
if selected_page == "Data Management":
    _show_data_management(
        parsed,
        workbook_name,
        workbook_bytes,
        int(selected_year),
        int(selected_month),
        excluded_branches,
        start_date,
        end_date,
    )
    if parsed.has_variance_column:
        st.caption("Variance values are available for analysis.")
    else:
        st.warning(
            "No daily Variances column was found. Missing sales can still be checked."
        )
    st.stop()

if result is None:
    st.markdown(
        """
        <div class="welcome-callout">
          <strong>Workbook ready for analysis</strong>
          <span>Set dates and branch exclusions above, then click Analyze selected dates or use Analyze full month.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()
if result.get("config") != config:
    st.warning(
        "Analysis options have changed. Select **Analyze selected dates** to refresh "
        "this view and its report."
    )
    st.stop()

records = result["records"]
issues = result["issues"]
coverage = result["coverage"]
st.caption(
    f"Workbook: **{workbook_name}** · "
    f"{result['start_date'].strftime('%d %b %Y')} – "
    f"{result['end_date'].strftime('%d %b %Y')} · "
    f"{len(parsed.branches) - len(excluded_branches)} branches included"
)

if not parsed.has_variance_column and selected_page != "Missing Sales Detector":
    st.warning("No daily Variances column was found; variance analysis is unavailable.")
if selected_page == "Executive Dashboard":
    _show_executive_dashboard(records, issues, coverage)
elif selected_page == "Missing Sales Detector":
    _show_missing_detector(records, issues, coverage)
elif selected_page == "Variance & Trends":
    _show_variance_analysis(
        records,
        issues,
        result["tolerance"],
        result["start_date"],
        result["end_date"],
        excluded_branches,
    )
elif selected_page == "Report Centre":
    _show_report_centre(records, issues, coverage, result)
