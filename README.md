# Sala Analysis Model by IT-19

A local Streamlit app for checking monthly branch sales workbooks. It detects
missing daily sales, highlights variance differences, and creates downloadable
Excel and PDF reports. Its hybrid corporate theme uses Farm & City forest green,
`#D9A441` warm gold, `#F5F7F3` soft white and `#26332D` charcoal. The sidebar
groups navigation into Overview, Sales Intelligence, Data and Reporting.
The interface is organized into five menu views. Upload or load a workbook in
the sidebar; the reporting period, readiness summary, and analysis controls then
appear in the main view. Use **Analyze full month** for a one-click monthly
analysis or adjust dates and exclusions before selecting **Analyze selected
dates**. The Executive Dashboard is the default view.
The app opens without sign-in or account controls. Anyone who can reach the
running app can use it, so run it only in a trusted environment.
The hero heading follows the selected navigation view; use **Back to Executive
Dashboard** to return. A sidebar toggle switches between light and dark display.

- **Executive Dashboard** — total sales, active branches, average sales per
  active branch, distinct branches with missing reports, sales and completeness
  trends, branch ranking, and review queue.
- **Missing Sales Detector** — branch/day heatmap, missing-entry counts, and
  follow-up table.
- **Variance & Trends** — daily variance ranges, distribution, branch
  sales trends, variance alerts, and same-day-range comparisons with another
  saved monthly workbook.
- **Data Management** — workbook import, branch inclusion/exclusion, import
  notes, and one-file-per-month history.
- **Report Centre** — report preview and Excel/PDF downloads.

## Run on Windows

From this folder in PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

The app opens in a browser at the local Streamlit address shown in the terminal.
Uploaded workbooks are analyzed on this computer; they are not sent to an online
service.

## Workbook support

- Upload an `.xlsx` workbook with worksheets named `Day 1`, `Day 2`, etc.
- Each daily worksheet should have a header row containing `Name`, `Branch`
  (the sales amount), and optionally `Variances`; `Code` is used when present.
- The `Consolidated` worksheet supplies the initial branch roster when available;
  branch rows found only on daily worksheets are included and noted. Otherwise,
  the daily worksheets supply the roster.
- Blank sales are reported as missing. A numeric zero is treated as an entered
  sale. Branches are discovered from the workbook rather than assumed to number
  79.
- Only Pomona, Mbare, FCC Msasa, Vet-Centre (including `THE VET-CENTRE`),
  Madokero, and Norton are expected to report on Sundays.
- Select the month, year, and inclusive start/end dates. Future end dates require
  an explicit confirmation before analysis.
- Variances are flagged when their absolute value exceeds the configurable
  tolerance, which defaults to `0.10`. Values at or below the tolerance are
  considered normal.
- When saving history, only one workbook is kept per `YYYY-MM`. Re-uploading an
  identical file does not create a duplicate. Replacing a different workbook
  requires choosing **Update the saved workbook**; the default is to reject it.

## Dashboard and reports

Use **Exclude branches** in the main analysis controls to omit selected
branches from all analysis views and exported reports. Workbook name, period,
size, load time and daily worksheet validation appear in the sidebar and Data
Management view. The sidebar also lets you clear the loaded workbook or load a
saved monthly workbook. The
dashboard uses interactive 2D charts with hover, zoom, and image-export
controls. Plotly chart exports are configured for 3840 × 2160 PNG images.
Workbook loading, analysis, and report preparation show a progress indicator.
There are no animated 3D charts.
Download the formatted Excel report or a paginated PDF report from **Report
Centre**. The Excel report can also be loaded into Power BI. The app does not
create a Power BI `.pbix` file. Streamlit runs in viewer toolbar mode; the local
main menu is hidden for a cleaner end-user view.

## Detection method

This version uses auditable completeness rules and the selected variance
tolerance. It also computes a median-absolute-deviation (MAD) score to give
variance alerts statistical context. It does not claim a trained predictive
model: a labeled history of confirmed errors is not available in the sample
workbook. Confirmed monthly reviews can later be used to evaluate a genuinely
trained anomaly model.

## Tests

```powershell
python -m pytest
```

The monthly history folder is local and excluded from version control.
