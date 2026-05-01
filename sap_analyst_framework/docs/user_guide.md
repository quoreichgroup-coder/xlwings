# User Guide — Creating a New SAP Report

This guide explains how a SAP analyst can create a new automated report by configuring three sheets — without writing any VBA or Python code.

## Overview

A "report" in this framework is defined entirely by:
1. **SETTINGS** — what SAP transaction to run, with which parameters.
2. **REPORT_CONFIG** — what KPIs to calculate from the exported data.
3. **UI_CONFIG** (optional) — where to display those KPIs on the HOME sheet.

---

## Step 1 — Configure SETTINGS

Open the SETTINGS sheet and fill in the values for your report.

**Example: Materials Withdrawn report**

| Key | Value |
|---|---|
| ProjectName | SAP Plant Analytics |
| ReportName | Materials Withdrawn |
| Plant | 1000 |
| TCode | MB51 |
| Variant | /MATS_WITHDRAWN |
| Layout | /STD_LAYOUT |
| DateMode | PREVIOUS_WEEK |
| DateFrom | (auto-filled by the framework) |
| DateTo | (auto-filled by the framework) |
| ExportFileName | MatWithdrawn_Export |
| UI_MODE | MANUAL |
| LOADER_TYPE | BAR |
| UsePython | NO |
| PythonPath | |
| SAPConnectionName | |

**DateMode options:**

| Value | Meaning |
|---|---|
| `CURRENT_WEEK` | Monday to today of this week |
| `PREVIOUS_WEEK` | Last Monday to last Sunday |
| `MONTH_TO_DATE` | 1st of this month to today |
| `YEAR_TO_DATE` | January 1st to today |
| `CUSTOM` | Use the DateFrom / DateTo values you enter manually |

---

## Step 2 — Define KPIs in REPORT_CONFIG

Each row in REPORT_CONFIG defines one calculated KPI.

**Column reference:**

| Column | Field | Description |
|---|---|---|
| A | KPI_Name | Descriptive name (for logging only) |
| B | Formula_Type | SUM, COUNT, COUNT_UNIQUE, AVERAGE, MIN, MAX, FORMULA, CUSTOM |
| C | Column_Name | Header name in the DATA sheet (exact match, case-insensitive) |
| D | Target_Sheet | Sheet where the result is written (e.g. HOME) |
| E | Target_Cell | Cell address (e.g. E10) |
| F | Filter_Column | Optional: column to filter by before calculating |
| G | Filter_Value | Optional: value to match in Filter_Column |
| H | Label | Human-readable label (used in AUTO mode KPI cards) |

**Example: Materials Withdrawn KPIs**

| KPI_Name | Formula_Type | Column_Name | Target_Sheet | Target_Cell | Filter_Column | Filter_Value | Label |
|---|---|---|---|---|---|---|---|
| Total Quantity | SUM | Quantity | HOME | E10 | | | Total Qty |
| Unique Materials | COUNT_UNIQUE | Material | HOME | E11 | | | Materials |
| Total Qty Plant 1000 | SUM | Quantity | HOME | E12 | Plant | 1000 | Qty P1000 |
| Movement Count | COUNT | Movement Type | HOME | E13 | | | Movements |
| Avg Quantity | AVERAGE | Quantity | HOME | E14 | | | Avg Qty |

---

## Step 3 — Design or Configure the HOME Sheet

### Option A — MANUAL mode (recommended for analysts)

1. Set `UI_MODE = MANUAL` in SETTINGS.
2. Design the HOME sheet however you like in Excel.
3. Put formulas or direct cell references in the cells listed in REPORT_CONFIG → Target_Cell.
4. The framework will write KPI values to those cells after each run.

### Option B — AUTO mode

1. Set `UI_MODE = AUTO` in SETTINGS.
2. Fill in the UI_CONFIG sheet (see [ui_components.md](ui_components.md)).
3. The framework will build KPI cards, buttons, and headers automatically.

---

## Step 4 — Run the Report

1. Open the workbook.
2. Go to the HOME sheet.
3. Click **Run Report**.
4. Watch the progress bar and status message.
5. When complete, the KPI values on HOME are updated.

---

## Step 5 — Check the LOG Sheet

After each run, the LOG sheet shows every pipeline step with its status (OK / ERROR) and a message.

If something went wrong, find the first ERROR row and read the message. Common issues and fixes:

| Error | Fix |
|---|---|
| "SAP connection failed" | Log in to SAP, then re-run |
| "Export file not found" | Check the SAP window — export dialog may have stalled |
| "Column not found: Quantity" | Check that the column header in DATA exactly matches Column_Name in REPORT_CONFIG |
| "Settings validation failed" | Check SETTINGS for missing or invalid values |

---

## Switching Between Reports

To run a different report (e.g. switch from Materials Withdrawn to MPU):

1. Update the SETTINGS sheet (TCode, Variant, Layout, ReportName, etc.).
2. Update REPORT_CONFIG with the new KPI definitions.
3. Update HOME cells to match the new Target_Cell values.
4. Click Run Report.

No VBA changes are needed. All logic is driven by configuration.

---

## Using Python for Advanced Analysis

If your report needs data cleaning or complex KPIs beyond what the formula engine supports:

1. Set `UsePython = YES` in SETTINGS.
2. Set `PythonPath` to your Python executable (or leave blank if `python` is on PATH).
3. Edit `python/kpi_engine.py` to add custom calculation logic.
4. The Python results are written back to the workbook before the VBA engine runs.

---

## Reviewing Historical Runs

The HISTORY sheet records every completed run:

| Column | Content |
|---|---|
| A | Run timestamp |
| B | Report name |
| C | Plant |
| D | Date From |
| E | Date To |
| F | Windows username |

Use this for audit trails, trend analysis, or debugging intermittent issues.
