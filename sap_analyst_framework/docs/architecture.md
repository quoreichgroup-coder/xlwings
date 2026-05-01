# Architecture — SAP Analyst Excel Framework

## Overview

The framework has four integrated layers that communicate through well-defined interfaces.

```
┌─────────────────────────────────────────────────────────┐
│                  Excel Workbook (.xlsm)                  │
│                                                          │
│  HOME  │  SETTINGS  │  REPORT_CONFIG  │  UI_CONFIG      │
│  DATA  │  LOG        │  HISTORY                          │
└────────────────────┬────────────────────────────────────┘
                     │ reads / writes sheets
┌────────────────────▼────────────────────────────────────┐
│                   VBA Modules                            │
│                                                          │
│  modMain (pipeline)  │  modSettings  │  modDates        │
│  modLog              │  modFolders   │  modDevSync       │
│  modSAPConnection    │  modSAPNav    │  modSAPExport     │
│  modImportData       │  modReportEngine │ modFormulaEngine│
│  modUI  │  modUIState │  modComponents │  modTheme       │
└──────┬────────────────────────────────────┬─────────────┘
       │ SAP GUI Scripting API              │ Shell / subprocess
       ▼                                    ▼
┌──────────────┐                 ┌──────────────────────────┐
│  SAP GUI     │                 │  Python Scripts          │
│  (GuiSession)│                 │                          │
│              │                 │  main.py                 │
│  TCode       │                 │  excel_bridge.py         │
│  Variant     │                 │  cleaner.py              │
│  Layout      │                 │  quality_check.py        │
│  ALV export  │                 │  kpi_engine.py           │
└──────────────┘                 └──────────────────────────┘
```

## Data Flow

```
Button on HOME
    ↓
modMain.RunMain()
    ↓ reads
SETTINGS sheet
    ↓
modDates — resolve DateMode → DateFrom / DateTo
    ↓
modSAPConnection — GetSAPSession()
    ↓
modSAPNavigation — OpenTransaction / ApplyVariant / ApplyLayout / SetDateRange
    ↓
modSAPExport — ExecuteReport / ExportALVToExcel / WaitForExportFile
    ↓
Exports/ folder  (local .xlsx file)
    ↓
modImportData — ImportExportToDataSheet → DATA sheet
    ↓
[Optional] Python pipeline:  cleaner → quality_check → kpi_engine → write_results
    ↓
modReportEngine — ReadReportConfig / CalculateConfiguredKPIs / WriteKPIResults
    ↓
HOME / REPORT_CONFIG target cells updated
    ↓
modUI.RefreshUI → dashboard redraws
    ↓
modLog — full run log in LOG sheet
```

## Module Responsibilities

| Module | Responsibility |
|---|---|
| `modMain` | Orchestrates the full pipeline with error handling |
| `modSettings` | Key-value store backed by SETTINGS sheet |
| `modDates` | DateMode resolution to concrete From/To values |
| `modLog` | Structured logging to LOG sheet + text file |
| `modFolders` | Manages Exports/, Archive/, Logs/ sub-folders |
| `modDevSync` | Export/import VBA modules to .bas files for VS Code |
| `modSAPConnection` | Acquire and cache a SAP GUI session |
| `modSAPNavigation` | Navigate TCode, variant, layout, date fields |
| `modSAPExport` | Execute report and export ALV grid to Excel |
| `modImportData` | Load exported file into DATA sheet |
| `modReportEngine` | Read REPORT_CONFIG, call FormulaEngine, write results |
| `modFormulaEngine` | SUM/COUNT/AVERAGE/MIN/MAX/FORMULA against DATA |
| `modUI` | Dispatch MANUAL vs AUTO UI build |
| `modUIState` | Loading bar/circle, status cell, error/success feedback |
| `modComponents` | Factory for buttons, KPI cards, status boxes, etc. |
| `modTheme` | Colour constants and theme application helpers |

## Sheet Responsibilities

| Sheet | Purpose |
|---|---|
| HOME | Main dashboard — buttons, KPI cards, status area |
| SETTINGS | Key-value configuration for every run parameter |
| REPORT_CONFIG | One row per KPI: formula type, column, filter, target cell |
| UI_CONFIG | One row per UI component (AUTO mode only) |
| DATA | Raw imported SAP data (refreshed each run) |
| LOG | Structured run log (Timestamp, Step, Status, Message) |
| HISTORY | One row per completed run for audit trail |

## No Add-In Requirement

The framework runs entirely from the .xlsm workbook using:
- VBA macros (built-in Excel capability)
- SAP GUI Scripting API (COM automation, no install required if enabled by Basis)
- Python via Shell() call to the user's own Python installation (or portable Python)
- openpyxl / pandas (pip install, no admin required for user-level install)

## Configuration-Driven KPI Design

KPI logic lives in REPORT_CONFIG, not in code. To add a new KPI:
1. Add a row to REPORT_CONFIG with the formula type and column name.
2. Point the target cell to a cell on HOME (or any sheet).
3. Run the pipeline — no VBA or Python changes needed.

This means a SAP analyst can create new reports by modifying Excel sheets, not by writing code.
