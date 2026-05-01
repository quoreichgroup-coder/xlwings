# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SAP Analyst Excel Framework — a configuration-driven pipeline that automates SAP GUI navigation → ALV export → data cleaning → KPI calculation → dashboard refresh, all within a standard `.xlsm` workbook. No add-ins, no admin rights required.

The framework has four layers:
1. **Excel workbook** (.xlsm) — 7 sheets, macros enabled
2. **VBA modules** (`sap_analyst_framework/vba/`) — pipeline orchestration and SAP GUI scripting
3. **Python scripts** (`sap_analyst_framework/python/`) — optional advanced data processing
4. **MCP server** (`sap_analyst_framework/mcp_server/`) — AI-callable tools registered in `.claude/settings.json`

## Development Commands

### Python pipeline (manual run)
```bash
python sap_analyst_framework/python/main.py <workbook_path> <export_file_path>
```

### MCP server
```bash
python sap_analyst_framework/mcp_server/server.py
```
The server is pre-registered in `.claude/settings.json`; tools are available immediately inside Claude Code.

### Python dependencies
```bash
pip install pandas openpyxl mcp
# or for MCP only:
pip install -r sap_analyst_framework/mcp_server/requirements.txt
```

### VBA development (DevSync workflow)
Edit `.bas` files in VS Code, then reimport via Excel VBE:
```vba
' Export VBA → .bas files (run inside Excel):
modDevSync.ExportVBAModules
' Reimport after editing:
modDevSync.ImportVBAModules
```

## VBA module import order

When importing `.bas` files into a new workbook via VBE (Alt+F11 → Import File), import **`modConfig.bas` first** — all other modules reference its Public constants. Order of remaining modules does not matter.

## Architecture

### Pipeline (17 steps in `modMain.RunMain`)
```
HOME button click
  → 1–5:  Init UI, validate SETTINGS, prepare folders, resolve dates
  → 6–11: Connect SAP, open TCode, apply variant/layout/dates, execute, export ALV
  → 12:   Import exported file → DATA sheet
  → 13:   (optional) Shell out to Python pipeline
  → 14:   Evaluate REPORT_CONFIG KPI rows → write results to HOME
  → 15–17: History, UI refresh, done
```

### Configuration sheets (no code changes needed to add KPIs)
- **SETTINGS** — key-value store (TCode, Plant, DateMode, SAPConnectionName, UsePython, PythonPath, …). Add `Environment=DEV|UAT|PROD` to activate per-environment key overrides (e.g. `TCode.DEV`).
- **REPORT_CONFIG** — one row per KPI; columns A–H (core) + I–O (enterprise extensions):
  - `I: Filters` — multi-condition filter `"MovType=261;Plant=1000"` (supersedes F/G)
  - `J/K/L: Threshold_Green/Orange/Red` — numeric thresholds for font colouring
  - `M: Format` — Excel number format string (e.g. `#,##0`, `0.0%`)
  - `N: DependsOn` — denominator KPI name when Formula_Type=`RATIO`
  - `O: DataSheet` — source sheet name (defaults to DATA; enables multi-source reports)
- **HISTORY** — now records columns G (Duration sec) and H (Status: SUCCESS / ERROR: …)
- **UI_CONFIG** — component layout for AUTO render mode

### VBA module responsibilities
| Module | Role |
|---|---|
| `modConfig` | **Central constants** — all sheet names, HISTORY column indices, pipeline defaults. Must be imported first. |
| `modMain` | 17-step pipeline. `g_BatchMode As Boolean` suppresses MsgBox for unattended runs. |
| `modSettings` | SETTINGS sheet read/write with **environment profiles**: set `Environment=DEV`, then `TCode.DEV=MB51` to override per environment. |
| `modReportEngine` | Parses REPORT_CONFIG (columns A–O), two-pass KPI calculation (RATIO handled in pass 2), threshold colouring, number format. |
| `modFormulaEngine` | Evaluates SUM/COUNT/COUNT_UNIQUE/AVERAGE/MIN/MAX/FORMULA/CUSTOM. Multi-filter via `filtersStr="col=val;col2=val2"`. DataSheet param for multi-source. |
| `modSAPNavigation` | Opens TCode, applies variant/layout/date fields |
| `modSAPExport` | Executes report, exports ALV to file |
| `modImportData` | Reads export file → DATA sheet |
| `modUIState` | Loading bar, progress steps, status/error/success messages. Respects `modMain.g_BatchMode` (no MsgBox in batch). |
| `modComponents` | Button, KPI card, progress bar, alert factories |
| `modUI` | Dispatches MANUAL vs AUTO UI rendering |
| `modDates` | Resolves DateMode (TODAY / PREVIOUS_WEEK / CUSTOM / …) to DateFrom/DateTo |
| `modLog` | Structured logging to LOG sheet + text file |
| `modDevSync` | Exports/imports .bas files for VS Code editing |
| `modTheme` | Gold/Charcoal colour palette constants |
| `modFolders` | Creates/manages Exports/, Archive/, Logs/ |
| `modSAPConnection` | Acquires and caches SAP GUI session |

### Python module responsibilities
| Module | Role |
|---|---|
| `main.py` | Orchestrates: load → clean → validate → calculate → write back |
| `excel_bridge.py` | openpyxl wrapper with `keep_vba=True` (never strips macros on write) |
| `cleaner.py` | 7-step SAP export normalisation (numeric format, dates, NaN, duplicates) |
| `quality_check.py` | Null rate, duplicate, outlier, date-range checks |
| `kpi_engine.py` | Python mirror of modFormulaEngine; writes results back via excel_bridge |

### MCP server tools (9 tools in `mcp_server/server.py`)
`setup_workbook`, `write_settings`, `read_settings`, `add_kpi`, `list_kpis`, `add_ui_component`, `build_kpi_dashboard`, `apply_theme`, `inspect_workbook`

## Key Constraints

- **Never modify `/xlwings/`** — it is an unmodified library fork; framework code sits on top of it.
- **Always use `keep_vba=True`** when opening `.xlsm` files with openpyxl, or all macros are silently stripped.
- **Python is optional** — `UsePython=YES/NO` in SETTINGS controls whether VBA shells out. The VBA pipeline is complete without Python.
- **No hard-coded paths** — all paths (Python executable, export folder, SAP connection name) come from SETTINGS sheet.
- **KPI logic lives in REPORT_CONFIG**, not in VBA/Python code. Adding a KPI = adding a row.

## Examples

Three reference implementations in `sap_analyst_framework/examples/`:
- `mpu_report/` — MCBA transaction, pure VBA
- `materials_withdrawn/` — MB51 with Python pipeline enabled
- `generic_report/` — blank template for any TCode

Each example ships with `SETTINGS.csv` and `REPORT_CONFIG.csv` showing the expected sheet structure.

## Docs

`sap_analyst_framework/docs/` contains:
- `architecture.md` — full data flow and design decisions
- `setup_without_addin.md` — SAP GUI scripting setup, Python install, security
- `dev_sync.md` — VS Code ↔ VBE round-trip workflow
- `mcp_server.md` — MCP integration and tool reference
- `ui_components.md` — Component factory reference
- `user_guide.md` — Step-by-step report creation walkthrough
