# SAP Analyst Excel Framework

A professional, reusable Excel framework for SAP reporting automation — built on standard Excel VBA with an optional Python layer. No Excel add-in, no admin rights required.

Excel is the user interface. VBA is the orchestration engine. Python is an optional power layer.

---

## Why This Framework?

SAP analysts spend significant time manually:
- Opening transactions, selecting variants and layouts
- Setting date ranges and exporting ALV grids
- Copying data into Excel and calculating KPIs by hand
- Reformatting dashboards for each report run

This framework automates the entire pipeline — from SAP GUI navigation to KPI calculation and dashboard refresh. It runs on locked-down corporate machines.

---

## Architecture

```
Excel (.xlsm)
  └── VBA Modules (modConfig, modMain, modSAP*, modReport*, modUI*, ...)
        ├── SAP GUI Scripting API  →  automates SAP transactions
        ├── Exports/ folder        →  temporary ALV export files
        └── Python (optional)      →  advanced cleaning & KPI logic
```

Full architecture: [docs/architecture.md](docs/architecture.md)

---

## Folder Structure

```
sap_analyst_framework/
├── README.md
├── vba/                         ← VBA source files (import into workbook)
│   ├── modConfig.bas            ← central constants — import FIRST
│   ├── modMain.bas              ← 17-step pipeline controller
│   ├── modSettings.bas          ← SETTINGS sheet + environment profiles
│   ├── modReportEngine.bas      ← REPORT_CONFIG driven KPI runner
│   ├── modFormulaEngine.bas     ← SUM/COUNT/AVERAGE/MIN/MAX/FORMULA
│   ├── modDates.bas             ← date mode resolution
│   ├── modLog.bas               ← structured logging
│   ├── modFolders.bas           ← folder management
│   ├── modDevSync.bas           ← VS Code export/import
│   ├── modSAPConnection.bas     ← SAP GUI session
│   ├── modSAPNavigation.bas     ← TCode / variant / layout / dates
│   ├── modSAPExport.bas         ← ALV export to Excel
│   ├── modImportData.bas        ← import export file into DATA sheet
│   ├── modUI.bas                ← MANUAL / AUTO UI dispatch
│   ├── modUIState.bas           ← loading bar, status, error/success
│   ├── modComponents.bas        ← button, KPI card, progress bar
│   └── modTheme.bas             ← Gold/Charcoal colour palette
├── python/                      ← optional Python scripts
│   ├── main.py                  ← pipeline entry point (called by VBA)
│   ├── excel_bridge.py          ← openpyxl read/write (.xlsm safe)
│   ├── cleaner.py               ← SAP export cleaning (dates, numbers)
│   ├── quality_check.py         ← null checks, duplicates, outliers
│   └── kpi_engine.py            ← REPORT_CONFIG driven KPI calculation
├── mcp_server/                  ← AI-callable tools (any MCP-compatible LLM)
│   ├── server.py                ← FastMCP server — 9 tools
│   ├── workbook.py              ← workbook manipulation helpers
│   └── requirements.txt         ← mcp, openpyxl
├── docs/
│   ├── architecture.md
│   ├── setup_without_addin.md
│   ├── dev_sync.md
│   ├── mcp_server.md
│   ├── ui_components.md
│   └── user_guide.md
└── examples/
    ├── mpu_report/
    ├── materials_withdrawn/
    └── generic_report/
```

---

## Getting Started

### 1. Prerequisites

- Microsoft Excel 2016+ with macros enabled
- SAP GUI for Windows with scripting enabled ([setup guide](docs/setup_without_addin.md))
- Python 3.9+ — only if `UsePython = YES`

### 2. Prepare the Workbook

Create a `.xlsm` workbook with these sheets:

| Sheet | Purpose |
|---|---|
| HOME | Dashboard — buttons, KPI cards, status |
| SETTINGS | Key-value configuration |
| REPORT_CONFIG | KPI definitions (columns A–O) |
| UI_CONFIG | UI component layout (AUTO mode) |
| DATA | Imported SAP export data |
| LOG | Pipeline run log |
| HISTORY | One row per completed run (incl. duration + status) |

### 3. Import VBA Modules

In the VBE (Alt+F11 → Import File), import **`modConfig.bas` first**, then all others:

```vba
modDevSync.ImportVBAModules   ' or import manually via File → Import
```

### 4. Configure SETTINGS

| Key | Example | Notes |
|---|---|---|
| TCode | MB51 | SAP transaction |
| Variant | /MY_VARIANT | Selection screen variant |
| Plant | 1000 | SAP plant code |
| DateMode | PREVIOUS_WEEK | TODAY / PREVIOUS_WEEK / MONTH_TO_DATE / YEAR_TO_DATE / CUSTOM |
| UsePython | NO | YES to enable Python cleaning layer |
| UI_MODE | MANUAL | MANUAL or AUTO |
| LOADER_TYPE | BAR | BAR or CIRCLE |
| Environment | DEV | Optional — activates key overrides (e.g. TCode.DEV) |

**Environment profiles** — add `Environment=DEV` then set `TCode.DEV=MB51_TEST` to override per environment without changing the base config.

### 5. Define KPIs in REPORT_CONFIG

One row per KPI. Core columns (A–H) + enterprise extensions (I–O):

| Col | Name | Example | Notes |
|---|---|---|---|
| A | KPI_Name | GI Count | Unique identifier |
| B | Formula_Type | COUNT | SUM / COUNT / COUNT_UNIQUE / AVERAGE / MIN / MAX / FORMULA / RATIO / CUSTOM |
| C | Column_Name | Quantity | Header in DATA sheet (case-insensitive) |
| D | Target_Sheet | HOME | Sheet to write result to |
| E | Target_Cell | E10 | Cell address |
| F | Filter_Column | Movement Type | Single-column filter (legacy) |
| G | Filter_Value | 261 | Value to match |
| H | Label | GI Movements | Display label on KPI card |
| I | Filters | MovType=261;Plant=1000 | Multi-condition AND filter (overrides F/G) |
| J | Threshold_Green | 1000 | value ≥ this → green font |
| K | Threshold_Orange | 500 | value ≥ this → orange font |
| L | Threshold_Red | 0 | set to activate red font below orange |
| M | Format | #,##0 | Excel NumberFormat string |
| N | DependsOn | Total Qty | Denominator KPI name (RATIO type only) |
| O | DataSheet | DATA | Source sheet (default: DATA) |

**RATIO example** — GI Rate % = GI Count / Total Qty × 100:

| KPI_Name | Formula_Type | Column_Name | DependsOn | Format |
|---|---|---|---|---|
| GI Count | COUNT | Quantity | | #,##0 |
| Total Qty | SUM | Quantity | | #,##0 |
| GI Rate % | RATIO | GI Count | Total Qty | 0.0% |

### 6. Run the Report

Click the Run button on HOME. The 17-step pipeline:
1. Validates SETTINGS
2. Resolves date range
3. Connects to SAP GUI
4. Opens TCode → applies variant / layout / dates
5. Executes report → exports ALV grid
6. Imports data → DATA sheet
7. (Optional) Python cleaning and quality checks
8. Calculates all KPIs from REPORT_CONFIG
9. Writes results → HOME, applies threshold colours and number formats
10. Appends run record to HISTORY (timestamp, duration, status)

**Batch / unattended mode** — set `modMain.g_BatchMode = True` before calling `RunMain` to suppress all MsgBox dialogs.

---

## MCP Server — Build Dashboards from Any AI Assistant

The `mcp_server/` module exposes the framework as **MCP tools**, compatible with any LLM that supports the Model Context Protocol: Claude, GPT-4, Gemini, Mistral, Cursor, Continue, Cody, and others.

```
You (in any AI chat):
  "Create a KPI dashboard for MB51 withdrawals, plant 1000, previous week,
   with Total Qty, Unique Materials, and GI Count filtered on MovType 261"

AI assistant:
  → build_kpi_dashboard(tcode="MB51", plant="1000", kpis=[...])
  → workbook created on disk with SETTINGS + REPORT_CONFIG + UI_CONFIG
  → open in Excel, import VBA modules, click Run Report
```

### Installation

```bash
# 1. Install dependencies (once)
pip install mcp openpyxl

# or using a virtual environment (recommended)
python -m venv .venv
.venv\Scripts\pip install mcp openpyxl
```

### Connecting to your AI assistant

The server uses **stdio transport** — it is launched automatically by the AI client, never manually.

#### Claude Code (VS Code extension or CLI)

Two files are required at the project root:

**`.mcp.json`** — declares the server:
```json
{
  "mcpServers": {
    "sap-analyst": {
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\sap_analyst_framework\\mcp_server\\server.py"]
    }
  }
}
```

**`~/.claude/settings.json`** — approves project servers automatically:
```json
{
  "enableAllProjectMcpServers": true
}
```

Then restart VS Code. Type `/mcp` to confirm `sap-analyst → connected`.

#### Cursor

Add to **Cursor Settings → MCP → Add Server**:
```json
{
  "sap-analyst": {
    "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
    "args": ["C:\\path\\to\\sap_analyst_framework\\mcp_server\\server.py"]
  }
}
```

Or edit `~/.cursor/mcp.json` directly with the same structure.

#### Continue (VS Code extension)

Add to `~/.continue/config.json`:
```json
{
  "mcpServers": [
    {
      "name": "sap-analyst",
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\sap_analyst_framework\\mcp_server\\server.py"]
    }
  ]
}
```

#### Any other MCP-compatible client

Point the client at:
- **Command:** path to `python.exe` in your venv
- **Args:** absolute path to `sap_analyst_framework/mcp_server/server.py`
- **Transport:** stdio

### Available tools

| Tool | Description |
|---|---|
| `build_kpi_dashboard` | **One-shot** — creates workbook, writes all settings and KPIs in one call |
| `setup_workbook` | Create workbook with all required sheets |
| `write_settings` | Write key-value pairs to SETTINGS sheet |
| `read_settings` | Read current settings |
| `add_kpi` | Add one KPI row to REPORT_CONFIG |
| `list_kpis` | List all defined KPIs |
| `add_ui_component` | Add component to UI_CONFIG |
| `apply_theme` | Apply Gold/Charcoal theme to sheet tabs and headers |
| `inspect_workbook` | Structural summary: sheets, row counts, settings, KPI list |

### Example prompt (works with any LLM)

> "Using the sap-analyst tools, create a workbook at C:\Reports\mb51_report.xlsx for transaction MB51, plant 1000, previous week. Add these KPIs: Total Quantity (SUM of Quantity), Unique Materials (COUNT_UNIQUE of Material), GI Count (COUNT of Quantity filtered on Movement Type = 261). Set threshold green at 500, orange at 200. Format quantities as #,##0."

---

## Key Design Decisions

**Configuration-driven KPIs.**
KPI logic lives in REPORT_CONFIG. Adding a new KPI = adding a row. No code changes.

**No Excel add-in.**
Uses only built-in Excel VBA and standard COM automation. Works on locked-down corporate machines.

**Python is opt-in.**
`UsePython = NO` for pure VBA. `UsePython = YES` to shell out to Python for advanced processing.

**Environment profiles.**
Add `Environment = DEV` and `TCode.DEV = MB51_TEST` to override settings per environment without editing base config.

**Batch-safe.**
Set `modMain.g_BatchMode = True` for unattended / scheduled execution — no dialog boxes.

**VS Code as VBA editor.**
`modDevSync` exports all VBA as `.bas` files for Git version control and VS Code editing.

---

## Examples

| Example | TCode | Date Mode | Python | Location |
|---|---|---|---|---|
| MPU Report | MCBA | PREVIOUS_WEEK | No | [examples/mpu_report/](examples/mpu_report/) |
| Materials Withdrawn | MB51 | PREVIOUS_WEEK | Yes | [examples/materials_withdrawn/](examples/materials_withdrawn/) |
| Generic Report | MB52 | CUSTOM | No | [examples/generic_report/](examples/generic_report/) |

Each ships with `SETTINGS.csv` and `REPORT_CONFIG.csv` ready to paste into Excel.

---

## Documentation

| Document | Topic |
|---|---|
| [Architecture](docs/architecture.md) | System design, data flow, module map |
| [Setup Without Add-In](docs/setup_without_addin.md) | SAP scripting, Python, macro security |
| [Dev Sync](docs/dev_sync.md) | Export/import VBA, Git, VS Code |
| [MCP Server](docs/mcp_server.md) | AI integration, tool reference |
| [UI Components](docs/ui_components.md) | Button, KPI card, progress bar reference |
| [User Guide](docs/user_guide.md) | Create a new report from scratch |

---

## License

Framework code: MIT.
xlwings (in `/xlwings/`): BSD 3-Clause — not modified by this framework.
