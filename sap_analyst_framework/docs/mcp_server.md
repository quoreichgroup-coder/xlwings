# MCP Server — Build Excel Dashboards from VS Code

The MCP (Model Context Protocol) server exposes the framework's workbook-building capabilities as AI-callable tools. Once registered, you can type a natural language instruction in VS Code and Claude will call the right tools to build the Excel file automatically.

```
VS Code (Claude)
  → "Create KPI dashboard for MB51 withdrawals"
  → calls build_kpi_dashboard(...)
  → Excel workbook created / updated on disk
  → you open it, import VBA, click Run Report
```

---

## Installation

### 1. Install server dependencies

No admin rights required — user-level pip install:

```bash
pip install --user mcp openpyxl
```

Or in a virtual environment:

```bash
cd sap_analyst_framework/mcp_server
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Register with Claude Code (project-level)

The file `.claude/settings.json` is already created at the root of this project with the correct Python path and user site-packages:

```json
{
  "mcpServers": {
    "sap-analyst": {
      "command": "C:\\Python314\\python.exe",
      "args": ["sap_analyst_framework/mcp_server/server.py"],
      "env": {
        "PYTHONPATH": "C:\\Users\\Abdouramane\\AppData\\Roaming\\Python\\Python314\\site-packages"
      }
    }
  }
}
```

> If you use a different Python installation or virtual environment, update `command` to the full path of that `python.exe` and remove the `PYTHONPATH` env var.

Restart Claude Code (or run `/mcp` in the chat) to load the server.

### 3. Register with VS Code (global, optional)

If you use the Claude extension in VS Code, add to your Claude config:

- Open the Command Palette → **Claude: Open Settings**
- Or edit `~/.claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "sap-analyst": {
      "command": "python",
      "args": ["C:\\Users\\YourName\\Documents\\SAP Studio\\Framework\\sap_analyst_framework\\mcp_server\\server.py"]
    }
  }
}
```

Use the full absolute path when registering globally.

---

## Available Tools

| Tool | Description |
|---|---|
| `setup_workbook` | Create/update workbook with all 7 required sheets |
| `write_settings` | Write SETTINGS key-value pairs from JSON |
| `read_settings` | Read current SETTINGS as JSON |
| `add_kpi` | Add one row to REPORT_CONFIG |
| `list_kpis` | List all current REPORT_CONFIG rows |
| `add_ui_component` | Add one row to UI_CONFIG |
| `build_kpi_dashboard` | **One-shot**: write SETTINGS + REPORT_CONFIG + UI_CONFIG |
| `apply_theme` | Apply Gold/Charcoal theme to all sheet headers and tabs |
| `inspect_workbook` | Return a full structural summary of the workbook |

---

## Usage Examples

### One-shot dashboard

Type in VS Code chat:

> "Create a KPI dashboard for transaction MB51 Materials Withdrawn, plant 1000, previous week, with KPIs: Total Quantity (SUM), Unique Materials (COUNT_UNIQUE), Goods Issues (COUNT filtered Movement Type=261), Reversals (COUNT filtered Movement Type=262), Average Qty (AVERAGE)"

Claude calls `build_kpi_dashboard` with the right parameters and creates/updates the workbook.

---

### Step-by-step

> "Set up the framework workbook at C:\Reports\MPU.xlsx"

→ `setup_workbook(path="C:\\Reports\\MPU.xlsx")`

> "Configure it for transaction MCBA, previous week, plant 1000"

→ `write_settings(path=..., settings_json='{"TCode":"MCBA","DateMode":"PREVIOUS_WEEK","Plant":"1000"}')`

> "Add a KPI: sum of Quantity column, target cell E10"

→ `add_kpi(path=..., kpi_name="Total Qty", formula_type="SUM", column_name="Quantity", target_cell="E10")`

> "What KPIs are currently configured?"

→ `list_kpis(path=...)`

---

## Tool Reference

### `build_kpi_dashboard` — the main tool

```
path         : full path to .xlsx or .xlsm file
report_name  : "Materials Withdrawn"
tcode        : "MB51"
date_mode    : CURRENT_WEEK | PREVIOUS_WEEK | MONTH_TO_DATE | YEAR_TO_DATE | CUSTOM
plant        : "1000"
kpis_json    : JSON array (see below)
variant      : "/MY_VARIANT"  (optional)
layout       : "/MY_LAYOUT"   (optional)
use_python   : false
loader_type  : BAR | CIRCLE
```

`kpis_json` format:

```json
[
  {
    "kpi_name":     "Total Qty",
    "formula_type": "SUM",
    "column_name":  "Quantity",
    "label":        "Total Qty"
  },
  {
    "kpi_name":     "GI Count",
    "formula_type": "COUNT",
    "column_name":  "Movement Type",
    "filter_column": "Movement Type",
    "filter_value":  "261",
    "label":        "GI Count"
  }
]
```

Supported `formula_type` values:

| Type | Behaviour |
|---|---|
| `SUM` | Numeric sum of the column |
| `COUNT` | Row count (with optional filter) |
| `COUNT_UNIQUE` | Distinct value count |
| `AVERAGE` | Mean value |
| `MIN` / `MAX` | Minimum / maximum |
| `FORMULA` | Raw pandas eval expression (Python mode) |
| `CUSTOM` | Placeholder — extend in kpi_engine.py |

---

## After the MCP Server Runs

The MCP server only writes configuration data to the workbook. It does NOT run SAP or import VBA. After calling `build_kpi_dashboard`:

1. Open the workbook in Excel.
2. Import VBA modules (Alt+F11 → File → Import, or run `modDevSync.ImportVBAModules`).
3. Fill in any remaining SETTINGS (Variant, SAPConnectionName, PythonPath).
4. Click **Run Report** on the HOME sheet.

---

## Troubleshooting

**"No module named 'mcp'"**  
Run: `pip install --user mcp openpyxl`

**"Server not found in Claude Code"**  
Check `.claude/settings.json` exists at the project root and the path to `server.py` is correct. Run `/mcp` in Claude Code to reload.

**"File not found: ..."**  
Use the full absolute path in all tool calls. Relative paths may resolve differently depending on how the server is launched.

**Changes not visible in Excel**  
Close the workbook in Excel before the MCP server writes to it. Excel locks the file while it is open.
