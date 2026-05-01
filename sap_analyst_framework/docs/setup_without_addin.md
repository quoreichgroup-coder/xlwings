# Setup Without Excel Add-In

The SAP Analyst Framework is designed to work in locked-down corporate environments where add-in installation is blocked. No add-in, no admin rights, and no Excel COM server installation are required.

## Prerequisites

| Requirement | Notes |
|---|---|
| Microsoft Excel 2016+ (.xlsm support) | Standard Office install |
| SAP GUI for Windows | Must already be installed and running |
| SAP GUI Scripting enabled | Ask your Basis team (no admin required on your PC) |
| Python 3.9+ (optional) | User-level pip install, or portable Python zip |

## Step 1 — Enable SAP GUI Scripting

SAP GUI Scripting must be enabled on both the SAP server and your SAP GUI client.

**SAP GUI client setting (no admin required):**
1. Open SAP GUI (not the workbook).
2. Go to **Options** (Alt+F12) → **Accessibility & Scripting** → **Scripting**.
3. Check **Enable scripting**.
4. Uncheck **Notify when a script attaches to SAP GUI** (optional, reduces popups).

**SAP server setting (Basis team):**  
Transaction `RZ11` → Parameter `sapgui/user_scripting` → set to `TRUE`.

## Step 2 — Enable VBA Macro Access to the VBA Project Object Model

This is required only for the DevSync export/import feature.

1. Open Excel → **File** → **Options** → **Trust Center**.
2. Click **Trust Center Settings** → **Macro Settings**.
3. Check **Trust access to the VBA project object model**.

> This setting allows VBA to read and write its own source code. It is not required for the main reporting pipeline.

## Step 3 — Configure Macro Security

The workbook uses VBA macros. Excel must be configured to allow them.

1. **File** → **Options** → **Trust Center** → **Trust Center Settings**.
2. **Macro Settings** → Select **Disable all macros with notification** (recommended).
3. When you open the workbook, click **Enable Content** in the yellow bar.

Alternatively, add the workbook folder to **Trusted Locations** (File → Options → Trust Center → Trusted Locations) to avoid the prompt every time.

## Step 4 — Install Python Dependencies (Optional)

If `UsePython = YES` in SETTINGS, the framework calls Python via Shell.

**User-level install (no admin):**
```
pip install --user pandas openpyxl
```

**Portable Python (no install at all):**
1. Download WinPython or embeddable Python from python.org.
2. Extract to a folder (e.g. `C:\Users\YourName\tools\python`).
3. Run `pip install pandas openpyxl` inside that folder.
4. Set `PythonPath` in SETTINGS to the full path of `python.exe`.

## Step 5 — Configure SETTINGS Sheet

Open the workbook and fill in the SETTINGS sheet:

| Key | Example Value |
|---|---|
| ProjectName | SAP Analytics |
| ReportName | MPU Report |
| Plant | 1000 |
| TCode | MB52 |
| Variant | /MY_VARIANT |
| Layout | /MY_LAYOUT |
| DateMode | PREVIOUS_WEEK |
| DateFrom | (auto-filled) |
| DateTo | (auto-filled) |
| ExportFileName | MPU_Export |
| UI_MODE | MANUAL |
| LOADER_TYPE | BAR |
| UsePython | NO |
| PythonPath | (leave blank if UsePython=NO) |
| SAPConnectionName | (leave blank to use first open connection) |

## Step 6 — Run the Pipeline

1. Open the .xlsm workbook and click **Enable Content**.
2. Go to the **HOME** sheet.
3. Click the **Run Report** button.
4. Monitor the status bar and LOG sheet for progress.

## Troubleshooting

**"SAP GUI is not running"**  
Log in to SAP before clicking Run Report.

**"Could not connect to SAP"**  
Check that SAP GUI Scripting is enabled on your SAP GUI client (Step 1).

**"Export file not found"**  
The SAP export dialog may have opened in an unexpected state. Check the SAP window manually and re-run.

**Python errors**  
Check the `Logs/` folder for a `python_*.log` file. Common fixes:  
- Wrong `PythonPath` in SETTINGS — use the full path to python.exe.  
- Missing packages — run `pip install --user pandas openpyxl`.
