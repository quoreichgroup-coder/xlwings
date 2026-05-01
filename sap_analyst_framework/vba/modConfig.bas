Attribute VB_Name = "modConfig"
Option Explicit

' =============================================================================
' modConfig — Single source of truth for sheet names, column indices, defaults.
' Rename a sheet or add a column here only; all modules reference these.
' =============================================================================

' ── Sheet names ──────────────────────────────────────────────────────────────
Public Const SHEET_HOME     As String = "HOME"
Public Const SHEET_DATA     As String = "DATA"
Public Const SHEET_RC       As String = "REPORT_CONFIG"
Public Const SHEET_UI       As String = "UI_CONFIG"
Public Const SHEET_LOG      As String = "LOG"
Public Const SHEET_HISTORY  As String = "HISTORY"
Public Const SHEET_SETTINGS As String = "SETTINGS"
Public Const SHEET_REGISTRY As String = "REPORT_REGISTRY"

' ── HISTORY column indices (1-based) ─────────────────────────────────────────
Public Const HIST_COL_DATE     As Long = 1   ' A
Public Const HIST_COL_REPORT   As Long = 2   ' B
Public Const HIST_COL_PLANT    As Long = 3   ' C
Public Const HIST_COL_DATEFROM As Long = 4   ' D
Public Const HIST_COL_DATETO   As Long = 5   ' E
Public Const HIST_COL_USER     As Long = 6   ' F
Public Const HIST_COL_DURATION As Long = 7   ' G — seconds
Public Const HIST_COL_STATUS   As Long = 8   ' H — "SUCCESS" or "ERROR: ..."

' ── Pipeline defaults ─────────────────────────────────────────────────────────
Public Const DEFAULT_LOADER_TYPE As String = "BAR"
Public Const DEFAULT_DATA_SHEET  As String = "DATA"
