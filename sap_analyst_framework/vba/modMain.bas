Attribute VB_Name = "modMain"
Option Explicit

' =============================================================================
' modMain — Full pipeline controller
' =============================================================================

' Set True before calling RunMain for unattended / batch execution.
' When True, ShowError writes to log only and never opens a MsgBox.
Public g_BatchMode As Boolean

Public Sub RunMain()
    On Error GoTo ErrHandler

    Dim sExportFile As String
    Dim bUsePython  As Boolean
    Dim tStart      As Double
    Dim sErrDesc    As String

    tStart = Timer

    ' ── 1. Start UI feedback ────────────────────────────────────────────────
    modUIState.StartLoading "Initialising framework..."
    modLog.ClearLog
    modLog.LogMessage "RunMain", "START", "Pipeline started"

    ' ── 2. Validate settings ────────────────────────────────────────────────
    modUIState.UpdateLoading 5, "Reading settings..."
    If Not modSettings.ValidateSettings() Then
        modUIState.ShowError "Settings validation failed. Check the SETTINGS sheet."
        modLog.LogMessage "RunMain", "ERROR", "Settings validation failed"
        GoTo CleanExit
    End If
    modLog.LogMessage "RunMain", "OK", "Settings validated"

    ' ── 3. Prepare folders ──────────────────────────────────────────────────
    modUIState.UpdateLoading 10, "Preparing folders..."
    modFolders.PrepareFolders
    modLog.LogMessage "RunMain", "OK", "Folders ready"

    ' ── 4. Resolve dates ────────────────────────────────────────────────────
    modUIState.UpdateLoading 15, "Resolving dates..."
    modDates.ResolveDates
    modLog.LogMessage "RunMain", "OK", "Dates resolved"

    ' ── 5. Connect to SAP ───────────────────────────────────────────────────
    modUIState.UpdateLoading 20, "Connecting to SAP..."
    If Not modSAPConnection.ConnectToSAP() Then
        modUIState.ShowError "Could not connect to SAP. Is SAP GUI running?"
        modLog.LogMessage "RunMain", "ERROR", "SAP connection failed"
        GoTo CleanExit
    End If
    modLog.LogMessage "RunMain", "OK", "SAP connected"

    ' ── 6. Open TCode ───────────────────────────────────────────────────────
    modUIState.UpdateLoading 30, "Opening transaction..."
    modSAPNavigation.OpenTransaction modSettings.GetSetting("TCode")
    modLog.LogMessage "RunMain", "OK", "TCode opened: " & modSettings.GetSetting("TCode")

    ' ── 7. Apply variant ────────────────────────────────────────────────────
    modUIState.UpdateLoading 38, "Applying variant..."
    modSAPNavigation.ApplyVariant modSettings.GetSetting("Variant")
    modLog.LogMessage "RunMain", "OK", "Variant applied"

    ' ── 8. Apply layout ─────────────────────────────────────────────────────
    modUIState.UpdateLoading 44, "Applying layout..."
    modSAPNavigation.ApplyLayout modSettings.GetSetting("Layout")
    modLog.LogMessage "RunMain", "OK", "Layout applied"

    ' ── 9. Set date range ───────────────────────────────────────────────────
    modUIState.UpdateLoading 50, "Setting date range..."
    modSAPNavigation.SetDateRange _
        modSettings.GetSetting("DateFrom"), _
        modSettings.GetSetting("DateTo")
    modLog.LogMessage "RunMain", "OK", "Date range set"

    ' ── 10. Execute report ──────────────────────────────────────────────────
    modUIState.UpdateLoading 55, "Executing SAP report..."
    modSAPExport.ExecuteReport
    modLog.LogMessage "RunMain", "OK", "Report executed"

    ' ── 11. Export ALV to Excel ─────────────────────────────────────────────
    sExportFile = modFolders.GetExportsPath() & "\" & _
                  modSettings.GetSetting("ExportFileName") & "_" & _
                  Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    modUIState.UpdateLoading 62, "Exporting ALV grid..."
    modSAPExport.ExportALVToExcel sExportFile
    modSAPExport.WaitForExportFile sExportFile
    modLog.LogMessage "RunMain", "OK", "ALV exported: " & sExportFile

    ' ── 12. Import into DATA sheet ──────────────────────────────────────────
    modUIState.UpdateLoading 70, "Importing data..."
    modImportData.ImportExportToDataSheet sExportFile, modConfig.SHEET_DATA
    modLog.LogMessage "RunMain", "OK", "Data imported"

    ' ── 13. Optional Python processing ──────────────────────────────────────
    bUsePython = (UCase(Trim(modSettings.GetSetting("UsePython"))) = "YES")
    If bUsePython Then
        modUIState.UpdateLoading 78, "Running Python scripts..."
        If Not RunPythonPipeline(sExportFile) Then
            modUIState.ShowError "Python pipeline failed. Check Logs/ for details."
            modLog.LogMessage "RunMain", "ERROR", "Python returned non-zero exit code"
            GoTo CleanExit
        End If
        modLog.LogMessage "RunMain", "OK", "Python pipeline complete"
    End If

    ' ── 14. Calculate KPIs from REPORT_CONFIG ───────────────────────────────
    modUIState.UpdateLoading 84, "Calculating KPIs..."
    modReportEngine.RunReportEngine
    modLog.LogMessage "RunMain", "OK", "KPIs calculated"

    ' ── 15. Update HISTORY ──────────────────────────────────────────────────
    modUIState.UpdateLoading 90, "Updating history..."
    UpdateHistory CLng(Timer - tStart), "SUCCESS"
    modLog.LogMessage "RunMain", "OK", "History updated"

    ' ── 16. Refresh dashboard ───────────────────────────────────────────────
    modUIState.UpdateLoading 95, "Refreshing dashboard..."
    modUI.RefreshUI
    modLog.LogMessage "RunMain", "OK", "Dashboard refreshed"

    ' ── 17. Done ────────────────────────────────────────────────────────────
    modUIState.StopLoading "Report complete."
    modUIState.ShowSuccess "Pipeline finished successfully."
    modLog.LogMessage "RunMain", "END", "Pipeline completed successfully"
    GoTo CleanExit

ErrHandler:
    sErrDesc = Err.Description
    modUIState.ShowError "Unexpected error in RunMain: " & sErrDesc
    modLog.LogMessage "RunMain", "ERROR", "Line " & Erl & " — " & sErrDesc
    modUIState.StopLoading "Pipeline failed."
    On Error Resume Next
    UpdateHistory CLng(Timer - tStart), "ERROR: " & sErrDesc
    On Error GoTo 0

CleanExit:
End Sub

' -----------------------------------------------------------------------------
' RunPythonPipeline — returns True on success (exit code 0), False otherwise.
' Uses WScript.Shell with WaitOnReturn so VBA blocks until Python finishes
' instead of sleeping a fixed 60 seconds.
' -----------------------------------------------------------------------------
Private Function RunPythonPipeline(ByVal exportFilePath As String) As Boolean
    Dim sPython   As String
    Dim sScript   As String
    Dim sWbPath   As String
    Dim sCmd      As String
    Dim exitCode  As Long
    Dim oShell    As Object

    sPython  = modSettings.GetSetting("PythonPath")
    sScript  = ThisWorkbook.Path & "\python\main.py"
    sWbPath  = ThisWorkbook.FullName

    If sPython = "" Then sPython = "python"

    sCmd = """" & sPython & """ """ & sScript & """ """ & sWbPath & """ """ & exportFilePath & """"

    Set oShell = CreateObject("WScript.Shell")
    exitCode = oShell.Run("cmd /c " & sCmd, 0, True)   ' WaitOnReturn = True

    If exitCode <> 0 Then
        modLog.LogMessage "RunPythonPipeline", "ERROR", "Python exited with code " & exitCode
        RunPythonPipeline = False
    Else
        RunPythonPipeline = True
    End If
End Function

' -----------------------------------------------------------------------------
' UpdateHistory — appends a run record to the HISTORY sheet.
' -----------------------------------------------------------------------------
Private Sub UpdateHistory(ByVal durationSec As Long, ByVal runStatus As String)
    Dim ws   As Worksheet
    Dim lRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_HISTORY)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lRow = ws.Cells(ws.Rows.Count, modConfig.HIST_COL_DATE).End(xlUp).Row + 1
    If lRow < 2 Then lRow = 2

    ws.Cells(lRow, modConfig.HIST_COL_DATE).Value     = Now
    ws.Cells(lRow, modConfig.HIST_COL_REPORT).Value   = modSettings.GetSetting("ReportName")
    ws.Cells(lRow, modConfig.HIST_COL_PLANT).Value    = modSettings.GetSetting("Plant")
    ws.Cells(lRow, modConfig.HIST_COL_DATEFROM).Value = modSettings.GetSetting("DateFrom")
    ws.Cells(lRow, modConfig.HIST_COL_DATETO).Value   = modSettings.GetSetting("DateTo")
    ws.Cells(lRow, modConfig.HIST_COL_USER).Value     = Environ("USERNAME")
    ws.Cells(lRow, modConfig.HIST_COL_DURATION).Value = durationSec
    ws.Cells(lRow, modConfig.HIST_COL_STATUS).Value   = runStatus
End Sub
