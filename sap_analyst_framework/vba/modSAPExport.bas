Attribute VB_Name = "modSAPExport"
Option Explicit

' =============================================================================
' modSAPExport — Execute SAP reports and export ALV grids to Excel
' =============================================================================

Private Const EXPORT_TIMEOUT_SEC As Long = 120

' =============================================================================
' ExecuteReport — presses F8 (Execute) on the current selection screen
' =============================================================================
Public Sub ExecuteReport()
    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    session.FindById("wnd[0]").SendVKey 8   ' F8 = Execute
    modSAPNavigation.WaitMS 1500
    Exit Sub
ErrHandler:
    modLog.LogMessage "ExecuteReport", "ERROR", Err.Description
End Sub

' =============================================================================
' ExportALVToExcel — exports the ALV result list to a local Excel file
' Uses the standard SAP menu: List > Export > Spreadsheet
' =============================================================================
Public Sub ExportALVToExcel(ByVal filePath As String)
    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    ' Menu path: System > List > Save > Local File (or List > Export > Spreadsheet)
    ' The exact menu ID varies; common path shown below.
    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select   ' List > Export > Spreadsheet
    modSAPNavigation.WaitMS 800

    HandleSAPExportDialogs filePath
    Exit Sub
ErrHandler:
    modLog.LogMessage "ExportALVToExcel", "ERROR", Err.Description
End Sub

' =============================================================================
' HandleSAPExportDialogs — handles the file-type and path dialogs
' =============================================================================
Public Sub HandleSAPExportDialogs(ByVal filePath As String)
    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    modSAPNavigation.WaitMS 500

    ' Dialog 1: choose file format (Spreadsheet = XLSX)
    Dim dlg1 As Object
    On Error Resume Next
    Set dlg1 = session.FindById("wnd[1]")
    On Error GoTo ErrHandler

    If Not dlg1 Is Nothing Then
        ' Select "Spreadsheet" radio button
        On Error Resume Next
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press  ' OK / Enter
        modSAPNavigation.WaitMS 500
        On Error GoTo ErrHandler
    End If

    ' Dialog 2: file path
    Dim pathBox As Object
    On Error Resume Next
    Set pathBox = session.FindById("wnd[1]/usr/ctxtDY_PATH")
    On Error GoTo ErrHandler

    If Not pathBox Is Nothing Then
        Dim fDir  As String
        Dim fName As String
        fDir  = Left(filePath, InStrRev(filePath, "\"))
        fName = Mid(filePath, InStrRev(filePath, "\") + 1)

        session.FindById("wnd[1]/usr/ctxtDY_PATH").Text     = fDir
        session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = fName
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press     ' Save
        modSAPNavigation.WaitMS 500
    End If

    ' Replace prompt if file exists
    On Error Resume Next
    Dim replaceBtn As Object
    Set replaceBtn = session.FindById("wnd[1]/tbar[0]/btn[11]")
    If Not replaceBtn Is Nothing Then replaceBtn.Press
    On Error GoTo 0
    Exit Sub

ErrHandler:
    modLog.LogMessage "HandleSAPExportDialogs", "ERROR", Err.Description
End Sub

' =============================================================================
' WaitForExportFile — polls until the file appears (or timeout)
' =============================================================================
Public Sub WaitForExportFile(ByVal filePath As String)
    Dim tEnd As Date
    tEnd = Now + TimeSerial(0, 0, EXPORT_TIMEOUT_SEC)

    Do While Dir(filePath) = ""
        DoEvents
        If Now > tEnd Then
            Err.Raise vbObjectError + 1002, "WaitForExportFile", _
                      "Export file not found within " & EXPORT_TIMEOUT_SEC & " seconds: " & filePath
        End If
    Loop
End Sub
