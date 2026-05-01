Attribute VB_Name = "modLog"
Option Explicit

' =============================================================================
' modLog — Structured logging to the LOG sheet and optionally a text file
' =============================================================================

Private Const LOG_SHEET  As String = "LOG"
Private Const HDR_ROW    As Long   = 1

' Column indices in LOG sheet
Private Const COL_TIME   As Long = 1  ' A — Timestamp
Private Const COL_STEP   As Long = 2  ' B — Pipeline step
Private Const COL_STATUS As Long = 3  ' C — OK | ERROR | START | END
Private Const COL_MSG    As Long = 4  ' D — Message

' =============================================================================
' LogMessage — appends one row to the LOG sheet
' =============================================================================
Public Sub LogMessage(ByVal step As String, ByVal status As String, ByVal message As String)
    Dim ws   As Worksheet
    Dim lRow As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(LOG_SHEET)

    ' Ensure header exists
    If ws.Cells(HDR_ROW, COL_TIME).Value = "" Then
        ws.Cells(HDR_ROW, COL_TIME).Value  = "Timestamp"
        ws.Cells(HDR_ROW, COL_STEP).Value  = "Step"
        ws.Cells(HDR_ROW, COL_STATUS).Value = "Status"
        ws.Cells(HDR_ROW, COL_MSG).Value   = "Message"
        ws.Rows(HDR_ROW).Font.Bold = True
    End If

    lRow = ws.Cells(ws.Rows.Count, COL_TIME).End(xlUp).Row + 1

    ws.Cells(lRow, COL_TIME).Value   = Now
    ws.Cells(lRow, COL_TIME).NumberFormat = "YYYY-MM-DD HH:MM:SS"
    ws.Cells(lRow, COL_STEP).Value   = step
    ws.Cells(lRow, COL_STATUS).Value = status
    ws.Cells(lRow, COL_MSG).Value    = message

    ' Colour-code the status cell
    Dim cel As Range
    Set cel = ws.Cells(lRow, COL_STATUS)
    Select Case UCase(status)
        Case "ERROR"
            cel.Interior.Color = RGB(220, 53, 69)
            cel.Font.Color     = RGB(255, 255, 255)
        Case "OK"
            cel.Interior.Color = RGB(40, 167, 69)
            cel.Font.Color     = RGB(255, 255, 255)
        Case "START", "END"
            cel.Interior.Color = RGB(0, 123, 255)
            cel.Font.Color     = RGB(255, 255, 255)
        Case Else
            cel.Interior.ColorIndex = xlNone
    End Select

    ' Also echo to Immediate window for VS Code / VBE debugging
    Debug.Print Format(Now, "HH:MM:SS") & " [" & status & "] " & step & " — " & message
    Exit Sub

ErrHandler:
    Debug.Print "modLog.LogMessage error: " & Err.Description
End Sub

' =============================================================================
' ClearLog — removes all data rows (keeps header)
' =============================================================================
Public Sub ClearLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(LOG_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TIME).End(xlUp).Row
    If lastRow > HDR_ROW Then
        ws.Rows(HDR_ROW + 1 & ":" & lastRow).Delete
    End If
End Sub

' =============================================================================
' ExportLogToText — saves LOG sheet as a plain text file in the Logs folder
' =============================================================================
Public Sub ExportLogToText()
    Dim ws      As Worksheet
    Dim fNum    As Integer
    Dim i       As Long
    Dim lastRow As Long
    Dim sPath   As String
    Dim sLine   As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(LOG_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, COL_TIME).End(xlUp).Row

    sPath = modFolders.GetLogsPath() & "\log_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    fNum = FreeFile
    Open sPath For Output As #fNum
    Print #fNum, "SAP Analyst Framework — Run Log"
    Print #fNum, "Generated: " & Now
    Print #fNum, String(80, "-")

    For i = HDR_ROW To lastRow
        sLine = ws.Cells(i, COL_TIME).Text & vbTab & _
                ws.Cells(i, COL_STEP).Text & vbTab & _
                ws.Cells(i, COL_STATUS).Text & vbTab & _
                ws.Cells(i, COL_MSG).Text
        Print #fNum, sLine
    Next i

    Close #fNum
    MsgBox "Log exported to:" & vbNewLine & sPath, vbInformation, "Export Log"
    Exit Sub

ErrHandler:
    If fNum > 0 Then Close #fNum
    MsgBox "Error exporting log: " & Err.Description, vbCritical
End Sub
