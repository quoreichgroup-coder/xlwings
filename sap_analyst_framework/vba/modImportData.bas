Attribute VB_Name = "modImportData"
Option Explicit

' =============================================================================
' modImportData — Imports SAP export files into the DATA sheet
' =============================================================================

' =============================================================================
' ImportExportToDataSheet — opens an XLSX export and pastes values into DATA
' Preserves workbook macros by NOT opening the target workbook for editing.
' =============================================================================
Public Sub ImportExportToDataSheet(ByVal filePath As String, ByVal targetSheetName As String)
    If Dir(filePath) = "" Then
        modLog.LogMessage "ImportExportToDataSheet", "ERROR", "File not found: " & filePath
        Exit Sub
    End If

    Dim wbSrc   As Workbook
    Dim wsSrc   As Worksheet
    Dim wsDst   As Worksheet
    Dim rngSrc  As Range
    Dim lastRow As Long
    Dim lastCol As Long

    On Error GoTo ErrHandler

    ' Open source file (read-only, no macros needed)
    Application.ScreenUpdating = False
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsSrc = wbSrc.Sheets(1)

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then
        wbSrc.Close SaveChanges:=False
        modLog.LogMessage "ImportExportToDataSheet", "ERROR", "Source file is empty"
        GoTo CleanExit
    End If

    Set rngSrc = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))

    ' Clear destination and paste values only
    ClearDataSheet targetSheetName

    On Error Resume Next
    Set wsDst = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo ErrHandler

    If wsDst Is Nothing Then
        wbSrc.Close SaveChanges:=False
        modLog.LogMessage "ImportExportToDataSheet", "ERROR", "Sheet not found: " & targetSheetName
        GoTo CleanExit
    End If

    wsDst.Cells(1, 1).Resize(rngSrc.Rows.Count, rngSrc.Columns.Count).Value = rngSrc.Value

    wbSrc.Close SaveChanges:=False
    modLog.LogMessage "ImportExportToDataSheet", "OK", _
        lastRow - 1 & " data rows imported into " & targetSheetName

    GoTo CleanExit

ErrHandler:
    modLog.LogMessage "ImportExportToDataSheet", "ERROR", Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False

CleanExit:
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' ClearDataSheet — removes all data (keeps sheet structure / formatting)
' =============================================================================
Public Sub ClearDataSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Cells.ClearContents
End Sub

' =============================================================================
' RefreshWorkbook — recalculates all formulas and refreshes pivot tables
' =============================================================================
Public Sub RefreshWorkbook()
    Application.CalculateFull

    Dim ws  As Worksheet
    Dim pt  As PivotTable
    For Each ws In ThisWorkbook.Sheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub

' =============================================================================
' GetDataLastRow — returns the last used row index in the DATA sheet
' =============================================================================
Public Function GetDataLastRow(Optional ByVal sheetName As String = "DATA") As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        GetDataLastRow = 0
        Exit Function
    End If
    GetDataLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

' =============================================================================
' GetDataRange — returns the full data range (including header row)
' =============================================================================
Public Function GetDataRange(Optional ByVal sheetName As String = "DATA") As Range
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow >= 1 And lastCol >= 1 Then
        Set GetDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
End Function
