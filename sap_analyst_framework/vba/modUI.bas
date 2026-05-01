Attribute VB_Name = "modUI"
Option Explicit

' =============================================================================
' modUI — Builds, refreshes, and protects the dashboard UI
' UI_MODE = MANUAL : only updates target cells, does not touch layout
' UI_MODE = AUTO   : builds components from UI_CONFIG sheet
' =============================================================================

Private Const UI_CONFIG_SHEET As String = "UI_CONFIG"
Private Const HOME_SHEET      As String = "HOME"

' UI_CONFIG layout (headers in row 1):
'   A: Component   B: Sheet   C: Row   D: Col   E: Width   F: Height
'   G: Label   H: Value_Cell   I: Style   J: Color_Override

' =============================================================================
' BuildUIFromConfig — entry point; reads UI_MODE and dispatches
' =============================================================================
Public Sub BuildUIFromConfig()
    Dim mode As String
    mode = UCase(Trim(modSettings.GetSetting("UI_MODE")))

    Select Case mode
        Case "AUTO"
            BuildAutoUI
        Case "MANUAL"
            ' Nothing to build; analyst has designed the sheet already
            modLog.LogMessage "BuildUIFromConfig", "OK", "MANUAL mode — UI layout unchanged"
        Case Else
            modLog.LogMessage "BuildUIFromConfig", "ERROR", "Unknown UI_MODE: " & mode
    End Select
End Sub

' =============================================================================
' BuildAutoUI — reads UI_CONFIG and creates each component
' =============================================================================
Private Sub BuildAutoUI()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(UI_CONFIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        modLog.LogMessage "BuildAutoUI", "ERROR", "Sheet not found: " & UI_CONFIG_SHEET
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        Dim comp    As String
        Dim tSheet  As String
        Dim tRow    As Long
        Dim tCol    As Long
        Dim w       As Long
        Dim h       As Long
        Dim lbl     As String
        Dim valCell As String
        Dim style   As String
        Dim colorOv As String

        comp    = Trim(CStr(ws.Cells(i, 1).Value))
        tSheet  = Trim(CStr(ws.Cells(i, 2).Value))
        tRow    = Val(ws.Cells(i, 3).Value)
        tCol    = Val(ws.Cells(i, 4).Value)
        w       = Val(ws.Cells(i, 5).Value)
        h       = Val(ws.Cells(i, 6).Value)
        lbl     = Trim(CStr(ws.Cells(i, 7).Value))
        valCell = Trim(CStr(ws.Cells(i, 8).Value))
        style   = Trim(CStr(ws.Cells(i, 9).Value))
        colorOv = Trim(CStr(ws.Cells(i, 10).Value))

        If comp = "" Or tSheet = "" Or tRow = 0 Or tCol = 0 Then GoTo NextComp

        Dim targetWS As Worksheet
        On Error Resume Next
        Set targetWS = ThisWorkbook.Sheets(tSheet)
        On Error GoTo 0
        If targetWS Is Nothing Then GoTo NextComp

        Select Case UCase(comp)
            Case "BUTTON"
                modComponents.CreateButton targetWS, tRow, tCol, lbl, w, h, style
            Case "KPI_CARD"
                modComponents.CreateKpiCard targetWS, tRow, tCol, lbl, valCell, w, h
            Case "STATUS_BOX"
                modComponents.CreateStatusBox targetWS, tRow, tCol, lbl, w, h, style
            Case "SECTION_HEADER"
                modComponents.CreateSectionHeader targetWS, tRow, tCol, lbl, w
            Case "NAV_BUTTON"
                modComponents.CreateNavigationButton targetWS, tRow, tCol, lbl, w, h, valCell
            Case "PROGRESS_BAR"
                modComponents.CreateProgressBar targetWS, tRow, tCol, 0, w, h
            Case "PROGRESS_CIRCLE"
                modComponents.CreateProgressCircle targetWS, tRow, tCol, 0, h
            Case "ALERT_BOX"
                modComponents.CreateAlertBox targetWS, tRow, tCol, lbl, w, h, style
        End Select

NextComp:
    Next i

    modLog.LogMessage "BuildAutoUI", "OK", "UI built from " & UI_CONFIG_SHEET
End Sub

' =============================================================================
' RefreshUI — updates value cells without rebuilding layout
' =============================================================================
Public Sub RefreshUI()
    ' Force recalculation so any cells referencing DATA are current
    Application.CalculateFull
    modLog.LogMessage "RefreshUI", "OK", "UI refreshed"
End Sub

' =============================================================================
' ProtectUI / UnprotectUI — lock the HOME sheet to prevent accidental edits
' =============================================================================
Public Sub ProtectUI()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOME_SHEET)
    If Not ws Is Nothing Then
        ws.Protect Password:="", UserInterfaceOnly:=True, DrawingObjects:=False
    End If
End Sub

Public Sub UnprotectUI()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOME_SHEET)
    If Not ws Is Nothing Then ws.Unprotect
End Sub
