Attribute VB_Name = "modUIState"
Option Explicit

' =============================================================================
' modUIState — Loading indicators, status messages, and error/success feedback
' LOADER_TYPE = BAR    : progress bar cells on HOME sheet
' LOADER_TYPE = CIRCLE : circle shape animation on HOME sheet
' =============================================================================

' Fallback cell addresses used when Named Ranges are not defined in the workbook.
Private Const STATUS_CELL_DEFAULT   As String = "B3"
Private Const PROGRESS_CELL_DEFAULT As String = "B4"
Private Const PROGRESS_BAR_SHAPE    As String = "ProgressBar_Fill"
Private Const LOADER_SHAPE          As String = "LoaderCircle"

Private g_LoaderType As String

' =============================================================================
' StartLoading — initialises loading state
' =============================================================================
Public Sub StartLoading(Optional ByVal message As String = "Loading...")
    g_LoaderType = UCase(Trim(modSettings.GetSetting("LOADER_TYPE")))
    If g_LoaderType = "" Then g_LoaderType = modConfig.DEFAULT_LOADER_TYPE

    SetStatus message
    SetProgress 0

    Select Case g_LoaderType
        Case "BAR"    : ShowProgressBar True
        Case "CIRCLE" : ShowCircleLoader True
    End Select

    DoEvents
End Sub

' =============================================================================
' UpdateLoading — sets percent complete and status text
' =============================================================================
Public Sub UpdateLoading(ByVal percent As Integer, Optional ByVal message As String = "")
    If percent < 0 Then percent = 0
    If percent > 100 Then percent = 100

    If message <> "" Then SetStatus message
    SetProgress percent

    Select Case g_LoaderType
        Case "BAR"    : UpdateProgressBar percent
        Case "CIRCLE" : AnimateCircle percent
    End Select

    DoEvents
End Sub

' =============================================================================
' StopLoading — hides all loading indicators
' =============================================================================
Public Sub StopLoading(Optional ByVal message As String = "Done.")
    SetStatus message
    SetProgress 100

    Select Case g_LoaderType
        Case "BAR"    : ShowProgressBar False
        Case "CIRCLE" : ShowCircleLoader False
    End Select

    DoEvents
End Sub

' =============================================================================
' ShowError — displays an error in the status area (red).
' When modMain.g_BatchMode is True, skips the MsgBox so unattended runs
' do not hang waiting for user input.
' =============================================================================
Public Sub ShowError(ByVal message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If Not ws Is Nothing Then
        GetStatusCell(ws).Value      = "ERROR: " & message
        GetStatusCell(ws).Font.Color = modTheme.ErrorRed
        GetStatusCell(ws).Font.Bold  = True
    End If
    On Error GoTo 0

    If Not modMain.g_BatchMode Then
        MsgBox message, vbCritical, "Error"
    End If
End Sub

' =============================================================================
' ShowSuccess — displays a success message (green)
' =============================================================================
Public Sub ShowSuccess(ByVal message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    GetStatusCell(ws).Value      = message
    GetStatusCell(ws).Font.Color = modTheme.SuccessGreen
    GetStatusCell(ws).Font.Bold  = True
End Sub

' =============================================================================
' SetStatus — updates the status text cell
' =============================================================================
Public Sub SetStatus(ByVal message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    GetStatusCell(ws).Value      = message
    GetStatusCell(ws).Font.Color = modTheme.Charcoal
    GetStatusCell(ws).Font.Bold  = False
End Sub

' =============================================================================
' GetStatusCell — resolves to the Named Range "StatusCell" if defined,
' otherwise falls back to the hard-coded default address.
' =============================================================================
Private Function GetStatusCell(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set GetStatusCell = ThisWorkbook.Names(modConfig.NR_STATUS_CELL).RefersToRange
    On Error GoTo 0
    If GetStatusCell Is Nothing Then
        Set GetStatusCell = ws.Range(STATUS_CELL_DEFAULT)
    End If
End Function

Private Function GetProgressCell(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set GetProgressCell = ThisWorkbook.Names(modConfig.NR_PROGRESS_CELL).RefersToRange
    On Error GoTo 0
    If GetProgressCell Is Nothing Then
        Set GetProgressCell = ws.Range(PROGRESS_CELL_DEFAULT)
    End If
End Function

' =============================================================================
' SetProgress — writes percent value to the progress cell
' =============================================================================
Private Sub SetProgress(ByVal percent As Integer)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If Not ws Is Nothing Then
        GetProgressCell(ws).Value = percent & "%"
    End If
End Sub

' =============================================================================
' UpdateProgressBar — resizes the fill shape to match percent
' =============================================================================
Private Sub UpdateProgressBar(ByVal percent As Integer)
    Dim ws  As Worksheet
    Dim shp As Shape
    On Error Resume Next
    Set ws  = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    Set shp = ws.Shapes(PROGRESS_BAR_SHAPE)
    If shp Is Nothing Then Exit Sub

    Dim bgShp As Shape
    Set bgShp = ws.Shapes("ProgressBar_BG")
    Dim maxWidth As Single
    maxWidth = IIf(Not bgShp Is Nothing, bgShp.Width, 400)

    shp.Width   = maxWidth * (percent / 100!)
    shp.Visible = True
End Sub

Private Sub ShowProgressBar(ByVal visible As Boolean)
    Dim ws  As Worksheet
    Dim shp As Shape
    On Error Resume Next
    Set ws  = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    Set shp = ws.Shapes(PROGRESS_BAR_SHAPE)
    If Not shp Is Nothing Then shp.Visible = visible
End Sub

Private Sub ShowCircleLoader(ByVal visible As Boolean)
    Dim ws  As Worksheet
    Dim shp As Shape
    On Error Resume Next
    Set ws  = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    Set shp = ws.Shapes(LOADER_SHAPE)
    If Not shp Is Nothing Then shp.Visible = visible
End Sub

Private Sub AnimateCircle(ByVal percent As Integer)
    Dim ws  As Worksheet
    Dim shp As Shape
    On Error Resume Next
    Set ws  = ThisWorkbook.Sheets(modConfig.SHEET_HOME)
    If ws Is Nothing Then Exit Sub
    Set shp = ws.Shapes(LOADER_SHAPE)
    If Not shp Is Nothing Then
        shp.Rotation = (percent / 100!) * 360
        shp.Visible  = True
    End If
End Sub
