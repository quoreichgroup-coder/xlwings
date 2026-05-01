Attribute VB_Name = "modReportEngine"
Option Explicit

' =============================================================================
' modReportEngine — Configuration-driven KPI engine
' Reads REPORT_CONFIG sheet and calculates all KPIs without hard-coded logic.
' =============================================================================
' REPORT_CONFIG column layout (headers in row 1):
'   A: KPI_Name      B: Formula_Type   C: Column_Name   D: Target_Sheet
'   E: Target_Cell   F: Filter_Column  G: Filter_Value  H: Label
'   I: Filters       J: Threshold_Green  K: Threshold_Orange  L: Threshold_Red
'   M: Format        N: DependsOn      O: DataSheet
'
' Column I (Filters) — multi-condition filter string, e.g. "MovType=261;Plant=1000"
'   When non-empty, takes precedence over F/G (Filter_Column / Filter_Value).
' Column N (DependsOn) — used by the RATIO formula type:
'   Formula_Type=RATIO, Column_Name=<numerator KPI name>, DependsOn=<denominator KPI name>
'   Result = numerator_result / denominator_result * 100
' Column O (DataSheet) — name of the sheet containing source data (default: DATA)
' =============================================================================

Private Const RC_HDR_ROW       As Long = 1
Private Const RC_KPI_NAME      As Long = 1    ' A
Private Const RC_FORMULA       As Long = 2    ' B
Private Const RC_COL_NAME      As Long = 3    ' C
Private Const RC_TGT_SHEET     As Long = 4    ' D
Private Const RC_TGT_CELL      As Long = 5    ' E
Private Const RC_FLT_COL       As Long = 6    ' F
Private Const RC_FLT_VAL       As Long = 7    ' G
Private Const RC_LABEL         As Long = 8    ' H
Private Const RC_FILTERS       As Long = 9    ' I — multi-filter
Private Const RC_THRESH_GREEN  As Long = 10   ' J
Private Const RC_THRESH_ORANGE As Long = 11   ' K
Private Const RC_THRESH_RED    As Long = 12   ' L
Private Const RC_FORMAT        As Long = 13   ' M — NumberFormat string
Private Const RC_DEPENDS_ON    As Long = 14   ' N — denominator KPI name for RATIO
Private Const RC_DATA_SHEET    As Long = 15   ' O — source data sheet

' =============================================================================
' KPIConfig — represents one row in REPORT_CONFIG
' =============================================================================
Private Type KPIConfig
    KPIName      As String
    FormulaType  As String
    ColName      As String
    TargetSheet  As String
    TargetCell   As String
    FilterCol    As String
    FilterVal    As String
    Label        As String
    Filters      As String   ' multi-condition "col=val;col2=val2"
    ThreshGreen  As String   ' numeric string; value >= this → green
    ThreshOrange As String   ' numeric string; value >= this → orange
    ThreshRed    As String   ' set (even empty string) to enable red coloring below orange
    FormatStr    As String   ' Excel number format, e.g. "#,##0" or "0.0%"
    DependsOn    As String   ' denominator KPI name (RATIO formula only)
    DataSheet    As String   ' source sheet; defaults to modConfig.DEFAULT_DATA_SHEET
    Result       As Variant
End Type

' =============================================================================
' RunReportEngine — public entry point
' =============================================================================
Public Sub RunReportEngine()
    Dim configs() As KPIConfig
    Dim n         As Long

    ReadReportConfig configs, n
    If n = 0 Then
        modLog.LogMessage "RunReportEngine", "ERROR", "No KPI configs found in " & modConfig.SHEET_RC
        Exit Sub
    End If

    CalculateConfiguredKPIs configs, n
    WriteKPIResults configs, n

    modLog.LogMessage "RunReportEngine", "OK", n & " KPI(s) calculated"
End Sub

' =============================================================================
' ReadReportConfig — populates the KPIConfig array from REPORT_CONFIG sheet
' =============================================================================
Public Sub ReadReportConfig(ByRef configs() As KPIConfig, ByRef count As Long)
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_RC)
    On Error GoTo 0
    If ws Is Nothing Then count = 0 : Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, RC_KPI_NAME).End(xlUp).Row
    count = lastRow - RC_HDR_ROW
    If count <= 0 Then Exit Sub

    ReDim configs(1 To count)
    For i = 1 To count
        Dim r As Long
        r = RC_HDR_ROW + i
        configs(i).KPIName      = Trim(CStr(ws.Cells(r, RC_KPI_NAME).Value))
        configs(i).FormulaType  = UCase(Trim(CStr(ws.Cells(r, RC_FORMULA).Value)))
        configs(i).ColName      = Trim(CStr(ws.Cells(r, RC_COL_NAME).Value))
        configs(i).TargetSheet  = Trim(CStr(ws.Cells(r, RC_TGT_SHEET).Value))
        configs(i).TargetCell   = Trim(CStr(ws.Cells(r, RC_TGT_CELL).Value))
        configs(i).FilterCol    = Trim(CStr(ws.Cells(r, RC_FLT_COL).Value))
        configs(i).FilterVal    = Trim(CStr(ws.Cells(r, RC_FLT_VAL).Value))
        configs(i).Label        = Trim(CStr(ws.Cells(r, RC_LABEL).Value))
        configs(i).Filters      = Trim(CStr(ws.Cells(r, RC_FILTERS).Value))
        configs(i).ThreshGreen  = Trim(CStr(ws.Cells(r, RC_THRESH_GREEN).Value))
        configs(i).ThreshOrange = Trim(CStr(ws.Cells(r, RC_THRESH_ORANGE).Value))
        configs(i).ThreshRed    = Trim(CStr(ws.Cells(r, RC_THRESH_RED).Value))
        configs(i).FormatStr    = Trim(CStr(ws.Cells(r, RC_FORMAT).Value))
        configs(i).DependsOn    = Trim(CStr(ws.Cells(r, RC_DEPENDS_ON).Value))
        configs(i).DataSheet    = Trim(CStr(ws.Cells(r, RC_DATA_SHEET).Value))
        If configs(i).DataSheet = "" Then configs(i).DataSheet = modConfig.DEFAULT_DATA_SHEET
        configs(i).Result = Empty
    Next i
End Sub

' =============================================================================
' CalculateConfiguredKPIs — two-pass evaluation
'   Pass 1: all non-RATIO KPIs (SUM, COUNT, etc.) — results stored in resultMap
'   Pass 2: RATIO KPIs — numerator and denominator resolved from resultMap
' =============================================================================
Public Sub CalculateConfiguredKPIs(ByRef configs() As KPIConfig, ByVal n As Long)
    Dim resultMap As Object
    Set resultMap = CreateObject("Scripting.Dictionary")
    resultMap.CompareMode = 1   ' vbTextCompare

    Dim i As Long

    ' Pass 1 — all formula types except RATIO
    For i = 1 To n
        If configs(i).FormulaType = "RATIO" Then GoTo NextPass1
        If configs(i).KPIName = "" Or configs(i).FormulaType = "" Then GoTo NextPass1

        configs(i).Result = modFormulaEngine.Calculate( _
            configs(i).FormulaType, _
            configs(i).ColName, _
            configs(i).FilterCol, _
            configs(i).FilterVal, _
            configs(i).Filters, _
            configs(i).DataSheet)

        resultMap(configs(i).KPIName) = configs(i).Result
        modLog.LogMessage "CalculateKPIs", "OK", configs(i).KPIName & " = " & CStr(configs(i).Result)
NextPass1:
    Next i

    ' Pass 2 — RATIO: numerator / denominator * 100
    For i = 1 To n
        If configs(i).FormulaType <> "RATIO" Then GoTo NextPass2
        If configs(i).KPIName = "" Then GoTo NextPass2

        Dim numerator   As Variant
        Dim denominator As Variant

        numerator   = IIf(resultMap.Exists(configs(i).ColName),   resultMap(configs(i).ColName),   0)
        denominator = IIf(resultMap.Exists(configs(i).DependsOn), resultMap(configs(i).DependsOn), 0)

        If IsNumeric(denominator) And CDbl(denominator) <> 0 Then
            configs(i).Result = CDbl(numerator) / CDbl(denominator) * 100
        Else
            configs(i).Result = 0
        End If

        resultMap(configs(i).KPIName) = configs(i).Result
        modLog.LogMessage "CalculateKPIs", "OK", configs(i).KPIName & " = " & CStr(configs(i).Result)
NextPass2:
    Next i
End Sub

' =============================================================================
' WriteKPIResults — writes results to target cells, applies format and
' threshold colours to the target cell's font.
' =============================================================================
Public Sub WriteKPIResults(ByRef configs() As KPIConfig, ByVal n As Long)
    Dim i   As Long
    Dim ws  As Worksheet
    Dim rng As Range

    For i = 1 To n
        If configs(i).TargetSheet = "" Or configs(i).TargetCell = "" Then GoTo NextItem

        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(configs(i).TargetSheet)
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextItem

        Set rng = ws.Range(configs(i).TargetCell)
        rng.Value = configs(i).Result

        If configs(i).FormatStr <> "" Then
            rng.NumberFormat = configs(i).FormatStr
        End If

        If IsNumeric(configs(i).Result) Then
            Dim clr As Long
            clr = GetThresholdColor(CDbl(configs(i).Result), configs(i))
            If clr <> -1 Then rng.Font.Color = clr
        End If

NextItem:
        Set ws = Nothing
    Next i
End Sub

' =============================================================================
' GetThresholdColor — returns an RGB color based on KPI thresholds,
' or -1 if no threshold is configured for this KPI.
' Rule: value >= ThreshGreen → green, >= ThreshOrange → orange, else → red.
' =============================================================================
Private Function GetThresholdColor(ByVal value As Double, ByRef cfg As KPIConfig) As Long
    GetThresholdColor = -1

    On Error Resume Next
    If cfg.ThreshGreen <> "" Then
        If value >= CDbl(cfg.ThreshGreen) Then
            GetThresholdColor = modTheme.SuccessGreen : Exit Function
        End If
    End If
    If cfg.ThreshOrange <> "" Then
        If value >= CDbl(cfg.ThreshOrange) Then
            GetThresholdColor = RGB(255, 140, 0) : Exit Function
        End If
    End If
    If cfg.ThreshRed <> "" Then
        GetThresholdColor = modTheme.ErrorRed
    End If
    On Error GoTo 0
End Function
