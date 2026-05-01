Attribute VB_Name = "modFormulaEngine"
Option Explicit

' =============================================================================
' modFormulaEngine — Generic formula evaluator for REPORT_CONFIG KPIs
' Supported formula types:
'   SUM, COUNT, COUNT_UNIQUE, AVERAGE, MIN, MAX, FORMULA, CUSTOM
'
' REPORT_CONFIG multi-filter syntax (filtersStr parameter):
'   "ColumnA=Value1;ColumnB=Value2"
'   When filtersStr is non-empty it takes precedence over filterCol/filterVal.
' =============================================================================

' =============================================================================
' Calculate — main entry point called by modReportEngine
' colName    : header name in DATA sheet (or Excel formula string for FORMULA type)
' filterCol  : optional single-column filter header (legacy; use filtersStr instead)
' filterVal  : value to match in filterCol (exact, case-insensitive)
' filtersStr : multi-condition filter "col1=val1;col2=val2" (takes precedence)
' dataSheet  : sheet containing the data; defaults to modConfig.DEFAULT_DATA_SHEET
' =============================================================================
Public Function Calculate(ByVal formulaType As String, _
                          ByVal colName     As String, _
                          Optional ByVal filterCol  As String = "", _
                          Optional ByVal filterVal  As String = "", _
                          Optional ByVal filtersStr As String = "", _
                          Optional ByVal dataSheet  As String = "") As Variant

    Dim ws           As Worksheet
    Dim dataCol      As Long
    Dim filterColIdx As Long
    Dim lastRow      As Long
    Dim sSheet       As String

    On Error GoTo ErrHandler
    Calculate = 0

    sSheet = dataSheet
    If sSheet = "" Then sSheet = modConfig.DEFAULT_DATA_SHEET

    Set ws = ThisWorkbook.Sheets(sSheet)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    dataCol = FindColumnByHeader(ws, colName)
    If dataCol = 0 And formulaType <> "FORMULA" Then
        Calculate = "COL_NOT_FOUND"
        Exit Function
    End If

    If filtersStr = "" And filterCol <> "" Then
        filterColIdx = FindColumnByHeader(ws, filterCol)
    End If

    ' Build filtered value array
    Dim values() As Variant
    Dim vCount   As Long
    ReDim values(1 To lastRow - 1)
    vCount = 0

    Dim r As Long
    For r = 2 To lastRow
        If Not RowPassesAllFilters(ws, r, filterColIdx, filterVal, filtersStr) Then GoTo NextRow
        vCount = vCount + 1
        values(vCount) = ws.Cells(r, dataCol).Value
NextRow:
    Next r

    Select Case UCase(Trim(formulaType))
        Case "SUM"
            Calculate = SumValues(values, vCount)
        Case "COUNT"
            Calculate = vCount
        Case "COUNT_UNIQUE"
            Calculate = CountUniqueValues(values, vCount)
        Case "AVERAGE"
            If vCount > 0 Then Calculate = SumValues(values, vCount) / vCount Else Calculate = 0
        Case "MIN"
            Calculate = MinValue(values, vCount)
        Case "MAX"
            Calculate = MaxValue(values, vCount)
        Case "FORMULA"
            Calculate = EvaluateFormula(colName, ws, lastRow)
        Case "CUSTOM"
            On Error Resume Next
            Calculate = Application.Run(colName, ws, filterCol, filterVal)
            If Err.Number <> 0 Then Calculate = "CUSTOM_ERROR: " & Err.Description
            On Error GoTo 0
        Case Else
            Calculate = "UNKNOWN_FORMULA"
    End Select
    Exit Function

ErrHandler:
    Calculate = "ERROR: " & Err.Description
End Function

' =============================================================================
' RowPassesAllFilters — single entry point for both filter modes
' =============================================================================
Private Function RowPassesAllFilters(ByVal ws As Worksheet, ByVal r As Long, _
                                      ByVal singleColIdx As Long, ByVal singleVal As String, _
                                      ByVal filtersStr As String) As Boolean
    If filtersStr <> "" Then
        RowPassesAllFilters = RowPassesMultiFilter(ws, r, filtersStr)
    ElseIf singleColIdx > 0 Then
        RowPassesAllFilters = (LCase(Trim(CStr(ws.Cells(r, singleColIdx).Value))) = _
                               LCase(Trim(singleVal)))
    Else
        RowPassesAllFilters = True
    End If
End Function

' =============================================================================
' RowPassesMultiFilter — evaluates "col1=val1;col2=val2" against a data row
' All conditions are AND-combined. Conditions with unknown column names are skipped.
' =============================================================================
Private Function RowPassesMultiFilter(ByVal ws As Worksheet, ByVal r As Long, _
                                       ByVal filtersStr As String) As Boolean
    Dim pairs() As String
    pairs = Split(filtersStr, ";")

    Dim i      As Long
    Dim eqPos  As Long
    Dim colNm  As String
    Dim colVal As String
    Dim colIdx As Long

    For i = 0 To UBound(pairs)
        Dim pair As String
        pair = Trim(pairs(i))
        If pair = "" Then GoTo NextPair

        eqPos = InStr(pair, "=")
        If eqPos < 2 Then GoTo NextPair

        colNm  = Trim(Left(pair, eqPos - 1))
        colVal = Trim(Mid(pair, eqPos + 1))
        colIdx = FindColumnByHeader(ws, colNm)
        If colIdx = 0 Then GoTo NextPair

        If LCase(Trim(CStr(ws.Cells(r, colIdx).Value))) <> LCase(colVal) Then
            RowPassesMultiFilter = False
            Exit Function
        End If
NextPair:
    Next i
    RowPassesMultiFilter = True
End Function

' =============================================================================
' FindColumnByHeader — 1-based column index, case-insensitive header match
' =============================================================================
Private Function FindColumnByHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c       As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, c).Value))) = LCase(Trim(headerName)) Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
    FindColumnByHeader = 0
End Function

' =============================================================================
' Aggregation helpers
' =============================================================================
Private Function SumValues(ByRef v() As Variant, ByVal n As Long) As Double
    Dim i As Long
    Dim s As Double
    For i = 1 To n
        If IsNumeric(v(i)) Then s = s + CDbl(v(i))
    Next i
    SumValues = s
End Function

Private Function MinValue(ByRef v() As Variant, ByVal n As Long) As Variant
    If n = 0 Then MinValue = 0 : Exit Function
    Dim i As Long
    Dim m As Variant
    m = v(1)
    For i = 2 To n
        If IsNumeric(v(i)) Then
            If CDbl(v(i)) < CDbl(m) Then m = v(i)
        End If
    Next i
    MinValue = m
End Function

Private Function MaxValue(ByRef v() As Variant, ByVal n As Long) As Variant
    If n = 0 Then MaxValue = 0 : Exit Function
    Dim i As Long
    Dim m As Variant
    m = v(1)
    For i = 2 To n
        If IsNumeric(v(i)) Then
            If CDbl(v(i)) > CDbl(m) Then m = v(i)
        End If
    Next i
    MaxValue = m
End Function

' O(n) using Scripting.Dictionary — replaces the former O(n²) nested loop.
Private Function CountUniqueValues(ByRef v() As Variant, ByVal n As Long) As Long
    If n = 0 Then CountUniqueValues = 0 : Exit Function
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' vbTextCompare — case-insensitive keys
    Dim i As Long
    For i = 1 To n
        dict(CStr(v(i))) = 1
    Next i
    CountUniqueValues = dict.Count
End Function

Private Function EvaluateFormula(ByVal formula As String, _
                                  ByVal ws As Worksheet, _
                                  ByVal lastDataRow As Long) As Variant
    formula = Replace(formula, "{LAST_ROW}", CStr(lastDataRow))
    formula = Replace(formula, "{SHEET}", "'" & ws.Name & "'!")

    On Error Resume Next
    EvaluateFormula = Evaluate(formula)
    If Err.Number <> 0 Then EvaluateFormula = "EVAL_ERROR"
    On Error GoTo 0
End Function
