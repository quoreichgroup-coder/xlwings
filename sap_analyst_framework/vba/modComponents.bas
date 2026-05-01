Attribute VB_Name = "modComponents"
Option Explicit

' =============================================================================
' modComponents — Reusable Excel UI component factory
' All dimensions are in Excel row/column units unless noted as points.
' =============================================================================

' =============================================================================
' CreateButton — inserts a styled button shape with a macro assignment
' macroName must match a Public Sub name in the workbook (e.g. "modMain.RunMain")
' =============================================================================
Public Sub CreateButton(ByVal ws As Worksheet, _
                        ByVal topRow As Long, ByVal leftCol As Long, _
                        ByVal caption As String, _
                        Optional ByVal colSpan As Long = 2, _
                        Optional ByVal rowSpan As Long = 1, _
                        Optional ByVal macroName As String = "")

    Dim shp   As Shape
    Dim rng   As Range
    Set rng = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + rowSpan - 1, leftCol + colSpan - 1))

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  rng.Left, rng.Top, rng.Width, rng.Height)

    With shp
        .Name = "Btn_" & caption
        .Fill.ForeColor.RGB      = modTheme.Gold
        .Line.Visible            = msoFalse
        .TextFrame2.TextRange.Text = caption
        With .TextFrame2.TextRange.Font
            .Size      = 11
            .Bold      = msoTrue
            .Fill.ForeColor.RGB = modTheme.Charcoal
        End With
        .TextFrame2.VerticalAnchor  = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        If macroName <> "" Then .OnAction = macroName
    End With
End Sub

' =============================================================================
' CreateKpiCard — a label + large value cell block
' valueCell : e.g. "HOME!C10" — the formula cell that holds the KPI value
' =============================================================================
Public Sub CreateKpiCard(ByVal ws As Worksheet, _
                         ByVal topRow As Long, ByVal leftCol As Long, _
                         ByVal label As String, _
                         ByVal valueCell As String, _
                         Optional ByVal colSpan As Long = 2, _
                         Optional ByVal rowSpan As Long = 3)

    Dim labelCell As Range
    Dim valCell   As Range
    Dim bgRng     As Range

    Set bgRng = ws.Range(ws.Cells(topRow, leftCol), _
                         ws.Cells(topRow + rowSpan - 1, leftCol + colSpan - 1))

    ' Background
    bgRng.Interior.Color = modTheme.Charcoal
    bgRng.Borders.LineStyle = xlNone

    ' Label
    Set labelCell = ws.Cells(topRow, leftCol)
    labelCell.Value      = label
    labelCell.Font.Color = modTheme.Gold
    labelCell.Font.Bold  = True
    labelCell.Font.Size  = 9
    labelCell.HorizontalAlignment = xlCenter

    ' Value (one row below label, spanning colSpan)
    Set valCell = ws.Cells(topRow + 1, leftCol).Resize(rowSpan - 1, colSpan)
    If valueCell <> "" Then
        valCell.Cells(1, 1).Formula = "=" & valueCell
    End If
    valCell.Font.Color  = modTheme.LightGray
    valCell.Font.Bold   = True
    valCell.Font.Size   = 18
    valCell.HorizontalAlignment = xlCenter
    valCell.VerticalAlignment   = xlCenter
End Sub

' =============================================================================
' CreateStatusBox — coloured box with a status label (style: OK / ERROR / WARN)
' =============================================================================
Public Sub CreateStatusBox(ByVal ws As Worksheet, _
                           ByVal topRow As Long, ByVal leftCol As Long, _
                           ByVal label As String, _
                           Optional ByVal colSpan As Long = 2, _
                           Optional ByVal rowSpan As Long = 1, _
                           Optional ByVal style As String = "OK")

    Dim rng   As Range
    Dim bgClr As Long

    Set rng = ws.Range(ws.Cells(topRow, leftCol), _
                       ws.Cells(topRow + rowSpan - 1, leftCol + colSpan - 1))

    Select Case UCase(style)
        Case "OK"    : bgClr = modTheme.SuccessGreen
        Case "ERROR" : bgClr = modTheme.ErrorRed
        Case "WARN"  : bgClr = RGB(255, 193, 7)
        Case Else    : bgClr = modTheme.LightGray
    End Select

    rng.Interior.Color       = bgClr
    rng.Value                = label
    rng.Font.Color           = modTheme.White
    rng.Font.Bold            = True
    rng.HorizontalAlignment  = xlCenter
    rng.VerticalAlignment    = xlCenter
End Sub

' =============================================================================
' CreateSectionHeader — wide header strip with Gold accent
' =============================================================================
Public Sub CreateSectionHeader(ByVal ws As Worksheet, _
                               ByVal topRow As Long, ByVal leftCol As Long, _
                               ByVal title As String, _
                               Optional ByVal colSpan As Long = 6)

    Dim rng As Range
    Set rng = ws.Cells(topRow, leftCol).Resize(1, colSpan)

    rng.Merge
    rng.Value              = "  " & UCase(title)
    rng.Interior.Color     = modTheme.Charcoal
    rng.Font.Color         = modTheme.Gold
    rng.Font.Bold          = True
    rng.Font.Size          = 12
    rng.RowHeight          = 28
    rng.VerticalAlignment  = xlCenter
End Sub

' =============================================================================
' CreateProgressBar — horizontal bar background + fill shape pair
' percent 0-100; caller updates fill shape width via modUIState.UpdateLoading
' =============================================================================
Public Sub CreateProgressBar(ByVal ws As Worksheet, _
                             ByVal topRow As Long, ByVal leftCol As Long, _
                             ByVal percent As Integer, _
                             Optional ByVal colSpan As Long = 6, _
                             Optional ByVal rowSpan As Long = 1)

    Dim rng    As Range
    Dim bgShp  As Shape
    Dim fillShp As Shape

    Set rng = ws.Range(ws.Cells(topRow, leftCol), _
                       ws.Cells(topRow + rowSpan - 1, leftCol + colSpan - 1))

    ' Background bar
    Set bgShp = ws.Shapes.AddShape(msoShapeRectangle, _
                                    rng.Left, rng.Top, rng.Width, rng.Height)
    bgShp.Name                   = "ProgressBar_BG"
    bgShp.Fill.ForeColor.RGB     = modTheme.LightGray
    bgShp.Line.Visible           = msoFalse

    ' Fill bar (starts at 0 width)
    Dim fillWidth As Single
    fillWidth = rng.Width * (percent / 100!)
    If fillWidth < 1 Then fillWidth = 1

    Set fillShp = ws.Shapes.AddShape(msoShapeRectangle, _
                                      rng.Left, rng.Top, fillWidth, rng.Height)
    fillShp.Name                 = "ProgressBar_Fill"
    fillShp.Fill.ForeColor.RGB   = modTheme.Gold
    fillShp.Line.Visible         = msoFalse
    fillShp.Visible              = (percent > 0)
End Sub

' =============================================================================
' CreateProgressCircle — circular loading indicator
' =============================================================================
Public Sub CreateProgressCircle(ByVal ws As Worksheet, _
                                ByVal topRow As Long, ByVal leftCol As Long, _
                                ByVal percent As Integer, _
                                Optional ByVal sizePts As Single = 40)

    Dim rng As Range
    Set rng = ws.Cells(topRow, leftCol)

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeOval, _
                                  rng.Left, rng.Top, sizePts, sizePts)
    shp.Name                 = "LoaderCircle"
    shp.Fill.ForeColor.RGB   = modTheme.Gold
    shp.Line.ForeColor.RGB   = modTheme.Charcoal
    shp.Line.Weight          = 2
    shp.Rotation             = (percent / 100!) * 360
    shp.Visible              = (percent > 0)
End Sub

' =============================================================================
' CreateAlertBox — popup-style alert merged cell block
' =============================================================================
Public Sub CreateAlertBox(ByVal ws As Worksheet, _
                          ByVal topRow As Long, ByVal leftCol As Long, _
                          ByVal message As String, _
                          Optional ByVal colSpan As Long = 4, _
                          Optional ByVal rowSpan As Long = 2, _
                          Optional ByVal style As String = "WARN")

    CreateStatusBox ws, topRow, leftCol, message, colSpan, rowSpan, style
End Sub

' =============================================================================
' CreateNavigationButton — a button that activates another sheet
' targetSheet : name of sheet to navigate to (used as macro name via wrapper)
' =============================================================================
Public Sub CreateNavigationButton(ByVal ws As Worksheet, _
                                  ByVal topRow As Long, ByVal leftCol As Long, _
                                  ByVal caption As String, _
                                  Optional ByVal colSpan As Long = 2, _
                                  Optional ByVal rowSpan As Long = 1, _
                                  Optional ByVal targetSheet As String = "")

    CreateButton ws, topRow, leftCol, caption, colSpan, rowSpan, _
                 IIf(targetSheet <> "", "modUI.NavigateTo_" & targetSheet, "")
End Sub
