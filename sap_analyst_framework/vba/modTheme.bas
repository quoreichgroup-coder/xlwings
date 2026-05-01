Attribute VB_Name = "modTheme"
Option Explicit

' =============================================================================
' modTheme — Central colour palette for all UI components
' Access colours as read-only properties so any component stays consistent.
' =============================================================================

' Primary palette
Public Property Get Gold() As Long
    Gold = RGB(204, 170, 102)       ' #CCAA66
End Property

Public Property Get Charcoal() As Long
    Charcoal = RGB(31, 31, 31)      ' #1F1F1F
End Property

Public Property Get LightGray() As Long
    LightGray = RGB(250, 250, 250)  ' #FAFAFA
End Property

Public Property Get White() As Long
    White = RGB(255, 255, 255)      ' #FFFFFF
End Property

' Semantic colours
Public Property Get ErrorRed() As Long
    ErrorRed = RGB(220, 53, 69)     ' Bootstrap danger-red
End Property

Public Property Get SuccessGreen() As Long
    SuccessGreen = RGB(40, 167, 69) ' Bootstrap success-green
End Property

Public Property Get WarnYellow() As Long
    WarnYellow = RGB(255, 193, 7)   ' Bootstrap warning-yellow
End Property

Public Property Get InfoBlue() As Long
    InfoBlue = RGB(0, 123, 255)     ' Bootstrap primary-blue
End Property

Public Property Get MidGray() As Long
    MidGray = RGB(108, 117, 125)    ' Bootstrap secondary
End Property

' =============================================================================
' ApplyThemeToRange — applies the default Charcoal/Gold palette to a range
' =============================================================================
Public Sub ApplyThemeToRange(ByVal rng As Range, _
                              Optional ByVal bgColor As Long = -1, _
                              Optional ByVal fgColor As Long = -1)
    If bgColor = -1 Then bgColor = Charcoal
    If fgColor = -1 Then fgColor = Gold

    rng.Interior.Color = bgColor
    rng.Font.Color     = fgColor
End Sub

' =============================================================================
' ApplyThemeToShape — fills a shape with theme colours
' =============================================================================
Public Sub ApplyThemeToShape(ByVal shp As Shape, _
                              Optional ByVal fillColor As Long = -1, _
                              Optional ByVal lineColor As Long = -1)
    If fillColor = -1 Then fillColor = Gold
    shp.Fill.ForeColor.RGB = fillColor

    If lineColor = -1 Then
        shp.Line.Visible = msoFalse
    Else
        shp.Line.ForeColor.RGB = lineColor
        shp.Line.Visible       = msoCTrue
    End If
End Sub

' =============================================================================
' ApplyThemeToWorkbook — applies basic theme formatting to all framework sheets
' =============================================================================
Public Sub ApplyThemeToWorkbook()
    Dim sheetNames As Variant
    Dim s          As Variant
    Dim ws         As Worksheet

    sheetNames = Array("HOME", "SETTINGS", "REPORT_CONFIG", "UI_CONFIG", _
                       "DATA", "LOG", "HISTORY")

    For Each s In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(s))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Tab.Color = Charcoal
        End If
    Next s
End Sub
