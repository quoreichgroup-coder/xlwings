Attribute VB_Name = "modDates"
Option Explicit

' =============================================================================
' modDates — Date resolution based on DateMode setting
' =============================================================================

' Resolved dates (written back to SETTINGS after resolution)
Public g_DateFrom As String
Public g_DateTo   As String

' =============================================================================
' ResolveDates — reads DateMode and populates DateFrom / DateTo
' =============================================================================
Public Sub ResolveDates()
    Dim mode     As String
    Dim dFrom    As Date
    Dim dTo      As Date

    mode = UCase(Trim(modSettings.GetSetting("DateMode")))

    Select Case mode
        Case "CURRENT_WEEK"
            dFrom = GetCurrentWeekMonday()
            dTo   = GetCurrentWeekSunday()

        Case "PREVIOUS_WEEK"
            dFrom = GetPreviousWeekMonday()
            dTo   = GetPreviousWeekSunday()

        Case "MONTH_TO_DATE"
            dFrom = DateSerial(Year(Date), Month(Date), 1)
            dTo   = Date

        Case "YEAR_TO_DATE"
            dFrom = DateSerial(Year(Date), 1, 1)
            dTo   = Date

        Case "CUSTOM"
            ' Use whatever the analyst entered in the SETTINGS sheet
            g_DateFrom = modSettings.GetSetting("DateFrom")
            g_DateTo   = modSettings.GetSetting("DateTo")
            Exit Sub

        Case Else
            Err.Raise vbObjectError + 1001, "modDates.ResolveDates", _
                      "Unknown DateMode: " & mode
    End Select

    g_DateFrom = FormatSAPDate(dFrom)
    g_DateTo   = FormatSAPDate(dTo)

    ' Write resolved dates back so analyst can see them
    modSettings.SetSetting "DateFrom", g_DateFrom
    modSettings.SetSetting "DateTo",   g_DateTo
End Sub

' =============================================================================
' Week helpers — week starts on Monday (vbMonday)
' =============================================================================
Public Function GetCurrentWeekMonday() As Date
    Dim today As Date
    today = Date
    GetCurrentWeekMonday = today - (Weekday(today, vbMonday) - 1)
End Function

Public Function GetCurrentWeekSunday() As Date
    GetCurrentWeekSunday = GetCurrentWeekMonday() + 6
End Function

Public Function GetPreviousWeekMonday() As Date
    GetPreviousWeekMonday = GetCurrentWeekMonday() - 7
End Function

Public Function GetPreviousWeekSunday() As Date
    GetPreviousWeekSunday = GetCurrentWeekMonday() - 1
End Function

' =============================================================================
' FormatSAPDate — returns DD.MM.YYYY (SAP GUI field format)
' =============================================================================
Public Function FormatSAPDate(ByVal d As Date) As String
    FormatSAPDate = Format(d, "DD.MM.YYYY")
End Function

' =============================================================================
' ParseSAPDate — converts DD.MM.YYYY string back to a VBA Date
' =============================================================================
Public Function ParseSAPDate(ByVal s As String) As Date
    Dim parts() As String
    parts = Split(s, ".")
    If UBound(parts) = 2 Then
        ParseSAPDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    Else
        ParseSAPDate = CDate(s)
    End If
End Function

' =============================================================================
' GetMonthRange — returns first and last day of a given month/year
' =============================================================================
Public Sub GetMonthRange(ByVal yr As Integer, ByVal mo As Integer, _
                         ByRef dFrom As Date, ByRef dTo As Date)
    dFrom = DateSerial(yr, mo, 1)
    dTo   = DateSerial(yr, mo + 1, 0)
End Sub
