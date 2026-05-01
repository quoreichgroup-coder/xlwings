Attribute VB_Name = "modSettings"
Option Explicit

' =============================================================================
' modSettings — Read / write / validate the SETTINGS sheet
' =============================================================================
' Sheet layout:  Column A = Key,  Column B = Value.  Row 1 = header.
'
' Environment profiles:
'   Add a key "Environment" with value DEV, UAT, or PROD.
'   Then add environment-specific overrides as "Key.DEV", "Key.UAT", "Key.PROD".
'   GetSetting("TCode") resolves to GetRawSetting("TCode.DEV") first when
'   Environment=DEV, falling back to GetRawSetting("TCode") if no override exists.
' =============================================================================

Private Const HEADER_ROW As Long = 1
Private Const KEY_COL    As Long = 1   ' A
Private Const VAL_COL    As Long = 2   ' B

Private ReadOnly Property RequiredKeys() As Variant
    RequiredKeys = Array( _
        "ProjectName", "ReportName", "Plant", "TCode", "Variant", "Layout", _
        "DateMode", "DateFrom", "DateTo", "ExportFileName", _
        "UI_MODE", "LOADER_TYPE", "UsePython", "PythonPath", "SAPConnectionName")
End Property

' =============================================================================
' GetSetting — environment-aware lookup.
' Tries "Key.ENV" first (where ENV = value of the "Environment" setting),
' then falls back to plain "Key".
' =============================================================================
Public Function GetSetting(ByVal key As String) As String
    Dim env As String
    env = UCase(Trim(GetRawSetting("Environment")))

    If env <> "" Then
        Dim envValue As String
        envValue = GetRawSetting(key & "." & env)
        If envValue <> "" Then
            GetSetting = envValue
            Exit Function
        End If
    End If

    GetSetting = GetRawSetting(key)
End Function

' =============================================================================
' GetRawSetting — direct lookup by key, no environment resolution.
' =============================================================================
Private Function GetRawSetting(ByVal key As String) As String
    Dim ws  As Worksheet
    Dim cel As Range
    Dim rng As Range

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_SETTINGS)
    Set rng = ws.Columns(KEY_COL)

    Set cel = rng.Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If Not cel Is Nothing Then
        GetRawSetting = Trim(CStr(ws.Cells(cel.Row, VAL_COL).Value))
    Else
        GetRawSetting = ""
    End If
    Exit Function
ErrHandler:
    GetRawSetting = ""
End Function

' =============================================================================
' SetSetting — writes a value back to the SETTINGS sheet.
' Creates the key as a new row if it does not exist.
' =============================================================================
Public Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim ws  As Worksheet
    Dim cel As Range
    Dim rng As Range

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_SETTINGS)
    Set rng = ws.Columns(KEY_COL)

    Set cel = rng.Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If Not cel Is Nothing Then
        ws.Cells(cel.Row, VAL_COL).Value = value
    Else
        Dim lLast As Long
        lLast = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row + 1
        ws.Cells(lLast, KEY_COL).Value = key
        ws.Cells(lLast, VAL_COL).Value = value
    End If
    Exit Sub
ErrHandler:
    Debug.Print "modSettings.SetSetting error: " & Err.Description
End Sub

' =============================================================================
' ValidateSettings — checks required keys and valid enum values.
' Returns True if valid.
' interactive: when False, errors are logged only (no MsgBox) for batch runs.
' =============================================================================
Public Function ValidateSettings(Optional ByVal interactive As Boolean = True) As Boolean
    Dim errors()  As String
    Dim errCount  As Long
    Dim k         As Variant
    Dim v         As String

    ReDim errors(0 To 20)
    errCount = 0

    For Each k In RequiredKeys
        v = GetSetting(CStr(k))
        If v = "" Then
            If CStr(k) = "PythonPath" And UCase(GetSetting("UsePython")) = "NO" Then
                ' PythonPath is not required when UsePython=NO
            Else
                errors(errCount) = "Missing or empty setting: " & CStr(k)
                errCount = errCount + 1
            End If
        End If
    Next k

    v = UCase(Trim(GetSetting("DateMode")))
    If v <> "CURRENT_WEEK" And v <> "PREVIOUS_WEEK" And v <> "CUSTOM" And _
       v <> "MONTH_TO_DATE" And v <> "YEAR_TO_DATE" Then
        errors(errCount) = "Invalid DateMode: " & v : errCount = errCount + 1
    End If

    v = UCase(Trim(GetSetting("UI_MODE")))
    If v <> "MANUAL" And v <> "AUTO" Then
        errors(errCount) = "Invalid UI_MODE: " & v : errCount = errCount + 1
    End If

    v = UCase(Trim(GetSetting("LOADER_TYPE")))
    If v <> "BAR" And v <> "CIRCLE" Then
        errors(errCount) = "Invalid LOADER_TYPE: " & v : errCount = errCount + 1
    End If

    v = UCase(Trim(GetSetting("UsePython")))
    If v <> "YES" And v <> "NO" Then
        errors(errCount) = "Invalid UsePython: " & v : errCount = errCount + 1
    End If

    ValidateSettings = (errCount = 0)

    If errCount > 0 Then
        Dim i   As Long
        Dim msg As String
        msg = "Settings errors found:" & vbNewLine
        For i = 0 To errCount - 1
            msg = msg & "  " & Chr(8226) & " " & errors(i) & vbNewLine
            modLog.LogMessage "ValidateSettings", "ERROR", errors(i)
        Next i
        If interactive And Not modMain.g_BatchMode Then
            MsgBox msg, vbCritical, "Settings Validation"
        End If
    End If
End Function

' =============================================================================
' DumpSettings — debug helper: prints all settings to the Immediate window
' =============================================================================
Public Sub DumpSettings()
    Dim ws   As Worksheet
    Dim i    As Long
    Dim last As Long

    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_SETTINGS)
    last = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row

    Dim env As String
    env = UCase(Trim(GetRawSetting("Environment")))
    If env <> "" Then Debug.Print "[ Environment: " & env & " ]"

    For i = HEADER_ROW + 1 To last
        Debug.Print ws.Cells(i, KEY_COL).Value & " = " & ws.Cells(i, VAL_COL).Value
    Next i
End Sub
