Attribute VB_Name = "modSAPNavigation"
Option Explicit

' =============================================================================
' modSAPNavigation — Navigate SAP transactions, variants, layouts, dates
' =============================================================================

Private Const WAIT_MS As Long = 500   ' default wait between GUI actions (ms)

' =============================================================================
' OpenTransaction — navigates to a transaction via the command field
' =============================================================================
Public Sub OpenTransaction(ByVal tcode As String)
    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    ' Use /n prefix to navigate without opening a new session
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n" & tcode
    session.FindById("wnd[0]").SendVKey 0   ' Enter
    WaitMS WAIT_MS
    Exit Sub
ErrHandler:
    modLog.LogMessage "OpenTransaction", "ERROR", "TCode=" & tcode & " — " & Err.Description
End Sub

' =============================================================================
' ApplyVariant — selects a selection screen variant
' =============================================================================
Public Sub ApplyVariant(ByVal variantName As String)
    If variantName = "" Then Exit Sub

    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    ' Standard toolbar: Goto > Variants > Get Variant
    session.FindById("wnd[0]/mbar/menu[3]/menu[1]/menu[0]").Select
    WaitMS WAIT_MS

    ' Type variant name in the popup and confirm
    Dim variantBox As Object
    On Error Resume Next
    Set variantBox = session.FindById("wnd[1]/usr/txtV-LOW")
    On Error GoTo ErrHandler
    If Not variantBox Is Nothing Then
        variantBox.Text = variantName
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press  ' Copy button
        WaitMS WAIT_MS
    End If
    Exit Sub
ErrHandler:
    modLog.LogMessage "ApplyVariant", "ERROR", variantName & " — " & Err.Description
End Sub

' =============================================================================
' ApplyLayout — selects an ALV layout (Change Layout popup)
' =============================================================================
Public Sub ApplyLayout(ByVal layoutName As String)
    If layoutName = "" Then Exit Sub

    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    ' Settings > Layout > Select (Ctrl+F8 on most reports)
    session.FindById("wnd[0]").SendVKey 67   ' F8 in some reports; adjust per TCode
    WaitMS WAIT_MS

    Dim layoutBox As Object
    On Error Resume Next
    Set layoutBox = session.FindById("wnd[1]/usr/txtLT-VARIANT")
    On Error GoTo ErrHandler

    If Not layoutBox Is Nothing Then
        layoutBox.Text = layoutName
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press
        WaitMS WAIT_MS
    End If
    Exit Sub
ErrHandler:
    modLog.LogMessage "ApplyLayout", "ERROR", layoutName & " — " & Err.Description
End Sub

' =============================================================================
' SetDateRange — fills the standard date From / To fields on a selection screen
' Adjust the field IDs to match the specific TCode's selection screen.
' =============================================================================
Public Sub SetDateRange(ByVal dateFrom As String, ByVal dateTo As String)
    Dim session As Object
    Set session = modSAPConnection.GetSAPSession()
    If session Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    ' These IDs are typical but must be verified for each transaction.
    ' Common patterns: P_DATUM, S_DATUM, PA_DATUM ...
    Dim dateFromField As Object
    Dim dateToField   As Object

    On Error Resume Next
    ' Try common field name patterns
    Set dateFromField = session.FindById("wnd[0]/usr/ctxtS_DATUM-LOW")
    If dateFromField Is Nothing Then
        Set dateFromField = session.FindById("wnd[0]/usr/ctxtP_DATUM")
    End If

    Set dateToField = session.FindById("wnd[0]/usr/ctxtS_DATUM-HIGH")
    On Error GoTo ErrHandler

    If Not dateFromField Is Nothing Then
        dateFromField.Text = dateFrom
    End If
    If Not dateToField Is Nothing Then
        dateToField.Text = dateTo
    End If

    WaitMS WAIT_MS
    Exit Sub
ErrHandler:
    modLog.LogMessage "SetDateRange", "ERROR", dateFrom & "–" & dateTo & " — " & Err.Description
End Sub

' =============================================================================
' WaitMS — pauses execution for a given number of milliseconds
' =============================================================================
Public Sub WaitMS(ByVal ms As Long)
    Dim tEnd As Single
    tEnd = Timer + (ms / 1000!)
    Do While Timer < tEnd
        DoEvents
    Loop
End Sub
