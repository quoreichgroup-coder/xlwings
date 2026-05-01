Attribute VB_Name = "modSAPConnection"
Option Explicit

' =============================================================================
' modSAPConnection — SAP GUI Scripting connection management
' Requires SAP GUI Scripting API to be enabled on the SAP system and client.
' =============================================================================

Public g_SAPSession As Object   ' GuiSession object (reused across modules)

' =============================================================================
' ConnectToSAP — obtains a SAP GUI session and stores it in g_SAPSession
' Returns True on success, False on failure.
' =============================================================================
Public Function ConnectToSAP() As Boolean
    Dim sapGUI   As Object
    Dim sapApp   As Object
    Dim conn     As Object
    Dim connName As String

    On Error GoTo ErrHandler
    ConnectToSAP = False

    ' SAP GUI must already be running with an open connection
    Set sapGUI = GetObject("SAPGUI")
    If sapGUI Is Nothing Then
        MsgBox "SAP GUI is not running. Please log in to SAP first.", _
               vbCritical, "SAP Connection"
        Exit Function
    End If

    Set sapApp = sapGUI.GetScriptingEngine()

    ' Try to use named connection from settings; fall back to first connection
    connName = modSettings.GetSetting("SAPConnectionName")
    Dim i As Integer

    If connName <> "" Then
        For i = 0 To sapApp.Connections.Count - 1
            If InStr(1, sapApp.Connections(i).Description, connName, vbTextCompare) > 0 Then
                Set conn = sapApp.Connections(i)
                Exit For
            End If
        Next i
    End If

    If conn Is Nothing Then
        If sapApp.Connections.Count = 0 Then
            MsgBox "No active SAP connections found.", vbCritical, "SAP Connection"
            Exit Function
        End If
        Set conn = sapApp.Connections(0)
    End If

    Set g_SAPSession = conn.Sessions(0)
    ConnectToSAP = True
    Exit Function

ErrHandler:
    MsgBox "Error connecting to SAP: " & Err.Description, vbCritical, "SAP Connection"
    ConnectToSAP = False
End Function

' =============================================================================
' GetSAPSession — returns the cached session (connects if not yet connected)
' =============================================================================
Public Function GetSAPSession() As Object
    If g_SAPSession Is Nothing Then
        If Not ConnectToSAP() Then
            Set GetSAPSession = Nothing
            Exit Function
        End If
    End If
    Set GetSAPSession = g_SAPSession
End Function

' =============================================================================
' CheckSAPSystem — verifies the connected system matches expectations
' =============================================================================
Public Function CheckSAPSystem() As Boolean
    Dim session As Object
    Set session = GetSAPSession()
    If session Is Nothing Then
        CheckSAPSystem = False
        Exit Function
    End If

    On Error Resume Next
    Dim sysInfo As Object
    Set sysInfo = session.FindById("wnd[0]/sbar")
    On Error GoTo 0

    CheckSAPSystem = Not (session Is Nothing)
End Function

' =============================================================================
' DisconnectSAP — releases the cached session reference
' =============================================================================
Public Sub DisconnectSAP()
    Set g_SAPSession = Nothing
End Sub
