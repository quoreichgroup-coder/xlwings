Attribute VB_Name = "modDevSync"
Option Explicit

' =============================================================================
' modDevSync — Export / import VBA modules as .bas files for VS Code editing
' Requires: Trust access to the VBA project object model
'   (File > Options > Trust Center > Trust Center Settings >
'    Macro Settings > check "Trust access to the VBA project object model")
' =============================================================================

Private Const VBA_DIR As String = "vba"   ' folder relative to workbook

' =============================================================================
' ExportVBAModules — saves every standard module as a .bas file
' =============================================================================
Public Sub ExportVBAModules()
    Dim vbComp  As Object   ' VBIDE.VBComponent
    Dim sDir    As String
    Dim sFile   As String
    Dim count   As Long

    sDir = ThisWorkbook.Path & "\" & VBA_DIR
    modFolders.EnsureFolder sDir

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' 1 = vbext_ct_StdModule, 2 = vbext_ct_ClassModule, 3 = vbext_ct_MSForm
        If vbComp.Type = 1 Then
            sFile = sDir & "\" & vbComp.Name & ".bas"
            vbComp.Export sFile
            count = count + 1
        End If
    Next vbComp

    MsgBox count & " module(s) exported to:" & vbNewLine & sDir, _
           vbInformation, "Export VBA Modules"
End Sub

' =============================================================================
' ImportVBAModules — imports .bas files from the vba folder into the workbook
' Existing modules with the same name are removed before import.
' =============================================================================
Public Sub ImportVBAModules()
    Dim sDir  As String
    Dim sFile As String
    Dim count As Long

    sDir  = ThisWorkbook.Path & "\" & VBA_DIR
    sFile = Dir(sDir & "\*.bas")

    Do While sFile <> ""
        Dim modName As String
        modName = Left(sFile, Len(sFile) - 4)   ' strip .bas

        ' Remove existing module if present
        On Error Resume Next
        Dim existing As Object
        Set existing = ThisWorkbook.VBProject.VBComponents(modName)
        If Not existing Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove existing
        End If
        On Error GoTo 0

        ThisWorkbook.VBProject.VBComponents.Import sDir & "\" & sFile
        count = count + 1
        sFile = Dir
    Loop

    MsgBox count & " module(s) imported from:" & vbNewLine & sDir, _
           vbInformation, "Import VBA Modules"
End Sub

' =============================================================================
' SyncVBAModules — exports first, then re-imports (round-trip consistency check)
' =============================================================================
Public Sub SyncVBAModules()
    ExportVBAModules
    ImportVBAModules
    MsgBox "Sync complete.", vbInformation, "Sync VBA Modules"
End Sub

' =============================================================================
' OpenProjectInVSCode — opens the vba folder in VS Code
' VS Code must be on the system PATH (code.cmd).
' =============================================================================
Public Sub OpenProjectInVSCode()
    Dim sDir As String
    sDir = ThisWorkbook.Path & "\" & VBA_DIR

    Dim sCmd As String
    sCmd = "cmd /c code """ & sDir & """"
    Shell sCmd, vbHide

    MsgBox "Opening VS Code at:" & vbNewLine & sDir & vbNewLine & vbNewLine & _
           "Tip: edit .bas files in VS Code, then run ImportVBAModules to reload.", _
           vbInformation, "Open in VS Code"
End Sub

' =============================================================================
' ExportCurrentModule — exports only the module that calls this sub
' (useful bound to a keyboard shortcut inside the VBE)
' =============================================================================
Public Sub ExportCurrentModule()
    Dim vbComp  As Object
    Dim sDir    As String
    Dim sFile   As String

    ' Identify the active component in the VBE
    On Error Resume Next
    Set vbComp = Application.VBE.ActiveCodePane.CodeModule.Parent
    On Error GoTo 0

    If vbComp Is Nothing Then
        MsgBox "No active module found.", vbExclamation
        Exit Sub
    End If

    sDir  = ThisWorkbook.Path & "\" & VBA_DIR
    modFolders.EnsureFolder sDir
    sFile = sDir & "\" & vbComp.Name & ".bas"
    vbComp.Export sFile
    MsgBox "Exported: " & sFile, vbInformation, "Export Module"
End Sub
