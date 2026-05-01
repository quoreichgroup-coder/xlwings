Attribute VB_Name = "modFolders"
Option Explicit

' =============================================================================
' modFolders — Manages the project folder structure
' All paths are relative to the workbook's parent folder.
' =============================================================================

Private Const EXPORTS_DIR As String = "Exports"
Private Const ARCHIVE_DIR As String = "Archive"
Private Const LOGS_DIR    As String = "Logs"

' =============================================================================
' PrepareFolders — creates all required sub-folders if they do not exist
' =============================================================================
Public Sub PrepareFolders()
    EnsureFolder GetExportsPath()
    EnsureFolder GetArchivePath()
    EnsureFolder GetLogsPath()
End Sub

' =============================================================================
' Path getters
' =============================================================================
Public Function GetExportsPath() As String
    GetExportsPath = ThisWorkbook.Path & "\" & EXPORTS_DIR
End Function

Public Function GetArchivePath() As String
    GetArchivePath = ThisWorkbook.Path & "\" & ARCHIVE_DIR
End Function

Public Function GetLogsPath() As String
    GetLogsPath = ThisWorkbook.Path & "\" & LOGS_DIR
End Function

' =============================================================================
' EnsureFolder — creates a folder (including nested) if it does not exist
' =============================================================================
Public Sub EnsureFolder(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

' =============================================================================
' ArchiveFile — moves a file from Exports to Archive with a timestamp suffix
' =============================================================================
Public Sub ArchiveFile(ByVal filePath As String)
    If Dir(filePath) = "" Then Exit Sub

    Dim fName    As String
    Dim ext      As String
    Dim baseName As String
    Dim dest     As String

    fName    = Mid(filePath, InStrRev(filePath, "\") + 1)
    ext      = Mid(fName, InStrRev(fName, "."))
    baseName = Left(fName, Len(fName) - Len(ext))
    dest     = GetArchivePath() & "\" & baseName & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ext

    FileCopy filePath, dest
    Kill filePath
End Sub

' =============================================================================
' CleanExportsFolder — removes files older than n days from Exports
' =============================================================================
Public Sub CleanExportsFolder(Optional ByVal daysOld As Integer = 7)
    Dim sPath As String
    Dim sFile As String
    sPath = GetExportsPath() & "\"
    sFile = Dir(sPath & "*.*")

    Do While sFile <> ""
        If FileDateTime(sPath & sFile) < Now - daysOld Then
            Kill sPath & sFile
        End If
        sFile = Dir
    Loop
End Sub
