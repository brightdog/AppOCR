Attribute VB_Name = "modClearHistoryFile"
Option Explicit



Public Sub ClearHistoryFile(ByVal strPath As String, ByVal intMinutes As Long)

        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject
        'Dim strPath As String

        'strPath = App.Path & "\SQL"
102     If Fso.FolderExists(strPath) Then
    
104         Call DelHistoryFile_Core(Fso.GetFolder(strPath), intMinutes)
        Else
    
106         Call WriteLog("*Path Not Exist @ " & strPath)
        End If
    
        'strPath = App.Path & "\log"
        'If FSo.FolderExists(strPath) Then
    
        '    Call DelHistoryFile_Core(FSo.GetFolder(strPath), intDays)
        '
        'End If
End Sub

Private Sub DelHistoryFile_Core(ByRef Fld As Scripting.Folder, ByVal intMinutes As Long)
        '<EhHeader>
        On Error GoTo DelHistoryFile_Core_Err
        '</EhHeader>

        Dim dtNowDate As Date
100     dtNowDate = VBA.Date()
    
        Dim iFile As Scripting.File

102     If Fld.Files.Count > 0 Then

104         For Each iFile In Fld.Files
    
                Dim dtFileDate As Date
106             dtFileDate = iFile.DateCreated
    
108             If DateDiff("n", dtFileDate, dtNowDate) > intMinutes Then
        
110                 Call iFile.Delete(True)
            
                End If

112             DoEvents
            Next

        Else
114         Call WriteLog("*Files.Count <= 0 @ " & Fld.Path)
        End If

        '<EhFooter>
        Exit Sub

DelHistoryFile_Core_Err:
        WriteLog Err.Description & vbCrLf & _
           "in DelHistoryFile_Core " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

