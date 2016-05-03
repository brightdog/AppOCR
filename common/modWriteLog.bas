Attribute VB_Name = "modWriteLog"
Option Explicit
Public strLogFileName As String
Public gstrLogFileName As String
Public Sub WriteLog(ByVal Str As String, Optional ForceWrite As Boolean = False, Optional ByVal LogFileName As String = "", Optional ByRef lstLog As VB.ListBox)
        '<EhHeader>
        On Error GoTo WriteLog_Err

        '</EhHeader>
100     If LogFileName = "" Then
102         If gstrLogFileName = "" Then
104             strLogFileName = "Log_xx_" & Format(Date, "YYYY-MM-DD") & ".txt"
            Else
106             strLogFileName = gstrLogFileName
            End If

        Else
108         strLogFileName = LogFileName
        End If
        
110     If lstLog Is Nothing Then
        
112         Set lstLog = frmMain.lstLog
        
        End If
        
        Dim i As Integer

114     If Not lstLog Is Nothing Then

116         With lstLog

118             If .ListCount > 100 Then
120                 .Visible = False

122                 For i = .ListCount To 30 Step -1

124                     .RemoveItem i - 1

126                     DoEvents

                    Next

128                 .Visible = True
                End If

130             lstLog.AddItem Str & "<-- " & VBA.Now(), 0

132             DoEvents
            End With

        End If

        '.Refresh
        Dim bolNeedWriteLog As Boolean
134     bolNeedWriteLog = ForceWrite '默认是都要纪录的。
        

        
140     If bolNeedWriteLog Then
            Dim iFile As Integer
        
142         iFile = VBA.FreeFile()
144         Open App.path & "\log\" & strLogFileName For Append As #iFile

146         Print #iFile, Str & "<-- " & Now()

148         Close #iFile
            '.Visible = True

150         Call ClearHistoryFile(App.path & "\log\", CLng(20 * 60 * 24))
        
        End If
        
        '<EhFooter>
        Exit Sub

WriteLog_Err:
        
        If Err.Number = 76 Then
            '路径未找到
            Err.Clear
            Dim Fso As Scripting.FileSystemObject
            Set Fso = New Scripting.FileSystemObject
            
            If Not Fso.FolderExists(App.path & "\log") Then
                Call Fso.CreateFolder(App.path & "\log")
            
            Else
                
            End If
        
            Set Fso = Nothing
        Else
            Err.Clear

            iFile = VBA.FreeFile()
            Open App.path & "\ERRLOG.txt" For Append As #iFile

            Print #iFile, "=================================================="
            Print #iFile, Err.Description & vbCrLf; Err.Number & "<-- " & Now()
            Print #iFile, Str & "<-- " & Now()
            Print #iFile, "--------------------------------------------------"

            Close #iFile
        
        End If

        '</EhFooter>
End Sub

