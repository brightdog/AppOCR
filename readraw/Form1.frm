VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1020
      TabIndex        =   0
      Top             =   1020
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim tmr As String

    tmr = VBA.Now()
    Dim iFile As Integer

    iFile = VBA.FreeFile()
    Dim InLine() As Byte
    Dim InHeight() As Byte
    ReDim InHeight(1920)
    Dim i As Long
    i = 0
    Open App.Path & "\screen.raw" For Binary Access Read As #iFile
    ReDim InLine(8294412)

    Get #iFile, , InLine

    Close #iFile
    Dim arrResult() As String
    ReDim arrResult(2073600)
    Dim j As Long
    j = 0
    
    For i = 16 To UBound(InLine) - 1 Step 8
    
        arrResult(j) = RGB(InLine(i), InLine(i + 1), InLine(i + 2))

        j = j + 1
    Next

    Dim sb As clsStringBuilder
    Set sb = New clsStringBuilder

    For j = 0 To UBound(arrResult)

        If arrResult(j) <> "" Then
            If arrResult(j) > 328965 And arrResult(j) < 13158600 Then
                sb.Append "*"
            Else
                sb.Append " "
            End If

            If j Mod 1080 / 2 = 0 Then
                sb.Append vbCrLf
                j = j + 1080
            End If
        End If

    Next

    iFile = VBA.FreeFile
    Open App.Path & "\result.txt" For Output As #iFile

    Print #iFile, sb.ToString

    Close #iFile

    Debug.Print DateDiff("s", tmr, VBA.Now)
End Sub

