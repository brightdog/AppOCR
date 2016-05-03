VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "AppOcr"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   12870
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5415
      Left            =   8340
      ScaleHeight     =   5355
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Timer tmrADB 
      Interval        =   100
      Left            =   7260
      Top             =   3480
   End
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   5340
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.ListBox lstLog 
      Height          =   1680
      Left            =   5460
      TabIndex        =   0
      Top             =   7380
      Width           =   6675
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   8955
      Left            =   180
      Stretch         =   -1  'True
      Top             =   180
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrADBpath As String
Dim mbolDeviceConnected As Boolean
Dim pngClass As New LoadPNG
Private Sub Form_Load()
Set pngClass = New LoadPNG
mstrADBpath = App.Path & "\platform-tools\"

Me.Show

Dim strDos As String
strDos = DosPrint(mstrADBpath & "adb devices -l", True, True)
Debug.Print strDos

    Dim strPattern As String
    
    strPattern = "([^\s]+)\s+device product:.*? model:(.*?) device:.*?"
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    Set Mc = RegExecute(strDos, strPattern)
    
    If Mc.Count > 0 Then
    
        Dim i As Integer
        Me.cboDevice.Clear
        For i = 0 To Mc.Count - 1
        
            Me.cboDevice.AddItem Mc.Item(i).SubMatches(1) & "--" & Mc.Item(i).SubMatches(0)
        
        Next
        
        Me.cboDevice.ListIndex = 0
        
    End If
    
End Sub

Private Sub tmrADB_Timer()
    If InStr(1, DosPrint(mstrADBpath & "adb get-state", True, True), "device", vbBinaryCompare) > 0 Then
        mbolDeviceConnected = True
    Else
        mbolDeviceConnected = False
    End If
    
    If mbolDeviceConnected Then
        Dim strTime As String
        Dim strTimeStart As String
        Dim strCommand As String
        Debug.Print "^^^"
        strTime = Now
        strTimeStart = strTime
        Debug.Print strTime
        strCommand = mstrADBpath & "adb shell screencap  /sdcard/Screen.raw"
        
        Call DosPrint(strCommand & vbCrLf, True, True)
        
        Debug.Print "½ØÍ¼£º" & DateDiff("s", strTime, Now)
        strTime = Now
        strCommand = mstrADBpath & "adb shell gzip /sdcard/screen.raw /sdcard/"
        
        Call DosPrint(strCommand & vbCrLf, True, True)
        Debug.Print "Ñ¹Ëõ£º" & DateDiff("s", strTime, Now)
        strTime = Now
        strCommand = mstrADBpath & "adb pull /sdcard/Screen.raw.gz"
        
        Call DosPrint(strCommand & vbCrLf, True, True)
        Debug.Print "´«Êä£º" & DateDiff("s", strTime, Now)
        strTime = Now
        'strCommand = mstrADBpath & "adb shell rm /sdcard/screen.raw.gz"
        
        'Call DosPrint(strCommand & vbCrLf, True, True)
        
       ' Debug.Print "É¾³ý£º" & DateDiff("s", strTime, Now)
        'strTime = Now
        Debug.Print "×Ü¼Æ£º"; DateDiff("s", strTimeStart, Now)
        'Call LoadPic(App.Path & "\screen.png")
        'Debug.Print Now
        'img.Picture = picScreen.Picture
        Debug.Print Now
        Debug.Print "$$$"
    End If
End Sub
Private Sub LoadPic(ByVal picName As String)

    

    If LCase(Right(picName, 3)) = "png" Then
    
        pngClass.PicBox = picScreen 'or Picturebox
        'pngClass.SetToBkgrnd True, 0, 0 'set to Background (True or false), x and y
        'pngClass.BackgroundPicture = Form1 'same Backgroundpicture
        'pngClass.SetAlpha = True 'when Alpha then alpha
        'pngClass.SetTrans = True 'when transparent Color then transparent Color
        pngClass.OpenPNG picName 'Open and display Picture
        
    Else
        picScreen.Picture = VB.LoadPicture(picName)
           
    End If

    'picOutput.Width = picSource.Width
End Sub
