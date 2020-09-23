VERSION 5.00
Begin VB.Form frmmedia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Shockwave Flash"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmmedia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check this box if you want your movie to loop continually"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm Selection"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   120
      Pattern         =   "*.swf"
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Select your movie file by double clicking items in the boxes below. Only files of .swf will be shown"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "File Name"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmmedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
   Dim st As String
   Dim source, dest
      st = GetSetting(App.Title, "Settings", "site", "")
      source = Dir1.Path & "\" & File1.filename
       dest = App.Path & "\" & st & "\images" & "\" & File1.filename
         FileCopy source, dest
        With File1
        Text1.Text = .filename
        End With
End Sub

Private Sub Command2_Click()
       Dim ex As String
       Dim fnm As String
       Dim lp As String
           ex = """"
           fnm = Text1.Text
      If Check1.Value = 1 Then
         lp = "true"
      Else
         lp = "false"
      End If
  frmMain.ActiveForm.txtText.SelItalic = False
  frmMain.ActiveForm.txtText.SelColor = &HC000C0
  frmMain.ActiveForm.txtText.SelText = _
    "<object classid= " & ex & "clsid : D27CDBE - AE6D - 11cf - 96B8 - 444553540000 " & ex & _
    "codebase= " & ex & "http://www.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version= 5,0,0,0" & ex & "width= " & ex & ex & " height= " & ex & ex & ">" & vbCrLf & _
    "<param name = movie value = " & ex & "images/" & fnm & ex & ">" & vbCrLf & _
    "<param name = quality value = high>" & vbCrLf & _
    "<param name = " & ex & "LOOP" & ex & " value = " & ex & lp & ex & ">" & vbCrLf & _
    "<embed src = " & ex & "images/" & fnm & ex & "quality = high pluginspage = " & ex & _
    "http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_version=ShockwaveFlash" & ex & _
    "type=" & ex & "application/x-shockwave-flash" & ex & "width= " & ex & ex & " height = " & ex & ex & " loop=" & ex & lp & ex & ">" & vbCrLf & _
    "</embed>" & vbCrLf & "</object>"
    Unload Me
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveHandler
    Dir1.Path = Drive1.Drive
    Exit Sub

DriveHandler:
       Drive1.Drive = Dir1.Path
    Exit Sub
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 43
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts a Flash movie. Press F1 for help"
End Sub
