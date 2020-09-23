VERSION 5.00
Begin VB.Form frmimage 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "frmimage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   3360
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   1230
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2280
      Pattern         =   "*.gif;*.jpg"
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   2760
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   3840
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   4635
      ScaleHeight     =   158
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   207
      TabIndex        =   1
      Top             =   255
      Width           =   3135
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2235
         Left            =   0
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   12
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   7080
      TabIndex        =   0
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label3 
      Caption         =   "Enter an alternative name, in case the image is not shown"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Width"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   240
      Left            =   4560
      TabIndex        =   6
      Top             =   2760
      Width           =   660
   End
End
Attribute VB_Name = "frmimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pth, im, ex, str As String

Private Sub Command2_Click()
On Error Resume Next
   Dim st As String
   Dim source, dest
     'Dim pth As String
     Dim wth As String
     Dim hth As String
     Dim al As String
     'Dim ex As String
     al = Text1.Text
     wth = Text2.Text
     hth = Text3.Text
     ex = """"
     im = "images/"
      st = GetSetting(App.Title, "Settings", "site", "")
      source = Dir1.Path & "\" & File1.filename
       dest = App.Path & "\" & st & "\images" & "\" & pth
         FileCopy source, dest
   str = "<img src = " & ex & im & pth & ex & " width =" & wth & " height = " & hth & " border = 0 alt = " & ex & al & ex & ">"
   frmMain.ActiveForm.txtText.SelItalic = False
   frmMain.ActiveForm.txtText.SelColor = &HC000C0
   frmMain.ActiveForm.txtText.SelText = str
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

Private Sub File1_Click()
On Error Resume Next
        With File1
        pth = .filename
        Picture2.Picture = LoadPicture(Dir1.Path & "\" & .filename)
        End With
        Text2.Text = Picture2.ScaleWidth
        Text3.Text = Picture2.ScaleHeight
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 30
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts an Image on the page. Press F1 for help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
      'Instruct (1)
End Sub
