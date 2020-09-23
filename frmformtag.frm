VERSION 5.00
Begin VB.Form frmformtag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Form Tags"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmformtag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2565
      TabIndex        =   5
      Top             =   2160
      Width           =   1170
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Do you want to submit via Email? Check the box and add your Email address to the box below"
      Height          =   615
      Left            =   975
      TabIndex        =   3
      Top             =   960
      Width           =   3465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   285
      Left            =   1185
      TabIndex        =   2
      Top             =   2160
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Your script here"
      Top             =   1680
      Width           =   3840
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmformtag.frx":08CA
      Left            =   2625
      List            =   "frmformtag.frx":08D4
      TabIndex        =   0
      Text            =   "post"
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Select Method"
      Height          =   255
      Left            =   1335
      TabIndex        =   4
      Top             =   345
      Width           =   1335
   End
End
Attribute VB_Name = "frmformtag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
      Text1.Text = "mailto:"
    Else
      Text1.Text = "Your script here"
    End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
      Dim prev As String
      Dim wes As String
      Dim myst As String
      Dim meth As String
      Dim act As String
      Dim ax As String
      
      prev = frmMain.ActiveForm.txtText.SelText
      meth = Combo1.Text
      act = Text1.Text
      ax = """"
      frmMain.ActiveForm.txtText.SelItalic = False
      frmMain.ActiveForm.txtText.SelText = _
      "<form method = " & ax & meth & ax & " action = " & ax & act & ax & ">" & vbCrLf & prev & vbCrLf & "</form>"
      myst = "<form method = " & ax & meth & ax & " action = " & ax & act & ax & ">"
      wes = "<form method = " & ax & meth & ax & " action = " & ax & act & ax & ">" & vbCrLf & prev & vbCrLf & "</form>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = vbRed
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 7)
        frmMain.ActiveForm.txtText.SelLength = 7
        frmMain.ActiveForm.txtText.SelColor = vbRed
      Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 33
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts Form Tags. Press F1 for help"
End Sub
