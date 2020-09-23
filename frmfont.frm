VERSION 5.00
Begin VB.Form frmfont 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Formatting"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmfont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2475
      TabIndex        =   8
      Top             =   1680
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2835
      ScaleHeight     =   255
      ScaleWidth      =   885
      TabIndex        =   7
      Top             =   1215
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1185
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmfont.frx":044A
      Left            =   1680
      List            =   "frmfont.frx":047E
      TabIndex        =   2
      Text            =   "aqua"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmfont.frx":04F0
      Left            =   1680
      List            =   "frmfont.frx":0503
      TabIndex        =   1
      Text            =   " 1"
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmfont.frx":0516
      Left            =   1680
      List            =   "frmfont.frx":0523
      TabIndex        =   0
      Text            =   "Helvetica,Arial,sans serif"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Font Color"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Font Family"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmfont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo3_Click()
         Select Case Combo3.Text
            Case Is = "black"
              Picture2.BackColor = &H0
            Case Is = "blue"
              Picture2.BackColor = &HFF0000
            Case Is = "fuchsia"
              Picture2.BackColor = &HFF00FF
            Case Is = "gray"
              Picture2.BackColor = &H808080
            Case Is = "green"
              Picture2.BackColor = &HC000&
            Case Is = "lime"
              Picture2.BackColor = &H80FF80
            Case Is = "maroon"
              Picture2.BackColor = &H80&
            Case Is = "navy"
              Picture2.BackColor = &HC00000
            Case Is = "olive"
              Picture2.BackColor = &H8000&
            Case Is = "purple"
              Picture2.BackColor = &H800080
            Case Is = "red"
              Picture2.BackColor = &HFF&
            Case Is = "silver"
              Picture2.BackColor = &HC0C0C0
            Case Is = "teal"
              Picture2.BackColor = &H808000
            Case Is = "yellow"
              Picture2.BackColor = &HFFFF&
            Case Is = "aqua"
              Picture2.BackColor = &HFFFF00
            Case Is = "white"
              Picture2.BackColor = &HFFFFFF
         End Select
End Sub

Private Sub Command1_Click()
On Error Resume Next
   Dim wes As String
   Dim wis As String
   Dim myst As String
   Dim fc
   Dim clr
   Dim sz
   Dim qt
   wes = frmMain.ActiveForm.txtText.SelText
   fc = Combo1.Text
   clr = Combo3.Text
   sz = Combo2.Text
   qt = """"
   frmMain.ActiveForm.txtText.SelItalic = False
   frmMain.ActiveForm.txtText.SelText = _
   "<font size = " & sz & " color = " & qt & clr & qt & " face = " _
   & qt & fc & qt & ">" & wes & "</font>"
   myst = "<font size = " & sz & " color = " & qt & clr & qt & " face = " _
   & qt & fc & qt & ">"
   wis = "<font size = " & sz & " color = " & qt & clr & qt & " face = " _
   & qt & fc & qt & ">" & wes & "</font>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8080&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 7)
        frmMain.ActiveForm.txtText.SelLength = 7
        frmMain.ActiveForm.txtText.SelColor = &H8080&
   Unload Me
End Sub

Private Sub Command2_Click()
      Unload Me
End Sub

Private Sub Form_Load()
          With Combo3
             .AddItem "aqua"
             .AddItem "black"
             .AddItem "blue"
             .AddItem "fuchsia"
             .AddItem "gray"
             .AddItem "green"
             .AddItem "lime"
             .AddItem "maroon"
             .AddItem "navy"
             .AddItem "olive"
             .AddItem "purple"
             .AddItem "red"
             .AddItem "silver"
             .AddItem "teal"
             .AddItem "yellow"
             .AddItem "white"
             .Text = "aqua"
          End With
          Picture2.BackColor = &HFFFF00
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 38
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Formats your font family,size and color. Press F1 for help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Instruct (1)
End Sub
