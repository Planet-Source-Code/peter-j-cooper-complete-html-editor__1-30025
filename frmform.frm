VERSION 5.00
Begin VB.Form frmform 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Elements"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmform.frx":0442
      Left            =   2280
      List            =   "frmform.frx":0461
      TabIndex        =   1
      Text            =   "Text Box"
      Top             =   240
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Form Element"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
   FmtFm
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
   frmMain.sbStatusBar.Panels(1).Text = "Inserts Form elements. Press F1 for help"
End Sub
