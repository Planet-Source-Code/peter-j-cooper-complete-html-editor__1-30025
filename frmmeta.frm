VERSION 5.00
Begin VB.Form frmmeta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Meta"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmmeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   885
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Meta Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Meta Keywords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Enter a short description of your site"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter keywords for your site as many as you like, seperated by comma's"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmmeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
       Dim ex As String
       Dim ky As String
       Dim ds As String
           ex = """"
           ky = Text1.Text
           ds = Text2.Text
 frmMain.ActiveForm.txtText.SelItalic = False
 frmMain.ActiveForm.txtText.SelColor = vbBlue
 frmMain.ActiveForm.txtText.SelText = _
  "<meta name = " & ex & "keywords" & ex & " content = " & ex & ky & ex & ">" & vbCrLf & _
  "<meta name = " & ex & "description" & ex & " content = " & ex & ds & ex & ">"
  Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 44
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts meta Information. Press F1 for help"
End Sub
