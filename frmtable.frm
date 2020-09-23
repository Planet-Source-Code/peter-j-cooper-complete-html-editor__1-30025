VERSION 5.00
Begin VB.Form frmtable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   Icon            =   "frmtable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Rows and Columns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   5055
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmtable.frx":0442
         Left            =   2640
         List            =   "frmtable.frx":0455
         TabIndex        =   14
         Text            =   "1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmtable.frx":0468
         Left            =   2640
         List            =   "frmtable.frx":047B
         TabIndex        =   13
         Text            =   "1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Rows"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Columns"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Table Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "100"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "100"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "As percentage of page"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In Pixels"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Table Height"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Table Width"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   3360
      Width           =   1140
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Uncheck this box if you don't want a border round your table"
      Height          =   450
      Left            =   585
      TabIndex        =   1
      Top             =   135
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   1170
   End
End
Attribute VB_Name = "frmtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
         
    Dim i
    Dim j
    i = Combo1.Text
    j = Combo2.Text
    
  If i > 5 Or j > 5 Then
    MsgBox "Please set number of columns and rows to 5 or less"
  Else
    fillTable
    Unload Me
  End If
End Sub

Private Sub Command2_Click()
       Unload Me
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 32
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts a Table. Press F1 for help"
End Sub

