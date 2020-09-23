VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   120
      Picture         =   "frmsplash.frx":0000
      ScaleHeight     =   2205
      ScaleWidth      =   3150
      TabIndex        =   0
      Top             =   120
      Width           =   3150
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2760
         Top             =   1620
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright "
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PearTek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2760
      TabIndex        =   4
      Top             =   2340
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A Web Page editor from "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   2460
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3540
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WebMagic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   3300
      TabIndex        =   1
      Top             =   900
      Width           =   2955
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
       Label2.Caption = _
       "Version : " & App.Major & " .. " & App.Minor & " .. " & App.Revision
       Label5.Caption = App.LegalCopyright
       Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Dim ste As String
      ste = GetSetting(App.Title, "Settings", "site", "")
    If ste = "" Then
       frmfolder1.Show
    Else
     Exit Sub
    End If
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
