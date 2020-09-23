VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WebMagic"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About WebMagic"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   4020
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   780
      ScaleWidth      =   1800
      TabIndex        =   6
      Top             =   60
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":522C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   180
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2820
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "www.peartek.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   600
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "A Web Page Editor from PearTek. Other software available on our website"
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   660
      TabIndex        =   4
      Tag             =   "App Description"
      Top             =   1200
      Width           =   4365
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   780
      TabIndex        =   3
      Tag             =   "Application Title"
      Top             =   180
      Width           =   3165
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   780
      Width           =   2865
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "copyright"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Tag             =   "Warning: ..."
      Top             =   2400
      Width           =   4470
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & " .. " & App.Minor & " ." & App.Revision
    lblTitle.Caption = App.Title
    lblDisclaimer.Caption = App.LegalCopyright
End Sub



Private Sub cmdOK_Click()
        Unload Me
End Sub


