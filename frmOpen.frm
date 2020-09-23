VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Page"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   210
      Pattern         =   "*.web"
      TabIndex        =   0
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Select Page from the Current Site"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "File Name"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    'Static lDocumentCount As Long
    Dim frmD As frmDocument
    Dim nm
    Dim nma
    Dim stf As String
        nm = Len(File1.filename) - 4
        nma = Left(File1.filename, nm)
        stf = GetSetting(App.Title, "Settings", "site", "")
    'lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    lDocumentCount = lDocumentCount + 1
        With File1
          frmD.txtText.LoadFile App.Path & "\" & stf & "\" & .filename
          frmD.Caption = nma
          'frmMain.Text1.Text = nma
        End With
      Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub File1_Click()
     Text1.Text = File1.filename
End Sub

Private Sub Form_Load()
On Error Resume Next
       Dim st As String
           st = GetSetting(App.Title, "Settings", "site", "")
       With File1
           .Path = App.Path & "\" & st
       End With
            App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 41

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Opens a page from your current site for editing. Press F1 for help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
     'Instruct (11)
End Sub
