VERSION 5.00
Begin VB.Form frmsave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Page"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter a Name for this page. Don't include a file extension, ie    NOT   index.html."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmsave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   'On Error Resume Next
    Dim sit As String
    Dim nme As String
     sit = GetSetting(App.Title, "Settings", "site", "")
     nme = Text1.Text
    'If frmMain.ActiveForm.Caption <> "New Web" Then
     'frmmes.Show
    ' frmMain.ActiveForm.RichTextBox1.SaveFile App.Path & "\MyWeb\" & nme, 1
    frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".html", 1
    frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".web"
    'lDocumentCount = lDocumentCount - 1
    'Else
    'lDocumentCount = lDocumentCount - 1
     'Exit Sub
    'End If
       Unload Me
    Unload frmMain.ActiveForm
End Sub

Private Sub Timer1_Timer()

End Sub

