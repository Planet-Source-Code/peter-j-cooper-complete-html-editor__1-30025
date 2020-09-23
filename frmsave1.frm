VERSION 5.00
Begin VB.Form frmsave1 
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
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a Name for this page. Don't include a file extension, ie mypage   NOT   mypage.html."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmsave1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
    Dim aa
    Dim pth As String
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
            pth = App.Path & "\" & sit
            frmMain.File1.Path = pth
            frmMain.File1.Refresh
       frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create a tree.
    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", pth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "C3", sit, "open")
    For aa = 0 To frmMain.File1.ListCount - 1
     Set nodX = frmMain.TreeView1.Nodes.Add("C3", tvwChild, , frmMain.File1.List(aa), "page")
    Next aa
    nodX.EnsureVisible  ' Show all nodes.
       Unload Me
       frmMain.ActiveForm.Caption = nme
End Sub

Private Sub Timer1_Timer()

End Sub
