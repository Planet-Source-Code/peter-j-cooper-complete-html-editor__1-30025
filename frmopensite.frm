VERSION 5.00
Begin VB.Form frmopensite 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Existing Site"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmopensite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Site Name"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select your site folder by clicking it with your mouse and then click OK"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmopensite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Newpth As String
Dim aa
Private Sub Command1_Click()
          If frmMain.ActiveForm.Caption <> "New Web" Then
             MsgBox "These pages are already saved to the current site. Please close them first"
             Exit Sub
        Else
            SaveSetting App.Title, "Settings", "site", Me.Text5.Text
            Newpth = App.Path & "\" & Me.Text5.Text
            frmMain.File1.Path = Newpth
            frmMain.File1.Refresh
       frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create a tree.
    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", Newpth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "C3", Text5.Text, "open")
    For aa = 0 To frmMain.File1.ListCount - 1
     Set nodX = frmMain.TreeView1.Nodes.Add("C3", tvwChild, , frmMain.File1.List(aa), "page")
    Next aa
    nodX.EnsureVisible  ' Show all nodes.
            Unload Me
          End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Click()
 On Error Resume Next
     Dim li As Integer
     Dim nm As Integer
     Dim an As Integer
     Dim bn As Integer
     li = Dir1.ListIndex
     an = Len(Dir1.List(li))
     bn = Len(Dir1.Path) + 1
     nm = an - bn
     Text5.Text = Right(Dir1.List(li), nm)
End Sub

Private Sub Form_Load()
      Dir1.Path = App.Path
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 25

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Opens an exsisting site folder. Press F1 for help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim labc As String
       labc = GetSetting(App.Title, "Settings", "site", "")
   If Text5.Text = "" Then
     Exit Sub
   Else
      frmMain.Label3.Caption = labc
      frmMain.Label6.Caption = "Pages from " & UCase(labc) & " Click on Page Name to edit"
   End If
End Sub
