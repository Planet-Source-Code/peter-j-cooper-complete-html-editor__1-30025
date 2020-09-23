VERSION 5.00
Begin VB.Form frmfolder1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create your first Site"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   Icon            =   "frmfolder1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2820
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   3045
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a name for your First Site"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   270
      Width           =   2175
   End
End
Attribute VB_Name = "frmfolder1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
     Dim nme, Newpth As String
     Dim aa, ff, gg, hh, ii
     nme = Text1.Text
     
         MkDir App.Path & "\" & nme
         MkDir App.Path & "\" & nme & "\images"
         
            Newpth = App.Path & "\" & nme
            frmMain.File1.Path = Newpth
            frmMain.File1.Refresh
       frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create a tree.
    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", Newpth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "C3", nme, "open")
    For aa = 0 To frmMain.File1.ListCount - 1
     Set nodX = frmMain.TreeView1.Nodes.Add("C3", tvwChild, , frmMain.File1.List(aa), "page")
    Next aa
    nodX.EnsureVisible  ' Show all nodes.
    frmMain.Dir1.Refresh
                frmMain.List1.Clear
   gg = Len(App.Path) + 2
    ff = frmMain.Dir1.ListIndex
      ff = ff + 1
      For ff = 0 To frmMain.Dir1.ListCount - 1
           hh = Len(frmMain.Dir1.List(ff))
           ii = (hh - gg) + 1
          frmMain.List1.AddItem Mid(frmMain.Dir1.List(ff), gg, ii)
      Next ff
    Unload Me

End Sub



Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 24
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Creates a new site folder. Press F1 for help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
          If Text1.Text = "" Then
            Exit Sub
          Else
            SaveSetting App.Title, "Settings", "site", Me.Text1.Text
            frmMain.Label3.Caption = Text1.Text
      frmMain.Label6.Caption = "Pages from " & UCase(Text1.Text) & " Click on Page Name to edit"
          End If
End Sub
