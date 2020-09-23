VERSION 5.00
Begin VB.Form frmlink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Links"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmlink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check5 
      Caption         =   "Link to a tag in this page"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Tag a destination in this page"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   2340
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Add an email link"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Link to another WebSite"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   270
      Left            =   1245
      TabIndex        =   6
      Top             =   3000
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2580
      Width           =   4455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Link to a page in this WebSite"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   4215
   End
End
Attribute VB_Name = "frmlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim es As String
Private Sub Check1_Click()


    If Check1.Value = 1 Then
     Text2.Text = "yourpage.html"
     Check2.Enabled = False
     Check3.Enabled = False
     Check4.Enabled = False
     Check5.Enabled = False
     Label1.Caption = " Type the name of the page"
    Else
     Text2.Text = ""
     Check2.Enabled = True
     Check3.Enabled = True
     Check4.Enabled = True
     Check5.Enabled = True
     Label1.Caption = ""
    End If
End Sub

Private Sub Check2_Click()


    If Check2.Value = 1 Then
     Text2.Text = "somesite.com"
     Check1.Enabled = False
     Check3.Enabled = False
     Check4.Enabled = False
     Check5.Enabled = False
     Label1.Caption = " Type the address of the site"
    Else
     Text2.Text = ""
     Check1.Enabled = True
     Check3.Enabled = True
     Check4.Enabled = True
     Check5.Enabled = True
     Label1.Caption = ""
    End If
End Sub

Private Sub Check3_Click()


    If Check3.Value = 1 Then
     Text2.Text = "you@youraddress"
     Check2.Enabled = False
     Check1.Enabled = False
     Check4.Enabled = False
     Check5.Enabled = False
     Label1.Caption = " Type the email address"
    Else
     Text2.Text = ""
     Check2.Enabled = True
     Check1.Enabled = True
     Check4.Enabled = True
     Check5.Enabled = True
     Label1.Caption = ""
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
     Text2.Text = "Tag Name"
     Check2.Enabled = False
     Check1.Enabled = False
     Check3.Enabled = False
     Check5.Enabled = False
     Label1.Caption = " Type a name for the tag"
    Else
     Text2.Text = ""
     Check2.Enabled = True
     Check1.Enabled = True
     Check3.Enabled = True
     Check5.Enabled = True
     Label1.Caption = ""
    End If
End Sub

Private Sub Check5_Click()
    If Check5.Value = 1 Then
     Text2.Text = "Destination tag name"
     Check2.Enabled = False
     Check1.Enabled = False
     Check4.Enabled = False
     Check3.Enabled = False
     Label1.Caption = " Type the name of the destination tag"
    Else
     Text2.Text = ""
     Check2.Enabled = True
     Check1.Enabled = True
     Check4.Enabled = True
     Check3.Enabled = True
     Label1.Caption = ""
    End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
   Dim wis As String
   Dim yp
   Dim myst As String
   Dim wes As String
   wis = frmMain.ActiveForm.txtText.SelText
   yp = Text2.Text
   es = """"
   frmMain.ActiveForm.txtText.SelItalic = False
   If Check1.Value = 1 Then
      frmMain.ActiveForm.txtText.SelText = _
      "<a href = " & es & yp & es & ">" & wis & "</a>"
      myst = "<a href = " & es & yp & es & ">"
      wes = "<a href = " & es & yp & es & ">" & wis & "</a>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8000&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 4)
        frmMain.ActiveForm.txtText.SelLength = 4
        frmMain.ActiveForm.txtText.SelColor = &H8000&
    ElseIf Check2.Value = 1 Then
      frmMain.ActiveForm.txtText.SelText = _
      "<a href = " & es & "http://www." & yp & es & ">" & wis & "</a>"
      myst = "<a href = " & es & "http://www." & yp & es & ">"
      wes = "<a href = " & es & "http://www." & yp & es & ">" & wis & "</a>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8000&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 4)
        frmMain.ActiveForm.txtText.SelLength = 4
        frmMain.ActiveForm.txtText.SelColor = &H8000&
    ElseIf Check3.Value = 1 Then
      frmMain.ActiveForm.txtText.SelText = _
      "<a href = " & es & "mailto:" & yp & es & ">" & wis & "</a>"
      myst = "<a href = " & es & "mailto:" & yp & es & ">"
      wes = "<a href = " & es & "mailto:" & yp & es & ">" & wis & "</a>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8000&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 4)
        frmMain.ActiveForm.txtText.SelLength = 4
        frmMain.ActiveForm.txtText.SelColor = &H8000&
    ElseIf Check4.Value = 1 Then
      frmMain.ActiveForm.txtText.SelText = _
      "<a name = " & es & yp & es & ">" & wis & "</a>"
      myst = "<a href = " & es & yp & es & ">"
      wes = "<a href = " & es & yp & es & ">" & wis & "</a>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8000&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 4)
        frmMain.ActiveForm.txtText.SelLength = 4
        frmMain.ActiveForm.txtText.SelColor = &H8000&
    ElseIf Check5.Value = 1 Then
      frmMain.ActiveForm.txtText.SelText = _
      "<a href = " & es & "#" & yp & es & ">" & wis & "</a>"
      myst = "<a href = " & es & "#" & yp & es & ">"
      wes = "<a href = " & es & "#" & yp & es & ">" & wis & "</a>"
        frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wes)
        frmMain.ActiveForm.txtText.SelLength = Len(myst)
        frmMain.ActiveForm.txtText.SelColor = &H8000&
        frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wes) - 4)
        frmMain.ActiveForm.txtText.SelLength = 4
        frmMain.ActiveForm.txtText.SelColor = &H8000&
   End If
   Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 42
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Creates Hyperlinks. Press F1 for help"
End Sub

Private Sub Text2_GotFocus()
     With Text2
        .SelStart = 0
        .SelLength = Len(.Text)
     End With
End Sub
