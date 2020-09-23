VERSION 5.00
Begin VB.Form frmlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert List"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Unordered List - bullets"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Odered List - 1,2,3 ect"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "List Heading name"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Bullet type"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ex As String
Private Sub Check2_Click()
         If Check2.Value = 1 Then
           Label1.Visible = True
           Combo1.Visible = True
         Else
           Label1.Visible = False
           Combo1.Visible = False
         End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
         Dim nme As String
         Dim ol As String
         Dim ul As String
         Dim bul As String
         nme = Text2.Text
         ex = """"
         bul = Combo1.Text
     frmMain.ActiveForm.RichTextBox1.SelItalic = False
      If Check1.Value = 1 And Check2.Value = 0 Then
          ol = "<ol>"
          ul = "</ol>"
      ElseIf Check1.Value = 0 And Check2.Value = 1 Then
          ol = "<ul type = " & ex & bul & ex & ">"
          ul = "</ul>"
      Else
        MsgBox "You can have only one type of List"
        Exit Sub
     End If
     frmMain.ActiveForm.txtText.SelColor = &H808000
     frmMain.ActiveForm.txtText.SelText = _
     ol & vbCrLf & "<lh>" & nme & "</lh>" & vbCrLf & _
     "<li>your item here" & vbCrLf & "<li>your item here" & vbCrLf & _
     ul
     Unload Me
End Sub

Private Sub Form_Load()
                With Combo1
                  .AddItem "circle"
                  .AddItem "disc"
                  .AddItem "square"
                  .ListIndex = 0
                End With
             App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 34
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts a bulleted or Ordered list. Press F1 for help"
End Sub
