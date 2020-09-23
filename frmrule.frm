VERSION 5.00
Begin VB.Form frmrule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Horizontal Rule"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmrule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No 3D Shading"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Text            =   """"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "300"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "4"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Line Color"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Alignment"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Height"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Width"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmrule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
         Select Case Combo2.Text
            Case Is = "black"
              Picture2.BackColor = &H0
            Case Is = "blue"
              Picture2.BackColor = &HFF0000
            Case Is = "fuchsia"
              Picture2.BackColor = &HFF00FF
            Case Is = "gray"
              Picture2.BackColor = &H808080
            Case Is = "green"
              Picture2.BackColor = &HC000&
            Case Is = "lime"
              Picture2.BackColor = &H80FF80
            Case Is = "maroon"
              Picture2.BackColor = &H80&
            Case Is = "navy"
              Picture2.BackColor = &HC00000
            Case Is = "olive"
              Picture2.BackColor = &H8000&
            Case Is = "purple"
              Picture2.BackColor = &H800080
            Case Is = "red"
              Picture2.BackColor = &HFF&
            Case Is = "silver"
              Picture2.BackColor = &HC0C0C0
            Case Is = "teal"
              Picture2.BackColor = &H808000
            Case Is = "yellow"
              Picture2.BackColor = &HFFFF&
            Case Is = "aqua"
              Picture2.BackColor = &HFFFF00
            Case Is = "white"
              Picture2.BackColor = &HFFFFFF
         End Select
End Sub

Private Sub Command1_Click()
On Error Resume Next
       Dim ht As String
       Dim wd As String
       Dim lft As String
       Dim shd As String
       Dim ex As String
       Dim cl As String
           ht = Text1.Text
           wd = Text2.Text
           lft = Combo1.Text
           ex = """"
        If Check1.Value = 1 Then
           shd = "noshade"
           cl = Combo2.Text
         Else
           shd = ""
           cl = ""
        End If
    frmMain.ActiveForm.txtText.SelItalic = False
    frmMain.ActiveForm.txtText.SelColor = vbRed
    frmMain.ActiveForm.txtText.SelText = _
    "<hr size = " & ht & " width = " & wd & " align = " & ex & lft & ex & " color = " & ex & cl & ex & " " & shd & ">"

     Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
        With Combo1
            .AddItem "left"
            .AddItem "center"
            .AddItem "right"
            .ListIndex = 0
        End With
        
          With Combo2
             .AddItem "aqua"
             .AddItem "black"
             .AddItem "blue"
             .AddItem "fuchsia"
             .AddItem "gray"
             .AddItem "green"
             .AddItem "lime"
             .AddItem "maroon"
             .AddItem "navy"
             .AddItem "olive"
             .AddItem "purple"
             .AddItem "red"
             .AddItem "silver"
             .AddItem "teal"
             .AddItem "yellow"
             .AddItem "white"
             .Text = "aqua"
          End With
          Picture2.BackColor = &HFFFF00
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 12
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts a horizontal bar. Press F1 for help"
End Sub
