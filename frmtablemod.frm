VERSION 5.00
Begin VB.Form frmtablemod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Table,cell or row  Properties"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Use Background Color"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "Combo3"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      ScaleHeight     =   225
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Select Background Color"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Alignment"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Vertical Alignment"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmtablemod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
         Select Case Combo1.Text
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
       Dim ex, aa, va, bg, str As String
          ex = """"
          aa = Combo3.Text
          va = Combo2.Text
          bg = Combo1.Text
          
    If Check1.Value = 0 Then
       str = " align = " & ex & aa & ex & " valign = " & _
             ex & va & ex
    ElseIf Check1.Value = 1 Then
       str = " bgcolor = " & ex & bg & ex & " align = " & _
             ex & aa & ex & " valign = " & _
             ex & va & ex
    End If
        frmMain.ActiveForm.txtText.SelItalic = False
        frmMain.ActiveForm.txtText.SelColor = &HC000&
        frmMain.ActiveForm.txtText.SelText = str
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
          With Combo1
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
          
        With Combo2
           .AddItem "top"
           .AddItem "middle"
           .AddItem "bottom"
           .Text = "middle"
        End With
        
        With Combo3
           .AddItem "Left"
           .AddItem "center"
           .AddItem "right"
           .Text = "left"
        End With
End Sub
