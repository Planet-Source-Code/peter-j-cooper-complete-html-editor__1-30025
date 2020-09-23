VERSION 5.00
Begin VB.Form frmframe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frameset Page"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmframe.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Frame with Border"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frame without Border"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add Side Column"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   2880
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   480
      Value           =   1
      Width           =   255
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   2535
      Index           =   0
      Left            =   3240
      ScaleHeight     =   2505
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin VB.PictureBox pac 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   2640
         ScaleHeight     =   945
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   360
         ScaleHeight     =   465
         ScaleWidth      =   1665
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   2385
         TabIndex        =   1
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Don't forget to remove the body tags before using Frames"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "col2"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "col1"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "row2"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "row1"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
       If Check1.Value = 1 Then
         HScroll1.Visible = True
         Text3.Visible = True
         Text4.Visible = True
       Else
         HScroll1.Visible = False
         Text3.Visible = False
         Text4.Visible = False
       End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
           Dim one
           Dim two
           Dim nmone
           Dim nmtwo
           Dim cone
           Dim ctwo
           Dim nmcone
           Dim nmctwo
           Dim ex As String
               one = Text1.Text
               two = Text2.Text
               nmone = Label1.Caption
               nmtwo = Label2.Caption
               cone = Text3.Text
               ctwo = Text4.Text
               nmcone = Label3.Caption
               nmctwo = Label4.Caption
               ex = """"
         frmMain.ActiveForm.txtText.SelItalic = False
If Check1.Value = 0 Then
           frmMain.ActiveForm.txtText.SelColor = vbBlue
           frmMain.ActiveForm.txtText.SelText = _
     "<frameset frameborder=no frameborder=0 border=0 rows=" & ex & one & "%," & two & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmone & ex & " noresize>" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmtwo & ex & " noresize>" & vbCrLf & _
     "</frameset>"
ElseIf Check1.Value = 1 Then
           frmMain.ActiveForm.txtText.SelColor = vbBlue
           frmMain.ActiveForm.txtText.SelText = _
     "<frameset frameborder=no frameborder=0 border=0 rows=" & ex & one & "%," & two & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmone & ex & " noresize>" & vbCrLf & _
     "<frameset cols= " & ex & cone & "%," & ctwo & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmcone & ex & " noresize>" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmctwo & ex & " noresize>" & vbCrLf & _
     "</frameset>" & vbCrLf & _
     "</frameset>"
  End If
Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next
           Dim one
           Dim two
           Dim nmone
           Dim nmtwo
           Dim cone
           Dim ctwo
           Dim nmcone
           Dim nmctwo
           Dim ex As String
               one = Text1.Text
               two = Text2.Text
               nmone = Label1.Caption
               nmtwo = Label2.Caption
               cone = Text3.Text
               ctwo = Text4.Text
               nmcone = Label3.Caption
               nmctwo = Label4.Caption
               ex = """"
        frmMain.ActiveForm.txtText.SelItalic = False
        If Check1.Value = 0 Then
           frmMain.ActiveForm.txtText.SelColor = vbBlue
           frmMain.ActiveForm.txtText.SelText = _
     "<frameset rows=" & ex & one & "%," & two & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmone & ex & " noresize>" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmtwo & ex & " noresize>" & vbCrLf & _
     "</frameset>"
        ElseIf Check1.Value = 1 Then
           frmMain.ActiveForm.txtText.SelColor = vbBlue
           frmMain.ActiveForm.txtText.SelText = _
     "<frameset rows=" & ex & one & "%," & two & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmone & ex & " noresize>" & vbCrLf & _
     "<frameset cols= " & ex & cone & "%," & ctwo & "%" & ex & ">" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmcone & ex & " noresize>" & vbCrLf & _
     "<frame scrolling=" & ex & "auto" & ex & " src =" & ex & ex & " name=" & ex & nmctwo & ex & " noresize>" & vbCrLf & _
     "</frameset>" & vbCrLf & _
     "</frameset>"
    End If
Unload Me

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 40
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Creates a page with different frames. Press F1 for help"
End Sub

Private Sub Form_Resize()
      Dim vl
      Dim bl
      vl = pic(0).Height
      bl = pic(0).Width
       pic(1).Left = 0
       pic(1).Top = 0
       pic(1).Width = pic(0).Width
       pic(2).Width = pic(0).Width
       pic(1).Height = vl / 2
       VScroll1.Max = vl
       VScroll1.Value = vl / 2
       HScroll1.Max = bl
       HScroll1.Value = 0
       pic(2).Top = pic(1).Top + pic(1).Height
       pic(2).Left = 0
       pac.Left = 0
       pac.Top = pic(1).Top + pic(1).Height
       pac.Width = HScroll1.Value
       pac.Height = VScroll1.Value
       Text3.Text = "0"
       Text4.Text = "100"

End Sub

Private Sub HScroll1_Change()
       Dim ht As Integer
       Dim hp As Integer
       Dim hs As Integer
       Dim hh As Integer
           ht = HScroll1.Value
           hp = pic(0).Width
           hs = (ht / hp) * 100
           hh = 100 - hs
          pac.Width = HScroll1.Value
          pic(2).Left = HScroll1.Value
          pic(2).Width = pic(0).Width - pac.Width
          Text3.Text = hs
          Text4.Text = hh
End Sub

Private Sub VScroll1_Change()
         Dim tp As Integer
         Dim op As Integer
         Dim sp As Integer
         Dim pp As Integer
             tp = VScroll1.Value
             op = pic(0).Height
             sp = (tp / op) * 100
             pp = 100 - sp
       pic(1).Height = VScroll1.Value
       pic(2).Top = VScroll1.Value
       pic(2).Height = pic(0).Height - pic(1).Height
       pac.Top = VScroll1.Value
       pac.Height = pic(0).Height - pic(1).Height
       Text1.Text = sp
       Text2.Text = pp
End Sub
