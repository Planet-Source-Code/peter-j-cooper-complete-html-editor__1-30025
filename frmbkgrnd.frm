VERSION 5.00
Begin VB.Form frmbkgrnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format  Page Background"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "frmbkgrnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6720
      TabIndex        =   23
      Text            =   "0"
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3240
      TabIndex        =   22
      Text            =   "Combo4"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5280
      TabIndex        =   21
      Text            =   "Combo3"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   20
      Text            =   "Combo2"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   18
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   5280
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   5280
      Width           =   375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Use Background Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   11
      Top             =   6720
      Width           =   2235
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2520
      ScaleHeight     =   240
      ScaleWidth      =   1305
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3960
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Top             =   6720
      Width           =   1230
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1755
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   2400
      Pattern         =   "*.gif;*.jpg"
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Use Background Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   6720
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   4680
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   1
      Top             =   480
      Width           =   2775
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   31
         Top             =   0
         Width           =   2355
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   0
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Set Page Hyperlink text colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1680
      TabIndex        =   30
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   7455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Set Page Margins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   29
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      Height          =   975
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Choose Background color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      TabIndex        =   28
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   735
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Choose Background Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   3135
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label9 
      Caption         =   "Top Margin"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Left Margin"
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Active link color"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Visited link color"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Link Color"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Use either a Background image or a Background color, not both"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   6360
      Width           =   6060
   End
   Begin VB.Label Label2 
      Caption         =   "Width"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   240
      Left            =   4680
      TabIndex        =   6
      Top             =   2280
      Width           =   660
   End
End
Attribute VB_Name = "frmbkgrnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pth, im, ex, str, bgcol, ll, vl, al, tm, lm As String

Private Sub Combo1_Click()
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
              bgcol = Combo1.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
   Dim whereat
   Dim source, dest
     Dim wth As String
     Dim hth As String
     wth = Text2.Text
     hth = Text3.Text
     ex = """"
     im = "images/"
     tm = Text4.Text
     lm = Text1.Text
      whereat = GetSetting(App.Title, "Settings", "site", "")
      source = Dir1.Path & "\" & File1.filename
      dest = App.Path & "\" & whereat & "\images" & "\" & File1.filename
         FileCopy source, dest
   str = " background = " & ex & im & pth & ex & " link = " & ex & ll & ex & _
         " vlink = " & ex & vl & ex & " alink = " & ex & al & ex & _
         " topmargin = " & tm & " leftmargin = " & lm
   frmMain.ActiveForm.txtText.SelItalic = False
   frmMain.ActiveForm.txtText.SelColor = vbBlue
   frmMain.ActiveForm.txtText.SelText = str
   Unload Me
End Sub

Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub Command4_Click()
     Dim sd As String
     Dim shade As String
         sd = """"
         shade = Combo1.Text
         tm = Text4.Text
         lm = Text1.Text
    frmMain.ActiveForm.txtText.SelItalic = False
    frmMain.ActiveForm.txtText.SelColor = vbBlue
    frmMain.ActiveForm.txtText.SelText = _
    " bgcolor = " & sd & shade & sd & " link = " & sd & ll & sd & _
      " vlink = " & sd & vl & sd & " alink = " & sd & al & sd & _
      " topmargin = " & tm & " leftmargin = " & lm
    Unload Me
End Sub

Private Sub Combo4_Click()
         Select Case Combo4.Text
            Case Is = "black"
              Picture5.BackColor = &H0
            Case Is = "blue"
              Picture5.BackColor = &HFF0000
            Case Is = "fuchsia"
              Picture5.BackColor = &HFF00FF
            Case Is = "gray"
              Picture5.BackColor = &H808080
            Case Is = "green"
              Picture5.BackColor = &HC000&
            Case Is = "lime"
              Picture5.BackColor = &H80FF80
            Case Is = "maroon"
              Picture5.BackColor = &H80&
            Case Is = "navy"
              Picture5.BackColor = &HC00000
            Case Is = "olive"
              Picture5.BackColor = &H8000&
            Case Is = "purple"
              Picture5.BackColor = &H800080
            Case Is = "red"
              Picture5.BackColor = &HFF&
            Case Is = "silver"
              Picture5.BackColor = &HC0C0C0
            Case Is = "teal"
              Picture5.BackColor = &H808000
            Case Is = "yellow"
              Picture5.BackColor = &HFFFF&
            Case Is = "aqua"
              Picture5.BackColor = &HFFFF00
            Case Is = "white"
              Picture5.BackColor = &HFFFFFF
         End Select
             al = Combo4.Text
End Sub

Private Sub Combo3_Click()
         Select Case Combo3.Text
            Case Is = "black"
              Picture4.BackColor = &H0
            Case Is = "blue"
              Picture4.BackColor = &HFF0000
            Case Is = "fuchsia"
              Picture4.BackColor = &HFF00FF
            Case Is = "gray"
              Picture4.BackColor = &H808080
            Case Is = "green"
              Picture4.BackColor = &HC000&
            Case Is = "lime"
              Picture4.BackColor = &H80FF80
            Case Is = "maroon"
              Picture4.BackColor = &H80&
            Case Is = "navy"
              Picture4.BackColor = &HC00000
            Case Is = "olive"
              Picture4.BackColor = &H8000&
            Case Is = "purple"
              Picture4.BackColor = &H800080
            Case Is = "red"
              Picture4.BackColor = &HFF&
            Case Is = "silver"
              Picture4.BackColor = &HC0C0C0
            Case Is = "teal"
              Picture4.BackColor = &H808000
            Case Is = "yellow"
              Picture4.BackColor = &HFFFF&
            Case Is = "aqua"
              Picture4.BackColor = &HFFFF00
            Case Is = "white"
              Picture4.BackColor = &HFFFFFF
         End Select
              vl = Combo3.Text
End Sub

Private Sub Combo2_Click()
         Select Case Combo2.Text
            Case Is = "black"
              Picture3.BackColor = &H0
            Case Is = "blue"
              Picture3.BackColor = &HFF0000
            Case Is = "fuchsia"
              Picture3.BackColor = &HFF00FF
            Case Is = "gray"
              Picture3.BackColor = &H808080
            Case Is = "green"
              Picture3.BackColor = &HC000&
            Case Is = "lime"
              Picture3.BackColor = &H80FF80
            Case Is = "maroon"
              Picture3.BackColor = &H80&
            Case Is = "navy"
              Picture3.BackColor = &HC00000
            Case Is = "olive"
              Picture3.BackColor = &H8000&
            Case Is = "purple"
              Picture3.BackColor = &H800080
            Case Is = "red"
              Picture3.BackColor = &HFF&
            Case Is = "silver"
              Picture3.BackColor = &HC0C0C0
            Case Is = "teal"
              Picture3.BackColor = &H808000
            Case Is = "yellow"
              Picture3.BackColor = &HFFFF&
            Case Is = "aqua"
              Picture3.BackColor = &HFFFF00
            Case Is = "white"
              Picture3.BackColor = &HFFFFFF
         End Select
            ll = Combo2.Text
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveHandler
    Dir1.Path = Drive1.Drive
    Exit Sub

DriveHandler:
       Drive1.Drive = Dir1.Path
    Exit Sub
End Sub

Private Sub File1_Click()
On Error Resume Next
        With File1
        pth = .filename
        Picture6.Picture = LoadPicture(Dir1.Path & "\" & .filename)
        End With
        Text2.Text = Picture6.ScaleWidth
        Text3.Text = Picture6.ScaleHeight
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
             .Text = "blue"
          End With
          
          With Combo3
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
             .Text = "maroon"
          End With
          
          With Combo4
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
             .Text = "maroon"
          End With
          Picture2.BackColor = &HFFFF00
          Picture3.BackColor = &HFF0000
          Picture4.BackColor = &H80&
          Picture5.BackColor = &H80&
          ll = Combo2.Text
          vl = Combo3.Text
          al = Combo4.Text
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 36

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Formats page background. Press F1 for help"
End Sub
