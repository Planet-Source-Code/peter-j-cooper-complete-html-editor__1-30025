VERSION 5.00
Begin VB.Form frmmap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image Map"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "frmmap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   717
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   2040
      TabIndex        =   26
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3960
      TabIndex        =   25
      Top             =   5520
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7080
      TabIndex        =   21
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   6600
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7320
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "O.K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   7080
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   8880
      Pattern         =   "*.gif;*.jpg"
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   0
      Top             =   120
      Width           =   6090
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   4200
         Width           =   5190
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3870
         Left            =   5640
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   0
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   329
         TabIndex        =   22
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "loaded Button Area Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Enter a destination for this Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   $"frmmap.frx":0442
      Height          =   975
      Left            =   7440
      TabIndex        =   20
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label11 
      Caption         =   "3 :- Enter Destination for Button Area"
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "2 :- Enter a name for the Map"
      Height          =   255
      Left            =   7440
      TabIndex        =   18
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "1 :- Choose an Image"
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Top"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Enter a Name for your Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choose Image for your Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2655
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Single
Dim OldY As Single
Dim X1 As Single
Dim Y1 As Single
Dim ex, fle, str, stra As String




Private Sub Command3_Click()
On Error Resume Next
            Dim st As String
            Dim source, dest
            Dim nme As String
            Dim str As String
            Dim stra As String
            Dim i
                ex = """"
                nme = Text5.Text
      st = GetSetting(App.Title, "Settings", "site", "")
      source = Dir1.Path & "\" & File1.filename
       dest = App.Path & "\" & st & "\images" & "\" & File1.filename
         FileCopy source, dest
            If Text5.Text = "" Then
               MsgBox "Please Enter a name for this Map"
               Exit Sub
            End If
           frmMain.ActiveForm.txtText.SelItalic = False
           frmMain.ActiveForm.txtText.SelColor = &HC000C0
           i = i + 1
      For i = 0 To List1.ListCount - 1
           str = "<area shape = " & ex & "rect" & ex & " coords = " & ex & List1.List(i) & ex & " href = " & ex & List2.List(i) & ex & ">" & vbCrLf
           stra = stra + str
      Next i

           frmMain.ActiveForm.txtText.SelText = _
    "<img src = " & ex & "images/" & fle & ex & " usemap = " & ex & "#" & nme & ex & " border = 0></img>" & vbCrLf & _
    "<map name = " & ex & nme & ex & ">" & vbCrLf & _
     stra & _
    "</map>"
     Unload Me
End Sub

Private Sub Command4_Click()
     Unload Me
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
        fle = .filename
        Picture2.Picture = LoadPicture(Dir1.Path & "\" & .filename)
        End With
End Sub

Private Sub Form_Load()
           HScroll1.Left = 0
           HScroll1.Top = Picture1.Height - HScroll1.Height
           HScroll1.Width = Picture1.Width
           VScroll1.Top = 0
           VScroll1.Left = Picture1.Width - VScroll1.Width
           VScroll1.Height = Picture1.Height - HScroll1.Height
           Picture2.Top = 0
           Picture2.Left = 0
           Picture2.Width = Picture1.Width - VScroll1.Width
           Picture2.Height = Picture1.Height - HScroll1.Height
           HScroll1.Max = Picture2.Width - Picture1.Width
           VScroll1.Max = Picture2.Height - Picture1.Height
           
          VScroll1.Visible = (Picture1.Height < _
          Picture2.Height)
          HScroll1.Visible = (Picture1.Width < _
          Picture2.Width)

     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 31

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts an Image Map. Press F1 for help"
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Text6.Text = "" Then
    MsgBox "Please enter a destination for this Button Area First!"
 Else
Picture2.Line (X1, Y1)-(X1, Y)
Picture2.Line (X1, Y1)-(X, Y1)
Picture2.Line (X, Y1)-(X, Y)
Picture2.Line (X1, Y)-(X, Y)
     Text1.Text = Y1
     Text2.Text = X1
     Text3.Text = Y
     Text4.Text = X

         Dim tp, lt, rt, bt
         tp = Text1.Text
         lt = Text2.Text
         rt = Text4.Text
         bt = Text3.Text
                 With List1
                   .AddItem lt & "," & tp & "," & rt & "," & bt
                 End With
                 With List2
                   .AddItem Text6.Text
                 End With
 End If
             Text6.Text = ""

End Sub


Private Sub Picture2_Resize()
           HScroll1.Max = Picture2.Width - Picture1.Width
           VScroll1.Max = Picture2.Height - Picture1.Height
          VScroll1.Visible = (Picture1.Height < _
          Picture2.Height)
          HScroll1.Visible = (Picture1.Width < _
          Picture2.Width)
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub
