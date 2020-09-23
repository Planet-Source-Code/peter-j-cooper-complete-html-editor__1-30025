VERSION 5.00
Begin VB.Form frmmarq 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Scrolling Text"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmmarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3240
      TabIndex        =   15
      Text            =   "Combo3"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "2"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "2"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Text            =   "20"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Text            =   "300"
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Direction of travel"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Type in the Text to scroll across the Page"
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
      Left            =   600
      TabIndex        =   16
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label7 
      Caption         =   "Height"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Width"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Scroll Delay"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Scroll Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Loop"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Behavior"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmmarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
        Dim ex As String
        Dim mes As String
        Dim sc As String
        Dim del As String
        Dim behv As String
        Dim lp As String
        Dim wd As String
        Dim ht As String
        Dim trav As String
            ex = """"
            mes = Text1.Text
            sc = Text5.Text
            del = Text6.Text
            behv = Combo1.Text
            lp = Combo2.Text
            wd = Text3.Text
            ht = Text4.Text
            trav = Combo3.Text
frmMain.ActiveForm.txtText.SelItalic = False
frmMain.ActiveForm.txtText.SelColor = &H800000
frmMain.ActiveForm.txtText.SelText = "<marquee behavior = " & ex & behv & ex & " scrollamount = " & sc & " scrolldelay = " & del & " loop = " & lp & " direction = " & trav & " width = " & wd & " height = " & ht & ">" & mes & "</marquee>"
    Unload Me
             
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub

Private Sub Form_Load()
          With Combo1
              .AddItem "Scroll"
              .AddItem "Slide"
              .ListIndex = 0
          End With
          
          With Combo2
               .AddItem "Infinite"
               .AddItem "1"
               .AddItem "2"
               .AddItem "3"
               .AddItem "5"
               .AddItem "6"
               .ListIndex = 0
          End With
          
          With Combo3
               .AddItem "Left"
               .AddItem "Right"
               .ListIndex = 0
          End With
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 35
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Inserts Scrolling text. Press F1 for help"
End Sub
