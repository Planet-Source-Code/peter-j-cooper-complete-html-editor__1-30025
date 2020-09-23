VERSION 5.00
Begin VB.Form frmcol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Panel Scheme"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmcol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1500
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Scheme"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   420
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   1635
      Left            =   3300
      TabIndex        =   2
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmcol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
ChgCol
End Sub

Private Sub Command1_Click()
      Unload Me
End Sub

Private Sub Form_Load()
    Label1.BackColor = GetSetting(App.Title, "Settings", "bcol", &H0&)
    Label1.ForeColor = GetSetting(App.Title, "Settings", "fcol", &HFF00&)
    Combo1.Text = GetSetting(App.Title, "Settings", "ctxt", " Techno Black")
    Label1.Caption = vbCrLf & " This is The Color"
    
          With Combo1
            .AddItem " Techno Black"
            .AddItem " Green Baize"
            .AddItem " Ice Blue"
            .AddItem " Deep Blue"
            .AddItem " Light Gold"
            .AddItem " Purple Haze"
            .AddItem " Standard Grey"
          End With
End Sub
Private Sub ChgCol()

           Select Case Combo1.Text
             Case Is = " Techno Black"
              Label1.BackColor = &H0&
              Label1.ForeColor = &HFF00&
            Case Is = " Green Baize"
              Label1.BackColor = &H8000&
              Label1.ForeColor = &HFFFF&
            Case Is = " Ice Blue"
              Label1.BackColor = &HFFFF00
              Label1.ForeColor = &HC00000
            Case Is = " Deep Blue"
              Label1.BackColor = &HC00000
              Label1.ForeColor = &HFFFF00
            Case Is = " Light Gold"
              Label1.BackColor = &H80C0FF
              Label1.ForeColor = &H0&
            Case Is = " Purple Haze"
              Label1.BackColor = &H800080
              Label1.ForeColor = &H80C0FF
            Case Is = " Standard Grey"
              Label1.BackColor = &HC0C0C0
              Label1.ForeColor = &H0&
          End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
        SaveSetting App.Title, "Settings", "bcol", Label1.BackColor
        SaveSetting App.Title, "Settings", "fcol", Label1.ForeColor
        SaveSetting App.Title, "Settings", "ctxt", Combo1.Text
        frmMain.Picture1.BackColor = Label1.BackColor
        frmMain.Picture2.BackColor = Label1.BackColor
        frmMain.Label9.BackColor = Label1.BackColor
        frmMain.Shape1.BorderColor = Label1.ForeColor
        frmMain.Label2.ForeColor = Label1.ForeColor
        frmMain.Label3.ForeColor = Label1.ForeColor
        frmMain.Label9.ForeColor = Label1.ForeColor
        frmMain.Label10.ForeColor = Label1.ForeColor
        frmMain.Picture3.BackColor = Label1.BackColor
        frmMain.SSTab1.BackColor = Label1.BackColor
        frmMain.ActiveForm.BackColor = Label1.BackColor
        frmMain.Label1.ForeColor = Label1.ForeColor
        frmMain.Label6.ForeColor = Label1.ForeColor
        
End Sub
