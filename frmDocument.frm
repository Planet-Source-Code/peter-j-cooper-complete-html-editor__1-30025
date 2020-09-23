VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   BackColor       =   &H00FFFFC0&
   Caption         =   "frmDocument"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   ForeColor       =   &H00000000&
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   6720
   Begin RichTextLib.RichTextBox txtText 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmDocument.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtAdd, txtLen, txtStrt, oldtxtStrt, oldtxtLen As Integer
Dim txtStr As String

Private Sub Form_Load()
    Me.WindowState = 2
    Me.BackColor = GetSetting(App.Title, "Settings", "bcol", &H0&)
    Form_Resize
    txtText.Text = "<html>" & vbCrLf & vbCrLf & "<head>" & vbCrLf & "<title>" & "Your Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & "Your Page Code goes here" & vbCrLf & "</body>" & vbCrLf & vbCrLf & "</html>"
              With txtText
                  .SelStart = 0
                  .SelLength = Len(.Text)
                  .SelColor = vbBlue
              End With
        txtText.SelStart = Len(txtText.Text) + 1
        txtAdd = Len(txtText.Text)
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    txtText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    txtText.RightMargin = txtText.Left + txtText.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim sit As String
    Dim nme As String
     sit = GetSetting(App.Title, "Settings", "site", "")
     nme = frmMain.ActiveForm.Caption
If Len(txtText.Text) <> txtAdd And frmMain.ActiveForm.Caption <> "New Web" Then
Dim Msg, Style, Response
Msg = "Do you want to Save Changes ?"   ' Define message.
Style = vbYesNo + vbQuestion
        ' Display message.
Response = MsgBox(Msg, Style)
If Response = vbYes Then    ' User chose Yes.
    frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".html", 1
    frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".web"
    Cancel = -1
    txtAdd = Len(txtText.Text)
Else    ' User chose No.
    Unload Me
End If
Else
 Cancel = 0
End If
If Len(txtText.Text) <> txtAdd And frmMain.ActiveForm.Caption = "New Web" Then
Dim Msga, Stylea, Responsea
Msga = "Do you want to Save ?"   ' Define message.
Stylea = vbYesNo + vbQuestion
        ' Display message.
Responsea = MsgBox(Msga, Stylea)
If Responsea = vbYes Then    ' User chose Yes.
    frmsave.Show
    Cancel = -1
    txtAdd = Len(txtText.Text)
Else    ' User chose No.
    Unload Me
End If
Else
 Cancel = 0
End If
 lDocumentCount = lDocumentCount - 1
End Sub

Private Sub txtText_Click()
        EdColor
End Sub
Private Sub EdColor()
'***********************************************************************
'  An attempt to keep the editing color the same unless the user is
'  setting tag attributes within a tag
'  Tag colors are formatted at module level which makes this
'  bit slightly easier
'***********************************************************************
'On Error Resume Next
    oldtxtStrt = txtText.SelStart  ' Initial position of insertion point
    oldtxtLen = txtText.SelLength  ' Initial selection length
    txtStrt = oldtxtStrt - 1       ' Sets selstart back one char
    txtLen = oldtxtLen + 1         ' Sets selLength to one char
    txtText.SelStart = txtStrt
    txtText.SelLength = txtLen
If oldtxtLen = 0 Then              ' Decide if sellength is nothing
  If txtText.SelText = ">" Then    ' Decide if edit color should be applied
     txtText.SelStart = txtStrt + 1
     txtText.SelLength = 0     ' Set selstart and sellength to where the user wanted
     txtText.SelColor = frmMain.Picture5.BackColor  ' Set the color
 ElseIf txtText.SelText = "" Then
     txtText.SelStart = oldtxtStrt
     txtText.SelLength = 0
     txtText.SelColor = frmMain.Picture5.BackColor
 ElseIf txtText.SelText <> ">" Then  ' Decide if the selection isn't outside a tag
     txtText.SelStart = oldtxtStrt   ' and don't change the color
     txtText.SelLength = 0
  End If
Else   ' If it's none of the above, leave well alone
    txtText.SelStart = oldtxtStrt
    txtText.SelLength = oldtxtLen
End If
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyDelete Or KeyCode = vbKeyReturn Then
            EdColor
         End If
End Sub
