Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global ctrparam As Control
Global lDocumentCount As Long

Dim Was As String
Dim wis As String
Dim tag As Single
Dim hlp As Single




Sub setTable()
   Set ctrparam = frmMain.ActiveForm.txtText
    frmtable.Show 1
End Sub
Sub fillTable()
      Dim row As String
      Dim rowa As String
      Dim rowb As String
      Dim rowc As String
      Dim rowd As String
      Dim rowe As String
      Dim cal As String
      Dim num As Integer
      Dim numa As Integer
      Dim bdr
      Dim szpix
      Dim szper
      Dim wdthp As String
      Dim hgthp As String

      num = frmtable.Combo1.Text
      numa = frmtable.Combo2.Text
      bdr = frmtable.Check1.Value
      szpix = frmtable.Option1.Value
      szper = frmtable.Option2.Value
      wdthp = frmtable.Text1.Text
      hgthp = frmtable.Text2.Text
      frmMain.ActiveForm.txtText.SelItalic = False
      rowa = "<td></td>"
    If num = 1 And numa = 1 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 1 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 1 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 1 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 1 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 2 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 2 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 2 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 2 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 2 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 3 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 3 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 3 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 3 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 3 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 4 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 4 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 4 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 4 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 4 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 5 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 5 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 5 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 5 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 5 And bdr = 0 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowc = "<table width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    End If
ctrparam.SelColor = &HC000&
ctrparam.SelText = rowc

    If num = 1 And numa = 1 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 1 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 1 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 1 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 1 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 2 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 2 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 2 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 2 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 2 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 3 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 3 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 3 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 3 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 3 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 4 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 4 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 4 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 4 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 4 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 5 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 5 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 5 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 5 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 5 And bdr = 0 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowd = "<table width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    End If
ctrparam.SelColor = &HC000&
ctrparam.SelText = rowd


    If num = 1 And numa = 1 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 1 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 1 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 1 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 1 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 2 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 2 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 2 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 2 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 2 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 3 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 3 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 3 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 3 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 3 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 4 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 4 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 4 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 4 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 4 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 5 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 5 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 5 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 5 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 5 And bdr = 1 And szper = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowb = "<table border = 1 width = " & wdthp & "%" & " height = " & hgthp & "%" & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    End If
ctrparam.SelColor = &HC000&
ctrparam.SelText = rowb

    If num = 1 And numa = 1 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 1 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 1 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 1 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 1 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 2 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 2 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 2 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 2 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 2 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 3 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 3 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 3 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 3 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 3 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 4 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 4 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 4 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 4 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 4 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 1 And numa = 5 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 2 And numa = 5 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 3 And numa = 5 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 4 And numa = 5 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    ElseIf num = 5 And numa = 5 And bdr = 1 And szpix = True Then
      row = "<tr>" & rowa & rowa & rowa & rowa & rowa & "</tr>"
      rowe = "<table border = 1 width = " & wdthp & " height = " & hgthp & " cellpadding = 1 cellspacing = 1>" & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & row & vbCrLf & "</table>"
    End If
ctrparam.SelColor = &HC000&
ctrparam.SelText = rowe

End Sub
Sub setImage()
   Set ctrparam = frmMain.ActiveForm.txtText
    frmimage.Show 1
End Sub
Sub SetBk()
   Set ctrparam = frmMain.ActiveForm.RichTextBox1
       frmbkgrnd.Show
End Sub
Sub TagFmt(tag)
      Was = frmMain.ActiveForm.txtText.SelText
      Dim ex As String
          ex = """"
     
           Select Case tag
              Case Is = 1 'Center open tag
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelText = "<center>"
              Case Is = 2 'Bold text tags
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<b>" & Was & "</b>"
                wis = "<b>" & Was & "</b>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 3
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 4)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = vbRed
              Case Is = 3 'Italic Text
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<i>" & Was & "</i>"
                wis = "<i>" & Was & "</i>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 3
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 4)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = vbRed
              Case Is = 4 'Underlined text
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<u>" & Was & "</u>"
                wis = "<u>" & Was & "</u>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 3
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 4)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = vbRed
              Case Is = 5 'Heading 1
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h1>" & Was & "</h1>"
                wis = "<h1>" & Was & "</h1>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 6 'Heading 2
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h2>" & Was & "</h2>"
                wis = "<h2>" & Was & "</h2>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 7 'Heading 3
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h3>" & Was & "</h3>"
                wis = "<h3>" & Was & "</h3>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 8 'Heading 4
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h4>" & Was & "</h4>"
                wis = "<h4>" & Was & "</h4>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 9 'Heading 5
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h5>" & Was & "</h5>"
                wis = "<h5>" & Was & "</h5>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 10 'Heading 6
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<h6>" & Was & "</h6>"
                wis = "<h6>" & Was & "</h6>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 4
                frmMain.ActiveForm.txtText.SelColor = &H80&
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 5)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = &H80&
              Case Is = 11 ' Line Break
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelText = "<br>"
              Case Is = 12 ' Space
                frmMain.ActiveForm.txtText.SelItalic = True
                frmMain.ActiveForm.txtText.SelText = Was & "&nbsp;"
                frmMain.ActiveForm.txtText.SelItalic = False
              Case Is = 14 'Paragraph
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelText = "<p>"
              Case Is = 15 ' Division
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<div>" & Was & "</div>"
                wis = "<div>" & Was & "</div>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 5
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 6)
                frmMain.ActiveForm.txtText.SelLength = 6
                frmMain.ActiveForm.txtText.SelColor = vbRed
              Case Is = 16 ' Blockquote
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelText = "<blockquote>" & Was & "</blockquote>"
                wis = "<blockquote>" & Was & "</blockquote>"
                frmMain.ActiveForm.txtText.SelStart = frmMain.ActiveForm.txtText.SelStart - Len(wis)
                frmMain.ActiveForm.txtText.SelLength = 12
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelStart = (frmMain.ActiveForm.txtText.SelStart + Len(wis) - 13)
                frmMain.ActiveForm.txtText.SelLength = 13
                frmMain.ActiveForm.txtText.SelColor = vbRed
              Case Is = 17 ' Javascript tags
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelColor = &H80FF&
                frmMain.ActiveForm.txtText.SelText = "<script language = " & ex & "JavaScript" & ex & ">" & vbCrLf & "Put your script here" & vbCrLf & "</script>"
                
              Case Is = 18 ' Comment
                frmMain.ActiveForm.txtText.SelItalic = True
                frmMain.ActiveForm.txtText.SelColor = &HC0C0C0
                frmMain.ActiveForm.txtText.SelText = "<!-- Put your comment here -->"
                frmMain.ActiveForm.txtText.SelItalic = False
              Case Is = 19 ' center closing tag
                frmMain.ActiveForm.txtText.SelItalic = False
                frmMain.ActiveForm.txtText.SelColor = vbRed
                frmMain.ActiveForm.txtText.SelText = "</center>"
               
                

           End Select
End Sub

Sub Instruct(hlp)

             Select Case hlp
             
                Case Is = 1
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Place the insertion point in the body tag, ( <body|> ), for this operation"
                Case Is = 2
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Select the text to be enclosed by these tags"
                Case Is = 3
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Place the insertion point where you want this to occur"
                Case Is = 4
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Place the insertion point where the paragraph is to start"
                Case Is = 5
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Place the insertion point in the table tag, td tag, or tr tag ( <table|> ), for this operation"
                Case Is = 6
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Place the insertion point on a new line in the head section"
                Case Is = 7
                 frmMain.sbStatusBar.Panels(1).Text = _
                 "Remove the <body> tags before inserting frame sets"
            End Select
                 
End Sub
Sub FmtFm()
     Dim item As String
     Dim ex As String
         item = frmform.Combo1.Text
         ex = """"
         frmMain.ActiveForm.txtText.SelItalic = False
        If item = "Text Box" Then
           frmMain.ActiveForm.txtText.SelColor = &H808000
           frmMain.ActiveForm.txtText.SelText = _
           "<input type = " & ex & "text" & ex & " name = " & ex & ex & " size = 30 cols = 30>"
        ElseIf item = "Submit Button" Then
           frmMain.ActiveForm.txtText.SelColor = &H808000
           frmMain.ActiveForm.txtText.SelText = _
           "<input type = " & ex & "submit" & ex & " name = " & ex & "submit" & ex & " value = Submit>"
        ElseIf item = "Reset Button" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<input type = " & ex & "reset" & ex & " name = " & ex & "reset" & ex & " value = Reset>"
        ElseIf item = "Drop down list" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<select name = " & ex & ex & ">" & vbCrLf & "<option value = " & ex & ex & ">" & " your list item name" & vbCrLf & "</select>"
        ElseIf item = "Radio Button" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<input name = " & ex & "any name" & ex & " type = " & ex & "radio" & ex & " value = " & ex & "value to return" & ex & ">"
        ElseIf item = "Check Box" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<input name = " & ex & "any name" & ex & " type = " & ex & "checkbox" & ex & " value = " & ex & "value to return" & ex & ">"
        ElseIf item = "Text Area" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<textarea name = " & ex & ex & " cols = 10 rows = 5></textarea>"
        ElseIf item = "Hidden data" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<input name = " & ex & ex & " type = " & ex & " hidden" & ex & " value = " & ex & " Some Data " & ex & ">"
        ElseIf item = "Password" Then
            frmMain.ActiveForm.txtText.SelColor = &H808000
            frmMain.ActiveForm.txtText.SelText = _
           "<input name = " & ex & ex & " type = " & ex & "password" & ex & " cols = 10 size = 10>"
        End If
             
End Sub
Sub ExTag()
             Set ctrparam = frmMain.ActiveForm.RichTextBox1
                            ctrparam.SelItalic = False


                 Select Case ctrparam.SelText
                     Case ""
                          frmMain.Label1.Caption = "Nothing"
                     Case "<body>"
                          frmMain.Label1.Caption = "DOCUMENT BODY <body>...</body>" & vbCrLf & "This section of the document contains the code that is to be displayed on the web browser screen. The first <body> tag is placed immediately after the </head> tag  and the last </body> tag is placed the just before the final </html> tag. The contents of the body tags are all the information to be displayed on screen."
                     Case "</body>"
                          frmMain.Label1.Caption = "DOCUMENT BODY <body>...</body>" & vbCrLf & "This section of the document contains the code that is to be displayed on the web browser screen. The first <body> tag is placed immediately after the </head> tag  and the last </body> tag is placed the just before the final </html> tag. The contents of the body tags are all the information to be displayed on screen."
                     Case "<html>"
                          frmMain.Label1.Caption = "HTML TAGS  <html> ...</html>" & vbCrLf & "These mark up tags enclose the entire document. They can be described as an overcoat. These tags are the first and last to be read by the browser. The basic task of <html>...</html> is to act as marker defining the script type within them."
                     Case "</html>"
                          frmMain.Label1.Caption = "HTML TAGS  <html> ...</html>" & vbCrLf & "These mark up tags enclose the entire document. They can be described as an overcoat. These tags are the first and last to be read by the browser. The basic task of <html>...</html> is to act as marker defining the script type within them."
                     Case "<head>"
                          frmMain.Label1.Caption = "DOCUMENT HEAD <head>...</head>" & vbCrLf & "The head element is used straight after the first<html> tag. Its job is to enclose the main information ABOUT the document. It contains the required <title>...</title> markup tags. The title of the document is placed within these title tags. It can also contain <script>...</script> tags."
                     Case "</head>"
                          frmMain.Label1.Caption = "DOCUMENT HEAD <head>...</head>" & vbCrLf & "The head element is used straight after the first<html> tag. Its job is to enclose the main information ABOUT the document. It contains the required <title>...</title> markup tags. The title of the document is placed within these title tags. It can also contain <script>...</script> tags."
                     Case "<title>"
                          frmMain.Label1.Caption = "HEAD <title>...</title>" & vbCrLf & "The title of the document is placed within these title tags. The title once in the body and title element tags will usually be rendered in the top toolbar of the browser and will be visible while the document is being scrolled through. The title should reflect accurately and concisely the content of the document."
                     Case "<table border = 1>"
                          frmMain.Label1.Caption = "TABLES <table>...</table>" & vbCrLf & "The table element is a feature that is very useful for laying out data in a ordered and readable way. A table can be rendered with multiple columns and rows and a good degree of control is given for the spacing and placement of the cells. The actual table tags enclose a range of elements that combine to define the final output."
                     Case "<table>"
                          frmMain.Label1.Caption = "TABLES <table>...</table>" & vbCrLf & "The table element is a feature that is very useful for laying out data in a ordered and readable way. A table can be rendered with multiple columns and rows and a good degree of control is given for the spacing and placement of the cells. The actual table tags enclose a range of elements that combine to define the final output."
                     Case "</table>"
                          frmMain.Label1.Caption = "TABLES <table>...</table>" & vbCrLf & "The table element is a feature that is very useful for laying out data in a ordered and readable way. A table can be rendered with multiple columns and rows and a good degree of control is given for the spacing and placement of the cells. The actual table tags enclose a range of elements that combine to define the final output."
                     Case "<br>"
                          frmMain.Label1.Caption = "LINE BREAKS: <BR>" & vbCrLf & "The function of the line break element BR is to force a break in a line of text. This element is classed as empty because it does not have an ending tag to operate on particular text segments."
                     Case "&nbsp;"
                          frmMain.Label1.Caption = "SPACE : &nbsp;" & vbCrLf & "This tag creates a non breaking space in a line of text. It is used to create a space between two words without causing a new line. You can use as many of them as you like, ie - Joe&nbsp;&nbsp;&nbsp;Bloggs. - This would generate three spaces"
                     Case "<div>"
                          frmMain.Label1.Caption = "DIVISIONS: <div>...</div>" & vbCrLf & "The DIV or division element is mainly used to define the content of a script section and to thus place it in some logical context or order."
                     Case "</div>"
                          frmMain.Label1.Caption = "DIVISIONS: <div>...</div>" & vbCrLf & "The DIV or division element is mainly used to define the content of a script section and to thus place it in some logical context or order."
                     Case "<blockquote>"
                          frmMain.Label1.Caption = "BLOCKQUOTE  ELEMENT <blockquote>..</blockquote>" & vbCrLf & "The blockquote element is used as a text formatting feature for defining an extended section of text that is a quoted passage. It is designed for somewhat larger blocks of text with many HTML markups. The BLOCKQUOTE text must be contained within other HTML formatting tags such as paragraphs."
                     Case "</blockquote>"
                          frmMain.Label1.Caption = "BLOCKQUOTE  ELEMENT <blockquote>..</blockquote>" & vbCrLf & "The blockquote element is used as a text formatting feature for defining an extended section of text that is a quoted passage. It is designed for somewhat larger blocks of text with many HTML markups. The BLOCKQUOTE text must be contained within other HTML formatting tags such as paragraphs."
                     Case "<p>"
                          frmMain.Label1.Caption = "PARAGRAPHS <P>..</P> " & vbCrLf & "The paragraph element is placed at the beginning of a new block of text.The web browsers interpret this and insert a line between the paragraph and the previous section. Some web browsers will add indentation spaces at the beginning of a new paragraph. The paragraph beginning tag <P> can be used alone at the head of a paragraph"
                     Case "<a href ="
                          frmMain.Label1.Caption = "HYPERTEXT LINKS:<a>..</a>" & vbCrLf & "A hypertext link is defined within the anchor tags <a>.</a>. The attributes that allow linking are the HREF and NAME. Anchor elements with the HREF attribute can contain text or images within the tag structure. These images or text become the label for the jump or link. The labels in a HTML script are usually rendered in a different way to the normal text (usually in a blue colour) and thus indicate to the viewer of the document that a jump can be made by clicking on the text with a mouse."
    
                  End Select
End Sub



