VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmmksite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Publish site"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ForeColor       =   &H00000000&
   Icon            =   "frmmksite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComCtl2.Animation Animation1 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   327680
      FullWidth       =   49
      FullHeight      =   25
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   4920
      Pattern         =   "*.gif;*.jpg;*.swf"
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      Pattern         =   "*.html;*.exe;*.js;*.zip"
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Drive to Save Site in"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "frmmksite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim szee As Double





Private Sub Command1_Click()
On Error GoTo fred
        Dim stf As String
        Dim source
        Dim dest
        Dim sourcea
        Dim desta
        Dim fl
        Dim ftl
        Dim maj
        Dim min
        fl = File1.ListIndex
        stf = GetSetting(App.Title, "Settings", "site", "")
     If szee > 1.3 And Drive1.Drive = "a:" Then
        Label4.Caption = "This file is too big for a 1.44 disk. Please select another drive or save with a zip utility."
        Exit Sub
     Else
               Label4.Caption = ""
               Animation1.Visible = True
       Animation1.Open App.Path & "\download.avi"
       Animation1.Play
        MkDir (Drive1.Drive & "\" & stf & "-Site")
        MkDir (Drive1.Drive & "\" & stf & "-Site" & "\images")
        maj = Drive1.Drive & "\" & stf & "-Site"
        For fl = 0 To File1.ListCount
        source = App.Path & "\" & stf & "\" & File1.List(fl)
        dest = maj & "\" & File1.List(fl)
        FileCopy source, dest
        Next
        min = Drive1.Drive & "\" & stf & "-Site" & "\images\"
        For ftl = 0 To File2.ListCount
        sourcea = App.Path & "\" & stf & "\images\" & File2.List(ftl)
        desta = min & File2.List(ftl)
        FileCopy sourcea, desta
        Next
        
            Animation1.Close
            Animation1.Visible = False
        Label2.Caption = stf & "  PUBLISHED!"
        Label4.Caption = "You will find your completed site in " & UCase(Drive1.Drive) & " Drive"
    End If
    'Unload Me
fred:
 
      If Err.Number = 71 Then
        Label4.Caption = "Please insert a formatted disk in the selected drive"
      ElseIf Err.Number = 61 Then
        Label4.Caption = "This Disk is Full, some files where not copied"
      Else
        Resume Next
      End If

End Sub





Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Label4.Caption = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
     Dim sf As String
     sf = GetSetting(App.Title, "Settings", "site", "")
     File1.Path = App.Path & "\" & sf
     File2.Path = App.Path & "\" & sf & "\images"
     Label2.Caption = "The Site you are about to publish is : " & sf
     
     App.HelpFile = App.Path & "\WEBMAGIC.HLP"
     Me.HelpContextID = 39
Dim bl
Dim cl
Dim dl
Dim dtl
Dim xt
Dim zt
Dim sze As Long

    bl = FileLen(App.Path & "\" & sf & "\" & File1.List(dl))
    cl = FileLen(App.Path & "\" & sf & "\images\" & File2.List(dtl))
    dl = File1.ListIndex
    dtl = File2.ListIndex
    xt = 0
    For dl = 0 To File1.ListCount - 1
     xt = xt + bl
     Next
     
     zt = 0
    For dtl = 0 To File2.ListCount - 1
     zt = zt + cl
    Next
    sze = (xt + zt) / 1024
    szee = Format(sze / 1024, "#.##")
     Label3.Caption = "Your Site size is : - " & szee & " MB"
     


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Publish your site to a local folder ready for uploading. Press F1 for help"
End Sub

Private Sub Timer1_Timer()

End Sub
