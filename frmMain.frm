VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "WebMagic"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   870
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   15180
      TabIndex        =   12
      Top             =   615
      Width           =   15240
      Begin TabDlg.SSTab SSTab1 
         Height          =   855
         Left            =   60
         TabIndex        =   13
         Top             =   0
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   1508
         _Version        =   327680
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Pa&ge Formatting"
         TabPicture(0)   =   "frmMain.frx":08CA
         Tab(0).ControlCount=   12
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdPback"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdLink"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdCen"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdLin"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdSpace"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdCom"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdDiv"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdPara"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdBlock"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdRule"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdFrame"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdCenc"
         Tab(0).Control(11).Enabled=   0   'False
         TabCaption(1)   =   "Fo&nt Formatting"
         TabPicture(1)   =   "frmMain.frx":08E6
         Tab(1).ControlCount=   10
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdH6"
         Tab(1).Control(0).Enabled=   -1  'True
         Tab(1).Control(1)=   "cmdH5"
         Tab(1).Control(1).Enabled=   -1  'True
         Tab(1).Control(2)=   "cmdH4"
         Tab(1).Control(2).Enabled=   -1  'True
         Tab(1).Control(3)=   "cmdH3"
         Tab(1).Control(3).Enabled=   -1  'True
         Tab(1).Control(4)=   "cmdH2"
         Tab(1).Control(4).Enabled=   -1  'True
         Tab(1).Control(5)=   "cmdH1"
         Tab(1).Control(5).Enabled=   -1  'True
         Tab(1).Control(6)=   "cmdUln"
         Tab(1).Control(6).Enabled=   -1  'True
         Tab(1).Control(7)=   "cmdItl"
         Tab(1).Control(7).Enabled=   -1  'True
         Tab(1).Control(8)=   "cmdBld"
         Tab(1).Control(8).Enabled=   -1  'True
         Tab(1).Control(9)=   "cmdFnt"
         Tab(1).Control(9).Enabled=   -1  'True
         TabCaption(2)   =   "&Insert Blocks"
         TabPicture(2)   =   "frmMain.frx":0902
         Tab(2).ControlCount=   9
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdScroll"
         Tab(2).Control(0).Enabled=   -1  'True
         Tab(2).Control(1)=   "cmdMed"
         Tab(2).Control(1).Enabled=   -1  'True
         Tab(2).Control(2)=   "cmdLst"
         Tab(2).Control(2).Enabled=   -1  'True
         Tab(2).Control(3)=   "cmdFormel"
         Tab(2).Control(3).Enabled=   -1  'True
         Tab(2).Control(4)=   "cmdForm"
         Tab(2).Control(4).Enabled=   -1  'True
         Tab(2).Control(5)=   "cmdImgm"
         Tab(2).Control(5).Enabled=   -1  'True
         Tab(2).Control(6)=   "cmdImg"
         Tab(2).Control(6).Enabled=   -1  'True
         Tab(2).Control(7)=   "cmdTblm"
         Tab(2).Control(7).Enabled=   -1  'True
         Tab(2).Control(8)=   "cmdTbl"
         Tab(2).Control(8).Enabled=   -1  'True
         TabCaption(3)   =   "Head &Section"
         TabPicture(3)   =   "frmMain.frx":091E
         Tab(3).ControlCount=   2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdJav"
         Tab(3).Control(0).Enabled=   -1  'True
         Tab(3).Control(1)=   "cmdMet"
         Tab(3).Control(1).Enabled=   -1  'True
         TabCaption(4)   =   "P&ublish Site"
         TabPicture(4)   =   "frmMain.frx":093A
         Tab(4).ControlCount=   2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdPub"
         Tab(4).Control(0).Enabled=   -1  'True
         Tab(4).Control(1)=   "Label7"
         Tab(4).Control(1).Enabled=   0   'False
         Begin VB.CommandButton cmdCenc 
            Caption         =   "Center Close"
            Height          =   405
            Left            =   3780
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   360
            Width           =   1275
         End
         Begin VB.CommandButton cmdFrame 
            Caption         =   "Frame Sets"
            Height          =   405
            Left            =   12075
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   360
            Width           =   1035
         End
         Begin VB.CommandButton cmdScroll 
            Caption         =   "Scrolling Text"
            Height          =   405
            Left            =   -65460
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdPub 
            Height          =   435
            Left            =   -72420
            Picture         =   "frmMain.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   360
            Width           =   555
         End
         Begin VB.CommandButton cmdRule 
            Caption         =   "Rule"
            Height          =   405
            Left            =   5130
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   360
            Width           =   645
         End
         Begin VB.CommandButton cmdMed 
            Caption         =   "Flash Movie"
            Height          =   405
            Left            =   -63720
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton cmdBlock 
            Caption         =   "Blockquote"
            Height          =   405
            Left            =   10815
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdPara 
            Caption         =   "Paragraph"
            Height          =   405
            Left            =   9660
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdDiv 
            Caption         =   "Div"
            Height          =   405
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdJav 
            Caption         =   "Javascript Tags"
            Height          =   405
            Left            =   -71610
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   360
            Width           =   1845
         End
         Begin VB.CommandButton cmdMet 
            Caption         =   "Meta Tags"
            Height          =   405
            Left            =   -74280
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton cmdLst 
            Caption         =   "List"
            Height          =   405
            Left            =   -66570
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   360
            Width           =   1065
         End
         Begin VB.CommandButton cmdFormel 
            Caption         =   "Form Elements"
            Height          =   405
            Left            =   -68130
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            Width           =   1485
         End
         Begin VB.CommandButton cmdForm 
            Caption         =   "Form Structure"
            Height          =   405
            Left            =   -69450
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   360
            Width           =   1275
         End
         Begin VB.CommandButton cmdImgm 
            Caption         =   "Image Map"
            Height          =   405
            Left            =   -70680
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   360
            Width           =   1125
         End
         Begin VB.CommandButton cmdImg 
            Caption         =   "Image"
            Height          =   405
            Left            =   -71850
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   360
            Width           =   1125
         End
         Begin VB.CommandButton cmdTblm 
            Caption         =   "Modify Table"
            Height          =   405
            Left            =   -73320
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   360
            Width           =   1305
         End
         Begin VB.CommandButton cmdTbl 
            Caption         =   "Table Structure"
            Height          =   405
            Left            =   -74820
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdCom 
            Caption         =   "Comment"
            Height          =   405
            Left            =   7980
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdSpace 
            Caption         =   "Space"
            Height          =   405
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   825
         End
         Begin VB.CommandButton cmdLin 
            Caption         =   "Line Break"
            Height          =   405
            Left            =   5850
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton cmdCen 
            Caption         =   "Center Open"
            Height          =   405
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "Links"
            Height          =   405
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   360
            Width           =   675
         End
         Begin VB.CommandButton cmdPback 
            Caption         =   "Page Background "
            Height          =   405
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   1650
         End
         Begin VB.CommandButton cmdH6 
            Caption         =   "H6"
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
            Left            =   -66480
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdH5 
            Caption         =   "H5"
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
            Left            =   -66960
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdH4 
            Caption         =   "H4"
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
            Left            =   -67440
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdH3 
            Caption         =   "H3"
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
            Left            =   -67920
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdH2 
            Caption         =   "H2"
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
            Left            =   -68400
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdH1 
            Caption         =   "H1"
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
            Left            =   -68880
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdUln 
            Height          =   375
            Left            =   -70680
            Picture         =   "frmMain.frx":0D98
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdItl 
            Height          =   375
            Left            =   -71160
            Picture         =   "frmMain.frx":0E9A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdBld 
            Height          =   375
            Left            =   -71640
            Picture         =   "frmMain.frx":0F9C
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdFnt 
            Caption         =   "Font Tags"
            Height          =   375
            Left            =   -74280
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Publish your completed site to a local folder ready for uploading to your Website"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -71730
            TabIndex        =   49
            Top             =   420
            Width           =   9405
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFF00&
      Height          =   8700
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   2940
      TabIndex        =   11
      Top             =   1485
      Width           =   3000
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         ItemData        =   "frmMain.frx":109E
         Left            =   660
         List            =   "frmMain.frx":10A0
         TabIndex        =   59
         Top             =   5460
         Width           =   1815
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   3195
         Left            =   60
         TabIndex        =   50
         Top             =   1500
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   5636
         _Version        =   327682
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         Pattern         =   "*.html"
         TabIndex        =   47
         Top             =   4320
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdPre 
         Height          =   705
         Left            =   60
         Picture         =   "frmMain.frx":10A2
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   705
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   1440
         TabIndex        =   58
         Top             =   4440
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Your Sites. Click on a site name to Open it"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   4980
         Width           =   2715
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pages from current site. Click on page name to edit"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   150
         TabIndex        =   46
         Top             =   990
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview page in default Browser"
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   840
         TabIndex        =   45
         Top             =   210
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H80000004&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      Begin VB.CommandButton cmdDel 
         Height          =   375
         Left            =   2220
         Picture         =   "frmMain.frx":14E4
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Delete Current Site"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSavep 
         Height          =   375
         Left            =   1680
         Picture         =   "frmMain.frx":15E6
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Save Page to Current Site"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdNewPage 
         Height          =   375
         Left            =   960
         Picture         =   "frmMain.frx":16E8
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Add a page to this site"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdExist 
         Height          =   375
         Left            =   540
         Picture         =   "frmMain.frx":1B2A
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Open an Existing Site"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCreate 
         Height          =   375
         Left            =   120
         Picture         =   "frmMain.frx":1C2C
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Create a New Site"
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8490
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   135
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6780
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   135
         Width           =   1575
      End
      Begin VB.CommandButton cmdPaste 
         Height          =   375
         Left            =   3990
         Picture         =   "frmMain.frx":24F6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   375
         Left            =   3600
         Picture         =   "frmMain.frx":2A28
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCut 
         Height          =   375
         Left            =   3210
         Picture         =   "frmMain.frx":2F5A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   2700
         Picture         =   "frmMain.frx":348C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Site Is : -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9300
         TabIndex        =   10
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11400
         TabIndex        =   9
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000004&
         Caption         =   "Change Editing color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4620
         TabIndex        =   8
         Top             =   150
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   420
         Left            =   5970
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   3180
      End
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   10185
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   21246
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "20/12/2001"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "23:45"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7950
      Top             =   4050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":39BE
            Key             =   "page"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3CD8
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3FF2
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewSite 
         Caption         =   "&Create New Site"
      End
      Begin VB.Menu mnuFileAddNewPage 
         Caption         =   "&Add New Page to Site"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenExSite 
         Caption         =   "&Open Existing Site"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import from other source"
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Page"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print Code"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditChg 
         Caption         =   "Change &Panel Scheme"
      End
   End
   Begin VB.Menu mnuPage 
      Caption         =   "&Page"
      Begin VB.Menu mnuPagePgBack 
         Caption         =   "&Page Background"
      End
      Begin VB.Menu mnuPageLinks 
         Caption         =   "&Links"
      End
      Begin VB.Menu mnuPageCopen 
         Caption         =   "Center &Open"
      End
      Begin VB.Menu mnuPageCclose 
         Caption         =   "Center &Close"
      End
      Begin VB.Menu mnuPageRule 
         Caption         =   "&Rule"
      End
      Begin VB.Menu mnuPageLineBrk 
         Caption         =   "Line &Break"
      End
      Begin VB.Menu mnuPageSpace 
         Caption         =   "&Space"
      End
      Begin VB.Menu mnuPageComment 
         Caption         =   "Co&mment"
      End
      Begin VB.Menu mnuPageDiv 
         Caption         =   "&Div"
      End
      Begin VB.Menu mnuPagePara 
         Caption         =   "P&aragraph"
      End
      Begin VB.Menu mnuPageBlock 
         Caption         =   "Block&quote"
      End
      Begin VB.Menu mnuPageFrame 
         Caption         =   "Fram&e Sets"
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "F&ont"
      Begin VB.Menu mnuFontTag 
         Caption         =   "F&ont Tags"
      End
      Begin VB.Menu mnuFontBld 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuFontItal 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuFontUnder 
         Caption         =   "&Underline"
      End
      Begin VB.Menu mnuFontH1 
         Caption         =   "H&1"
      End
      Begin VB.Menu mnuFontH2 
         Caption         =   "H&2"
      End
      Begin VB.Menu mnuFontH3 
         Caption         =   "H&3"
      End
      Begin VB.Menu mnuFontH4 
         Caption         =   "H&4"
      End
      Begin VB.Menu mnuFontH5 
         Caption         =   "H&5"
      End
      Begin VB.Menu mnuFontH6 
         Caption         =   "H&6"
      End
   End
   Begin VB.Menu mnuBlock 
      Caption         =   "&Blocks"
      Begin VB.Menu mnuBlockTblStruct 
         Caption         =   "&Table Structure"
      End
      Begin VB.Menu mnuBlockModTbl 
         Caption         =   "&Modify Table"
      End
      Begin VB.Menu mnuBlockImg 
         Caption         =   "&Image"
      End
      Begin VB.Menu mnuBlockImgMap 
         Caption         =   "Im&age Map"
      End
      Begin VB.Menu mnuBlockFrmStruct 
         Caption         =   "&Form Structure"
      End
      Begin VB.Menu mnuBlockFrmEls 
         Caption         =   "Form &Elements"
      End
      Begin VB.Menu mnuBlockLst 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuBlockScroll 
         Caption         =   "&Scrolling Text"
      End
      Begin VB.Menu mnuBlockFlsh 
         Caption         =   "Flas&h"
      End
   End
   Begin VB.Menu mnuHead 
      Caption         =   "He&ad"
      Begin VB.Menu mnuHeadMet 
         Caption         =   "&Meta Tags"
      End
      Begin VB.Menu mnuHeadJav 
         Caption         =   "&Javascript Tags"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About WebMagic..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim nscp, sit, sitpth As String
Dim aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk, mm, nn


Private Sub cmdBld_Click()
        TagFmt (2)
End Sub

Private Sub cmdBld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (2)
End Sub

Private Sub cmdBlock_Click()
      TagFmt (16)
End Sub

Private Sub cmdBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdCen_Click()
       TagFmt (1)
End Sub

Private Sub cmdCen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         Instruct (3)
End Sub

Private Sub cmdCenc_Click()
        TagFmt (19)
End Sub

Private Sub cmdCenc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdCom_Click()
       TagFmt (18)
End Sub

Private Sub cmdCom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Instruct (3)
End Sub

Private Sub cmdCopy_Click()
        mnuEditCopy_Click
End Sub

Private Sub cmdCreate_Click()
        frmfolder.Show
End Sub

Private Sub cmdCut_Click()
       mnuEditCut_Click
End Sub

Private Sub cmdDel_Click()
On Error Resume Next
Dim DelPth As String              ' Define the folder to be deleted
    DelPth = App.Path & "\" & Label3.Caption
If lDocumentCount <> 0 Then
   MsgBox "Please close all pages before deleting"
Else
Dim Msg, Style, Response
Msg = "Are you sure you want to delete this site ?"   ' Define message.
Style = vbYesNo + vbExclamation
        ' Display message.
Response = MsgBox(Msg, Style)
If Response = vbYes Then    ' User chose Yes.

                Kill sitpth & "\*.html" ' remove html pages
                Kill sitpth & "\*.web"  ' remove .web pages
                Kill sitpth & "\images\" & "*.*" ' remove any images
                RmDir DelPth & "\images" ' Remove images folder
                RmDir DelPth             ' Remove site folder
                Dir1.Refresh

                List1.Clear
   gg = Len(App.Path) + 2
    ff = Dir1.ListIndex
      ff = ff + 1
      For ff = 0 To Dir1.ListCount - 1
           hh = Len(Dir1.List(ff))
           ii = (hh - gg) + 1
          List1.AddItem Mid(Dir1.List(ff), gg, ii)
      Next ff
      jj = List1.ListCount
             If jj = 0 Then
               kk = ""
             Else
               kk = List1.List(kk)
             End If
             Label3.Caption = kk
             Label6.Caption = "Pages from " & UCase(kk) & " Click on Page Name to edit"
             sit = kk
             sitpth = App.Path & "\" & sit
             frmMain.File1.Path = sitpth
             frmMain.File1.Refresh
             frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create a tree.
    Set nodX = TreeView1.Nodes.Add(, , "r", sitpth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = TreeView1.Nodes.Add("r", tvwChild, "C3", sit, "open")
    For aa = 0 To File1.ListCount - 1
     Set nodX = TreeView1.Nodes.Add("C3", tvwChild, , File1.List(aa), "page")
    Next aa
    nodX.EnsureVisible  ' Show all nodes.
             SaveSetting App.Title, "Settings", "site", kk
            
Else
    Exit Sub
End If
End If
          
End Sub

Private Sub cmdDiv_Click()
       TagFmt (15)
End Sub

Private Sub cmdDiv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdExist_Click()
        frmopensite.Show
End Sub



Private Sub cmdFnt_Click()
    frmfont.Show
End Sub



Private Sub cmdFnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Instruct (2)
End Sub

Private Sub cmdForm_Click()
      frmformtag.Show
End Sub

Private Sub cmdForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdFormel_Click()
      frmform.Show
End Sub

Private Sub cmdFormel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdFrame_Click()
         frmframe.Show
End Sub

Private Sub cmdFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         Instruct (7)
End Sub

Private Sub cmdH1_Click()
       TagFmt (5)
End Sub

Private Sub cmdH1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (2)
End Sub

Private Sub cmdH2_Click()
       TagFmt (6)
End Sub

Private Sub cmdH2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         Instruct (2)
End Sub

Private Sub cmdH3_Click()
      TagFmt (7)
End Sub

Private Sub cmdH3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdH4_Click()
       TagFmt (8)
End Sub

Private Sub cmdH4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (2)
End Sub

Private Sub cmdH5_Click()
      TagFmt (9)
End Sub

Private Sub cmdH5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdH6_Click()
       TagFmt (10)
End Sub

Private Sub cmdH6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdImg_Click()
       frmimage.Show
End Sub

Private Sub cmdImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdImgm_Click()
        frmmap.Show
End Sub

Private Sub cmdImgm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdItl_Click()
       TagFmt (3)
End Sub

Private Sub cmdItl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (2)
End Sub

Private Sub cmdJav_Click()
       TagFmt (17)
End Sub

Private Sub cmdJav_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (6)
End Sub

Private Sub cmdLin_Click()
       TagFmt (11)
End Sub

Private Sub cmdLin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdLink_Click()
       frmlink.Show
End Sub

Private Sub cmdLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Instruct (2)
End Sub

Private Sub cmdLst_Click()
      frmlist.Show
End Sub

Private Sub cmdLst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdMed_Click()
       frmmedia.Show
End Sub

Private Sub cmdMed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdMet_Click()
      frmmeta.Show
End Sub

Private Sub cmdMet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (6)
End Sub

Private Sub cmdNewPage_Click()
        LoadNewDoc
        frmsave1.Show
End Sub

Private Sub cmdPara_Click()
         TagFmt (14)
End Sub

Private Sub cmdPara_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (4)
End Sub

Private Sub cmdPaste_Click()
     mnuEditPaste_Click
End Sub

Private Sub cmdPback_Click()
       frmbkgrnd.Show
End Sub



Private Sub cmdPback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Instruct (1)
End Sub

Private Sub cmdPrint_Click()
    mnuFilePrint_Click
End Sub

Private Sub cmdPub_Click()
        frmmksite.Show
End Sub

Private Sub cmdRule_Click()
       frmrule.Show
End Sub

Private Sub cmdRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Instruct (3)
End Sub

Private Sub cmdSavep_Click()
      mnuFileSave_Click
End Sub

Private Sub cmdScroll_Click()
      frmmarq.Show
End Sub

Private Sub cmdScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (3)
End Sub

Private Sub cmdSpace_Click()
       TagFmt (12)
End Sub



Private Sub cmdSpace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (3)
End Sub

Private Sub cmdTbl_Click()
     setTable
End Sub

Private Sub cmdTbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (3)
End Sub

Private Sub cmdTblm_Click()
         frmtablemod.Show
End Sub

Private Sub cmdTblm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Instruct (5)
End Sub

Private Sub cmdUln_Click()
       TagFmt (4)
End Sub

Private Sub cmdUln_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Instruct (2)
End Sub

Private Sub Combo1_Click()
        Select Case Combo1.Text
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
         End Select
         Me.ActiveForm.txtText.SelColor = Picture5.BackColor
End Sub




Private Sub cmdPre_Click()
On Error Resume Next
      Dim m As Long
           Dim bgn As String
     Dim sts As String
          bgn = frmMain.ActiveForm.Caption
          sts = GetSetting(App.Title, "Settings", "site", "")
    If Me.ActiveForm.Caption <> "New Web" Then
           frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sts & "\" & bgn & ".html", 1
    End If
If Me.ActiveForm.Caption = "New Web" Then
   MsgBox "Please save your page before previewing"
   Exit Sub
Else
m = ShellExecute(Me.hwnd, "open", App.Path & "\" & sts & "\" & bgn & ".html", "", App.Path, 1)
End If

End Sub
Private Sub File1_Click()
    sit = GetSetting(App.Title, "Settings", "site", "")
    sitpth = App.Path & "\" & sit
        Dim frmD As frmDocument
        Dim nm
        Dim nma
            nm = Len(File1.filename) - 5
            nma = Left(File1.filename, nm)
    Set frmD = New frmDocument
    lDocumentCount = lDocumentCount + 1
        With File1
          frmD.txtText.LoadFile App.Path & "\" & sit & "\" & nma & ".web"
          frmD.Caption = nma
          'frmMain.Text1.Text = nma
        End With
End Sub

Private Sub List1_Click()
On Error Resume Next
   Dim li
   Dim Newpth
       li = List1.ListIndex
          If frmMain.ActiveForm.Caption = "New Web" Then
            SaveSetting App.Title, "Settings", "site", List1.List(li)
            Newpth = App.Path & "\" & List1.List(li)
            frmMain.File1.Path = Newpth
            frmMain.File1.Refresh
       frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create a tree.
    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", Newpth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "C3", List1.List(li), "open")
    For mm = 0 To frmMain.File1.ListCount - 1
     Set nodX = frmMain.TreeView1.Nodes.Add("C3", tvwChild, , frmMain.File1.List(mm), "page")
    Next mm
    nodX.EnsureVisible  ' Show all nodes.
       Label3.Caption = List1.List(li)
       Label6.Caption = "Pages from " & UCase(List1.List(li)) & " Click on Page Name to edit"
   Else
             MsgBox "These pages are already saved to the current site. Please close them first"
             Exit Sub
  End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMain.sbStatusBar.Panels(1).Text = "Opens existing site."
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 13485)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 10230)
    Label3.Caption = GetSetting(App.Title, "Settings", "site", "")
    Picture5.BackColor = GetSetting(App.Title, "Settings", "edcol", &H0&)
    bb = GetSetting(App.Title, "Settings", "bcol", &H0&)
    cc = GetSetting(App.Title, "Settings", "fcol", &HFF00&)
    nn = GetSetting(App.Title, "Settings", "edtxt", "black")
    sit = GetSetting(App.Title, "Settings", "site", "")
    sitpth = App.Path & "\" & sit
 
       Picture1.BackColor = bb
       Picture2.BackColor = bb
       Picture3.BackColor = bb
       Label9.BackColor = bb
       Shape1.BorderColor = cc
       Label2.ForeColor = cc
       Label3.ForeColor = cc
       Label9.ForeColor = cc
       Label10.ForeColor = cc
       SSTab1.BackColor = bb
       Label1.ForeColor = cc
       Label6.ForeColor = cc

            With Combo1
             .AddItem "black"
             .AddItem "blue"
             .AddItem "fuchsia"
             .AddItem "gray"
             .AddItem "green"
             .AddItem "maroon"
             .AddItem "navy"
             .AddItem "olive"
             .AddItem "purple"
             .AddItem "red"
             .AddItem "silver"
             .AddItem "teal"
             .Text = nn
          End With
          
          File1.Path = sitpth
          Dir1.Path = App.Path

    Dim nodX As Node    ' Create a tree.
    Set nodX = TreeView1.Nodes.Add(, , "r", sitpth, "open")
    nodX.EnsureVisible  ' Show all nodes.
    Set nodX = TreeView1.Nodes.Add("r", tvwChild, "C3", sit, "open")
    For aa = 0 To File1.ListCount - 1
     Set nodX = TreeView1.Nodes.Add("C3", tvwChild, , File1.List(aa), "page")
    Next aa
    nodX.EnsureVisible  ' Show all nodes.
    
    gg = Len(App.Path) + 2
    ff = Dir1.ListIndex
      ff = ff + 1
      For ff = 0 To Dir1.ListCount - 1
           hh = Len(Dir1.List(ff))
           ii = (hh - gg) + 1
          List1.AddItem Mid(Dir1.List(ff), gg, ii)
      Next ff
      jj = List1.ListCount
      
      Label6.Caption = "Pages from " & UCase(sit) & " Click on Page Name to edit"
 
       LoadNewDoc
       frmsplash.Show
End Sub


Private Sub LoadNewDoc()
    'Static lDocumentCount As Long
    Dim frmD As frmDocument


    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "New Web"
    frmD.Show
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lDocumentCount > 1 Then
      MsgBox "Please close All open pages before Closing"
       Cancel = -1
    Else
      Cancel = 0
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "site", Label3.Caption
        SaveSetting App.Title, "Settings", "edcol", Picture5.BackColor
        SaveSetting App.Title, "Settings", "edtxt", Combo1.Text
    End If

End Sub



Private Sub mnuBlockFlsh_Click()
        cmdMed_Click
End Sub

Private Sub mnuBlockFrmEls_Click()
        cmdFormel_Click
End Sub

Private Sub mnuBlockFrmStruct_Click()
      cmdForm_Click
End Sub

Private Sub mnuBlockImg_Click()
      cmdImg_Click
End Sub

Private Sub mnuBlockImgMap_Click()
       cmdImgm_Click
End Sub

Private Sub mnuBlockLst_Click()
       cmdLst_Click
End Sub

Private Sub mnuBlockModTbl_Click()
       cmdTblm_Click
End Sub

Private Sub mnuBlockScroll_Click()
        cmdScroll_Click
End Sub

Private Sub mnuBlockTblStruct_Click()
        cmdTbl_Click
End Sub

Private Sub mnuEditChg_Click()
     frmcol.Show
End Sub

Private Sub mnuFileAddNewPage_Click()
      cmdNewPage_Click
End Sub

Private Sub mnuFileImport_Click()
On Error GoTo fred
        With dlgCommonDialog
          .Filter = "html Files(*.html)|*.html|htm Files(*.htm)|*.htm"
          .FilterIndex = 0
          .ShowOpen
          Me.ActiveForm.txtText.LoadFile .filename, 1
        End With
fred:
Exit Sub
End Sub

Private Sub mnuFileNewSite_Click()
      cmdCreate_Click
End Sub

Private Sub mnuFileOpenExSite_Click()
      cmdExist_Click
End Sub

Private Sub mnuFontBld_Click()
      cmdBld_Click
End Sub

Private Sub mnuFontH1_Click()
     cmdH1_Click
End Sub

Private Sub mnuFontH2_Click()
     cmdH2_Click
End Sub

Private Sub mnuFontH3_Click()
     cmdH3_Click
End Sub

Private Sub mnuFontH4_Click()
     cmdH4_Click
End Sub

Private Sub mnuFontH5_Click()
     cmdH5_Click
End Sub

Private Sub mnuFontH6_Click()
     cmdH6_Click
End Sub

Private Sub mnuFontItal_Click()
     cmdItl_Click
End Sub

Private Sub mnuFontTag_Click()
      cmdFnt_Click
End Sub

Private Sub mnuFontUnder_Click()
    cmdUln_Click
End Sub

Private Sub mnuHeadJav_Click()
      cmdJav_Click
End Sub

Private Sub mnuHeadMet_Click()
      cmdMet_Click
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuEditCopy_Click()
       Clipboard.Clear
       Clipboard.SetText Me.ActiveForm.txtText.SelRTF
End Sub


Private Sub mnuEditCut_Click()
Clipboard.Clear
       Clipboard.SetText Me.ActiveForm.txtText.SelRTF
       Me.ActiveForm.txtText.SelRTF = ""
End Sub


Private Sub mnuEditPaste_Click()
    Me.ActiveForm.txtText.SelRTF = Clipboard.GetText()
End Sub

Private Sub mnuFileSave_Click()
On Error Resume Next
    Dim sit As String
    Dim nme As String
     sit = GetSetting(App.Title, "Settings", "site", "")
     nme = frmMain.ActiveForm.Caption
    If frmMain.ActiveForm.Caption <> "New Web" Then
       frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".html", 1
       frmMain.ActiveForm.txtText.SaveFile App.Path & "\" & sit & "\" & nme & ".web"
    Else
       frmsave1.Show
    End If
End Sub



Private Sub mnuFilePrint_Click()
    dlgCommonDialog.CancelError = True
    On Error GoTo cease
    dlgCommonDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If Me.ActiveForm.txtText.SelLength = 0 Then
        dlgCommonDialog.Flags = dlgCommonDialog.Flags + cdlPDAllPages
    Else
        dlgCommonDialog.Flags = dlgCommonDialog.Flags + cdlPDSelection
    End If
    dlgCommonDialog.ShowPrinter
    Printer.Print ""
    Me.ActiveForm.txtText.SelPrint dlgCommonDialog.hDC
    Printer.EndDoc
cease:
    Exit Sub
End Sub



Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub




Private Sub mnuPageBlock_Click()
         cmdBlock_Click
End Sub

Private Sub mnuPageCclose_Click()
        cmdCenc_Click
End Sub

Private Sub mnuPageComment_Click()
        cmdCom_Click
End Sub

Private Sub mnuPageCopen_Click()
         cmdCen_Click
End Sub

Private Sub mnuPageDiv_Click()
      cmdDiv_Click
End Sub

Private Sub mnuPageFrame_Click()
     cmdFrame_Click
End Sub

Private Sub mnuPageLineBrk_Click()
      cmdLin_Click
End Sub

Private Sub mnuPageLinks_Click()
     cmdLink_Click
End Sub

Private Sub mnuPagePara_Click()
      cmdPara_Click
End Sub

Private Sub mnuPagePgBack_Click()
      cmdPback_Click
End Sub

Private Sub mnuPageRule_Click()
        cmdRule_Click
End Sub

Private Sub mnuPageSpace_Click()
        cmdSpace_Click
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       frmMain.sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       frmMain.sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       frmMain.sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub Picture3_Resize()
      If Me.WindowState <> 1 Then
        SSTab1.Width = (Picture3.Width - SSTab1.Left) - 50
      Else
        Exit Sub
      End If
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       frmMain.sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
     If Node.Key = "C3" Then
        Node.Image = "closed"
     End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
     If Node.Key = "C3" Then
        Node.Image = "open"
     End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
On Error GoTo fred
    sit = GetSetting(App.Title, "Settings", "site", "")
    sitpth = App.Path & "\" & sit
        Dim frmD As frmDocument
        Dim nm
        Dim nma
            nm = Len(TreeView1.SelectedItem.Text) - 5
            nma = Left(TreeView1.SelectedItem.Text, nm)
If nma <> Me.ActiveForm.Caption Then
    Set frmD = New frmDocument
    lDocumentCount = lDocumentCount + 1
        'With File1
          frmD.txtText.LoadFile App.Path & "\" & sit & "\" & nma & ".web"
          frmD.Caption = nma
          'frmMain.Text1.Text = nma
        'End With
  Else
          MsgBox "You already have this page open"
         Exit Sub

End If
fred:
       If Err.Number = 75 Then
          Unload frmD
       Else
         Resume Next
       End If
End Sub
