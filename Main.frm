VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DD32A320-6E5E-44C8-BCE6-5908CA400231}#1.0#0"; "agRichEdit.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "(New File)"
   ClientHeight    =   8460
   ClientLeft      =   135
   ClientTop       =   690
   ClientWidth     =   11175
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11175
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   6360
      Left            =   3000
      ScaleHeight     =   6360
      ScaleWidth      =   7575
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7572
      Begin agRichEditBox.agRichEdit agEditor 
         Height          =   5352
         Left            =   2520
         TabIndex        =   34
         Top             =   1200
         Width           =   5856
         _ExtentX        =   10319
         _ExtentY        =   9446
         Version         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         ViewMode        =   1
         TextLimit       =   100000
         TrapTab         =   0   'False
         AutoURLDetect   =   0   'False
         TextOnly        =   -1  'True
         DisableNoScroll =   -1  'True
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   3600
      End
   End
   Begin VB.Timer tmrRevertTips 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   9240
      Top             =   7920
   End
   Begin MSComctlLib.ImageList ilsFileIcons 
      Left            =   720
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":179A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BEC
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":203E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2490
            Key             =   "textfile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":28E2
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3610
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4804
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBrowser 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   7560
      Left            =   0
      ScaleHeight     =   7560
      ScaleWidth      =   2415
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   2412
      Begin VB.CommandButton btnScrollToTop 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   1848
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":495E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Scroll To Top"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnCurrentDirectory 
         Appearance      =   0  'Flat
         Caption         =   "<>"
         Height          =   264
         Left            =   1584
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Jump to the directory containing your open file... (Ctrl+F5)"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnDeleteSelected 
         Appearance      =   0  'Flat
         Caption         =   "X"
         Height          =   264
         Left            =   1320
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Delete File (Del)"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnRefresh 
         Appearance      =   0  'Flat
         Caption         =   "R"
         Height          =   264
         Left            =   1056
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Refresh Files (F5)"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnSort 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   792
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":4AA8
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Reverse the sort order (Ctrl+H)"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnFolderUp 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   528
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":4BAA
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Go up a directory (Left arrow key or Ctrl+F6)"
         Top             =   420
         Width           =   264
      End
      Begin VB.CommandButton btnPathForward 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   264
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":4F34
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Go forward a directory (Alt+Right)"
         Top             =   420
         Width           =   264
      End
      Begin VB.ComboBox cboPath 
         Height          =   315
         ItemData        =   "Main.frx":507E
         Left            =   0
         List            =   "Main.frx":5080
         TabIndex        =   4
         Text            =   "*"
         ToolTipText     =   "Type a directory into here, or select one below.  You can even specify a file extension.  Example:   c:\windows\*.dll"
         Top             =   100
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvwBrowser 
         Height          =   4335
         Left            =   0
         TabIndex        =   5
         Tag             =   "c:\test\"
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7646
         SortKey         =   1
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         Icons           =   "ilsFileIcons"
         SmallIcons      =   "ilsFileIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "filename"
            Object.Tag             =   "0"
            Text            =   "Um"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "IsDirectory"
            Text            =   "IsDirectory"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton btnPathBack 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   0
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":5082
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Go back a directory (Alt+Left)"
         Top             =   420
         Width           =   264
      End
      Begin VB.Label lblDivider 
         BackStyle       =   0  'Transparent
         Height          =   25005
         Left            =   2295
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   9480
      Top             =   1680
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   12
   End
   Begin VB.PictureBox picToolBox 
      Align           =   1  'Align Top
      ClipControls    =   0   'False
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   11115
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton btnOptions 
         Appearance      =   0  'Flat
         Caption         =   "Opts."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Options..."
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnNextFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4200
         Picture         =   "Main.frx":51CC
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Next file down"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton btnPrevFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3600
         Picture         =   "Main.frx":560E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Next file up"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton btnFont 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3000
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Set Font (Shift+Ctrl+F)"
         Top             =   0
         Width           =   615
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "read only"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Toggle Read-Only mode"
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox picQuery 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   4800
         ScaleHeight     =   555
         ScaleWidth      =   3975
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   3975
         Begin VB.TextBox txtQueryBox 
            Height          =   375
            Left            =   0
            MaxLength       =   50
            OLEDropMode     =   1  'Manual
            TabIndex        =   25
            ToolTipText     =   "Type something here, and search the file or on the web (F9)"
            Top             =   245
            Width           =   2895
         End
         Begin VB.CommandButton btnQueryExecute 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            MaskColor       =   &H80000000&
            Picture         =   "Main.frx":5A50
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Find next match (F3)"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton btnFindPrev 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            MaskColor       =   &H80000000&
            Picture         =   "Main.frx":5B9A
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Find next match backwards"
            Top             =   0
            Width           =   495
         End
         Begin VB.CheckBox chkQueryDotDotDot 
            Caption         =   "..."
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "More search options..."
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optQueries 
            Caption         =   "google"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Search google"
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optQueries 
            Caption         =   "thes"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Look it up at thesaurus.reference.com (Ctrl+T)"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optQueries 
            Caption         =   "dict"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Look it up at dictionary.reference.com (Ctrl+D)"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optQueries 
            Caption         =   "find"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Search within file (Ctrl+F)"
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.CheckBox chkWordWrap 
         Caption         =   "W"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Toggle Word Wrap (Ctrl+W)"
         Top             =   0
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CommandButton btnSave 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1200
         Picture         =   "Main.frx":5CE4
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Save File (Ctrl+S)"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton btnNewFile 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   600
         Picture         =   "Main.frx":6126
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "New File (Ctrl+N)"
         Top             =   0
         Width           =   615
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         Height          =   570
         Left            =   0
         Picture         =   "Main.frx":6330
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Show/Hide the File Browser (F8)"
         Top             =   0
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CommandButton btnFileForward 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   5400
         Picture         =   "Main.frx":6772
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnFileBack 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   9720
         Picture         =   "Main.frx":6BB4
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.Slider sliZoom 
         Height          =   570
         Left            =   3000
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Adjust picture zoom"
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1005
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SmallChange     =   10
         Max             =   500
         SelStart        =   100
         TickStyle       =   1
         TickFrequency   =   100
         Value           =   100
         TextPosition    =   1
      End
      Begin VB.CommandButton btnZoomOut 
         Appearance      =   0  'Flat
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   0
         Width           =   252
      End
      Begin VB.CommandButton btnZoomIn 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3360
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   290
         Width           =   252
      End
   End
   Begin MSComctlLib.StatusBar staTusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "0000 Files, 000000 Bytes Total"
            TextSave        =   "0000 Files, 000000 Bytes Total"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
            TextSave        =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Modified"
            TextSave        =   "Modified"
            Key             =   "statModified"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Key             =   "seltext"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   "statTips"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "&Rename Open File"
      End
      Begin VB.Menu mnuFileDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuFileDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditcut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuEditIncFont 
         Caption         =   "&Increase Font Size\tAlt+F6"
      End
   End
   Begin VB.Menu mnuBrowser 
      Caption         =   "Br&owser"
      Begin VB.Menu mnuBrowserRename 
         Caption         =   "Re&name Selected"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBrowserDelete 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu mnuFileCurrentDirectory 
         Caption         =   "Go To &Current Directory"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFileParentDirectory 
         Caption         =   "Parent Directo&ury"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuBrowserRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBrowserDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowserSort 
         Caption         =   "Reverse &Sort Order"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuBrowserOpenDefault 
         Caption         =   "&Open With Default Program"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status bar"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "T&oolbar"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewFilebrowser 
         Caption         =   "File &Browser"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDictionary 
         Caption         =   "&Dictionary..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuViewThesaurus 
         Caption         =   "&Thesaurus..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "Find &Next Forward"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewFindBackwards 
         Caption         =   "Find N&ext Backwards"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuViewDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReadOnly 
         Caption         =   "&Read Only"
      End
      Begin VB.Menu mnuViewWordWrap 
         Caption         =   "&Word Wrap"
         Checked         =   -1  'True
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuViewDiv5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmarksAdd 
         Caption         =   "&Add Current File"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuBookmarksAddPath 
         Caption         =   "Add Current &Path"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBookmarksManage 
         Caption         =   "&Manage"
      End
      Begin VB.Menu mnuBookmarksDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuWindowMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuWindowSize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuWindowMinimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuWindowMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuWindowDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowSaveSettings 
         Caption         =   "Save Se&ttings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&Readme.txt"
      End
      Begin VB.Menu mnuHelpDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuListDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuListDelete 
         Caption         =   "&Delete File..."
      End
      Begin VB.Menu mnuListDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListCopyPath 
         Caption         =   "&Copy Full File Name"
      End
      Begin VB.Menu mnuListOpenDefault 
         Caption         =   "Open In Default &Application..."
      End
      Begin VB.Menu mnuListShowOnly 
         Caption         =   "&Show only this file type"
      End
      Begin VB.Menu mnuListProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuListDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListCancel 
         Caption         =   "Canc&el"
      End
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "Write"
      Visible         =   0   'False
      Begin VB.Menu mnuWriteCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuWriteCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuWritePaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuWriteDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWriteToggleCaps 
         Caption         =   "Toggle Caps"
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "Query"
      Visible         =   0   'False
      Begin VB.Menu mnuQueryMatchCase 
         Caption         =   "Match Case"
      End
      Begin VB.Menu mnuQueryWholeWord 
         Caption         =   "Match Whole Word Only"
      End
      Begin VB.Menu mnuQueryDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "groups.google.com"
         Index           =   6
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "www.download.com"
         Index           =   7
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "froogle.google.com"
         Index           =   8
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "news.google.com"
         Index           =   9
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "images.google.com"
         Index           =   10
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "www.feedster.com"
         Index           =   11
      End
      Begin VB.Menu mnuQueries 
         Caption         =   "www.acronymfinder.com"
         Index           =   12
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' I've used tags like this, hoping it's not too terribly counterintuitive.
' (It is, though.  I think I may change it.)
'
' agEditor.tag -- full file & pathname of successfully loaded file
' lvwBrowser.tag -- full path of directory displayed
' mnuBookmark(x).tag -- exact filename of bookmark.

' TODO: maybe if I can find where I'm testing for (something), then
' I can something.

Option Compare Text
Option Explicit

Dim long1 As Long, long2 As Long
Dim msTest As String * 512
Dim mbTestByte() As Byte                  ' Scratch variables.  Comment them out later.
Dim msTestArray() As String


Const Debugging = True
Const mfSkipMouseEventCrap = True
Const msSettingsVersion = "0.8.8" ' Not the current build number, but the last time I changed the registry structure.
Const MAX_HISTORY = 15
Const MoveIncrement = 512
'Const ZoomIncrement = 5

Dim msPhlegmDate As String
Dim mudtStats As TStatType
'Dim FuckIHateThis As Boolean
Dim mfValidCboPath As Boolean
Dim mfStartLabelEditFromButton As Boolean
Dim mfSkipFormResize As Boolean
Dim mfEditorLoading As Boolean

Dim mfBrowserItemClicked As Boolean
Dim mfBrowserButtonPressed As Boolean
Dim miBrowserMouseButton As Integer
Dim miBrowserShift As Integer

Dim miPathRecent As Integer

Dim miCurrentQuery As Integer
Dim msLastFindQuery As String
Dim mfFindQueryChanged As Boolean
Dim miFindResults As Integer
Dim mlBeginFindPos As Long
Dim mfQueryMenuOpen As Boolean

Dim msPhlegmKey As String

'Dim EditorAccelTable() As ACCEL
'Dim ControlInfoData As CONTROLINFO
'Dim ctrlInfo1 As CTRLINFO

Dim mudtSettings As TWindowPrefs
Dim mudtCurrentFileSettings As TEditorPrefs

Dim miEditorMode As Integer
Dim miImageWidthDefault As Integer, miImageHeightDefault As Integer
Dim mfImageDragging As Boolean
'Dim miImageZoom As Integer
Dim miPrevX As Integer, miPrevY As Integer

Enum EFileType
      Directory = 1
      Drive = 3
      Text = 4
      Other = 5
      Picture = 6
      Error = 7
      Bookmark = 8
      Floppy = 9
      Network = 10
      Cdrom = 11
      rtf = 12
End Enum

Enum EStat
      BrowserStats = 1
      Stats = 2
      Modified = 3
      SelText = 4
      Tips = 5
End Enum

Enum EDirection
      Forward = 1
      back = -1
End Enum

Enum EQuery
      Find = 0
      Dictionary = 1
      thesaurus = 2
      Google = 3
      ' reserved = 4
      ' reserved = 5
      Groups = 6
      CNET = 7
      Froogle = 8
      News = 9
      Images = 10
      Feedster = 11
      Acronymfinder = 12
End Enum



Private Sub AddToBookmarks(ByVal sNewBookmark As String)
      Dim iIndex As Integer

      sNewBookmark = CstringToVBstring(sNewBookmark)
      If sNewBookmark = "" Then Exit Sub
     
      iIndex = mnuBookmark.UBound + 1
      Load mnuBookmark(iIndex)
      With mnuBookmark(iIndex)
            .Tag = sNewBookmark  ' exact path here, for safe keeping
            .Caption = iIndex & "   " & sNewBookmark ' here, to make it look all nice
            If iIndex <= 10 Then .Caption = "&" & .Caption
            .Visible = True
      End With

End Sub

Private Sub AddToHistory(ByVal sNewHistory As String)
      Dim iIndex As Integer
      Dim sPrevTag As String

      sNewHistory = CstringToVBstring(sNewHistory)
      If mnuFileHistory.UBound > 0 Then
            If sNewHistory = mnuFileHistory(1).Tag Then Exit Sub
      End If
      If sNewHistory = "" Then Exit Sub
     
      If mnuFileHistory.UBound < MAX_HISTORY Then
            Load mnuFileHistory(mnuFileHistory.UBound + 1)
            mnuFileHistory(mnuFileHistory.UBound).Visible = True
      End If
      
      For iIndex = mnuFileHistory.UBound To 2 Step -1
            sPrevTag = mnuFileHistory(iIndex - 1).Tag
            mnuFileHistory(iIndex).Tag = sPrevTag
            If iIndex < 10 Then
                  mnuFileHistory(iIndex).Caption = "&" & iIndex & " " & sPrevTag
            ElseIf iIndex = 10 Then
                  mnuFileHistory(iIndex).Caption = "1&0 " & sPrevTag
            Else
                  mnuFileHistory(iIndex).Caption = iIndex & " " & sPrevTag
            End If
      Next iIndex
      
      With mnuFileHistory(1)
            .Tag = sNewHistory ' exact path here, for safe keeping
            .Caption = "&1 " & sNewHistory ' here, to make it look all nice
      End With
      
End Sub

Private Sub BookmarkSaveChanges()
      Dim iIndex As Integer
      
      For iIndex = 1 To lvwBrowser.ListItems.Count
            mnuBookmark(iIndex).Tag = lvwBrowser.ListItems(iIndex)
            mnuBookmark(iIndex).Caption = iIndex & "   " & lvwBrowser.ListItems(iIndex)
      Next iIndex
      
      For iIndex = iIndex To mnuBookmark.UBound
            Unload mnuBookmark(iIndex)
      Next iIndex
End Sub

'   BrowserAutoSelectListItem:
'
'
'
Private Function BrowserAutoSelectListItem(ByRef BD As TBrowserData)
      Dim litCurrentItem As ListItem
      
      If BD.ListEmpty Or BD.BookmarkMode Then Exit Function
      
      ' Auto-select first filename to match partialfilename, if given.
      
      If BD.PartialFileName <> "" Then
            Set litCurrentItem = lvwBrowser.FindItem(BD.PartialFileName, , , lvwPartial)
            If Not (litCurrentItem Is Nothing) Then litCurrentItem.Selected = True
      
      ' Auto-select the directory we just moved out of, if doing a ParentDirectory.
            
      ElseIf BD.GoingToParent Then
      
            Set litCurrentItem = lvwBrowser.FindItem(gFSO.GetBaseName(BD.DirPrev))
            If Not (litCurrentItem Is Nothing) Then litCurrentItem.Selected = True
            
      ' Auto-select the item previously selected, for a refresh.
      
      ElseIf BD.DirUnchanged Then
            Set litCurrentItem = lvwBrowser.FindItem(BD.SelTextPrev)

            If (litCurrentItem Is Nothing) Then
                  lvwBrowser.ListItems(1).Selected = True
            Else
                  litCurrentItem.Selected = True
            End If
            
      ' Otherwise, auto-select the first item.
            
      Else
            lvwBrowser.ListItems(1).Selected = True
      End If
                  
      DoEvents
      lvwBrowser.SelectedItem.EnsureVisible ' Just doesn't seem to work without DoEvents first.
End Function

'   BrowserExecuteNext
'   Select the next item after the selection, and open it.
'
Private Sub BrowserExecuteNext()
      Dim iIndex As Integer

       ' Selecting the item next to the open file, not next to whatever random thing is currently selected.
      If (agEditor.Tag <> "") And (Not gBrowserData.BookmarkMode) Then btnCurrentDirectory_Click
        ' TODO: that should still do the sync if in bookmark mode and *the open file is not a bookmark*.
      
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      iIndex = lvwBrowser.SelectedItem.Index
      If iIndex < lvwBrowser.ListItems.Count Then
            If lvwBrowser.ListItems(iIndex + 1).Icon <> EFileType.Directory Then
                  lvwBrowser.ListItems(iIndex + 1).EnsureVisible
                  lvwBrowser.ListItems(iIndex + 1).Selected = True
                  BrowserExecuteItem lvwBrowser.ListItems(iIndex + 1)
                  DoEvents
            End If
      End If
End Sub

'   BrowserExecutePrev
'   Select the item previous to the one selected, and open it.
'
Private Sub BrowserExecutePrev()
      Dim iIndex As Integer
            
       ' Selecting the item next to the open file, not next to whatever random thing is currently selected.
      If (agEditor.Tag <> "") And (Not gBrowserData.BookmarkMode) Then btnCurrentDirectory_Click
        ' TODO: that should still do the sync if in bookmark mode and *the open file is not a bookmark*.
        
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      If ActiveControl.name = "lvwBrowser" Then
            SendKeys "{UP}{ENTER}"
      Else
            iIndex = lvwBrowser.SelectedItem.Index
            If iIndex > 1 Then
                  If lvwBrowser.ListItems(iIndex - 1).Icon <> EFileType.Directory Then
                        lvwBrowser.ListItems(iIndex - 1).EnsureVisible
                        lvwBrowser.ListItems(iIndex - 1).Selected = True
                        DoEvents
                        BrowserExecuteItem lvwBrowser.ListItems(iIndex - 1)
                  End If
            End If
      End If
End Sub


' *********************************************
' *
' *  BrowserGetFilesAndFolders
' *
' *  Takes the parsed data (BD) and fills lvwBrowser with the appropriate files.
' *
' *
' *********************************************
'
Private Sub BrowserGetFilesAndFolders(ByRef BD As TBrowserData)
            
      Dim iIcon As Integer
      Dim curTotalBytes As Currency
      Dim litCurrentItem As ListItem
      Dim hNextFile As Long, sFileName As String, sEx As String
      Dim WFD As WIN32_FIND_DATA
      Dim fHadFocus As Boolean ', fDirUnchanged As Boolean
      'Dim sOldSelectedItem As String
      'Dim sngStartTime As Single
      
      
      On Error Resume Next    ' there won't be an active control during form_load, so skip this part.
      fHadFocus = (ActiveControl.name = "lvwBrowser")
      On Error GoTo 0
      
      
      lvwBrowser.Tag = BD.Dir
      
      lvwBrowser.Visible = False  ' a nice idea, but we don't want to lose focus while under.  OR DO WE ?
      lvwBrowser.ListItems.Clear
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      
      'sngStartTime = Timer
      If BD.Filter = "" Then BD.Filter = "*"
      hNextFile = FindFirstFile(BD.Dir & BD.Filter, WFD)
      
      Do
            On Error Resume Next
            
            ' Divide the file types up slightly for icon selection
            
            sFileName = CstringToVBstring(WFD.cFileName) ' Lots of junk past the null character.
            sEx = gFSO.getextensionname(sFileName)
            
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                  iIcon = EFileType.Directory
            Else
                  iIcon = FileTypeFromExtension(sEx)
            End If

            If Err > 0 Then
                  iIcon = EFileType.Error
                  Debug.Print Err & ": " & Err.Description
            End If
            On Error GoTo 0
            
            
            ' Add that file!
            
            If sFileName <> "." And sFileName <> "" Then ' just what is the point in providing them with a "." folder?
                  Set litCurrentItem = lvwBrowser.ListItems.Add(, , sFileName, iIcon, iIcon)
                  
                  ' here, let's keep an invisible second column for sorting by directory later
                  If iIcon = EFileType.Directory Then
                        litCurrentItem.ListSubItems.Add , , 0
                  Else
                        litCurrentItem.ListSubItems.Add , , 1
                        curTotalBytes = curTotalBytes + WFD.nFileSizeLow
                  End If
            End If
      
      Loop While FindNextFile(hNextFile, WFD) <> 0
           
      
      If BD.Filter = "*" Then BD.Filter = ""
      BD.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = 1
      lvwBrowser.Visible = True
      If fHadFocus Then lvwBrowser.SetFocus
      
      staTusBar1.Panels(EStat.BrowserStats).Text = FormatNumber(curTotalBytes, 0) & " bytes in " & _
            lvwBrowser.ListItems.Count & " objects"
      
      'Debug.Print Timer - sngStartTime
End Sub


Private Function BrowserResizeHorizontal(ByVal iSupposedWidth As Integer) As Integer
      ' This is like a miniature RearrangeControls() for just picBrowser and everything within,
      ' and it happens to only affect their horizontal components.
      
      ' The return value is the difference (in twips) that picBrowser has grown.  Can be negative.
      
      Dim iOffset As Integer, iRealWidth As Integer, iScrollButtonX As Integer
      
      
      If iSupposedWidth < 1000 Then
            iRealWidth = 1000
      ElseIf picBrowser.Left + iSupposedWidth + 1500 > frmMain.ScaleWidth Then
            iRealWidth = frmMain.ScaleWidth - picBrowser.Left - 1500
      Else
            iRealWidth = iSupposedWidth
      End If
      
      iOffset = iRealWidth - picBrowser.Width
      
      picBrowser.Width = iRealWidth
      lvwBrowser.Width = lvwBrowser.Width + iOffset
      lblDivider.Left = lvwBrowser.Width
      lvwBrowser.ColumnHeaders(1).Width = lvwBrowser.Width - 100
      cboPath.Width = cboPath.Width + iOffset
      
      iScrollButtonX = lvwBrowser.Left + lvwBrowser.Width - btnScrollToTop.Width - 30
      If btnCurrentDirectory.Left + btnCurrentDirectory.Width <= iScrollButtonX Then
            btnScrollToTop.Left = iScrollButtonX
      Else
            btnScrollToTop.Left = btnCurrentDirectory.Left + btnCurrentDirectory.Width
      End If
      
      BrowserResizeHorizontal = iOffset
End Function

Private Sub EditorSetMode(iMode As Integer)
      ' When we change the sort of data to display (text, picture, more to be determined),
      ' there are some things that have to be set, hidden, etc.
      
      ' Use EFileType types for iMode
      
      ' Other routines that may be curious about the mode may use miEditorMode.
      
      If iMode = miEditorMode Then Exit Sub
      
      Select Case iMode
            Case EFileType.Text, EFileType.rtf, EFileType.Other
      
                  miEditorMode = iMode
                  agEditor.Visible = True
                  Image1.Visible = False
'                  btnZoomIn.Visible = False
'                  btnZoomOut.Visible = False
'                  sliZoom.Visible = False
                  btnFont.Visible = True
      
            Case EFileType.Picture
                  
                  miEditorMode = EFileType.Picture
                  agEditor.Visible = False
                  Image1.Visible = True
'                  btnZoomIn.Visible = True
'                  btnZoomOut.Visible = True
'                  sliZoom.Visible = True
                  btnFont.Visible = False
      End Select

End Sub

Private Function FileTypeFromExtension(sEx As String) As String
      ' This function takes an extension (DO NOT INCLUDE DOT) and returns a mode
      ' which can be fed into EditorSetMode.
      
      ' Current possible modes:   EFileType.text, EFileType.rtf, EFileType.picture
      
      Select Case sEx
            Case "bmp", "gif", "jpg", "jpeg", "ico", "cur"
                  FileTypeFromExtension = EFileType.Picture
            Case "rtf"
                  FileTypeFromExtension = EFileType.rtf
            Case "txt"
                  FileTypeFromExtension = EFileType.Text
            Case Else
                  FileTypeFromExtension = EFileType.Other
      End Select
End Function

Private Sub ImageZoomIn(iStep As Integer)
      ' goes up to the next zoom divisible by iStep
      If sliZoom.Value >= sliZoom.Max Then Exit Sub
      sliZoom.Value = sliZoom.Value + (iStep - (sliZoom.Value Mod iStep))
End Sub

Private Sub ImageZoomOut(iStep As Integer)
      ' Sets zoom to the next lowest integer divisibly by iStep.
      
      If sliZoom.Value <= 0 Then Exit Sub
      
      If sliZoom.Value Mod iStep = 0 Then
            sliZoom.Value = sliZoom.Value - iStep
      Else
            sliZoom.Value = sliZoom.Value - (sliZoom.Value Mod iStep)
      End If
End Sub


Private Sub ListMenuDisable()

      If Not mnuListOpenDefault.Enabled Then Exit Sub
      
      mnuListOpenDefault.Enabled = False
      mnuListOpen.Enabled = False
      mnuListDelete.Enabled = False
      mnuListRename.Enabled = False
      mnuListCopyPath.Enabled = False
      mnuListShowOnly.Enabled = False
      mnuListProperties.Enabled = False
End Sub

Private Sub ListMenuEnable(litHoverItem As ListItem)
      If mnuListOpenDefault.Enabled Then Exit Sub  ' don't wanna bother doing all this on every single mousemove!
      
      mnuListOpenDefault.Enabled = True
      mnuListOpen.Enabled = True
      mnuListOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
      mnuListDelete.Enabled = True
      mnuListRename.Enabled = True
      mnuListCopyPath.Enabled = True
      mnuListShowOnly.Enabled = True
      mnuListProperties.Enabled = True
      
      If litHoverItem.Icon = EFileType.Directory Or litHoverItem.Icon = EFileType.Drive Then
            mnuListOpenDefault.Caption = "Explore..." & vbTab & "Shift+Ctrl+Enter"
            mnuListDelete = False
            If litHoverItem.Text = ".." Or litHoverItem.Icon = EFileType.Drive Then mnuListRename = False
      End If

End Sub

Private Function ParentDirectoryOf(ByVal sPath As String)
      Dim iSlash As Integer
      
      If sPath = "\" Then
            ParentDirectoryOf = ""
      Else
            iSlash = InStrRev(sPath, "\", Len(sPath) - 1)
            ParentDirectoryOf = Left(sPath, iSlash)
      End If
End Function

'   Much can be learned that is locked within cboPath.
'   Turn that data into a structure, that we can use and abuse from anywhere, anytime!
'
'   ParsePath translates input string sInput into referenced data structure BD.     TODO TODO TODO:
'   BD will hold the working directory, filter, previous directory, mode,
'   ...and much, much more!
'
Private Sub ParsePath(ByVal sInput As String, ByRef BD As TBrowserData)
      
      Dim sFileName As String
      
      sInput = Trim(sInput)
      
      With BD
      
            .BookmarkMode = False
            .DrivesMode = False
            .ListEmpty = (lvwBrowser.ListItems.Count = 0)
            .DirPrev = .Dir
            .FilterPrev = .Filter
            
            
            
            If sInput = "(Bookmarks)" Then  ' We are in Manage Bookmarks mode.
                  .BookmarkMode = True
                  .Dir = "(Bookmarks)"  ' Just so that (.Dir = X) never accidentally returns true.
                  .Filter = ""
                  .PartialFileName = ""
                  .ValidPath = False
            
            Else
                  If Not (sInput Like "*:\*") Then  ' Drives mode, root of the file system.
                        .ValidPath = False
                        .DrivesMode = True
                        .PartialFileName = sInput
                        .Dir = ""
                  Else                                            ' Ordinary (folder) mode.
                        .ValidPath = True
                        .Dir = SnipFileName(sInput)
                        If Not gFSO.FolderExists(.Dir) Then .ValidPath = False
                  End If
                  .DirUnchanged = (.Dir = .DirPrev)
                  .GoingToParent = (.Dir = ParentDirectoryOf(.DirPrev)) And Not .DirUnchanged
            End If
            
            sFileName = SnipPath(sInput)
            
            If .ValidPath Then
                  
                  .PartialFileName = ""
                  If Right(sInput, 1) = "\" Then  ' c:\temp\   (just a plain old directory)
                        .Filter = ""
                  ElseIf sFileName Like ".*" And Not (sFileName Like "*.") Then  ' c:\temp\.txt  (wildcard implied)
                        .Filter = "*." & gFSO.getextensionname(sFileName)
                        
                  ElseIf sFileName Like "*[?*]*" Then  ' c:\temp\peni*   (contains wildcard(s) after the directory)
                        .Filter = sFileName
                        
                  ElseIf Not .ListEmpty Then  ' c:\temp\peni   (some trailing characters, but no wildcard)
                        .Filter = ""
                        .PartialFileName = sFileName
                  End If
            End If
            .FilterUnchanged = (.Filter = .FilterPrev)
            
            If Not .ListEmpty Then .SelTextPrev = lvwBrowser.SelectedItem.Text
            
            
            .InputPrev = sInput
      End With


End Sub

Private Sub PathAddRecent(ByVal sPath As String)
      ' Supplement recent paths list, unless we are currently scrolling through them.
      ' Top of the List = Lowest of the ListIndeces = Forward(recent)most of the paths.
            
      Dim iIndex As Integer
      
      With cboPath
      
            If .ListIndex = -1 Then  ' (not scrolling through them)
                  
                  ' Delete forward history, if any, and insert current path.
                  
                  For iIndex = 0 To miPathRecent - 1
                        .RemoveItem 0
                  Next iIndex
                  
                  ' May contain repeats, but we don't need any *consecutive* repeats.
                  
                  If .ListCount = 0 Or .List(0) <> sPath Then
                        ' It's either empty, or it DOESN'T match the previous path.
                        .AddItem sPath, 0
                  End If
                  
                  miPathRecent = 0
            End If
            
      End With
End Sub

Private Sub BrowserGetBookmarks()
      Dim iIndex As Integer
      Dim litCurrentItem As ListItem
      
      lvwBrowser.ListItems.Clear
      lvwBrowser.Tag = "(Bookmarks)"
      For iIndex = 1 To mnuBookmark.UBound
            Set litCurrentItem = lvwBrowser.ListItems.Add(, , mnuBookmark(iIndex).Tag, _
                  EFileType.Bookmark, EFileType.Bookmark)
            litCurrentItem.ListSubItems.Add 1, , 1
      Next iIndex
      
      staTusBar1.Panels(EStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " bookmarks"
End Sub

Private Function PathBack() As Boolean
      ' Go back in the recent paths list
      
      PathBack = False
      
      With cboPath
            If .ListCount = 0 Then
                  Exit Function
            ElseIf .ListCount = 1 Then
                  .ListIndex = 0
                  Exit Function
            ElseIf .ListIndex = -1 Then
                  .ListIndex = 1
                  PathBack = True
            ElseIf .ListIndex < .ListCount - 1 Then
                  .ListIndex = .ListIndex + 1
                  PathBack = True
            End If
      End With
      
      PathBack = True
End Function

Private Sub PathForward()
      With cboPath
            If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
      End With
End Sub

Private Sub ShowFileProperties(ByVal sPath As String)
      ' SImply calls the Explorer file properties dialog.  Hope this works.
      
      Dim seeEx As SHELLEXECUTEINFO
            
      seeEx.cbSize = LenB(seeEx)
      seeEx.lpFile = sPath
      seeEx.lpVerb = "properties"
      seeEx.fMask = SEE_MASK_INVOKEIDLIST
      
      ShellExecuteEx seeEx
End Sub

'Private Sub StatusBarToolTip(ByVal sTip As String)
'      staTusBar1.Panels(EStat.Tips) = sTip
'
'      tmrRevertTips.Enabled
'End Sub

Private Sub agEditor_ProgressStatus(ByVal lAmount As Long, ByVal lTotal As Long)
'      Debug.Print "PROGRESS: "; lAmount & " " & lTotal

      ' TODO: if a second file is told to load, it cancels this one but won't remove it from the editor first.
      
      DoEvents
End Sub

Private Sub btnCurrentDirectory_Click()
      mnuFileCurrentDirectory_Click
End Sub

Private Sub btnDeleteSelected_Click()
      BrowserDeleteSelected
End Sub

Private Sub btnFindPrev_Click()
      Dim lFoundMin As Long, lFoundMax As Long
      Dim iFindOptions As Integer
                  
      'iFindOptions = ERECFindTypeOptions.FR_DOWN
      If mnuQueryWholeWord.Checked Then iFindOptions = iFindOptions + ERECFindTypeOptions.FR_WHOLEWORD
      If mnuQueryMatchCase.Checked Then iFindOptions = iFindOptions + ERECFindTypeOptions.FR_MATCHCASE
      ' TODO: search within selection menu item
      
      agEditor.GetSelection lFoundMin, lFoundMax
      agEditor.SetSelection lFoundMin, lFoundMin
      agEditor.FindText txtQueryBox, iFindOptions, True, False, lFoundMin, lFoundMax
      agEditor.SetSelection lFoundMin, lFoundMax
End Sub

Private Sub btnFolderUp_Click()
      mnuFileParentDirectory_Click
End Sub


Private Sub btnFont_Click()
      Dim fntTemp As New StdFont ' StdFont is a Class
      
      With dlgFont 'make the dialog choices begin with what the agEditor shows
            .flags = cdlCFBoth + cdlCFApply ' and allow for all font types.
            .FontName = agEditor.GetFont.name                    ' btw, Apply doesn't work
            .FontBold = agEditor.GetFont.Bold
            .FontUnderline = agEditor.GetFont.Underline
            .FontSize = CSng(agEditor.GetFont.Size)  ' one uses Single, the other Currency
      End With

      On Error Resume Next 'trap the error. if they hit cancel, do nothing and exit
      dlgFont.ShowFont
      If Err.Number = cdlCancel Then Exit Sub
      On Error GoTo 0  'btw, I think this has the effect of err.Clear
      
      With fntTemp
            .name = dlgFont.FontName
            .Bold = dlgFont.FontBold
            .Italic = dlgFont.FontItalic
            .Underline = dlgFont.FontUnderline
            .Size = CCur(dlgFont.FontSize)
      End With
      agEditor.SetFont fntTemp, , , , ercSetFormatAll
      Me.Caption = agEditor.GetFont.name & " " & agEditor.GetFont.Charset & " " & agEditor.GetFont.Size
End Sub


Private Sub btnNewFile_Click()
      mnuFileNew_Click
End Sub

Private Sub btnNextFile_Click()
      BrowserExecuteNext
End Sub

Private Sub btnOptions_Click()
      frmOptions.Show
End Sub

Private Sub btnPathBack_Click()
      PathBack
End Sub

Private Sub btnPathForward_Click()
      PathForward
End Sub


Private Sub btnPrevFile_Click()
      BrowserExecutePrev
End Sub

Private Sub btnZoomIn_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      ' Goes to the next zoom divisible by the zoom increment
      If sliZoom.Value < 100 Then
            ImageZoomIn 25
      Else
            ImageZoomIn sliZoom.LargeChange
      End If
End Sub

Private Sub btnZoomOut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      ' Goes to the next lowest zoom % divisible by the zoom increment
      
      If sliZoom.Value <= 100 Then
            ImageZoomOut 25
      Else
            ImageZoomOut sliZoom.LargeChange
      End If
End Sub

Private Sub chkQueryDotDotDot_Click()
      Debug.Print "chkQuery..._Click"
      
      If chkQueryDotDotDot.Value = vbChecked Then
            PopupMenu mnuQuery, vbPopupMenuRightAlign, AbsoluteRight(chkQueryDotDotDot), _
                  AbsoluteBottom(chkQueryDotDotDot)
      End If

End Sub

Private Sub btnRefresh_Click()
      ' A lot like invoking cboPath_Change, but with the distinction that
      ' there is no check for gBrowserData.DirUnchanged, and
      ' there is no need to re-parse gBrowserData.
            
      With gBrowserData
            .DirPrev = .Dir
            .FilterPrev = .Filter
            .DirUnchanged = True
            .FilterUnchanged = True
            .GoingToParent = False
            
            If .BookmarkMode Then
                  BrowserGetBookmarks
                  
            ElseIf .DrivesMode Then
                  BrowserGetDrives
            Else
                  BrowserGetFilesAndFolders gBrowserData
            End If
      End With
      
      BrowserAutoSelectListItem gBrowserData
      
End Sub

Private Sub btnSave_Click()
      mnuFileSave_Click
End Sub

Private Sub btnScrollToTop_Click()
      If lvwBrowser.ListItems.Count > 0 Then lvwBrowser.ListItems(1).EnsureVisible
End Sub

Private Sub btnSort_Click()
      
      ' List remains sorted at all times.  Only the order can be reversed.
      
      With lvwBrowser
            .SortKey = 0
            .SortOrder = Abs(.SortOrder - 1)
            .SortKey = 1
      End With
      
      If gBrowserData.BookmarkMode Then BookmarkSaveChanges
End Sub


Private Sub btnScrolltotop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnScrollToTop.ToolTipText
End Sub

Private Sub btnSort_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnSort.ToolTipText
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnSave.ToolTipText
End Sub

Private Sub btnrefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnRefresh.ToolTipText
End Sub

Private Sub chkQueryDotDotDot_LostFocus()
      chkQueryDotDotDot.Value = vbUnchecked
End Sub

'Private Sub chkQueryDotDotDot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      Debug.Print "chkQuery..._Mousedown"
'      'PopupMenu mnuQuery, vbPopupMenuRightAlign
'      If mfQueryMenuOpen = False Then
'            PopupMenu mnuQuery, vbPopupMenuRightAlign, AbsoluteRight(chkQueryDotDotDot), _
'                  AbsoluteBottom(chkQueryDotDotDot)
'      End If
'
'End Sub



Private Sub chkQueryDotdotdot_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = chkQueryDotDotDot.ToolTipText
End Sub

Private Sub btnprevfile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnPrevFile.ToolTipText
End Sub

Private Sub btnpathforward_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnPathForward.ToolTipText
End Sub

Private Sub btnpathback_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnPathBack.ToolTipText
End Sub

Private Sub btnnextfile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnNextFile.ToolTipText
End Sub

Private Sub btnnewfile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnNewFile.ToolTipText
End Sub

Private Sub btnfont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnFont.ToolTipText
End Sub

Private Sub btnfolderup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnFolderUp.ToolTipText
End Sub

Private Sub btnfindprev_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnFindPrev.ToolTipText
End Sub

Private Sub btnQueryExecute_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnQueryExecute.ToolTipText
End Sub

Private Sub btnfileforward_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnFileForward.ToolTipText
End Sub

Private Sub btnfileback_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnFileBack.ToolTipText
End Sub

Private Sub btndeleteselected_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnDeleteSelected.ToolTipText
End Sub

Private Sub btncurrentdirectory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = btnCurrentDirectory.ToolTipText
End Sub

Private Sub ageditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      ' TODO: this needs to not happen, if frmMain is not the front window
      
      'Debug.Print Screen.ActiveForm.name
      On Error Resume Next
      If Screen.ActiveForm.name = "frmMain" And Not (ActiveControl.name = "agEditor") Then
            agEditor.SetFocus
      End If
      On Error GoTo 0
      staTusBar1.Panels(EStat.Tips).Text = ""
End Sub

Private Sub chkQueryDotDotDot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      Debug.Print "chkquery..._mouseup"
End Sub

Private Sub chkReadOnly_Click()
      
      mnuViewReadOnly.Checked = chkReadOnly.Value
      agEditor.ReadOnly = chkReadOnly.Value
      If chkReadOnly.Value = vbChecked Then
            agEditor.BackColor = &H8000000F
      Else
            agEditor.BackColor = &H80000005
      End If
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = ""
End Sub


Private Sub cboPath_Change()
      
      ParsePath cboPath, gBrowserData
      
      If gBrowserData.BookmarkMode Then
            BrowserGetBookmarks
            PathAddRecent "(Bookmarks)"
            
      ElseIf gBrowserData.DrivesMode Then
            BrowserGetDrives
            PathAddRecent ""
            
      ElseIf Not (gBrowserData.DirUnchanged And gBrowserData.FilterUnchanged) Then
            BrowserGetFilesAndFolders gBrowserData
            ' Add to recent paths only if filtration was fruitful.
            If Not gBrowserData.ListEmpty Then PathAddRecent gBrowserData.Dir & gBrowserData.Filter
      End If
      
      BrowserAutoSelectListItem gBrowserData
End Sub

Private Sub cboPath_Click()
      'debug.print "cboPath_CLICK!!!! " & cboPath.ListIndex
      
      ' So as it turns out, this is the event that fires when you select another
      '   item from the combobox list (via keyboard or mouse).  It is better thought
      '   of as a Change event for the combobox acting as a drop-down list.
      '   Naturally, it requires no "click" of the mouse.  Why should it?
      
      ' Combobox's actual Change event is associated with combobox acting as a textbox,
      '   and does not occur when combobox acts as a drop-down list.
      
      ' ComboBox DropDown event would be more aptly named the Click event
      '   for the dropdown arrow button.  It doesn't care what you do with the dropdown
      '   later.  Just fires once on the click (or probably an F4).
      
      miPathRecent = cboPath.ListIndex
     
      cboPath_Change
End Sub

Private Sub cboPath_DropDown()
      'debug.print "cboPath_DROPDOWN"
End Sub

Private Sub cboPath_Scroll()
      'debug.print "cboPath_SCROLL"
End Sub

Private Sub chkFileBrowser_Click()
      
      picBrowser.Visible = chkFileBrowser.Value
      mnuViewFilebrowser.Checked = chkFileBrowser.Value
      mnuBrowser.Enabled = chkFileBrowser.Value
      staTusBar1.Panels(EStat.BrowserStats).Visible = chkFileBrowser.Value
      
      RearrangeControls
      'agEditor.SetFocus
End Sub

'
'   cbopath_GotFocus
'
'   When focus is obtained, put the cursor right where we would have moved it anyway:
'   At the end of the path, before the extension if one exists.
'
Private Sub cboPath_GotFocus()
      If cboPath <> "(Bookmarks)" Then
            
            Dim iExtensionLength As Integer
            
            iExtensionLength = Len(gFSO.getextensionname(cboPath))
            If iExtensionLength > 0 Then iExtensionLength = iExtensionLength + 1 ' include the dot
            cboPath.SelStart = Len(cboPath) - iExtensionLength
      End If
End Sub

Private Sub cboPath_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim iSlash As Integer
      Debug.Print "cbopath.selstart" & cboPath.SelStart
      Select Case KeyCode
            Case vbKeyBack
                  If Shift = vbCtrlMask Then
                        ' ctrl+backspace = delete to the previous slash.
                        With cboPath
                              iSlash = InStrRev(.Text, "\", .SelStart + .SelLength)
                              
                              If iSlash > 0 Then .Text = Mid(.Text, iSlash, .SelStart + .SelLength - iSlash)
                        End With
                  End If
            Case vbKeyReturn
                  lvwBrowser.SetFocus
            
            Case vbKeyDown
                  If cboPath.ListIndex = -1 And cboPath.ListCount > 1 Then
                        cboPath.ListIndex = 1
                  End If
            
            Case vbKeyLeft
                  If Shift = vbAltMask Then PathBack
                  
            Case vbKeyRight
                  If Shift = vbAltMask Then PathForward
      End Select
End Sub

Private Sub chkFileBrowser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      staTusBar1.Panels(EStat.Tips).Text = chkFileBrowser.ToolTipText
End Sub

Private Sub chkWordWrap_Click()
      
      mnuViewWordWrap.Checked = chkWordWrap.Value
      agEditor.ViewMode = chkWordWrap.Value
      
End Sub

Private Sub btnQueryExecute_Click()
      
      If txtQueryBox = "" Then txtQueryBox = agEditor.SelectedText
            
      
      Select Case miCurrentQuery
            Case EQuery.Find
            
                  Dim lFoundMin As Long, lFoundMax As Long, lStartMin As Long, lStartMax As Long
                  Dim lFindRetval As Long, fFindInSelection As Boolean
                  Dim iFindOptions As Integer
                  
                  ' TODO: keep track of this start position (lStartMax)
                  agEditor.GetSelection lStartMin, lStartMax
                  
                  If (txtQueryBox <> msLastFindQuery) Or (mfFindQueryChanged = True) Then
                        mfFindQueryChanged = True
                        miFindResults = 0
                        msLastFindQuery = txtQueryBox
                        mlBeginFindPos = lStartMax
                  End If
                              
                  
                  lFindRetval = EditorFindText(txtQueryBox, Forward, lStartMax, agEditor.CharacterCount, lFoundMin, lFoundMax)
                  If lFindRetval >= 0 Then
                        agEditor.SetSelection lFoundMin, lFoundMax
                        miFindResults = miFindResults + 1
                        staTusBar1.Panels(EStat.Tips) = "Search results: " & miFindResults & " found"
                  End If
                  
            Case EQuery.Dictionary
                  ShellExecute 0, "open", "http://dictionary.reference.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.thesaurus
                  ShellExecute 0, "open", "http://thesaurus.reference.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.Google
                  ShellExecute 0, "open", "http://www.google.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.Froogle
                  ShellExecute 0, "open", "http://froogle.google.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.Groups
                  ShellExecute 0, "open", "http://groups.google.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.Images
                  ShellExecute 0, "open", "http://images.google.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.News
                  ShellExecute 0, "open", "http://news.google.com/search?q=" & txtQueryBox, 0, "", 8
            Case EQuery.Acronymfinder
                  ShellExecute 0, "open", _
                        "http://www.acronymfinder.com/af-query.asp?String=exact&Acronym=" _
                        & txtQueryBox & "&Find=Find", 0, "", 8
                              
            Case EQuery.Feedster
                  ShellExecute 0, "open", _
                        "http://feedster.com/search.php?q=" & txtQueryBox & "&hl=en&ie=UTF-8&sort=date", 0, "", 8
                        
            Case EQuery.CNET
                  ShellExecute 0, "open", _
                        "http://www.download.com/3120-20-0.html?qt=" & txtQueryBox & _
                        "&tg=dl-2001&part=opera&subj=windows&tag=search", 0, "", 8
      End Select

      txtQueryBox.SetFocus
End Sub

' EditorFindText
'   Finds the search string sFindMe in agEditor between values of lRangeStart and lRangeEnd.
'
'  The way EM_FINDTEXTEX works is that it goes from lRangeStart to lRangeEnd in the
'  specified direction.  That means the start position has to come first.  NOT the lower of the values first.
'
'  lFoundMin and lFoundMax receive the start and end positions of the found string.
'  Returns -1 if nothing found, returns lFoundMin if successful.

Private Function EditorFindText( _
            ByVal sFindme As String, ByVal iDirection As EDirection, _
            ByVal lRangeStart As Long, ByVal lRangeEnd As Long, _
            ByRef lFoundMin As Long, ByRef lFoundMax As Long) As Long
      
      Const FR_MATCHCASE As Long = &H4
      Const FR_WHOLEWORD As Long = &H2
      Const FR_DOWN As Long = &H1
'      Const EM_FINDTEXT As Long = (WM_USER + 56)
      Const EM_FINDTEXTEX As Long = (WM_USER + 79)

      Dim fFindNext As Boolean, fFindInSelection As Boolean
      Dim lFindOptions As Long
      Dim fexFindData As FINDTEXTEX
      
      If iDirection = Forward Then lFindOptions = FR_DOWN ' fr_down = go from lStartMin to end of editor.
      If mnuQueryWholeWord.Checked Then lFindOptions = lFindOptions + FR_WHOLEWORD
      If mnuQueryMatchCase.Checked Then lFindOptions = lFindOptions + FR_MATCHCASE
      
      fexFindData.chrg.cpMin = lRangeStart
      fexFindData.chrg.cpMax = lRangeEnd
      fexFindData.lpstrText = sFindme & Chr(0) ' it wants a C string
      
      EditorFindText = SendMessage(agEditor.RichEdithWnd, EM_FINDTEXTEX, ByVal lFindOptions, fexFindData)
      
      lFoundMin = fexFindData.chrgText.cpMin
      lFoundMax = fexFindData.chrgText.cpMax
End Function

Private Sub Form_Unload(Cancel As Integer)
      If Not Debugging Then
            SetWindowLong lvwBrowser.hwnd, GWL_WNDPROC, gpOldLvwBrowserProc
            gpOldLvwBrowserProc = 0
      End If
      
      SaveWindowSettings
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      miPrevX = x
      miPrevY = y
      If Button = vbLeftButton Then
            mfImageDragging = True
            picEditor.SetFocus
      End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      On Error Resume Next
      If Not (ActiveControl.name = "picEditor") Then picEditor.SetFocus
      On Error GoTo 0
            
      If mfImageDragging Then
            Image1.Move Image1.Left + x - miPrevX, Image1.Top + y - miPrevY, Image1.Width, Image1.Height
      End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      mfImageDragging = False
End Sub


'
'  lvwBrowser_AfterLabelEdit (in other words, "rename")
'
Private Sub lvwBrowser_AfterLabelEdit(Cancel As Integer, NewString As String)

      ' TODO: finish this, and make it work for directories.
      
      Dim sFolder As String, sOldPath As String
      
      sFolder = gBrowserData.Dir
      
      If sFolder = "(Bookmarks)" Then
            ' We'll take their renamed bookmark, and if it's not a valid file, let that be
            ' a problem when they try to open the bookmark.
            With mnuBookmark(lvwBrowser.SelectedItem.Index)
                  .Tag = NewString
                  .Caption = lvwBrowser.SelectedItem.Index & "   " & NewString
            End With
            
            Exit Sub
      End If
            
      
      If NewString Like "*:\*" Then sFolder = ""  ' If it looks like a full path, treat it like one.
      sOldPath = gBrowserData.Dir & lvwBrowser.SelectedItem
      
      If FileExists(sOldPath) = False Then
            Caption = "Can't rename what's not there: " & sOldPath
            btnRefresh_Click
            Cancel = True
      
      ElseIf lvwBrowser.SelectedItem = NewString Or sOldPath = NewString Then
            
            If StrComp(lvwBrowser.SelectedItem, NewString, vbBinaryCompare) = 0 Or _
                  StrComp(sOldPath, NewString, vbBinaryCompare) = 0 Then
                  
                  Cancel = True  ' No change whatsoever.
            
            Else ' Change in caps only.  We'll rename it anyway, just to be a sport.
                  
                  On Error Resume Next
                  Name sOldPath As sFolder & NewString
                  If Err > 0 Then
                        Caption = Err.Number & ": " & Err.Description
                        Cancel = True
                  ElseIf sOldPath = agEditor.Tag Then
                        Caption = "Adjusted the capitalization of open file to: " & sFolder & NewString
                        agEditor.Tag = sFolder & NewString
                  Else
                        Caption = "Renamed.  Even though all you changed was the capitalization.  Freak."
                        agEditor.Tag = sFolder & NewString
                  End If
                  On Error GoTo 0
            
            End If
      
      ElseIf FileExists(sFolder & NewString) Then
            Caption = "This name sucks: " & Chr(34) & sFolder & NewString & Chr(34) & ".  Change it."
            Cancel = True
            
      Else   ' ...and finally, we rename a file.
            On Error Resume Next
            Name sOldPath As sFolder & NewString
            If Err > 0 Then
                  Caption = Err.Number & ": " & Err.Description
                  Cancel = True
            ElseIf sOldPath = agEditor.Tag Then
                  Caption = "Renamed open file: " & sFolder & NewString
                  agEditor.Tag = sFolder & NewString
            Else
                  Caption = "Rename successful: " & sFolder & NewString
                  agEditor.Tag = sFolder & NewString
            End If
            On Error GoTo 0
      
      End If
      
      btnRefresh_Click
      btnCurrentDirectory_Click
End Sub

Private Sub lvwBrowser_BeforeLabelEdit(Cancel As Integer)
      'debug.print "lvwBrowser_Before " & Cancel
End Sub

Private Sub lvwBrowser_Click()
      'debug.print "lvwBrowser_CLICK"
      
      
      miBrowserMouseButton = 0  ' These probably an overcaution --
      miBrowserShift = 0                  ' They are reset in the next MouseDown anyway.
End Sub

Private Sub BrowserExecuteItem(ByVal Item As MSComctlLib.ListItem)
      If (lvwBrowser.ListItems.Count = 0) Then Exit Sub
      
      Select Case Item.Icon
      
            Case EFileType.Directory, EFileType.Drive, EFileType.Floppy, EFileType.Cdrom, EFileType.Network
                  ' Open the folder, or go up a folder.
                  If Item.Text = ".." Then
                        mnuFileParentDirectory_Click
                  Else
                        cboPath = gBrowserData.Dir & Item.Text & "\"
                  End If
            
            Case EFileType.Bookmark
                  ' Open the bookmarked file.  TODO: make it work for folders
                  EditorLoadFile Item.Text, FileTypeFromExtension(gFSO.getextensionname(Item.Text))
            
            Case Else
                  ' Open the file.  EditorLoadFile knows what to do.
                  
                  EditorLoadFile gBrowserData.Dir & Item.Text, Item.Icon
      End Select
End Sub

Private Sub lvwBrowser_DblClick()
      'debug.print "lvwBrowser_DBLCLICK"
End Sub

Private Sub lvwBrowser_ItemClick(ByVal Item As MSComctlLib.ListItem)
      ' ItemClick fires every time the selection changes, or a selection is clicked.
      'Debug.Print "itemclick " & Item.Index
            
      mfBrowserItemClicked = True
      
'      mnuListOpenDefault.Enabled = True
'      mnuListOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
'      mnuListDelete.Enabled = True
'      mnuListRename.Enabled = True
'      mnuListCopyPath.Enabled = True
'      mnuListShowOnly.Enabled = True
'      mnuListProperties.Enabled = True
'
'      If Item.Icon = EFileType.Directory Or Item.Icon = EFileType.Drive Then
'            mnuListOpenDefault.Caption = "Explore..." & vbTab & "Shift+Ctrl+Enter"
'            mnuListDelete = False
'            If Item.Text = ".." Or Item.Icon = EFileType.Drive Then mnuListRename = False
'      End If
End Sub

Private Sub lvwBrowser_KeyPress(KeyAscii As Integer)
      'debug.print "lvwBrowser_KEYPRESS " & KeyAscii
      Select Case KeyAscii
            
            Case vbKeyReturn
                  BrowserExecuteItem lvwBrowser.SelectedItem
      End Select
End Sub

Private Sub lvwBrowser_KeyUp(KeyCode As Integer, Shift As Integer)
      'debug.print "lvwBrowser_KEYUP"
End Sub

Private Sub lvwBrowser_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      'debug.print "lvwBrowser_MOUSEDOWN "; Button & " " & Shift
      
      lvwBrowser_MouseMove Button, Shift, x, y
      mfBrowserItemClicked = False
      miBrowserMouseButton = Button
      miBrowserShift = Shift
      
End Sub

Private Sub lvwBrowser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'      Debug.Print "lvwBROWSER MOUSEMOVE, X: " & x

      Dim litHoverItem As ListItem
      
      On Error Resume Next
      If Not (ActiveControl.name = "lvwBrowser") Then lvwBrowser.SetFocus
      On Error GoTo 0
      
      Set litHoverItem = lvwBrowser.HitTest(x, y)  ' To see if we're over an item.
      
      If Not (litHoverItem Is Nothing) Then
            staTusBar1.Panels(EStat.Tips).Text = litHoverItem.Text
            ListMenuEnable litHoverItem
            
            If Button = vbLeftButton Or Button = vbRightButton Then
                  litHoverItem.Selected = True
            End If
      
      Else
            staTusBar1.Panels(EStat.Tips).Text = ""
            ListMenuDisable
      
      End If
      
'      If GetCapture <> lvwBrowser.hwnd Then
'            SetCapture (lvwBrowser.hwnd)
'      End If
      'Caption = x & " " & y
End Sub

Private Sub mnuBookmark_Click(Index As Integer)
      Dim sEx As String
      
      sEx = gFSO.getextensionname(mnuBookmark(Index).Tag)
      EditorLoadFile mnuBookmark(Index).Tag, FileTypeFromExtension(sEx)
      
      mnuFileCurrentDirectory_Click
End Sub

Private Sub mnuBookmarksAdd_Click()
      Dim iBookm As Integer
      
      ' TODO: ctrl+M doesn't work from the Editor
      ' find a better shortcut, and see what else doesn't work from the editor.
      
      For iBookm = 1 To mnuBookmark.UBound
            If mnuBookmark(iBookm).Tag = agEditor.Tag Then
                              ' Oops, got that bookmark already.
                  Exit Sub  ' Nothing left to do here!
            End If
      Next iBookm
      
      AddToBookmarks agEditor.Tag
      
      If gBrowserData.BookmarkMode Then btnRefresh_Click
End Sub

Private Sub mnuBookmarksAddPath_Click()
      Dim iBookm As Integer
      
      For iBookm = 1 To mnuBookmark.UBound
            If mnuBookmark(iBookm).Tag = cboPath Then
                              ' Oops, got that bookmark already.
                  Exit Sub  ' Nothing left to do here!
            End If
      Next iBookm
      
      AddToBookmarks cboPath
End Sub


Private Sub mnuBookmarksManage_Click()
      ' Basically, the brains for the entire program rest within cboPath_Change.
      
      cboPath = "(Bookmarks)"
      
End Sub

Private Sub mnuBrowser_Click()
      
      If lvwBrowser.ListItems.Count = 0 Then
            mnuBrowserOpenDefault.Enabled = False
            mnuBrowserOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
            mnuBrowserDelete.Enabled = False
            mnuBrowserRename.Enabled = False
            Exit Sub
      Else
            mnuBrowserOpenDefault.Enabled = True
            mnuBrowserOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
            mnuBrowserDelete.Enabled = True
            mnuBrowserRename.Enabled = True
      End If
      
      With lvwBrowser.SelectedItem
            If .Icon = EFileType.Directory Or .Icon = EFileType.Drive Then
                  mnuBrowserOpenDefault.Caption = "Explore Selected..." & vbTab & "Shift+Ctrl+Enter"
                  mnuBrowserDelete = False
                  If .Text = ".." Or .Icon = EFileType.Drive Then mnuBrowserRename = False
            End If
      End With

End Sub

Private Sub mnuBrowserDelete_Click()
      BrowserDeleteSelected
End Sub

Private Sub BrowserDeleteSelected()
      Dim iBookm As Integer, iRetVal As Integer
      Dim sTheDamned As String
      
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      sTheDamned = gBrowserData.Dir & lvwBrowser.SelectedItem
      
      If gBrowserData.BookmarkMode Then
            
            iBookm = lvwBrowser.SelectedItem.Index      ' TODO: FIIIIIIIIXXXXXXXXX
            lvwBrowser.ListItems.Remove iBookm
            
            BookmarkSaveChanges
            
      ElseIf gBrowserData.DrivesMode Then
            Caption = "I WILL NOT DELETE YOUR DISK.  FIND SOMEONE ELSE."
      ElseIf Not FileExists(sTheDamned) Then
            Caption = "Can't delete what isn't there: " & sTheDamned
'      ElseIf sTheDamned = agEditor.Tag Then  'TODO: refresh or something here.
'            Caption = "Can't delete your open file.  Sorry."
      ElseIf GetAttr(sTheDamned) And vbDirectory Then
            Caption = "This program would rather not be held responsible for mass deletions.  Please use another."

'                  RmDir sTheDamned
'                  Caption = "Folder deleted successfully: " & sTheDamned
'                  btnRefresh_Click
      Else
            On Error Resume Next
            iRetVal = RecycleFile(sTheDamned)
            If Err > 0 Then
                  Caption = Err.Number & ": " & Err.Description
                  
            ElseIf iRetVal <> 0 Then
                  Caption = "Error " & iRetVal
            Else
                  If sTheDamned = agEditor.Tag Then mnuFileNew_Click
                  Caption = "File deleted successfully: " & sTheDamned
                  btnRefresh_Click
            End If
            On Error GoTo 0
      End If
End Sub

Private Sub mnuBrowserOpenDefault_Click()
      mnuListOpenDefault_Click
End Sub

Private Sub mnuBrowserRefresh_Click()
      btnRefresh_Click
End Sub

Private Sub mnuBrowserSort_Click()
      btnSort_Click
End Sub

Private Sub mnuEdit_Click()
      mnuEditUndo.Enabled = agEditor.CanUndo
      mnuEditRedo.Enabled = agEditor.CanRedo
End Sub

Private Sub mnuEditFont_Click()
      btnFont_Click
End Sub

Private Sub mnuEditRedo_Click()
      agEditor.Redo
End Sub

Private Sub mnuEditUndo_Click()
      agEditor.Undo
End Sub

'
'   What this really does is:
'      1. go to directory containing open file
'      2. select open file from list
'
Private Sub mnuFileCurrentDirectory_Click()
      Dim litCurrentFile As ListItem
      
      If agEditor.Tag = "" Then Exit Sub
      
      Set litCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.Tag))
      
      If litCurrentFile Is Nothing Then
            cboPath = SnipFileName(agEditor.Tag)
            Set litCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.Tag))
            If litCurrentFile Is Nothing Then
                  MsgBox "It seems that your file was deleted by another application." & _
                        "  If you wish to keep it, save at once!"
                  Exit Sub
            End If
      End If
      litCurrentFile.Selected = True
      litCurrentFile.EnsureVisible
End Sub

Private Sub mnuFileExit_Click()
      Unload Me
End Sub

Private Sub mnuFileHistory_Click(Index As Integer)
      Dim sEx As String
      
      sEx = gFSO.getextensionname(mnuFileHistory(Index).Tag)
      EditorLoadFile mnuFileHistory(Index).Tag, FileTypeFromExtension(sEx)
      
      mnuFileCurrentDirectory_Click
End Sub


Private Sub mnuFileOpen_Click()
      If mnuViewFilebrowser.Checked = False Then
            mnuViewFilebrowser.Checked = True
            mnuViewFilebrowser_Click
      End If
      lvwBrowser.SetFocus
End Sub

' ******************************************************
'   mnuFileParentDirectory
'
'   Take the browser up a folder.
'   Preserve filter except in a drives list.
' ******************************************************
Private Sub mnuFileParentDirectory_Click()
      Dim sParentDir As String
      
      If gBrowserData.DrivesMode Or gBrowserData.BookmarkMode Then Exit Sub
      
      sParentDir = ParentDirectoryOf(gBrowserData.Dir)
      
      If gBrowserData.Error Or sParentDir = "" Then
            cboPath = sParentDir
      Else
            cboPath = sParentDir & gBrowserData.Filter
      End If
End Sub

Private Sub mnuBrowserRename_Click()
      lvwBrowser.StartLabelEdit
End Sub


Private Sub mnuFileRename_Click()

      ' Rename an open file without saving as a new file or deleting anything.
      ' About thirty percent of my nervous ticks come from not having this simple
      ' feature at my disposal in other applications.
      
      ' PLEASE NOTE: this is not a "save as" with extras.  What has already been saved as
      ' sOldPath will be renamed sNewPath.  Your unsaved progress will not be tampered with,
      ' but NOR WILL IT BE SAVED, until you save it.
      
      ' TODO: Auto-select new file after the rename.
      ' Currently fucked because it's looking for the old name in btnRefresh_Click.

      Dim sOldPath As String, sNewPath As String
      
      sOldPath = agEditor.Tag
      sNewPath = InputBox("To what name would you rechristen this document, your majesty?", _
            "Rename", sOldPath)
      
      If Dir(SnipFileName(sNewPath), vbDirectory) = "" Then
            Caption = "Can't rename due to invalid directory: " & SnipFileName(sNewPath)
      
      ElseIf Not FileExists(sOldPath) Then
            Caption = "Can't rename what's not there: " & sOldPath
            btnRefresh_Click
      
      ElseIf sOldPath = sNewPath Then
            
            If StrComp(sOldPath, sNewPath, vbBinaryCompare) = 0 Then
                  
                  ' No change whatsoever.
            
            Else ' Change in caps only.  We'll rename it anyway, just to be a sport.
                  
                  On Error Resume Next
                  Name sOldPath As sNewPath
                  If Err > 0 Then
                        Caption = Err.Number & ": " & Err.Description
                  Else
                        Caption = "Renamed.  Even though all you changed was the capitalization.  Freak."
                        agEditor.Tag = sNewPath
                        btnRefresh_Click
                  End If
                  On Error GoTo 0
            
            End If
      
      ElseIf FileExists(sNewPath) Then
            Caption = "This name sucks: " & Chr(34) & sNewPath & Chr(34) & ".  Change it."
            
      Else   ' ...and finally, we rename a file.
            On Error Resume Next
            Name sOldPath As sNewPath
            If Err > 0 Then
                  Caption = Err.Number & ": " & Err.Description
            Else
                  Caption = "Rename successful: " & sNewPath
                  agEditor.Tag = sNewPath
                  btnRefresh_Click
                  btnCurrentDirectory_Click
            End If
            On Error GoTo 0
      
      End If
End Sub
Private Sub mnuFileSaveAs_Click()
      Dim sDefaultPath As String, sFileName As String
      Dim fSuccess As Boolean
      Dim dteSaveTime As Date
      Dim vDate As Variant

      If Not agEditor.Visible Then
            Caption = "ERROR: can't save in image mode."
            Exit Sub
      ElseIf chkReadOnly.Value = vbChecked Then
            Caption = "ERROR: can't save in Read Only mode."
            Exit Sub
      End If
      
      vDate = Date
      msPhlegmDate = Year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
     
      ' here we decide on a default file name to suggest to the user,
      ' based on a whether the editor.tag is empty, and whether the file browser is at a valid folder.
      If agEditor.Tag <> "" Then
            sDefaultPath = agEditor.Tag  ' It means this is not a new file we're saving.  Default to old name.
            
      ElseIf gBrowserData.ValidPath Then
            sDefaultPath = gBrowserData.Dir & msPhlegmDate & ".txt"  ' New file, good directory in browser.
      Else
            sDefaultPath = CurDir & "\" & msPhlegmDate & ".txt"  ' New file, no good directory present.
      End If
      
      sFileName = InputBox("File name:", "Save", sDefaultPath)
      fSuccess = agEditor.SaveToFile(sFileName, SF_TEXT)
      dteSaveTime = Now
      
      If Not fSuccess Then  ' That SaveToFile gives an error after successfully saving a blank.  Make special case for it.
            If FileExists(sDefaultPath) And agEditor.Text = "" Then fSuccess = True
      End If  ' TODO: should have checked for zero file length to match agEditor emptiness.

      If fSuccess Then
            staTusBar1.Panels(EStat.Modified) = ""
            agEditor.Tag = sFileName
            Caption = sFileName & "  (" & agEditor.CharacterCount & " bytes saved on " & dteSaveTime & ")"
            btnRefresh_Click
            btnCurrentDirectory_Click
      
      ElseIf sFileName <> "" Then  ' Empty string would have meant the user hit "Cancel".
            frmMain.Caption = "ERROR: cannot save to " & sFileName
      End If
End Sub

Private Sub mnuHelpAbout_Click()
      MsgBox "phlegmoirs " & App.Major & "." & App.Minor & "." & App.Revision, , ""
End Sub

Private Sub mnuHelpReadme_Click()
      EditorLoadFile CurDir & "\progress.txt"
End Sub

Private Sub mnuList_Click()
      
      ' This is the popup menu for lvwBrowser.  Click fires whenever the menu is popped up.
      
      ' Most menu items are enabled/disabled in lvwBrowser_ItemClick.
      ' Here, we un-set some of them if the user has clicked somewhere that is not a list item.
      
      ' Events happen in this order: lvwBrowser_MouseDown, lvwBrowser_ItemClick, mnuList_Click.
      
      ' mfBrowserItemClicked is set to False on the MouseDown, and True on the ItemClick.
      ' So if it gets here as False, that means ItemClick did not happen on this mouse event.
      
'      If Not mfBrowserItemClicked Then
'            mnuListOpenDefault.Enabled = False
'            mnuListDelete.Enabled = False
'            mnuListRename.Enabled = False
'            mnuListCopyPath.Enabled = False
'            mnuListShowOnly.Enabled = False
'            mnuListProperties.Enabled = False
'      End If
End Sub

Private Sub mnuListCancel_Click()
      SendKeys "{ESC}"
End Sub

Private Sub mnuListCopyPath_Click()
      Clipboard.Clear
      If gBrowserData.BookmarkMode Then
            Clipboard.SetText lvwBrowser.SelectedItem
      Else
            Clipboard.SetText gBrowserData.Dir & lvwBrowser.SelectedItem
      End If
End Sub

Private Sub mnuListDelete_Click()
      ' TODO: remove bookmark from list here, but also need buttons,
      
      ' menu disabled in lvwBrowser_MouseUp unless item is clicked on.
            
      BrowserDeleteSelected
      
End Sub

Private Sub lblDivider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
      If lblDivider.MousePointer = vbSizeWE And lblDivider.Tag = "" Then
            
            lblDivider.Tag = "Resizing"
      End If
End Sub

Private Sub lblDivider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim iOffset As Integer

      If lblDivider.MousePointer = vbSizeWE And lblDivider.Tag = "Resizing" Then
            With picEditor
'                  .Visible = False
                  iOffset = BrowserResizeHorizontal(x + lblDivider.Left)
                  .Move .Left + iOffset, .Top, .Width - iOffset, .Height
                  agEditor.Move 0, 0, picEditor.Width, picEditor.Height
'                  .Visible = True
            End With
      Else
            lblDivider.MousePointer = vbSizeWE
      End If
End Sub

Private Sub lblDivider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      If lblDivider.MousePointer = vbSizeWE Then
            lblDivider.MousePointer = 0
            lblDivider.Tag = ""
'            agEditor.Width = picBrowser.Width + 160
'            agEditor.Left = frmMain.Width - agEditor.Width - 150
      End If

End Sub


Private Sub mnuListOpen_Click()
      BrowserExecuteItem lvwBrowser.SelectedItem
End Sub

Private Sub mnuListOpenDefault_Click()
      Dim sPath As String
      
      If lvwBrowser.ListItems.Count > 0 Then
            If gBrowserData.BookmarkMode Then
                  sPath = lvwBrowser.SelectedItem.Text
            Else
                  sPath = gBrowserData.Dir & lvwBrowser.SelectedItem.Text
            End If
            ShellExecute 0, "open", sPath, 0, "", SW_RESTORE
      End If
End Sub

Private Sub mnuListProperties_Click()
      ' SImply calls the Explorer file properties dialog.  Hope this works.
      
      ShowFileProperties gBrowserData.Dir & lvwBrowser.SelectedItem
End Sub

Private Sub mnuListRename_Click()
      lvwBrowser.StartLabelEdit
      
End Sub

'   Show only files of extension sEx.
'
Private Sub mnuListShowOnly_Click()
      Dim sEx As String
      
      sEx = gFSO.getextensionname(lvwBrowser.SelectedItem)
      If sEx <> "" Then sEx = "." & sEx
      cboPath = gBrowserData.Dir & sEx
End Sub

Private Sub mnuQueries_Click(Index As Integer)
      Dim iQuery As Integer
      
      ' Uncheck everything except the menu item clicked.
      ' This sub can be called with a bogus Index as the parameter, to deselect everything.
      
      For iQuery = mnuQueries.LBound To mnuQueries.UBound
            mnuQueries(iQuery).Checked = (iQuery = Index)
      Next iQuery
      
      If Index >= mnuQueries.LBound Then  ' if it's not a bogus index...
            
            ' Deselect the option buttons, too, since all the queries are mutually exclusive.
            
            For iQuery = optQueries.LBound To optQueries.UBound
                  optQueries(iQuery).Value = False
            Next iQuery
            
            miCurrentQuery = Index
      End If
End Sub

Private Sub mnuQuery_Click()
      Debug.Print "mnuQuery_Click"
End Sub

Private Sub mnuQueryMatchCase_Click()
      mnuQueryMatchCase.Checked = Not mnuQueryMatchCase.Checked
End Sub

Private Sub mnuQueryWholeWord_Click()
      mnuQueryWholeWord.Checked = Not mnuQueryWholeWord.Checked
End Sub


Private Sub mnuViewFind_Click()
      If ActiveControl.name = "agEditor" Then txtQueryBox = Trim(agEditor.SelectedText)
      
      optQueries(EQuery.Find).Value = True
      txtQueryBox.SetFocus
End Sub

Private Sub mnuViewFindBackwards_Click()
      btnFindPrev_Click
End Sub

Private Sub mnuViewFindNext_Click()
      btnQueryExecute_Click
End Sub

Private Sub mnuViewOptions_Click()
      frmOptions.Show
End Sub

Private Sub mnuViewReadOnly_Click()
      chkReadOnly.Value = Abs(chkReadOnly.Value - 1)
End Sub

Private Sub optQueries_Click(Index As Integer)
      '
      
      mnuQueries_Click (-1) '  Deselect the menu queries.
      
      miCurrentQuery = Index
      'If txtQueryBox <> "" Then btnQueryExecute_Click
      txtQueryBox.SetFocus
End Sub


Private Sub optQueries_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
      Select Case Index
      
            Case EQuery.Find
                  'btnQueryExecute_Click
                  
      End Select
End Sub




Private Sub picEditor_KeyDown(KeyCode As Integer, Shift As Integer)
      
      'Debug.Print KeyCode
      Select Case KeyCode
            Case 107, 187 ' "+" and Keypad "+"
                  If Shift = 0 Then
                        ImageZoomIn sliZoom.SmallChange
                  ElseIf Shift = vbCtrlMask Then
                        ImageZoomIn sliZoom.LargeChange
                  End If
            Case 109, 189 ' "-" and Keypad "-"
                  If Shift = 0 Then
                        ImageZoomOut sliZoom.SmallChange
                  ElseIf Shift = vbCtrlMask Then
                        ImageZoomOut sliZoom.LargeChange
                  End If
            Case vbKey0, 106 ' 0 and Keypad "*" -- reset position and size.
                  sliZoom.Value = 100
                  Image1.Move 0, 0, miImageWidthDefault, miImageHeightDefault
            Case 107, 55   ' 7 and Keypad 7
                  sliZoom.Value = sliZoom.Value / 2
            Case 104, 56   ' 8 and Keypad 8
                  sliZoom.Value = sliZoom.Value * 2
            Case vbKeyDown
                  Image1.Top = Image1.Top + MoveIncrement
            Case vbKeyUp
                  Image1.Top = Image1.Top - MoveIncrement
            Case vbKeyLeft
                  Image1.Left = Image1.Left - MoveIncrement
            Case vbKeyRight
                  Image1.Left = Image1.Left + MoveIncrement
                  
            Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                  If Shift = 0 Then BrowserExecuteNext
            Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                  If Shift = 0 Then BrowserExecutePrev
      End Select
End Sub

Private Sub picEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      On Error Resume Next
      If Not (ActiveControl.name = "picEditor") Then picEditor.SetFocus
      On Error GoTo 0
End Sub

Private Sub sliZoom_Change()
      ImageSetZoom (sliZoom.Value)
End Sub

Private Sub sliZoom_Scroll()
      ImageSetZoom (sliZoom.Value)
End Sub


'Private Sub tmrMouseDown_Timer()
'      tmrMouseDown.Tag = CInt(tmrMouseDown.Tag) + 1
'      'if tmrmousedown.Tag >
'End Sub

Private Sub tmrRevertTips_Timer()
      staTusBar1.Panels(EStat.Tips) = staTusBar1.Panels(EStat.Tips).Tag
      tmrRevertTips.Enabled = False
End Sub


'Private Sub txtQueryBox_Change()
'      Dim pos As Integer
'      Dim quickkey As String
'      Dim NewQuery As URLQueryType
'
'      pos = InStr(0, txtQueryBox, " ", )
'      NewQuery.key = Left(txtQueryBox, pos)
'      NewQuery.URL = Right(txtQueryBox, pos)
'End Sub

Private Sub txtQueryBox_GotFocus()
      txtQueryBox.SelStart = 0
      txtQueryBox.SelLength = Len(txtQueryBox)
End Sub

Private Sub txtQueryBox_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn And Shift = 0 Then
            If txtQueryBox <> "" Then btnQueryExecute_Click
      ElseIf KeyCode = vbKeyReturn And Shift = vbShiftMask And miCurrentQuery = EQuery.Find Then
            If txtQueryBox <> "" Then btnFindPrev_Click
      End If
End Sub

Private Sub txtQueryBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
      txtQueryBox = Data.GetData(vbCFText)
      txtQueryBox_KeyDown vbKeyReturn, 0
End Sub

Private Sub txtQueryBox_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
      txtQueryBox.SelStart = 0
      txtQueryBox.SelLength = Len(txtQueryBox)
End Sub

Private Sub agEditor_KeyDown(KeyCode As Integer, Shift As Integer)
'      gpOldProc = SetWindowLong(agEditor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)
      Select Case KeyCode
            Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                  If Shift = 0 And chkReadOnly.Value = vbChecked Then BrowserExecuteNext
            Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                  If Shift = 0 And chkReadOnly.Value = vbChecked Then BrowserExecutePrev
'            Case vbKeyAdd
'                  If Shift = vbAltMask Then
'                        SendMessage agEditor.RichEdithWnd, EM_SETFONTSIZE, ByVal 1, 0
''                        Set fntTemp = agEditor.GetFont
''                        fntTemp.Size = fntTemp.Size + 1
''                        agEditor.SetFont fntTemp, , , , ercSetFormatAll
'                        Me.Caption = agEditor.GetFont.Name & " " & agEditor.GetFont.Charset & " " & agEditor.GetFont.Size
'                  End If
'
'            Case vbKeySubtract
'                  If Shift = vbAltMask Then
'                        Set fntTemp = agEditor.GetFont
'                        fntTemp.Size = fntTemp.Size - 1
'                        agEditor.SetFont fntTemp, , , , ercSetFormatAll
'                        Me.Caption = agEditor.GetFont.Name & " " & agEditor.GetFont.Charset & " " & agEditor.GetFont.Size
'                  End If
                  
      End Select
'      retval = SetWindowLong(agEditor.RichEdithWnd, GWL_WNDPROC, gpOldProc)
'      gpOldProc = 0
End Sub

Private Sub agEditor_SelectionChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As agricheditbox.ERECSelectionTypeConstants)
      ' Update a few items on the status bar.
      
      Dim lLineIndex As Long, lCharIndex As Long
      Dim chrSelection As CHARRANGE
      
      lLineIndex = agEditor.CurrentLine
      lCharIndex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lLineIndex, 0)
      
      If staTusBar1.Visible Then
            With mudtStats
                  
                .y = lLineIndex + 1
                
                ' We want mudtStats.i to count CRs and LFs both, since agEditor.CharacterCount does that.
                .i = lMin
                SendMessage agEditor.RichEdithWnd, EM_EXGETSEL, 0, chrSelection
                .x = lMin - lCharIndex + 1
                .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
            End With
        
            FillStats
            'staTusBar1.Panels(EStat.SelText) = Len(agEditor.SelectedContents(SF_TEXT))
            staTusBar1.Panels(EStat.SelText) = lMax - lMin
      End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'      Dim lControlHWnd As Long
'      Dim poiCursor As POINTAPI
'
'      'gpOldProc = SetWindowLong(agEditor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)
'      Select Case KeyCode
'            Case vbKeyTab
'                  If Shift = vbCtrlMask And Me.ActiveControl.name = "agEditor" Then
'                        lvwBrowser.SetFocus
'                  ElseIf Me.ActiveControl.name = "agEditor" Then
'
'                  Else
'
'                  End If
'
'            Case 93
'                  GetCursorPos poiCursor
'                  lControlHWnd = WindowFromPoint(poiCursor.x, poiCursor.y)
'                  staTusBar1.Panels(EStat.LastSaved) = "c: " & lControlHWnd
'
'            Case vbKeyB
'                  Me.Caption = Me.ActiveControl.name
'
'            Case Else
'                  staTusBar1.Panels(EStat.LastSaved) = KeyCode
'      End Select
''      retval = SetWindowLong(agEditor.RichEdithWnd, GWL_WNDPROC, gpOldProc)
''      gpOldProc = 0

      Select Case KeyCode
            Case 220 '  Backslash.  Making Alt+\ into a spare Tab key for the right side of the keyboard.
                  If Shift And vbAltMask Then
                        SendKeys "{TAB}"
                  End If
                  
            Case vbKeyReturn           ' File Properties window from Explorer!
                  If Shift And vbAltMask Then
                        If ActiveControl.name = "lvwBrowser" And lvwBrowser.ListItems.Count > 0 Then
                              ShowFileProperties gBrowserData.Dir & lvwBrowser.SelectedItem
                        Else
                              ShowFileProperties agEditor.Tag
                        End If
                  End If
                  
            Case vbKeyF
                  If Shift = vbCtrlMask + vbShiftMask Then ' TODO: this is still writing letters to the editor.
                        btnFont_Click
                  End If
            
            Case 221 ' Right Bracket "]"
                  If Shift = vbCtrlMask Then BrowserExecuteNext
            
            Case 219 ' Left Bracket "["
                  If Shift = vbCtrlMask Then BrowserExecutePrev
            
            Case 190 ' Period (".") -- opens popup menu for find options
                  If Shift = vbAltMask Then
                        If chkQueryDotDotDot.Value = vbUnchecked Then
                              chkQueryDotDotDot.SetFocus
                              chkQueryDotDotDot.Value = vbChecked
                              
                        ElseIf chkQueryDotDotDot.Value = vbChecked Then
                              ' Same button closes menu, if already opened
                              chkQueryDotDotDot.Value = vbUnchecked
                        End If
                  End If
            
            Case vbKeyEscape  ' Hotfix, popup menu doesn't wanna die by itself
                  If Shift = 0 Then
                        If chkQueryDotDotDot.Value = vbChecked Then
                              chkQueryDotDotDot.Value = vbUnchecked
                        End If
                  End If
      End Select
End Sub



Private Sub Form_Load()
      Dim vDate As Variant
      Dim sCommandFile As String

      InitializeMenus
            
      Set gFSO = CreateObject("Scripting.FileSystemObject") ' Just so I'll never have to do this again.
      
      gBrowserData.ListEmpty = True
      
      miEditorMode = EFileType.Text

'      miImageZoom = 100
      
      'Debug.Print "command line sayeth: [" & Command() & "]"
      sCommandFile = Trim(Command())
      If Left(sCommandFile, 1) = Chr(34) Then sCommandFile = Mid(sCommandFile, 2, Len(sCommandFile) - 2)
      If sCommandFile <> "" And Not (sCommandFile Like "*:\*") Then sCommandFile = CurDir & "\" & sCommandFile
      agEditor.Tag = sCommandFile
      
      msPhlegmKey = "Software\" & App.Title & "\" & msSettingsVersion
      
      vDate = Date
      msPhlegmDate = Year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
      
      GetWindowSettings
      mudtStats.imax = CharacterCount(agEditor)
      FillStats
      staTusBar1.Panels(EStat.Modified) = ""

      If Not Debugging Then
            gpOldLvwBrowserProc = SetWindowLong(lvwBrowser.hwnd, GWL_WNDPROC, _
                  AddressOf SuppressArrowKeysProc)
      End If
      
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      ' If we open a popupmenu, and then right click off into space,
      '   the mousedown event is called for the form (not for the control we are
      '   hovering over nor the menu itself.)
      ' Our form doesn't need it.  We'll have him pass it to the control it's over.

      Dim ctrlhWnd As Long
      Dim retval As Long
      Dim poiCursor As POINTAPI

      'debug.print "Form_Mousedown " & Button & " " & Shift
      
      If Button <> vbRightButton Or Shift <> 0 Then Exit Sub

      GetCursorPos poiCursor
      ctrlhWnd = WindowFromPoint(poiCursor.x, poiCursor.y)
      
      On Error Resume Next
      If Screen.ActiveControl.Container.name = "picQuery" Then
            txtQueryBox.SetFocus
      End If
      On Error GoTo 0
      
'      If Not mfSkipMouseEventCrap And ctrlhWnd = agEditor.RichEdithWnd Then
'            mouse_event MOUSEEVENTF_LEFTDOWN, poiCursor.x, poiCursor.y, 0, 0
'            mouse_event MOUSEEVENTF_LEFTUP, poiCursor.x, poiCursor.y, 0, 0
'      ElseIf ctrlhWnd = lvwBrowser.hwnd Then
'            FuckIHateThis = True
'            mouse_event MOUSEEVENTF_LEFTDOWN, poiCursor.x, poiCursor.y, 0, 0
'            mouse_event MOUSEEVENTF_RIGHTUP, poiCursor.x, poiCursor.y, 0, 0
'      ElseIf ctrlhWnd = txtQueryBox.hwnd Then
'            mouse_event MOUSEEVENTF_LEFTDOWN, poiCursor.x, poiCursor.y, 0, 0
'            'mouse_event MOUSEEVENTF_LEFTUP, poiCursor.X, poiCursor.Y, 0, 0
'      Else
'            SendMessage FMain.hwnd, WM_CANCELMODE, 0, 0
'      End If
End Sub

Private Sub Form_Resize()
      If mfSkipFormResize Then
'            Beep
      Else
            RearrangeControls
      End If
End Sub

Private Sub lvwBrowser_KeyDown(KeyCode As Integer, Shift As Integer)
      ' Left = up folder.  Right = open folder.
      ' Trying to copy the functionality of explorer,
      ' somehow, but without a visible tree.
      
      Dim iIndex As Integer
      
      'debug.print "KEYDOWN " & KeyCode & " " & Shift
      
      Select Case KeyCode
                        
            Case vbKeyLeft
                  If Shift = vbAltMask Then ' Alt+Left = go back in the recent paths list
                        PathBack
                  ElseIf (Shift And vbCtrlMask) Then
                        ' Ctrl+left = scroll left.  No additional coding needed.
                  Else
                        mnuFileParentDirectory_Click   ' Ordinary left arrow...
                  End If
                                                 
                                                 
            Case vbKeyF13 ' F13, but contains code for it and for right arrow.
                              ' See SuppressArrowKeysProc for details.
                              
                  ' Right = open a folder or a drive, but leave a file alone.
                  '     ...and don't fucking scroll anywhere.
                  
                  If lvwBrowser.ListItems.Count > 0 Then
                        If lvwBrowser.SelectedItem.ListSubItems(1) = 0 And Shift = 0 Then
                              BrowserExecuteItem lvwBrowser.SelectedItem
                        End If
                  End If
                  
            
            Case vbKeyRight
            
                  ' Alt+Right = go forward in the recent paths list
                  
                  If Shift = vbAltMask Then PathForward
                  
                  ' Ctrl+right = scroll right.  (Requires no additional code.)
                  
            Case vbKeyDelete
                  If Shift = 0 Then BrowserDeleteSelected
                  
            Case 219 ' Left Bracket [
                  If Shift = 0 Then BrowserExecutePrev
            
            Case 221 ' Right Bracket ]
                  If Shift = 0 Then BrowserExecuteNext
                  
            Case 220 ' Backslash \
                  If Shift = 0 Then SendKeys "{TAB}"
            
            Case 93 ' That keyboard button that usually means right click.
                  Dim iItemX As Integer, iItemY As Integer
                  
                  iItemX = picBrowser.Left + lvwBrowser.Left + lvwBrowser.SelectedItem.Left + lvwBrowser.SelectedItem.Width
                  iItemY = picBrowser.Top + lvwBrowser.Top + lvwBrowser.SelectedItem.Top + lvwBrowser.SelectedItem.Height
                  Me.PopupMenu mnuList, , iItemX, iItemY, mnuListOpen
                  
      End Select
End Sub

Private Sub lvwBrowser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      'debug.print "lvwBrowser_MOUSEUP " & Button & " " & Shift
'      If FuckIHateThis Then
'            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'            FuckIHateThis = True
'      End If
      
      Dim litHoverItem As ListItem
      
      Set litHoverItem = lvwBrowser.HitTest(x, y)  ' To see if we're over an item.
      
      If (Button = vbRightButton And Shift = 0) Then
            Me.PopupMenu mnuList
      
      ElseIf Button = vbLeftButton And Shift = 0 Then
            
            If Not (litHoverItem Is Nothing) Then
                  ' Open the file/folder on an ordinary left click.
                  BrowserExecuteItem lvwBrowser.SelectedItem
            Else
                  ' Clicking on empty space deselects the selected item.
                  If Not gBrowserData.ListEmpty Then lvwBrowser.SelectedItem.Selected = False
            End If
      
      End If
      
      ' For use in the click event, so we know what was clicked.  OBSOLETE.
      miBrowserMouseButton = Button
      miBrowserShift = Shift
End Sub

Private Sub mnuFileNew_Click()
      
      ' TODO: this needs default behavior.
      
      Dim sDefaultName As String
      
      sDefaultName = msPhlegmDate & ".txt"
      If FileExists(sDefaultName) = False Then
            
      End If
      
      Image1.Visible = False
'      Image1.Picture = LoadPicture
      agEditor.Text = ""
      agEditor.Tag = ""
      EditorSetMode EFileType.Text
      frmMain.Caption = "(New File)"
End Sub

Private Sub mnuFileSave_Click()
      Dim fSuccess As Boolean
      Dim dteSaveTime As Date
      
      If agEditor.Tag = "" Then  ' Saving a nameless New File
            mnuFileSaveAs_Click
            Exit Sub
      
      ElseIf Not agEditor.Visible Then
            Caption = "ERROR: can't save in image mode."
            Exit Sub
      End If
      
      If chkReadOnly.Value = vbChecked Then
            Caption = "ERROR: can't save in Read Only mode."
            Exit Sub
      End If
      
      fSuccess = agEditor.SaveToFile(agEditor.Tag, SF_TEXT)
      dteSaveTime = Now

      If fSuccess Then
            staTusBar1.Panels(EStat.Modified) = ""
            Caption = agEditor.Tag & "  (" & agEditor.CharacterCount & " bytes saved on " & dteSaveTime & ")"
      Else
            frmMain.Caption = "ERROR: cannot save to " & agEditor.Tag
      End If
End Sub

Private Sub mnuViewDictionary_Click()
      If ActiveControl.name = "agEditor" And Trim(agEditor.SelectedText) <> "" Then
            txtQueryBox = Trim(agEditor.SelectedText)
      End If
      optQueries(EQuery.Dictionary).Value = True
      
      
'      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
'      If txtQueryBox.Visible Then
'            txtQueryBox.SetFocus
'      Else
'            txtQueryBox_GotFocus
'      End If
'      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
'      txtQueryBox.SetFocus
'      If agEditor.SelectedText <> "" Then txtQueryBox_KeyPress vbKeyReturn
End Sub

Private Sub mnuViewFilebrowser_Click()
    chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub agEditor_Change()

      staTusBar1.Panels(EStat.Modified) = "Modified"
      
      If staTusBar1.Visible Then
            With mudtStats
                .imax = CharacterCount(agEditor)
                .ymax = agEditor.LineCount
            End With
            
            FillStats
      End If
      
      mfFindQueryChanged = True
      miFindResults = 0
End Sub

Private Sub agEditor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      If (Button = vbRightButton And Shift = 0) Then
          Me.PopupMenu mnuWrite
      End If
End Sub

Private Sub FillStats()

      staTusBar1.Panels(EStat.Stats) = "Char: " & mudtStats.i & "/" & mudtStats.imax _
            & "  Ln: " & mudtStats.y & "/" & mudtStats.ymax & "  Col: " & mudtStats.x _
            & "/" & mudtStats.xmax
End Sub


Private Sub RearrangeControls()

      ' TODO: clean up these godawful variable names!

      ' Put the various controls where they need to be.
      '   agEditor, lvwBrowser
      ' Made to go on a window resize or when showing or hiding a control
      
      Dim iEdHeight As Integer, iEdWidth As Integer, iEdTop As Integer, iEdLeft As Integer
      Dim lineindex As Long, charindex As Long, lMin As Long, lMax As Long
      Dim fValidWindowSize As Boolean, iRedoResizeX As Integer, iRedoResizeY As Integer
      Dim iPicBoxMarginsY As Integer, iFormMarginsX As Integer, iFormMarginsY As Integer
      Dim sHadFocus As String
      
      Const topmargin = 100
      Const leftmargin = 0
      Const rightmargin = 150
      Const midspace = 100
      Const bottommargin = 60
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      fValidWindowSize = True ' ...until proven guilty.
      iRedoResizeY = frmMain.Height
      iRedoResizeX = frmMain.Width
      
      If Not (ActiveControl Is Nothing) Then  ' activecontrol is nothing if image1 is up front...
            sHadFocus = ActiveControl.name                               ' images cannot take focus.
            picEditor.Visible = False ' MUCH faster if you turn him off while thinking (unless he's empty).
      End If
      
      ' Calculate control positions...
      
      iEdTop = topmargin
      If mnuViewToolbar.Checked Then iEdTop = iEdTop + picToolBox.Height
      
      iEdHeight = frmMain.ScaleHeight - iEdTop - bottommargin
      If mnuViewStatusBar.Checked Then iEdHeight = iEdHeight - staTusBar1.Height
      
      iEdLeft = leftmargin
      If mnuViewFilebrowser.Checked Then iEdLeft = iEdLeft + picBrowser.Width
      
      iEdWidth = frmMain.ScaleWidth - iEdLeft
      
      
      ' Check to see if we've gone out of bounds...
      
            ' Caution: iEdWidth would come back around a second time as 1499.
            ' I *think* that I fixed it with that 1510 down there, rather than 1500.
      If iEdWidth < 1500 And WindowState = vbMaximized Then
            BrowserResizeHorizontal picBrowser.Width
            RearrangeControls
            Exit Sub
      ElseIf iEdWidth < 1500 Then
            fValidWindowSize = False
            iFormMarginsX = frmMain.Width - frmMain.ScaleWidth
            iRedoResizeX = iEdLeft + iFormMarginsX + 1510
      End If
      
      If iEdHeight < 1500 Then
            fValidWindowSize = False
            iFormMarginsY = frmMain.Height - frmMain.ScaleHeight
            iPicBoxMarginsY = picBrowser.Height - picBrowser.ScaleHeight
            iRedoResizeY = iEdTop + lvwBrowser.Top + iPicBoxMarginsY + iFormMarginsY + 1510
      End If
      
      If Not fValidWindowSize Then
            frmMain.Move Left, Top, iRedoResizeX, iRedoResizeY
            Exit Sub
      End If
      
      ' It's all good.  Move the controls now!
      
      picEditor.Move iEdLeft, iEdTop, iEdWidth, iEdHeight
      agEditor.Move 0, 0, iEdWidth, iEdHeight
      
      lvwBrowser.Height = iEdHeight - lvwBrowser.Top + topmargin
            
            
      ' a few things in the statusbar could change in a window resize:
      '   x, xmax, y, ymax
      ' and some shouldn't change:
      '   i, imax,   (we're not adding or deleting characters or moving the cursor)
      '   sellength
      
      agEditor.GetSelection lMin, lMax
      lineindex = agEditor.CurrentLine
      charindex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lineindex, 0)
      
      If staTusBar1.Visible Then
            With mudtStats
                .x = lMin - charindex + 1
                .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal charindex, 0) + 1
                .y = lineindex + 1
                .ymax = agEditor.LineCount
            End With
            FillStats
      End If
      
      picEditor.Visible = True
      'If sHadFocus = "agEditor" Then agEditor.SetFocus
End Sub

Private Sub mnuViewStatusBar_Click()
      staTusBar1.Visible = Not staTusBar1.Visible
      mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
      RearrangeControls
End Sub

Private Sub mnuViewThesaurus_Click()
      If ActiveControl.name = "agEditor" And Trim(agEditor.SelectedText) <> "" Then
            txtQueryBox = Trim(agEditor.SelectedText)
      End If
      optQueries(EQuery.thesaurus).Value = True
      
'      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
'      If txtQueryBox.Visible Then
'            txtQueryBox.SetFocus
'      Else
'            txtQueryBox_GotFocus
'      End If
'      If agEditor.SelectedText <> "" Then
'            ShellExecute 0, "open", "http://thesaurus.reference.com/search?q=" & txtQueryBox, 0, "", 8
'            Me.SetFocus
'      End If
End Sub

Private Sub InitializeMenus()
'      Dim tempinfo As MENUITEMINFO
'      Dim hMenu As Long, retval As Long
'
'      hMenu = GetMenu(hwnd)
'      hMenu = GetSubMenu(hMenu, 2)
'      retval = ModifyMenu(hMenu, 0, MF_STRING + MF_BYPOSITION, 2, "&Penis" + vbTab + "Ctrl+P")
      
      mnuEditIncFont.Caption = "&Increase Font Size" & vbTab & "Alt+="
      mnuEditUndo.Caption = "Undo" & vbTab & "Ctrl+Z"
      mnuEditRedo.Caption = "Redo" & vbTab & "Ctrl+Y"
      mnuEditFont.Caption = "Font..." & vbTab & "Shift+Ctrl+F"
      
      
      mnuWriteCut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnuWriteCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuEditCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuWritePaste.Caption = "&Paste" & vbTab & "Ctrl+V"
      
      mnuWindowMinimize.Caption = mnuWindowMinimize.Caption & vbTab & "Alt+F10"
      mnuWindowMaximize.Caption = mnuWindowMaximize.Caption & vbTab & "Alt+F12"
      mnuWindowRestore.Caption = mnuWindowRestore.Caption & vbTab & "Alt+F11"
      
      mnuBrowserDelete.Caption = mnuBrowserDelete.Caption & vbTab & "Del"
      
      mnuListDelete.Caption = mnuListDelete.Caption & vbTab & "Del"
      mnuListProperties.Caption = "&Properties" & vbTab & "Alt+Enter"
      mnuListCancel.Caption = "&Cancel" & vbTab & "Esc"
End Sub

Private Function BrowserGetDrives() As Integer
      ' Find all logical drives and display them in lvwBrowser
      ' Returns the number of logical drives found.
      
      Dim sDrivesFixed As String * 255
      Dim sDriveString As String
      Dim sDriveArray() As String
      Dim sNextDrive As String, iDriveIcon As Integer
      Dim lLength As Long
      Dim iIndex As Integer
      Dim litCurrentItem As ListItem
      
            
      lLength = GetLogicalDriveStrings(100, sDrivesFixed)
      sDriveString = Left(sDrivesFixed, lLength)
      sDriveArray = Split(sDriveString, Chr(0)) ' "(x,x, , )" is an error.  don't put in more commas unless
      lvwBrowser.ListItems.Clear          ' they lead to something.
      lvwBrowser.Tag = ""
      
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      
      iIndex = LBound(sDriveArray)
      sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      
      Do While (sNextDrive <> "") And (sNextDrive <> Chr(0))
            
            Select Case gFSO.getdrive(sNextDrive).drivetype
                  Case 1: iDriveIcon = EFileType.Floppy
                  Case 2: iDriveIcon = EFileType.Drive
                  Case 3: iDriveIcon = EFileType.Network
                  Case 4: iDriveIcon = EFileType.Cdrom
            End Select
            Set litCurrentItem = lvwBrowser.ListItems.Add( _
                  1, , sNextDrive, iDriveIcon, iDriveIcon)
            litCurrentItem.ListSubItems.Add , , 0
            
            iIndex = iIndex + 1
            sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      Loop
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = 1
      BrowserGetDrives = iIndex - 1
      
      staTusBar1.Panels(EStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " drives"
End Function
Private Sub mnuViewToolbar_Click()
      picToolBox.Visible = Not picToolBox.Visible
      mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
      RearrangeControls
End Sub

Private Sub mnuViewWordWrap_Click()
      chkWordWrap.Value = Abs(chkWordWrap.Value - 1)
End Sub

Private Sub mnuWindowMaximize_Click()
      Me.WindowState = vbMaximized
End Sub

Private Sub mnuWindowMinimize_Click()
      Me.WindowState = vbMinimized
End Sub

Private Sub mnuWindowMove_Click()
      'Me.Move
End Sub

Private Sub mnuWindowRestore_Click()
      Me.WindowState = vbNormal
End Sub

Private Sub mnuWindowSaveSettings_Click()
      SaveWindowSettings
End Sub

Private Sub mnuWriteCopy_Click()
      agEditor.Copy
End Sub

Private Sub mnuWriteCut_Click()
      agEditor.Cut
End Sub

Private Sub mnuWritePaste_Click()
      agEditor.Paste
End Sub

Private Function EditorLoadFile(sFileName As String, Optional iMode As Integer = EFileType.Text) As Boolean
      
      ' TODO: (un)IMPORTANT:
      ' To manage file loading properly, I cannot use LoadFromFile.
      ' For one thing, it returns false for an empty file.  I'd rather it make a distinction between an error
      ' and an empty (FIXED).  And there's the issue of interrupting the display.  It's great that I can
      ' interrupt the loading of the file to a buffer of some sort, but displaying that buffer is slower.
      
      Dim fLoadSuccess As Boolean
      
      If mfEditorLoading Then agEditor.Text = ""
      mfEditorLoading = True
      
      If Trim(sFileName) = "" Then  ' Blank means start a new file.  For when the registry settings come up empty.
            mnuFileNew_Click
      
      ElseIf Not FileExists(sFileName) Then
            frmMain.Caption = "ERROR: file does not exist."
            agEditor.Tag = ""
            
      Else ' Normal file load.
            
            EditorSetMode iMode
            
            Select Case iMode
                  
                  ' TODO: MUCH BIGGER PROBLEM WITH EDITORLOADFILE:
                  ' can't handle too long a filename.  127 characters was too long.  I dunno the limit just yet.
                  
                  Case EFileType.Text, EFileType.Other, EFileType.rtf
                        ' pass along the boolean return value, if anyone wants it.
                        fLoadSuccess = agEditor.LoadFromFile(sFileName, SF_TEXT)
                  
'                  Case EFileType.rtf
'                        fLoadSuccess = agEditor.LoadFromFile(sFileName, SF_RTF)
                        
                  Case EFileType.Picture
                        fLoadSuccess = True
                        On Error Resume Next
                        Image1.Picture = LoadPicture(sFileName)
                        If Err > 0 Then
                              Caption = "ERROR: " & sFileName & ", picture couldn't load"
                              fLoadSuccess = False
                        End If
                        On Error GoTo 0
                        
                        miImageWidthDefault = Image1.Picture.Width * 0.567  ' HiMetric to Twip conversion
                        miImageHeightDefault = Image1.Picture.Height * 0.567
                        ImageSetZoom (sliZoom.Value)
            End Select
                  
            If (fLoadSuccess = True) Or (FileSize(sFileName) = 0) Then  ' Success!
                  
                  agEditor.Tag = sFileName
                  frmMain.Caption = sFileName
                  staTusBar1.Panels(EStat.Modified) = ""
                  agEditor.SetSelection 0, 0
                  AddToHistory sFileName
            
            Else  ' Miscellaneous Failure!  agEditor returns no clues as to the problem.
                  frmMain.Caption = "WEIRD ERROR.  command() = " & Chr(34) & Command() & Chr(34) _
                        & "; Tag = " & Chr(34) & sFileName & Chr(34)
                  agEditor.Tag = ""
            End If
      End If
      
      EditorLoadFile = fLoadSuccess
      mfEditorLoading = False
End Function

Private Sub ImageSetZoom(iZoom As Integer)
      Image1.Stretch = True
      Image1.Move Image1.Left, Image1.Top, miImageWidthDefault * CSng(iZoom) / 100#, _
            miImageHeightDefault * CSng(iZoom) / 100#
'      miImageZoom = iZoom
      Caption = agEditor.Tag & "  (" & iZoom & "%)"
End Sub

Private Sub SaveWindowSettings()
      Dim lMin As Long, lMax As Long, lKey As Long, lRetVal As Long
      Dim lNewOrUsed As Long, lValueSize As Long
      Dim iIndex As Integer
      
'      Dim wnpPlacement As WINDOWPLACEMENT'
'      Dim rectRestored As RECT
      Dim fntTemp As New StdFont
'      Dim poiTemp As POINTAPI
      
      
      With mudtSettings
            .WNP.Length = LenB(.WNP)
            GetWindowPlacement hwnd, .WNP
            If .WNP.showCmd = SW_MINIMIZE Then
                  .WNP.showCmd = SW_RESTORE
            ElseIf .WNP.showCmd = SW_SHOWMINIMIZED Then  '  <-- It'll be this one, not SW_MINIMIZE.
                  .WNP.showCmd = SW_SHOWNORMAL                ' Including the other for paranoia.
            End If
            
            .BrowserWidth = picBrowser.Width
            .ShowFileBrowser = picBrowser.Visible
            .ShowStatusBar = staTusBar1.Visible
            .ShowToolBar = picToolBox.Visible
            .SortMethod = lvwBrowser.SortOrder
            .AutoLoadFile = agEditor.Tag
            .cboPath = cboPath
            .BookmarkCount = mnuBookmark.UBound
            .HistoryCount = mnuFileHistory.UBound
      End With
      
      agEditor.GetSelection lMin, lMax
      
      With mudtCurrentFileSettings
            .FirstVisibleLine = agEditor.FirstVisibleLine
            .SelEnd = lMax
            .SelStart = lMin
            .WordWrap = chkWordWrap.Value
            
            Set fntTemp = agEditor.GetFont
            .FontBold = fntTemp.Bold
            .FontCharset = fntTemp.Charset
            .FontItalic = fntTemp.Italic
            .FontName = fntTemp.name
            .FontSize = fntTemp.Size
            .FontStrikethrough = fntTemp.Strikethrough
            .FontUnderline = fntTemp.Underline
            
            SendMessage agEditor.RichEdithWnd, EM_GETSCROLLPOS, 0, .ScrollPos
      End With
      
            
      ' Create storage key.
      
      lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, msPhlegmKey, 0, "", 0, _
                  KEY_ALL_ACCESS, ByVal 0, lKey, lNewOrUsed)
      If lRetVal <> 0 Then MsgBox "RegCreateKey Failed: " & lKey
      
      ' Store the Window Settings.
      
      lValueSize = LenB(mudtSettings)
      lRetVal = RegSetValueExAny(lKey, "Settings", 0, REG_NONE, _
                  ByVal mudtSettings, lValueSize)
      If lRetVal <> 0 Then MsgBox "RegSetValueEx Failed.  settings: " & _
                  LenB(mudtSettings) & " " & lKey, , App.Title
      
      ' Store the File Settings.
      
      lValueSize = LenB(mudtCurrentFileSettings)
      lRetVal = RegSetValueExAny(lKey, "agEditor", 0, REG_NONE, _
                  ByVal mudtCurrentFileSettings, lValueSize)
      If lRetVal <> 0 Then MsgBox "RegSetValueEx Failed.  mudtCurrentFileSettings: " & _
                  LenB(mudtCurrentFileSettings) & " " & lKey, , App.Title
      
      ' Store Bookmarks.
      
      For iIndex = 1 To mnuBookmark.UBound
            lValueSize = LenB(mnuBookmark(iIndex).Tag)
            lRetVal = RegSetValueExString(lKey, "Bookmark" & CStr(iIndex), 0, REG_SZ, _
                        ByVal mnuBookmark(iIndex).Tag, lValueSize)
      Next iIndex
      
      For iIndex = mnuBookmark.UBound + 1 To mudtSettings.BookmarkCount
            RegDeleteValue lKey, "Bookmark" & CStr(iIndex)
      Next iIndex
      
      ' Store History.
      
      For iIndex = 1 To mnuFileHistory.UBound
            lValueSize = LenB(mnuFileHistory(iIndex).Tag)
            lRetVal = RegSetValueExString(lKey, "History" & CStr(iIndex), 0, REG_SZ, _
                  ByVal mnuFileHistory(iIndex).Tag, lValueSize)
      Next iIndex
      
      For iIndex = mnuFileHistory.UBound + 1 To mudtSettings.HistoryCount
            RegDeleteValue lKey, "History" & CStr(iIndex)
      Next iIndex
      
      ' TODO: Gotta remember to delete bookmarks in the regisy that were
      ' deleted in the program!

      lRetVal = RegCloseKey(lKey)
End Sub

Private Sub GetWindowSettings()
      Dim lRetVal As Long, lKey As Long
      Dim lDataType As Long ' receiving only
      Dim lValueSize As Long ' in/out
      Dim poiFirstLine As POINTAPI
      Dim sTemp As String * 255
      Dim fntTemp As New StdFont
      Dim iBookm As Integer, iHistIndex As Integer
      Dim sEx As String
      
      Dim udtWindowPlacement As WINDOWPLACEMENT
      Dim rectRestored As RECT
      Dim poiTemp As POINTAPI
      
      lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, msPhlegmKey, 0, KEY_QUERY_VALUE, lKey)
      
      lValueSize = LenB(mudtSettings)
      lRetVal = RegQueryValueExAny(lKey, "Settings", 0, lDataType, ByVal mudtSettings, lValueSize)
      If lRetVal = 0 Then
            With mudtSettings
                  mfSkipFormResize = True
                  BrowserResizeHorizontal .BrowserWidth
                  
                  .WNP.Length = LenB(.WNP)
                  SetWindowPlacement hwnd, .WNP
                  
                  lvwBrowser.SortOrder = .SortMethod
                  If agEditor.Tag = "" Then agEditor.Tag = Trim(CstringToVBstring(.AutoLoadFile))
                  
                  For iBookm = 1 To .BookmarkCount ' TODO: THIS BEFORE SETTINGS... SOMEHOW...
                        lValueSize = 255 * 2
                        lRetVal = RegQueryValueExString(lKey, "Bookmark" & CStr(iBookm), 0, lDataType, _
                                    ByVal sTemp, lValueSize)
                        AddToBookmarks Left(sTemp, lValueSize - 1) ' size included the null
                  Next iBookm
                  
                  For iHistIndex = 1 To .HistoryCount
                        lValueSize = 255 * 2
                        lRetVal = RegQueryValueExString(lKey, "History" & CStr(iHistIndex), 0, lDataType, ByVal sTemp, lValueSize)
                        AddToHistory Left(sTemp, lValueSize - 1)
                  Next iHistIndex
      
                  cboPath = Trim(CstringToVBstring(.cboPath))
      
                  chkFileBrowser.Value = -CInt(.ShowFileBrowser)
                  chkFileBrowser_Click
                  'picBrowser.Visible = .ShowFileBrowser
                 ' mnuViewFilebrowser.Checked = .ShowFileBrowser
                  
                  staTusBar1.Visible = .ShowStatusBar
                  mnuViewStatusBar.Checked = .ShowStatusBar
                  picToolBox.Visible = .ShowToolBar
                  mnuViewToolbar.Checked = .ShowToolBar
            
                  mfSkipFormResize = False
                  RearrangeControls
            End With
      Else
            cboPath = ""
      End If
      
'      If Trim(Command()) = "" Then ' no command line argument was given
            lValueSize = LenB(mudtCurrentFileSettings)
            lRetVal = RegQueryValueExAny(lKey, "agEditor", 0, lDataType, ByVal mudtCurrentFileSettings, lValueSize)
            If lRetVal = 0 Then
                  With mudtCurrentFileSettings
                        chkWordWrap.Value = .WordWrap
                        chkWordWrap_Click
                  
                        fntTemp.Bold = .FontBold
                        fntTemp.Charset = .FontCharset
                        fntTemp.Italic = .FontItalic
                        fntTemp.name = Trim(CstringToVBstring(.FontName))
                        fntTemp.Size = .FontSize
                        fntTemp.Strikethrough = .FontStrikethrough
                        fntTemp.Underline = .FontUnderline
                        agEditor.SetFont fntTemp, , , , ercSetFormatAll
                        
                        ' It's important to set the above prior to loading a file.
                        ' Otherwise agEditor's display routines are called again and again for an entire file,
                        ' rather than for a blank editor.
                        
                        sEx = gFSO.getextensionname(agEditor.Tag)
                        EditorLoadFile agEditor.Tag, FileTypeFromExtension(sEx)
                        
                        ' If the file has been changed so that selection and scroll positions are meaningless,
                        ' just skip them...
                        
                        On Error Resume Next
                        agEditor.SetSelection .SelStart, .SelEnd
                        SendMessage agEditor.RichEdithWnd, EM_SETSCROLLPOS, 0, .ScrollPos
                        On Error GoTo 0
                  End With
            End If
'      Else
'
'            EditorLoadFile agEditor.Tag ' If there's a command line argument, just load it plain here
'                                                            ' and let the command line decide the rest.
'      End If
      
      RegCloseKey lKey
End Sub

