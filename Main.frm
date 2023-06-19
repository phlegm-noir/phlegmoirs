VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DD32A320-6E5E-44C8-BCE6-5908CA400231}#1.0#0"; "agRichEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "(New File)"
   ClientHeight    =   8250
   ClientLeft      =   135
   ClientTop       =   675
   ClientWidth     =   11760
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11760
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   6960
      Left            =   2640
      ScaleHeight     =   6960
      ScaleWidth      =   8535
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   840
      Width           =   8535
      Begin TabDlg.SSTab sstProperties 
         Height          =   6375
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   11245
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "File Properties"
         TabPicture(0)   =   "Main.frx":179A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraID3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraProperties"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame fraProperties 
            Caption         =   "File Name:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   240
            TabIndex        =   37
            Top             =   480
            Width           =   6135
            Begin VB.CommandButton btnOpenDefault 
               Caption         =   "&Open"
               Height          =   375
               Left            =   4440
               TabIndex        =   57
               Top             =   2400
               Width           =   1215
            End
            Begin VB.Label lblPropValue 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1440
               TabIndex        =   44
               Top             =   2400
               UseMnemonic     =   0   'False
               Width           =   2895
            End
            Begin VB.Label lblPropValue 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   45
            End
            Begin VB.Label lblPropValue 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1440
               TabIndex        =   45
               Top             =   840
               UseMnemonic     =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPropValue 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1440
               TabIndex        =   46
               Top             =   1200
               UseMnemonic     =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPropValue 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   1440
               TabIndex        =   47
               Top             =   1560
               UseMnemonic     =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPropValue 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   1440
               TabIndex        =   48
               Top             =   1920
               UseMnemonic     =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Size:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   42
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Created:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   240
               TabIndex        =   41
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Modified:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   240
               TabIndex        =   40
               Top             =   1560
               Width           =   1005
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Accessed:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   240
               TabIndex        =   39
               Top             =   1920
               Width           =   1005
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Opens With:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   8
               Left            =   240
               TabIndex        =   38
               Top             =   2400
               Width           =   1125
            End
         End
         Begin VB.Frame fraID3 
            Caption         =   "ID3 tag info"
            Height          =   2415
            Left            =   240
            TabIndex        =   36
            Top             =   3720
            Width           =   6135
            Begin VB.Label lblPropValue 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   9
               Left            =   1440
               TabIndex        =   56
               Top             =   1800
               UseMnemonic     =   0   'False
               Width           =   2655
            End
            Begin VB.Label lblPropValue 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   8
               Left            =   1440
               TabIndex        =   55
               Top             =   1320
               UseMnemonic     =   0   'False
               Width           =   2655
            End
            Begin VB.Label lblPropValue 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   1440
               TabIndex        =   54
               Top             =   840
               UseMnemonic     =   0   'False
               Width           =   2655
            End
            Begin VB.Label lblPropValue 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   1440
               TabIndex        =   53
               Top             =   360
               UseMnemonic     =   0   'False
               Width           =   2655
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Album:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   480
               TabIndex        =   52
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Artist:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   51
               Top             =   840
               Width           =   735
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Year:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   50
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblPropTitle 
               Alignment       =   1  'Right Justify
               Caption         =   "Title:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   49
               Top             =   360
               Width           =   735
            End
         End
      End
      Begin agRichEditBox.agRichEdit agEditor 
         Height          =   5535
         Left            =   5520
         TabIndex        =   34
         Top             =   -240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9763
         Version         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         ViewMode        =   0
         TextLimit       =   9999999
         TrapTab         =   0   'False
         AutoURLDetect   =   0   'False
         TextOnly        =   -1  'True
         ScrollBars      =   0
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   0
         MousePointer    =   15  'Size All
         Top             =   0
         Visible         =   0   'False
         Width           =   3600
      End
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
            Picture         =   "Main.frx":17B6
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C08
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":205A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24AC
            Key             =   "textfile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":28FE
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":362C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4506
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4820
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBrowser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   6960
      Left            =   0
      ScaleHeight     =   6960
      ScaleWidth      =   2415
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   2412
      Begin VB.CommandButton btnScrollToTop 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   2088
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":497A
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
         Left            =   1824
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
         Left            =   1560
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
         Left            =   1296
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
         Left            =   1032
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":4AC4
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
         Picture         =   "Main.frx":4BC6
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Go up a directory (Left arrow key or Ctrl+F6)"
         Top             =   420
         Width           =   504
      End
      Begin VB.CommandButton btnPathForward 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   264
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":4F50
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Go forward a directory (Alt+Right)"
         Top             =   420
         Width           =   264
      End
      Begin VB.ComboBox cboPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Main.frx":509A
         Left            =   0
         List            =   "Main.frx":509C
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
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsFileIcons"
         SmallIcons      =   "ilsFileIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Main.frx":509E
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Object.Tag             =   "0"
            Text            =   "[N]ame"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Key             =   "Type"
            Text            =   "[T]ype"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "Size"
            Text            =   "Si[z]e"
            Object.Width           =   2090
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Modified"
            Text            =   "[M]odified"
            Object.Width           =   3651
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Key             =   "SortSize"
            Object.Tag             =   "Adding 0s (000054) to make #s text-sortable"
            Text            =   "SortSize"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton btnPathBack 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   0
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":53B8
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
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      ClipControls    =   0   'False
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   11700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11760
      Begin VB.PictureBox picQuery 
         ClipControls    =   0   'False
         Height          =   600
         Left            =   4800
         ScaleHeight     =   540
         ScaleWidth      =   4035
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   -25
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtReplace 
            Height          =   288
            Left            =   480
            MaxLength       =   50
            OLEDropMode     =   1  'Manual
            TabIndex        =   65
            ToolTipText     =   "Replace"
            Top             =   290
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtFind 
            Height          =   288
            Left            =   480
            MaxLength       =   50
            OLEDropMode     =   1  'Manual
            TabIndex        =   64
            ToolTipText     =   "Search within file (Ctrl+F)"
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton btnCloseFind 
            Appearance      =   0  'Flat
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   175
            Left            =   3840
            TabIndex        =   63
            TabStop         =   0   'False
            ToolTipText     =   "Hide Toolbar (F7)"
            Top             =   0
            Width           =   175
         End
         Begin VB.CommandButton btnFindNext 
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
            Height          =   300
            Left            =   480
            MaskColor       =   &H80000000&
            Picture         =   "Main.frx":5502
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "Find Next (F3)"
            Top             =   270
            Width           =   1095
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
            Height          =   300
            Left            =   1560
            MaskColor       =   &H80000000&
            Picture         =   "Main.frx":564C
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Find Previous (Shift+F3)"
            Top             =   270
            Width           =   1095
         End
         Begin VB.CheckBox chkFindOptions 
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
            Height          =   285
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   60
            TabStop         =   0   'False
            ToolTipText     =   "More search options (Alt+period)"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton btnReplace 
            Appearance      =   0  'Flat
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            MaskColor       =   &H80000000&
            Picture         =   "Main.frx":5796
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "Replace (Ctrl+R)"
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblFindResult 
            Caption         =   "not found"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   67
            Top             =   120
            Width           =   570
         End
         Begin VB.Label lblFind 
            Alignment       =   2  'Center
            Caption         =   "Find:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   30
            TabIndex        =   66
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.CommandButton btnToolbarClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   175
         Left            =   6480
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Hide Toolbar (F7)"
         Top             =   120
         Visible         =   0   'False
         Width           =   175
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
         Left            =   6240
         Picture         =   "Main.frx":58E0
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnEdit 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Edit This File"
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
         Picture         =   "Main.frx":5D22
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Next file down (Ctrl+""]"")"
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
         Picture         =   "Main.frx":6164
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Next file up (Ctrl+""["")"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton btnZoomIn 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3240
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   320
         Width           =   375
      End
      Begin VB.CommandButton btnZoomDefault 
         Appearance      =   0  'Flat
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2400
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Reset Zoom"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnZoomOut 
         Appearance      =   0  'Flat
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2400
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   320
         Width           =   375
      End
      Begin VB.CommandButton btnFont 
         Caption         =   "Lucida Console"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Set Font (Shift+Ctrl+F)"
         Top             =   0
         Width           =   1215
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Toggle Read-Only mode"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Options..."
         Top             =   0
         Visible         =   0   'False
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
         Left            =   6240
         Picture         =   "Main.frx":65A6
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Slider sliZoom 
         Height          =   372
         Left            =   1800
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Adjust picture zoom"
         Top             =   -24
         Width           =   1836
         _ExtentX        =   3228
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Max             =   500
         SelStart        =   100
         TickFrequency   =   100
         Value           =   100
         TextPosition    =   1
      End
      Begin VB.CommandButton btnFullScreen 
         Appearance      =   0  'Flat
         Caption         =   "Full Screen"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Full Screen (F11)"
         Top             =   0
         Visible         =   0   'False
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
         Picture         =   "Main.frx":69E8
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
         Picture         =   "Main.frx":6E2A
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "New File (Ctrl+N)"
         Top             =   0
         Width           =   615
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         Height          =   570
         Left            =   0
         Picture         =   "Main.frx":7034
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Show/Hide the File Browser (F8)"
         Top             =   0
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label lblFontSize 
         Alignment       =   2  'Center
         Caption         =   "22.3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   960
      End
   End
   Begin MSComctlLib.StatusBar staTusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7950
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "0000 Files, 000000 Bytes Total"
            TextSave        =   "0000 Files, 000000 Bytes Total"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
            Text            =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
            TextSave        =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "statModified"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Key             =   "seltext"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnuPlus 
      Caption         =   "="
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
      Begin VB.Menu mnuBrowserOpenDefault 
         Caption         =   "Open With &Default Program"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "&Rename Open File"
      End
      Begin VB.Menu mnuBrowserSort 
         Caption         =   "Reverse &Sort Order"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileDiv5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNext 
         Caption         =   "Next &File"
      End
      Begin VB.Menu mnuFilePrev 
         Caption         =   "&Previous File"
      End
      Begin VB.Menu mnuFileCurrentDirectory 
         Caption         =   "Sync &Contents"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFileParentDirectory 
         Caption         =   "Parent Directo&ury"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuFileDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowSaveSettings 
         Caption         =   "Save Se&ttings"
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
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditFindBackwards 
         Caption         =   "Find &Previous"
         Shortcut        =   +{F3}
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
      Begin VB.Menu mnuViewFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuViewZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mnuViewZoomOut 
         Caption         =   "Zoom &Out"
      End
      Begin VB.Menu mnuViewDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReadOnly 
         Caption         =   "Read &Only"
      End
      Begin VB.Menu mnuViewWordWrap 
         Caption         =   "&Word Wrap"
         Checked         =   -1  'True
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuViewDiv5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "Show &History"
      End
      Begin VB.Menu mnuViewDiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowserRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmarksAdd 
         Caption         =   "&Add Current File"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBookmarksAddPath 
         Caption         =   "Add Current &Path"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBookmarksManage 
         Caption         =   "&Manage"
         Shortcut        =   ^M
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
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&README.md"
      End
      Begin VB.Menu mnuHelpDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPrev 
      Caption         =   "<<"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuNext 
      Caption         =   ">>"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListCancel 
         Caption         =   "Canc&el"
      End
      Begin VB.Menu mnuListOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuListOpenDefault 
         Caption         =   "Open In Default &Application..."
      End
      Begin VB.Menu mnuListDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListCopyPath 
         Caption         =   "&Copy Full File Name"
      End
      Begin VB.Menu mnuListRename 
         Caption         =   "&Rename"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuListDelete 
         Caption         =   "&Delete File..."
      End
      Begin VB.Menu mnuListDiv2 
         Caption         =   "-"
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
      Begin VB.Menu mnuListHideFileBrowser 
         Caption         =   "&Hide File Browser"
      End
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "Write"
      Visible         =   0   'False
      Begin VB.Menu mnuWriteDelete 
         Caption         =   "Delete"
      End
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
      Begin VB.Menu mnuWriteFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuWriteDiv3 
         Caption         =   "-"
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
      Begin VB.Menu mnuQueryReplace 
         Caption         =   "Show &Replace"
      End
      Begin VB.Menu mnuQueryClose 
         Caption         =   "&Close"
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
' mnuBookmark(x).tag -- exact filename of bookmark.

' TODO: maybe if I can find where I'm testing for (something), then
' I can something.

Option Compare Text
Option Explicit

' *************************************************************
' Scratch or deceased variables.  Comment them out later.
' *************************************************************
Dim long1 As Long, long2 As Long
Dim msTest As String * 512
Dim mbTestByte() As Byte
Dim mbTestByte256(0 To 255) As Byte
Dim msTestArray() As String
'Const mfSkipMouseEventCrap = True
'Dim FuckIHateThis As Boolean
'Dim EditorAccelTable() As ACCEL
'Dim ControlInfoData As CONTROLINFO
'Dim ctrlInfo1 As CTRLINFO


' *************************************************************
' App related variables
' *************************************************************
Const Debugging = False
Const msSettingsVersion = "0.10.0" ' Not the current build number, but the last time I changed the registry structure.

' *************************************************************
' General/Settings variables
' *************************************************************
Dim mStats As TStatType
Dim mWindowPrefs As TWindowPrefs
Dim mEditorPrefs As TEditorPrefs

Dim msPhlegmKey As String
Dim msPhlegmDate As String
Dim mfSkipFormResize As Boolean
Dim mfEditorLoading As Boolean
Const MAX_HISTORY = 10
Const BIGFILESIZE As Long = 2097152   ' Files larger than this won't be opened as text. TODO MAKE THIS DO.
'Const ZoomIncrement = 5

' *************************************************************
' File Browser related variables
' *************************************************************
Dim mfBrowserDoubleClick As Boolean
Dim mfBrowserItemClicked As Boolean  ' here,
Dim mfBrowserButtonPressed As Boolean ' here,
Dim miBrowserMouseButton As Integer  ' here,
Dim miBrowserShift As Integer      ' and here, TODO: make it use gBrowserData instead of these.
Dim mfStartLabelEditFromButton As Boolean

' *************************************************************
' cboPath related variables
' *************************************************************
Dim miPathRecent As Integer
Dim mfValidCboPath As Boolean

' *************************************************************
' Find related variables
' *************************************************************
Dim mfHideFind As Boolean
Dim mfReplaceMode As Boolean
Dim miCurrentQuery As Integer
Dim msLastFindQuery As String
Dim mfFindQueryChanged As Boolean
Dim miFindResult As Integer
Dim mlFirstResultPos As Long
Dim mfQueryMenuOpen As Boolean
Dim mfFinding As Boolean
Dim miTotalResults As Integer





Private Sub AddToBookmarks(ByVal sNewBookmark As String)
      Dim iIndex As Integer

      sNewBookmark = CstringToVBstring(sNewBookmark)
      If sNewBookmark = "" Then Exit Sub
     
      iIndex = mnuBookmark.UBound + 1
      Load mnuBookmark(iIndex)
      With mnuBookmark(iIndex)
            .tag = sNewBookmark  ' exact path here, for safe keeping
            .Caption = iIndex & "   " & sNewBookmark ' here, to make it look all nice
            If iIndex <= 10 Then .Caption = "&" & .Caption
            .Visible = True
      End With

End Sub

Private Sub AddToHistory(ByVal sNewHistory As String)
      Dim iIndex As Integer
      Dim sPrevTag As String, sTempTag As String
      Dim fFoundSame As Boolean, fHistoryGrew As Boolean

      sNewHistory = CstringToVBstring(sNewHistory)
      
      ' Reloading the same file causes no history changes.
            
      If mnuFileHistory.UBound > 0 Then
            If sNewHistory = mnuFileHistory(1).tag Then Exit Sub
      End If
      If sNewHistory = "" Then Exit Sub
     
     ' Add us a new menu item, assuming we're not full yet.
     
      If mnuFileHistory.UBound < MAX_HISTORY Then
            Load mnuFileHistory(mnuFileHistory.UBound + 1)
            mnuFileHistory(mnuFileHistory.UBound).Visible = True
            fHistoryGrew = True
      End If
      
      ' What it SHOULD do:
      '   Put current file at the top.
      '   Start shifting the rest down.
      '   If current file was already in History, remove that one.
      '   Stop shifting.  Don't shift anything below that one.
      sPrevTag = mnuFileHistory(1).tag
      mnuFileHistory(1).tag = sNewHistory
      mnuFileHistory(1).Caption = "&1 " & mnuFileHistory(1).tag
      
      For iIndex = 2 To mnuFileHistory.UBound
            If mnuFileHistory(iIndex).tag = sNewHistory Then
                  mnuFileHistory(iIndex).tag = sPrevTag
                  fFoundSame = True
            Else
                  sTempTag = mnuFileHistory(iIndex).tag
                  mnuFileHistory(iIndex).tag = sPrevTag
                  sPrevTag = sTempTag
            End If
            
            ' Now, we figure out the numbering of the caption, and
            ' which digit to underline.
            If iIndex < 10 Then
                  mnuFileHistory(iIndex).Caption = "&" & iIndex & " " & mnuFileHistory(iIndex).tag
            ElseIf iIndex = 10 Then
                  mnuFileHistory(iIndex).Caption = "1&0 " & mnuFileHistory(iIndex).tag
            Else
                  mnuFileHistory(iIndex).Caption = iIndex & " " & mnuFileHistory(iIndex).tag
            End If
            
            If fFoundSame Then Exit For
      Next iIndex
      
      If fFoundSame And fHistoryGrew Then Unload mnuFileHistory(mnuFileHistory.UBound)
      If gBrowserData.HistoryMode Then btnRefresh_Click
End Sub

'Private Sub AdjustFindArea()
'      ' TODO: this is not even started yet.  Try not to, ah, use it.
'      ' AdjustFindArea is an attempt to do everything that needs to be done,
'      ' to sort out this area.  This includes:
'      '     setting the proper button to default
'      '     enabling or disabling the Replace button
'      '
'
'      If mfReplaceMode And txtFind = agEditor.SelectedText And txtFind <> "" Then
'            btnReplace.Enabled = True
'            If ActiveControl.Name = "txtReplace" Then btnReplace.Default = True
'      End If
'
'      ' This'll sort out the enabling of btnReplace from anywhere, at any time.
'      ' Shouldn't depend on the picFind being visible.
'      If mnuViewReadOnly.Checked And mfReplaceMode Then
'            mnuQueryReplace_Click ' this turns it off
'
'      ElseIf mnuViewReadOnly.Checked And Not mfReplaceMode Then
'            btnReplace.Enabled = False ' it was already off, now it can't be turned on.
'
'      ElseIf Not mnuViewReadOnly.Checked And Not mfReplaceMode Then
'            btnReplace.Enabled = (txtFind <> "") ' gotta have something to search for, at the very least.
'
'      ElseIf Not mnuViewReadOnly.Checked And mfReplaceMode Then
'            ' Here, we enable him if there's text ready to be replaced.
'            btnReplace.Enabled = (txtFind = agEditor.SelectedText)
'      End If
'End Sub

Private Sub BookmarkSaveChanges()
      Dim iIndex As Integer
      
      For iIndex = 1 To lvwBrowser.ListItems.Count
            mnuBookmark(iIndex).tag = lvwBrowser.ListItems(iIndex)
            mnuBookmark(iIndex).Caption = iIndex & "   " & lvwBrowser.ListItems(iIndex)
      Next iIndex
      
      For iIndex = iIndex To mnuBookmark.UBound
            Unload mnuBookmark(iIndex)
      Next iIndex
End Sub


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
                  
      DoEvents ' Just doesn't seem to work without DoEvents first.
      If Not gfFullScreenMode And lvwBrowser.SelectedItem <> Null Then
            lvwBrowser.SelectedItem.EnsureVisible
      End If
End Function

'   BrowserExecuteNext
'   Select the next item after the selection, and open it.
'
Public Sub BrowserExecuteNext()
      DoEvents
      Dim iIndex As Integer, objNextItem As Object

       ' Selecting the item next to the open file, not next to whatever random thing is currently selected.
      If (agEditor.tag <> "") And (Not gBrowserData.BookmarkMode) And (Not gBrowserData.HistoryMode) Then
            btnCurrentDirectory_Click
      End If
        ' TODO: that should still do the sync if in bookmark mode and *the open file is not a bookmark*.
      
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      iIndex = lvwBrowser.SelectedItem.Index
      
      If iIndex < lvwBrowser.ListItems.Count Then
            Set objNextItem = lvwBrowser.ListItems(iIndex + 1)
            If objNextItem.Icon <> eMode.Directory Then
                  If Not gfFullScreenMode Then objNextItem.EnsureVisible
                  objNextItem.Selected = True
                  BrowserExecuteItem objNextItem
            End If
      End If
      Set objNextItem = Nothing
End Sub

'   BrowserExecutePrev
'   Select the item previous to the one selected, and open it.
'
Public Sub BrowserExecutePrev()
      DoEvents
      Dim iIndex As Integer
            
       ' Selecting the item next to the open file, not next to whatever random thing is currently selected.
      If (agEditor.tag <> "") And (Not gBrowserData.BookmarkMode) And (Not gBrowserData.HistoryMode) Then
            btnCurrentDirectory_Click
      End If
        ' TODO: that should still do the sync if in bookmark mode and *the open file is not a bookmark*.
        
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      iIndex = lvwBrowser.SelectedItem.Index
      If iIndex > 1 Then
            If lvwBrowser.ListItems(iIndex - 1).Icon <> eMode.Directory Then
                  If Not gfFullScreenMode Then lvwBrowser.ListItems(iIndex - 1).EnsureVisible
                  lvwBrowser.ListItems(iIndex - 1).Selected = True
                  BrowserExecuteItem lvwBrowser.ListItems(iIndex - 1)
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
            
      Dim iIcon As Integer, iTempKey As Integer
      Dim curTotalBytes As Currency
      Dim litCurrentItem As ListItem
      Dim hNextFile As Long, sFileName As String, sEx As String
      Dim WFD As WIN32_FIND_DATA
      Dim fHadFocus As Boolean ', fDirUnchanged As Boolean
      'Dim sOldSelectedItem As String
      'Dim sngStartTime As Single
      QUOT = Chr(34)
      
      On Error Resume Next    ' there won't be an active control during form_load, so skip this part.
      fHadFocus = (ActiveControl.Name = "lvwBrowser")
      On Error GoTo 0
      
      
      lvwBrowser.tag = BD.Dir
      
      lvwBrowser.Visible = False  ' a nice idea, but we don't want to lose focus while under.  OR DO WE ?
      lvwBrowser.ListItems.Clear
      iTempKey = lvwBrowser.SortKey
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
                  iIcon = eMode.Directory
            Else
                  iIcon = FileTypeFromExtension(sEx)
            End If

            If Err > 0 Then
                  iIcon = eMode.ERROR
                  Debug.Print Err & ": " & Err.Description
            End If
            On Error GoTo 0
            
            
            ' Add that file!
            
            If sFileName <> "." And sFileName <> "" Then ' just what is the point in providing them with a "." folder?
                  Set litCurrentItem = lvwBrowser.ListItems.Add(, , sFileName, iIcon, iIcon)
                  
                  ' here, let's keep an invisible second column for sorting by directory later
                  If iIcon = eMode.Directory Then
                        'litCurrentItem.ListSubItems.Add iFileTypeHeader, "Type", ""
                  Else
                        litCurrentItem.ListSubItems.Add , "Type", sEx  ' Oh deary, column keys are case sensitive!
                        litCurrentItem.ListSubItems.Add , "Size", Format(WFD.nFileSizeLow, "#,#0")
                        litCurrentItem.ListSubItems.Add , "Modified", FormatNonLocalFileTime(WFD.ftLastWriteTime)
                        litCurrentItem.ListSubItems.Add , "SortSize", Format(WFD.nFileSizeLow, "000000000000000")
                        curTotalBytes = curTotalBytes + WFD.nFileSizeLow
                  End If
            End If
      
      Loop While FindNextFile(hNextFile, WFD) <> 0
           
      
      If BD.Filter = "*" Then BD.Filter = ""
      BD.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = iTempKey
      
      ' Leaving Name and Date columns alone, normally.
      'SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 3, LVSCW_AUTOSIZE
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 1, LVSCW_AUTOSIZE '_USEHEADER
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 2, LVSCW_AUTOSIZE
      lvwBrowser.Visible = True
      If fHadFocus Then lvwBrowser.SetFocus
      
      staTusBar1.Panels(eStat.BrowserStats).Text = FormatBytes(curTotalBytes, 1) & " in " & _
            lvwBrowser.ListItems.Count & " objects"
      
      'Debug.Print Timer - sngStartTime
End Sub


Private Sub BrowserGetHistory()
      Dim iIndex As Integer
      Dim litCurrentItem As ListItem
      
      lvwBrowser.ListItems.Clear
      lvwBrowser.tag = "(History)"
      lvwBrowser.Sorted = False
      
      For iIndex = 1 To mnuFileHistory.UBound
            ' TODO: get icon from file extension
            Set litCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuFileHistory(iIndex).tag, _
                  eMode.Bookmark, eMode.Bookmark)
            litCurrentItem.ListSubItems.Add 1, , gFSO.getextensionname(mnuFileHistory(iIndex).tag)
            ' TODO: add file info to subitems
      Next iIndex
      gBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      If Not gBrowserData.ListEmpty Then
            SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 0, LVSCW_AUTOSIZE
            SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 1, LVSCW_AUTOSIZE
            SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 2, LVSCW_AUTOSIZE
            SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 3, LVSCW_AUTOSIZE
      End If
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " most recent files"

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
      'lvwBrowser.ColumnHeaders(1).Width = lvwBrowser.Width - 100
      cboPath.Width = cboPath.Width + iOffset
      
      iScrollButtonX = lvwBrowser.Left + lvwBrowser.Width - btnScrollToTop.Width - 30
      If btnCurrentDirectory.Left + btnCurrentDirectory.Width <= iScrollButtonX Then
            btnScrollToTop.Left = iScrollButtonX
      Else
            btnScrollToTop.Left = btnCurrentDirectory.Left + btnCurrentDirectory.Width
      End If
      
      BrowserResizeHorizontal = iOffset
End Function

Private Sub EditorSetMode(iMode As eMode)

      ' When we change the sort of data to display (text, picture, more to be determined),
      ' there are some things that have to be set, hidden, etc.
      
      ' Use EMode types for iMode
      
      ' Other routines that may be curious about the mode may use gieditormode.
      
      If iMode = giEditorMode Then Exit Sub
      
      Select Case iMode
            Case eMode.Text, eMode.rtf, eMode.other
      
                  giEditorMode = iMode
                  agEditor.Visible = True
                  Image1.Visible = False
                  Image1.Picture = LoadPicture
                  sstProperties.Visible = False
'                  btnZoomIn.Visible = False
'                  btnZoomOut.Visible = False
                  btnFont.Visible = True
                  chkWordWrap.Visible = True
                  btnFullScreen.Visible = False
'                  chkReadOnly.Visible = True
                  If Not mfHideFind And mnuViewToolbar.Checked Then picQuery.Visible = True
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3240, 320, 375, 252
                  btnZoomIn.Caption = "+"
                  btnZoomOut.Move 2400, 320, 375, 252
                  btnZoomOut.Caption = "-"
                  btnZoomDefault.Visible = False
                  
                  staTusBar1.Panels(eStat.encoding).Visible = True
                  staTusBar1.Panels(eStat.Modified).Visible = True
                  staTusBar1.Panels(eStat.Stats).Visible = True
                  staTusBar1.Panels(eStat.SelText).Visible = True
      
            Case eMode.Picture
                  
                  giEditorMode = eMode.Picture
                  agEditor.Visible = False
                  agEditor.Text = ""
                  Image1.Visible = True
                  sstProperties.Visible = False
'                  btnZoomIn.Visible = True
'                  btnZoomOut.Visible = True
                  btnFont.Visible = False
                  chkWordWrap.Visible = False
                  btnFullScreen.Visible = True
 '                 chkReadOnly.Visible = False
                  If Not mfHideFind Then picQuery.Visible = False
                  
                  sliZoom.Visible = True
                  btnZoomIn.Move 3000, 360, 615, 252
                  btnZoomIn.Caption = "z+"
                  btnZoomOut.Move 1800, 360, 615, 252
                  btnZoomOut.Caption = "z-"
                  btnZoomDefault.Visible = True
                  
                  If gpOldpicEditorProc = 0 Then
                        gpOldpicEditorProc = SetWindowLong(picEditor.hwnd, GWL_WNDPROC, _
                              AddressOf TrackMouseWheel)
                  End If
                  
                  staTusBar1.Panels(eStat.encoding).Visible = False
                  staTusBar1.Panels(eStat.Modified).Visible = False
                  staTusBar1.Panels(eStat.Stats).Visible = False
                  staTusBar1.Panels(eStat.SelText).Visible = False
                  
            Case eMode.Properties
                  giEditorMode = Properties
                  agEditor.Visible = False
                  agEditor.Text = ""
                  Image1.Visible = False
                  Image1.Picture = LoadPicture
                  sstProperties.Visible = True
                  
                  btnFont.Visible = True
                  chkWordWrap.Visible = True
                  btnFullScreen.Visible = False
'                  chkReadOnly.Visible = True
                  If Not mfHideFind Then picQuery.Visible = False
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3240, 320, 375, 252
                  btnZoomIn.Caption = "+"
                  btnZoomOut.Move 2400, 320, 375, 252
                  btnZoomOut.Caption = "-"
                  btnZoomDefault.Visible = False
                  
                  staTusBar1.Panels(eStat.encoding).Visible = False
                  staTusBar1.Panels(eStat.Modified).Visible = False
                  staTusBar1.Panels(eStat.Stats).Visible = False
                  staTusBar1.Panels(eStat.SelText).Visible = False
      End Select
      RearrangeControls

End Sub

Private Function FileTypeFromExtension(sEx As String) As String
      ' This function takes an extension (DO NOT INCLUDE DOT) and returns a mode
      ' which can be fed into EditorSetMode.
      
      ' Current possible modes:   EMode.text, EMode.rtf, EMode.picture
      
      Select Case sEx
            Case "bmp", "gif", "jpg", "jpeg", "png", "ico", "cur"
                  FileTypeFromExtension = eMode.Picture
            Case "rtf"
                  FileTypeFromExtension = eMode.rtf
            Case "txt"
                  FileTypeFromExtension = eMode.Text
            Case Else
                  FileTypeFromExtension = eMode.other
      End Select
End Function

Private Sub GetFileProperties(ByVal sFileName As String)
      Dim WFD As WIN32_FIND_DATA
      Dim hFile As Long
      Dim sEx As String
      
      hFile = FindFirstFile(sFileName, WFD)
      fraProperties.Caption = WFD.cFileName
      lblPropValue(2) = Format(WFD.nFileSizeLow, "#,#0")
      lblPropValue(4) = FormatNonLocalFileTime(WFD.ftLastWriteTime)
      lblPropValue(3) = FormatNonLocalFileTime(WFD.ftCreationTime)
      lblPropValue(5) = FormatNonLocalFileTime(WFD.ftLastAccessTime)
      FindClose hFile

      sEx = gFSO.getextensionname(sFileName)
'      If sEx = "mp3" Then
      Dim mp3info As MP3TagInfo
      
      GetMP3Info sFileName, mp3info
      With mp3info
            lblPropValue(6) = mp3info.title
            lblPropValue(7) = mp3info.artist
            lblPropValue(8) = mp3info.album
            lblPropValue(9) = mp3info.year
      
      End With
 '     End If
End Sub

Public Sub ImageZoomIn(iStep As Integer)
      ' goes up to the next zoom divisible by iStep
      If sliZoom.Value >= sliZoom.Max Then Exit Sub
      sliZoom.Value = sliZoom.Value + (iStep - (sliZoom.Value Mod iStep))
End Sub

Public Sub ImageZoomOut(iStep As Integer)
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
      ' This will be called when a listitem is clicked, and it will enable or disable parts
      ' of the right click menu, based on the sort of listitem is passed to it.
      
      'If mnuListOpenDefault.Enabled Then Exit Sub  ' don't wanna bother doing all this on every single mousemove!
      
      mnuListOpenDefault.Enabled = True
      mnuListOpen.Enabled = True
      mnuListOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
      mnuListCopyPath.Enabled = True
      mnuListProperties.Enabled = True
      
      If gBrowserData.BookmarkMode Then
            mnuListShowOnly.Enabled = False
            mnuListDelete.Caption = "&Delete Bookmark" & vbTab & "Del"
      
      ElseIf gBrowserData.HistoryMode Then
            mnuListShowOnly.Enabled = False
            mnuListRename.Enabled = False
            mnuListDelete.Enabled = False
            mnuListDelete.Caption = "&Delete File..." & vbTab & "Del"
      Else
            mnuListDelete.Enabled = True
            mnuListRename.Enabled = True
            mnuListShowOnly.Enabled = True
            mnuListDelete.Caption = "&Delete File..." & vbTab & "Del"
      End If
      
      If litHoverItem.Icon = eMode.Directory Or litHoverItem.Icon = eMode.Drive Then
            mnuListOpenDefault.Caption = "Explore..." & vbTab & "Shift+Ctrl+Enter"
            mnuListDelete = False
            If litHoverItem.Text = ".." Or litHoverItem.Icon = eMode.Drive Then mnuListRename = False
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
'   ParsePath translates input string sInput into referenced data structure BD.
'   BD holds the working directory, filter, previous directory, mode,
'   ...and much, much more!
'
Private Sub ParsePath(ByVal sInput As String, ByRef BD As TBrowserData)
      ' (Bookmarks)      (that means bookmark mode, of course!)
      ' (History)           (History mode -- TODO, NOT IN HERE YET)
      '                            (a blank is intrepreted as "root" / drives list mode)
      ' c:\temp\   (just a plain old directory)
      ' c:\temp\.txt  (wildcard implied)
      ' c:\temp\peni*   (contains wildcard(s) after the directory)
      ' c:\temp\peni   (some trailing characters, but no wildcard)
      
      Dim sFileName As String
      
      sInput = Trim(sInput)
      
      With BD
      
            .BookmarkMode = False
            .DrivesMode = False
            .HistoryMode = False
            .ListEmpty = (lvwBrowser.ListItems.Count = 0)
            .DirPrev = .Dir
            .FilterPrev = .Filter
            
            
            If sInput = "(Bookmarks)" Then  ' We are in Manage Bookmarks mode.
                  .BookmarkMode = True
                  .Dir = "(Bookmarks)"  ' Just so that (.Dir = X) never accidentally returns true.
                  .Filter = ""
                  .PartialFileName = ""
                  .ValidPath = False
            
            ElseIf sInput = "(History)" Then
                  .HistoryMode = True
                  .Dir = "(History)"
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
      lvwBrowser.tag = "(Bookmarks)"
      ' I'm adding the index as a Key, to avoid using indeces.
      ' (Dumb workaround so that I can use API functions that desynchronize listitem indexing.)
      ' Edit: I'm not really doing that.  Still using bookmarks as a test case on whether that might
      ' be accomplished one day.
      For iIndex = 1 To mnuBookmark.UBound
            Set litCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuBookmark(iIndex).tag, _
                  eMode.Bookmark, eMode.Bookmark)
            litCurrentItem.ListSubItems.Add 1, , gFSO.getextensionname(mnuBookmark(iIndex).tag)
            ' I'm thinking bookmarks don't need subitems?
            ' TODO: Sure they do... just list them like they were files!  Later, perhaps.
      Next iIndex
      gBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 0, LVSCW_AUTOSIZE
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 1, LVSCW_AUTOSIZE
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 2, LVSCW_AUTOSIZE
      SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal 3, LVSCW_AUTOSIZE
      
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " bookmarks"
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

Public Sub WheelInput(iWheelTurn As Integer, iVirtKeys As Integer, lx As Long, ly As Long)
      ' This is called from modPhlegmoirs.TrackMouseWheel
      ' It acts on picEditor while in picture mode.
      
      Dim iWheelMoveIncrement As Integer
      ' iWheelMoveIncrement will be the positive distance that the wheel moves a picture.
      iWheelMoveIncrement = -MoveIncrement * 3 * Abs(iWheelTurn) * sliZoom.Value / 100
      
      With gImageData.OutPic
            ' Wheel scroll up = move picture down = make Top value HIGHER
            ' ...but not to rise above zero.
            If iVirtKeys = 0 And iWheelTurn > 0 Then
                  If .Top < -iWheelMoveIncrement Then
                        .Top = .Top + iWheelMoveIncrement
                  ElseIf .Top < 0 Then
                        .Top = 0
                  End If
            
            ' Wheel scroll down = move picture up = make Top value LOWER.
            ' ...the bottom value not to fall below the bottom value of its container.
            ElseIf iVirtKeys = 0 And iWheelTurn < 0 Then
                  If .Top + .Height > .Container.Height + iWheelMoveIncrement Then
                        .Top = .Top - iWheelMoveIncrement
                  ElseIf .Top + .Height > .Container.Height Then
                        .Top = .Container.Height - .Height
                  End If
                  
            ElseIf iVirtKeys = MK_MBUTTON Then
                  ' Hold down wheel while turning = move picture right/left
                  .Left = .Left - iWheelTurn * MoveIncrement * 3
                  
            ElseIf iVirtKeys = MK_RBUTTON Then
                  ' Right mouse button + wheel scroll = Picture Zoom
                  Dim iPresses As Integer
                  ' So we'll be lazy and just press the appropriate zoom button once for each mouse turn.
                  For iPresses = 1 To Abs(iWheelTurn)
                        If iWheelTurn > 0 Then
                              btnZoomIn_Click
                                          
                        ElseIf iWheelTurn < 0 Then
                              btnZoomOut_Click
                        End If
                  Next iPresses
                  gImageData.Zoomed = True
                  If gfFullScreenMode Then frmFullScreen.lblFileNameZoom = Caption & "  "
                  
            ElseIf iVirtKeys = MK_LBUTTON Then
                  ' Left mouse button + wheel scroll = Picture zoom, small increment
                  sliZoom.Value = sliZoom.Value - iWheelTurn * sliZoom.SmallChange
                  gImageData.Zoomed = True
                  If gfFullScreenMode Then frmFullScreen.lblFileNameZoom = Caption & "  "
                  
            End If
      End With
End Sub


Private Sub agEditor_ProgressStatus(ByVal lAmount As Long, ByVal lTotal As Long)
'      Debug.Print "PROGRESS: "; lAmount & " " & lTotal

      ' TODO: if a second file is told to load, it cancels this one but won't remove it from the editor first.
      
      DoEvents
End Sub

Private Sub btnCloseFind_Click()
      mnuQueryClose_Click
End Sub

Private Sub btnCurrentDirectory_Click()
      mnuFileCurrentDirectory_Click
End Sub

Private Sub btnDeleteSelected_Click()
      BrowserDeleteSelected
End Sub

Private Sub btnEdit_Click()
      mnuViewReadOnly_Click
End Sub

Private Sub btnFindPrev_Click()
      ' So I've decided that it's possible to have negative numbers of find results.
      ' This is what happens when you click "Find Previous",
      ' and there wasn't a previous find, but there is a match.
      ' We can't just call it #N, where N is the total number of matches in the document,
      ' because we haven't searched the entire document!  That would take too long.
      ' So instead, it's #-1.
      
      ' No searching text within a picture or a properties tab.
      If giEditorMode = eMode.Picture Or giEditorMode = Properties Then Exit Sub
      
      If txtFind = "" Then txtFind = agEditor.SelectedText
      
      Dim lFoundMin As Long, lFoundMax As Long, lStartMin As Long, lStartMax As Long
      Dim lFindRetval As Long, fFindInSelection As Boolean
      Dim iFindOptions As Integer
      
      agEditor.GetSelection lStartMin, lStartMax
      
      lFindRetval = EditorFindText(txtFind, back, lStartMin, 0, lFoundMin, lFoundMax)
      
      If lFindRetval = -1 Then
            ' Nothing found upward.  Search from end of file.
            lFindRetval = EditorFindText(txtFind, back, agEditor.CharacterCount, _
                  lStartMin, lFoundMin, lFoundMax)
      End If
            
      If lFindRetval > -1 Then
            ' Found something!
            mfFinding = True ' make sure the find count doesn't reset when we highlight a find result!
            agEditor.SetSelection lFoundMin, lFoundMax
            mfFinding = False
            
            If mlFirstResultPos = lFoundMin And miFindResult < 0 Then
                  ' -8, -9, -10 => 10
                  ' When we reach the starting point again WHILE going in reverse, we now know
                  ' how many results exist.  So rather than wrap from -N up to -1 again,
                  ' we'll call the next one up from -N, simply N.
                  ' No more need for negative search results unless the count is reset.
                  miTotalResults = Abs(miFindResult)
                  miFindResult = miTotalResults
            ElseIf miFindResult = 1 And miTotalResults = 0 Then
                  ' 3, 2, 1 => -1
                  ' when counting backwards, there shall be no zeroth result
                  miFindResult = -1
            ElseIf miFindResult = 1 And miTotalResults > 0 Then
                  ' 3, 2, 1 => 10
                  ' when the total is known, we don't use negatives.
                  miFindResult = miTotalResults
            Else
                  ' -7, -6, -5 => -4
                  ' This is the typical case.
                  miFindResult = miFindResult - 1
            End If
            If miFindResult = -1 Then mlFirstResultPos = lFoundMin
                  
            lblFindResult.ForeColor = vbButtonText
            lblFindResult = "# " & miFindResult
            'staTusBar1.Panels(EStat.Tips) = "Search results: " & miFindResult & " found"
      Else
            agEditor.SetSelection lStartMin, lStartMin
            miFindResult = 0
            lblFindResult.ForeColor = vbRed
            lblFindResult = "not found"
      End If

      'txtFind.SetFocus
End Sub

Private Sub btnFolderUp_Click()
      mnuFileParentDirectory_Click
End Sub


Private Sub btnFont_Click()
      Dim fntTemp As New StdFont ' StdFont is a Class
      Dim lRetVal As Long, lTextColor As Long
      Const cdlCFScreenFonts As Long = &H1
      Const cdlCFScalableOnly As Long = &H20000
      Const cdlCFEffects As Long = &H100
      
      Set fntTemp = GetRealStdFont(agEditor, lTextColor)
      
      'make the dialog choices begin with what the agEditor shows
      With dlgFont
            .flags = cdlCFScreenFonts + cdlCFApply + cdlCFEffects ' btw, Apply doesn't work
            .FontName = fntTemp.Name
            .FontBold = fntTemp.Bold
            .FontUnderline = fntTemp.Underline
            .FontSize = fntTemp.Size  ' one uses Single, the other Currency
            .FontStrikethru = fntTemp.Strikethrough
            .Color = lTextColor
      End With

      On Error Resume Next 'trap the error. if they hit cancel, do nothing and exit
      dlgFont.ShowFont
      If Err.Number = cdlCancel Then Exit Sub
      On Error GoTo 0  'btw, I think this has the effect of err.Clear
      
      With fntTemp
            .Name = dlgFont.FontName
            ' IMPORTANT TO REMEMBER:
            ' when you set a StdFont's NAME property, you've ALSO set its CHARSET property.
            ' AUTOMATICALLY.  Same for weight.
            ' The problem (one of many) with this agRichEdit control is that even though my
            ' commondialog set the font name, and therefore the charset, of the StdFont object,
            ' the editor's stupid SetFont method DOES NOT PASS THE CHARSET on to the
            ' rich edit control.  It probably uses a CHARFORMAT2, and it must have slipped their
            ' minds to give its dwMask property the CFM_CHARSET flag!  So that even if they
            ' DID set the bCharset property to the stdfont.charset value, it would not have
            ' even SEEN it!
            
            ' And it assumes charset = 0, which it is for most fonts.
            ' And that's why the agEditor won't work with symbol fonts, which have charset = 2.
            ' UNTIL NOW.  BECAUSE I WENT AROUND THE STUPID SETFONT METHOD!
            
            .Bold = dlgFont.FontBold
            .Italic = dlgFont.FontItalic
            .Underline = dlgFont.FontUnderline
            .Strikethrough = dlgFont.FontStrikethru
            .Size = CCur(dlgFont.FontSize)
            ' Weight is set automatically.  (It seems that) 400 = plain, 700 = bold.
      End With
      'agEditor.SetFont fntTemp, , , , ercSetFormatAll
      lRetVal = SetRealStdFont(agEditor, fntTemp, dlgFont.Color)
      btnFont.Caption = GetRealStdFont(agEditor).Name
      If Len(btnFont.Caption) > 11 Then
            btnFont.Caption = Left(Trim(btnFont.Caption), 10) & "..."
      End If
      lblFontSize = Round(GetRealStdFont(agEditor).Size, 1)
End Sub


Private Sub btnFullScreen_Click()
      Hide
      frmFullScreen.Show
End Sub

Private Sub btnNewFile_Click()
      mnuFileNew_Click
End Sub

Private Sub btnNextFile_Click()
      BrowserExecuteNext
End Sub

Private Sub btnOptions_Click()
'      frmOptions.Show
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

Private Sub btnReplace_Click()
      ' The replace button puts us in replace mode if we aren't already there.
      If Not mfReplaceMode Then
            mnuQueryReplace_Click

      ' Otherwise, if we were already in replace mode, it replaces (if legal).
      ElseIf btnReplace.Enabled Then
            agEditor.InsertContents SF_TEXT, txtReplace
            btnFindNext_Click
      End If
End Sub

Private Sub btnToolbarClose_Click()
      mnuViewToolbar_Click
End Sub

Private Sub btnZoomDefault_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      sliZoom.Value = 100
      Image1.Move 0, 0, gImageData.DefaultWidth, gImageData.DefaultHeight
End Sub

 
Private Sub btnZoomIn_Click()
      Select Case giEditorMode
      
            Case eMode.Picture
                  ' Go to the next zoom divisible by the zoom increment
                  If sliZoom.Value < 100 Then
                        ImageZoomIn 25
                  Else
                        ImageZoomIn sliZoom.LargeChange
                  End If
                  
            Case Else  ' Increase the Font Size
                  Dim iFontSize As Integer
                  
                  iFontSize = CInt(GetRealFontSize(agEditor))
                  
                  If iFontSize < 12 Then
                        iFontSize = iFontSize + 1
                  ElseIf iFontSize < 28 Then
                        iFontSize = iFontSize + (2 - iFontSize Mod 2)   ' rounding to the previous even number
                  ElseIf iFontSize < 36 Then
                        iFontSize = 36
                  ElseIf iFontSize < 48 Then
                        iFontSize = 48
                  ElseIf iFontSize < 72 Then
                        iFontSize = 72
                  Else
                        iFontSize = iFontSize + (10 - iFontSize Mod 10)
                  End If
                  
                  SetRealFontSize agEditor, iFontSize
                  lblFontSize = iFontSize
      End Select
End Sub

Private Sub btnZoomOut_Click()
      Select Case giEditorMode
            
            Case eMode.Picture
                  ' Go to the next lowest zoom % divisible by the zoom increment
                  If sliZoom.Value <= 100 Then
                        ImageZoomOut 25
                  Else
                        ImageZoomOut sliZoom.LargeChange
                  End If
            
            Case Else   ' Decrease the Font Size
                  Dim iFontSize As Integer
                  
                  iFontSize = CInt(GetRealFontSize(agEditor))
                  
                  If iFontSize <= 1 Then
                        iFontSize = 1
                  ElseIf iFontSize <= 13 Then
                        iFontSize = iFontSize - 1
                  ElseIf iFontSize <= 28 Then
                        iFontSize = iFontSize - (2 - iFontSize Mod 2)  ' rounding to the previous even number
                  ElseIf iFontSize <= 36 Then
                        iFontSize = 28
                  ElseIf iFontSize <= 48 Then
                        iFontSize = 36
                  ElseIf iFontSize <= 72 Then
                        iFontSize = 48
                  ElseIf iFontSize <= 80 Then
                        iFontSize = 72
                  Else
                        iFontSize = iFontSize - (10 - iFontSize Mod 10)
                  End If
                  
                  SetRealFontSize agEditor, iFontSize
                  lblFontSize = iFontSize
      End Select
End Sub

Private Sub chkFindOptions_Click()
      Debug.Print "chkQuery..._Click"
      
      If chkFindOptions.Value = vbChecked Then
            PopupMenu mnuQuery, vbPopupMenuRightAlign, AbsoluteRight(chkFindOptions), _
                  AbsoluteBottom(chkFindOptions)
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
            ElseIf .HistoryMode Then
                  BrowserGetHistory
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
      Dim iTempKey As Integer
      
      ' List remains sorted at all times.  Only the order can be reversed.
      
      If gBrowserData.HistoryMode Then Exit Sub
      
      With lvwBrowser
            .Sorted = True
            iTempKey = .SortKey
            .SortKey = 0
            .SortOrder = Abs(.SortOrder - 1)
            .SortKey = iTempKey
      End With
            
      If gBrowserData.BookmarkMode Then BookmarkSaveChanges
End Sub


Private Sub btnScrolltotop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnScrollToTop.ToolTipText
End Sub

Private Sub btnSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnSort.ToolTipText
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnSave.ToolTipText
End Sub

Private Sub btnrefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnRefresh.ToolTipText
End Sub

Private Sub chkFindOptions_LostFocus()
      chkFindOptions.Value = vbUnchecked
End Sub

'Private Sub chkFindOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      Debug.Print "chkQuery..._Mousedown"
'      'PopupMenu mnuQuery, vbPopupMenuRightAlign
'      If mfQueryMenuOpen = False Then
'            PopupMenu mnuQuery, vbPopupMenuRightAlign, AbsoluteRight(chkFindOptions), _
'                  AbsoluteBottom(chkFindOptions)
'      End If
'
'End Sub



Private Sub chkFindOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = chkFindOptions.ToolTipText
End Sub

Private Sub btnprevfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPrevFile.ToolTipText
End Sub

Private Sub btnpathforward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPathForward.ToolTipText
End Sub

Private Sub btnpathback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPathBack.ToolTipText
End Sub

Private Sub btnnextfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnNextFile.ToolTipText
End Sub

Private Sub btnnewfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnNewFile.ToolTipText
End Sub

Private Sub btnfont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFont.ToolTipText
End Sub

Private Sub btnfolderup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFolderUp.ToolTipText
End Sub

Private Sub btnfindprev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFindPrev.ToolTipText
End Sub

Private Sub btnFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFindNext.ToolTipText
End Sub

Private Sub btnfileforward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFileForward.ToolTipText
End Sub

Private Sub btnfileback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFileBack.ToolTipText
End Sub

Private Sub btndeleteselected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnDeleteSelected.ToolTipText
End Sub

Private Sub btncurrentdirectory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnCurrentDirectory.ToolTipText
End Sub

Private Sub ageditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' TODO: this needs to not happen, if frmMain is not the front window
      
'      Debug.Print "agEditor: " & agEditor.CharFromPos(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
'
'      Dim poiTemp As POINTAPI, lRetVal As Long
'      poiTemp.x = x / Screen.TwipsPerPixelX
'      poiTemp.y = y / Screen.TwipsPerPixelY
'      lRetVal = SendMessage(agEditor.RichEdithWnd, EM_CHARFROMPOS, ByVal 0, poiTemp)
'      Debug.Print "API: " & lRetVal & " " & Timer
      
      'Debug.Print Screen.ActiveForm.name & " "; Screen.ActiveControl.name & " " & Timer
      'Debug.Print GetForegroundWindow & "   " & frmMain.hwnd & " " & agEditor.RichEdithWnd
      
      On Error Resume Next
      If GetForegroundWindow = frmMain.hwnd And Not (ActiveControl.Name = "agEditor") And _
            Not ActiveControl.Name = "txtFind" And Not ActiveControl.Name = "txtReplace" Then
            agEditor.SetFocus
      End If
      On Error GoTo 0
      
      ' Here, I'm throwing in a feature where a tooltip comes up with your character code...
      '   * If there's ONLY ONE character highlighted, and
      '   * If the mouse is hovering over that one character.
      
      If staTusBar1.Panels(eStat.SelText) = "1" Then
            Dim lMin As Long, lMax As Long
            agEditor.GetSelection lMin, lMax

            If agEditor.CharFromPos(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) = lMax Then
                  agEditor.ToolTipText = "Char: " & Asc(agEditor.SelectedText)
            Else
                  agEditor.ToolTipText = ""
            End If
      End If
      
      staTusBar1.Panels(eStat.Tips).Text = ""
End Sub

Private Sub chkFindOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Debug.Print "chkquery..._mouseup"
End Sub

Private Sub chkReadOnly_Click()
      
      mnuViewReadOnly.Checked = chkReadOnly.Value
      agEditor.ReadOnly = chkReadOnly.Value
      If chkReadOnly.Value = vbChecked Then
            agEditor.BackColor = &H8000000F
            btnEdit.Visible = True
            If mfReplaceMode Then mnuQueryReplace_Click
      Else
            btnEdit.Visible = False
            agEditor.BackColor = &H80000005
      End If
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = ""
End Sub


Private Sub cboPath_Change()
      
      ParsePath cboPath, gBrowserData
      
      If gBrowserData.BookmarkMode Then
            BrowserGetBookmarks
            PathAddRecent "(Bookmarks)"
      
      ElseIf gBrowserData.HistoryMode Then
            BrowserGetHistory
            PathAddRecent "(History)"
      
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
      staTusBar1.Panels(eStat.BrowserStats).Visible = chkFileBrowser.Value
      
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
      If cboPath <> "(Bookmarks)" And cboPath <> "(History)" Then
            
            Dim iExtensionLength As Integer
            
            iExtensionLength = Len(gFSO.getextensionname(cboPath))
            If iExtensionLength > 0 Then iExtensionLength = iExtensionLength + 1 ' include the dot
            cboPath.SelStart = Len(cboPath) - iExtensionLength
      End If
End Sub

Private Sub cboPath_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim iSlash As Integer
'      Debug.Print "cbopath.selstart" & cboPath.SelStart
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

Private Sub chkFileBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = chkFileBrowser.ToolTipText
End Sub

Private Sub chkWordWrap_Click()
      
      Dim lineindex As Long, charindex As Long, lMin As Long, lMax As Long
      
      mnuViewWordWrap.Checked = chkWordWrap.Value
      agEditor.ViewMode = chkWordWrap.Value
      
      ' a few things in the statusbar could change in a word wrap:
      '   x, xmax, y, ymax
      ' and some shouldn't change:
      '   i, imax,   (we're not adding or deleting characters or moving the cursor)
      '   sellength
      
      If agEditor.CharacterCount = 0 Then Exit Sub
      agEditor.GetSelection lMin, lMax
      lineindex = agEditor.CurrentLine
      charindex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lineindex, 0)
      
      If staTusBar1.Visible Then
            With mStats
                .X = lMin - charindex + 1
                .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal charindex, 0) + 1
                .Y = lineindex + 1
                .ymax = agEditor.LineCount
            End With
            FillStats
      End If
End Sub

Private Sub btnFindNext_Click()
      If giEditorMode = eMode.Picture Or giEditorMode = Properties Then Exit Sub
      
      If txtFind = "" Then txtFind = agEditor.SelectedText
      
      Dim lFoundMin As Long, lFoundMax As Long, lStartMin As Long, lStartMax As Long
      Dim lFindRetval As Long, fFindInSelection As Boolean
      Dim iFindOptions As Integer
      
      agEditor.GetSelection lStartMin, lStartMax
      
'      If txtFind <> msLastFindQuery Then
'            ' Reset result count when the query changes.
'            miFindResult = 0
'            lblFindResult = ""
'            msLastFindQuery = txtFind
'      End If
'
'      If miFindResult = 0 Then
'            'mlFirstResultPos = lStartMax
'      End If
            
      
      lFindRetval = EditorFindText(txtFind, Forward, lStartMax, _
            agEditor.CharacterCount, lFoundMin, lFoundMax)
      
      If lFindRetval = -1 Then
            ' Nothing found downward.  Search from beginning.
            lFindRetval = EditorFindText(txtFind, Forward, 0, _
                  lStartMax, lFoundMin, lFoundMax)
      End If
            
      If lFindRetval > -1 Then
            ' Found something!
            mfFinding = True ' make sure the find count doesn't reset when we highlight a find result!
            agEditor.SetSelection lFoundMin, lFoundMax
            mfFinding = False
            
            If miFindResult = miTotalResults And miTotalResults > 0 Then
                  miFindResult = 1
            ElseIf mlFirstResultPos = lFoundMin And miFindResult > 0 Then
                  ' Reset find count when we reach the starting point again, going forward.
                  miTotalResults = miFindResult
                  miFindResult = 1
            ElseIf miFindResult = -1 Then
                  ' When counting up, after a backwards search which resulted in negative numbers,
                  ' there shall be no zeroth match.  Skip to #1.
                  miFindResult = 1
            Else
                  miFindResult = miFindResult + 1
            End If
            If miFindResult = 1 Then mlFirstResultPos = lFoundMin
                  
            lblFindResult.ForeColor = vbButtonText
            lblFindResult = "# " & miFindResult
            'staTusBar1.Panels(EStat.Tips) = "Search results: " & miFindResult & " found"
      Else
            agEditor.SetSelection lStartMax, lStartMax
            miFindResult = 0
            lblFindResult.ForeColor = vbRed
            lblFindResult = "not found"
      End If

      'txtFind.SetFocus
End Sub

' EditorFindText
'   Finds the search string sFindMe in agEditor between values of lRangeStart and lRangeEnd.
'   This function DOES NOT HIGHLIGHT ANYTHING OR MOVE THE CURSOR.
'
'  lFoundMin and lFoundMax receive the start and end positions of the found string.
'  Returns -1 if nothing found, returns lFoundMin if successful.
'
'  The way EM_FINDTEXTEX works is that it goes from lRangeStart to lRangeEnd in the
'  specified direction.  That means the start position has to come first.  NOT the lower of the values first.

Private Function EditorFindText( _
            ByVal sFindme As String, _
            ByVal iDirection As eDirection, _
            ByVal lRangeStart As Long, _
            ByVal lRangeEnd As Long, _
            ByRef lFoundMin As Long, _
            ByRef lFoundMax As Long) As Long
      
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
            SetWindowLong lvwBrowser.hwnd, GWL_WNDPROC, gpOldLvwProc
            gpOldLvwProc = 0
      End If
      
      If gpOldpicEditorProc <> 0 Then
            SetWindowLong picEditor.hwnd, GWL_WNDPROC, gpOldpicEditorProc
            gpOldpicEditorProc = 0
      End If
      SaveWindowSettings
End Sub


Private Sub Image1_DblClick()
      ' This needs to (effectively) call an Image1_mousedown... but with what parameters???
      Dim poiPrev As POINTAPI
      Dim recPicBox As RECT
      
      GetCursorPos poiPrev
      GetWindowRect picEditor.hwnd, recPicBox
      
      gImageData.PrevX = (poiPrev.X - recPicBox.Left) * Screen.TwipsPerPixelX - Image1.Left
      gImageData.PrevY = (poiPrev.Y - recPicBox.Top) * Screen.TwipsPerPixelY - Image1.Top
      gImageData.Dragging = True
      picEditor.SetFocus
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      gImageData.PrevX = X
      gImageData.PrevY = Y
      If Button = vbLeftButton Then
            gImageData.Dragging = True
            picEditor.SetFocus
      End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error Resume Next
      If GetForegroundWindow = frmMain.hwnd And Not (ActiveControl.Name = "picEditor") Then
            picEditor.SetFocus
      End If
      On Error GoTo 0
            
      If gImageData.Dragging Then
            Image1.Move Image1.Left + X - gImageData.PrevX, Image1.Top + Y - gImageData.PrevY, Image1.Width, Image1.Height
            If X <> gImageData.PrevX Or Y <> gImageData.PrevY Then gImageData.Moved = True
      End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' Mouse button lifted?  Stop the drag!
      gImageData.Dragging = False
      
      If Not gImageData.Moved And Not gImageData.Zoomed And Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            btnNextFile_Click
      ElseIf Not gImageData.Moved And Not gImageData.Zoomed And Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            btnPrevFile_Click
      End If
      
      gImageData.Moved = False
      gImageData.Zoomed = False
End Sub


Private Sub lblFind_DblClick()
'      If Debugging Then
'            Dim penis(4) As Collection
'            penis(1) = btnFindNext
'            penis(2) = btnFindPrev
'            penis(3) = chkFindOptions
'            penis(4) = btnReplace
'
'            For p = 1 To 4
'                  Debug.Print
'            Next p
'      End If
      Debug.Print "btnfindnext.move " & btnFindNext.Left & "," & _
            btnFindNext.Top & "," & btnFindNext.Width & "," & btnFindNext.Height
      Debug.Print "btnfindprev.move " & btnFindPrev.Left & "," & _
            btnFindPrev.Top & "," & btnFindPrev.Width & "," & btnFindPrev.Height
      Debug.Print "chkfindoptions.move " & chkFindOptions.Left & "," & _
            chkFindOptions.Top & "," & chkFindOptions.Width & "," & chkFindOptions.Height
      Debug.Print "btnreplace.move " & btnReplace.Left & "," & _
            btnReplace.Top & "," & btnReplace.Width & "," & btnReplace.Height
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
                  .tag = NewString
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
                  ElseIf sOldPath = agEditor.tag Then
                        Caption = "Adjusted the capitalization of open file to: " & sFolder & NewString
                        agEditor.tag = sFolder & NewString
                  Else
                        Caption = "Renamed.  Even though all you changed was the capitalization.  Freak."
                        agEditor.tag = sFolder & NewString
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
            ElseIf sOldPath = agEditor.tag Then
                  Caption = "Renamed open file: " & sFolder & NewString
                  agEditor.tag = sFolder & NewString
            Else
                  Caption = "Rename successful: " & sFolder & NewString
                  agEditor.tag = sFolder & NewString
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
      
            Case eMode.Directory, eMode.Drive, eMode.Floppy, eMode.Cdrom, eMode.Network
                  ' Open the folder, or go up a folder.
                  If Item.Text = ".." Then
                        mnuFileParentDirectory_Click
                  Else
                        cboPath = gBrowserData.Dir & Item.Text & "\"
                  End If
            
            Case eMode.Bookmark
                  ' Open the bookmarked file.  TODO: make it work for folders
                  If FileSize(Item.Text) > BIGFILESIZE Then
                        EditorLoadFile Item.Text, Properties
                  Else
                        EditorLoadFile Item.Text, FileTypeFromExtension(gFSO.getextensionname(Item.Text))
                  End If
                  
            Case Else
                  ' Unless it's too big, open the file.  EditorLoadFile knows what to do.
                  
                  Dim lSize As Long
                  lSize = CLng(Item.ListSubItems("Size"))
                  
                  If lSize < BIGFILESIZE Then
                        EditorLoadFile gBrowserData.Dir & Item.Text, Item.Icon
                  Else
'                        Caption = "WARNING: FILE VERY BIG: " & Item.Text & ", " & _
'                              Format(lSize, "#,#0") & " bytes.   Try another program."
                        EditorLoadFile gBrowserData.Dir & Item.Text, Properties
                  End If
      End Select
End Sub

Private Sub lvwBrowser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      If gBrowserData.HistoryMode Then Exit Sub
      
      Dim iNewKey As Integer
      
      With lvwBrowser
            ' This overhead maneuver can't be used without major, major, major overhaul...
'            If ColumnHeader.key = "Size" Then
'                  lRetVal = SendMessage(.hwnd, LVM_SORTITEMSEX, ByVal .SortOrder, _
'                        AddressOf CompareLong)

            If ColumnHeader.key = "Size" Then
                  iNewKey = 4  ' Doing the switch... 5th column stores size invisibly, with leading zeroes for text sorting.
            Else
                  iNewKey = ColumnHeader.Index - 1
            End If
                  
            If .SortKey = iNewKey Then
                  .Sorted = True
                  .SortKey = 0
                  .SortOrder = Abs(.SortOrder - 1)
                  .SortKey = iNewKey
            Else
                  .Sorted = True
                  .SortKey = iNewKey
            End If
      End With
      
      If gBrowserData.BookmarkMode Then BookmarkSaveChanges

End Sub

Private Sub lvwBrowser_DblClick()
'      Debug.Print "lvwBrowser_DBLCLICK"
      mfBrowserDoubleClick = True
      If Not mfBrowserItemClicked Then
            btnFolderUp_Click
      End If
End Sub

Private Sub lvwBrowser_ItemClick(ByVal Item As MSComctlLib.ListItem)
      mfBrowserItemClicked = True
      ListMenuEnable lvwBrowser.SelectedItem
End Sub


Private Sub lvwBrowser_KeyUp(KeyCode As Integer, Shift As Integer)
      'debug.print "lvwBrowser_KEYUP"
End Sub

Private Sub lvwBrowser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      Debug.Print "lvwBrowser_MOUSEDOWN "; Button & " " & Shift
      
      lvwBrowser_MouseMove Button, Shift, X, Y
      mfBrowserItemClicked = False
      miBrowserMouseButton = Button
      miBrowserShift = Shift
      
End Sub

Private Sub lvwBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      Debug.Print "lvwBROWSER MOUSEMOVE, X: " & x
      
      ' TODO: scroll on dragging the left click.
      ' Store XY as a POINTAPI on mousedown,
      '
      
      Dim litHoverItem As ListItem
      
      ' Autofocus on the file browser.
      ' But we don't do that from within cboPath, because it would be very annoying to
      ' have your typing of a directory interrupted by stray movement of the mouse.
      On Error Resume Next
      If GetForegroundWindow = frmMain.hwnd And Not (ActiveControl.Name = "lvwBrowser") _
            And Not (ActiveControl.Name = "cboPath") Then
            lvwBrowser.SetFocus
      End If
      On Error GoTo 0
      
      ' See if we're over an item.
      Set litHoverItem = lvwBrowser.HitTest(X, Y)
      
      ' Show file names in statusbar on mouseover.
      If Not (litHoverItem Is Nothing) Then
            staTusBar1.Panels(eStat.Tips).Text = litHoverItem.Text
            lvwBrowser.MousePointer = ccCustom
            
            If Button = vbLeftButton Or Button = vbRightButton Then
                  litHoverItem.Selected = True
            End If
      Else
            staTusBar1.Panels(eStat.Tips).Text = ""
            lvwBrowser.MousePointer = ccDefault
      End If
      
'      If GetCapture <> lvwBrowser.hwnd Then
'            SetCapture (lvwBrowser.hwnd)
'      End If
      'Caption = x & " " & y
End Sub

Private Sub mnuBookmark_Click(Index As Integer)
      Dim sEx As String
      
      If FileSize(mnuBookmark(Index).tag) > BIGFILESIZE Then
            EditorLoadFile mnuBookmark(Index).tag, Properties
      Else
             sEx = gFSO.getextensionname(mnuBookmark(Index).tag)
            EditorLoadFile mnuBookmark(Index).tag, FileTypeFromExtension(sEx)
      End If
      
      mnuFileCurrentDirectory_Click
End Sub

Private Sub mnuBookmarksAdd_Click()
      Dim iBookm As Integer
      
      ' TODO: ctrl+M doesn't work from the Editor
      ' find a better shortcut, and see what else doesn't work from the editor.
      
      For iBookm = 1 To mnuBookmark.UBound
            If mnuBookmark(iBookm).tag = agEditor.tag Then
                              ' Oops, got that bookmark already.
                  Exit Sub  ' Nothing left to do here!
            End If
      Next iBookm
      
      AddToBookmarks agEditor.tag
      
      If gBrowserData.BookmarkMode Then btnRefresh_Click
End Sub

Private Sub mnuBookmarksAddPath_Click()
      ' DOES NOTHING YET.  DON'T USE.
      ' TODO: this.
      
      Dim iBookm As Integer
      
      For iBookm = 1 To mnuBookmark.UBound
            If mnuBookmark(iBookm).tag = cboPath Then
                              ' Oops, got that bookmark already.
                  Exit Sub  ' Nothing left to do here!
            End If
      Next iBookm
      
      AddToBookmarks cboPath
End Sub


Private Sub mnuBookmarksManage_Click()
      ' Basically, the brains for the entire program rest within cboPath_Change.
      
      If mnuViewFilebrowser.Checked = False Then mnuViewFilebrowser_Click
      cboPath = "(Bookmarks)"
      
End Sub

Private Sub BrowserDeleteSelected()
      Dim sBookKey As String, iRetVal As Integer
      Dim sTheDamned As String
      
      ' No deletion of history.  If you'd like to delete a file you see in the history,
      ' do it some other way.  (For now, at least).
      
      If lvwBrowser.ListItems.Count = 0 Or gBrowserData.HistoryMode Then Exit Sub
      
      sTheDamned = gBrowserData.Dir & lvwBrowser.SelectedItem
      
      If gBrowserData.BookmarkMode Then
            
            sBookKey = lvwBrowser.SelectedItem.key      ' TODO: FIIIIIIIIXXXXXXXXX
            lvwBrowser.ListItems.Remove sBookKey
            
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
                  If sTheDamned = agEditor.tag Then mnuFileNew_Click
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

Private Sub mnuViewHistory_Click()
      If mnuViewFilebrowser.Checked = False Then mnuViewFilebrowser_Click
      cboPath = "(History)"
End Sub

Private Sub mnuviewzoomout_Click()
'      btnZoomOut_MouseDown vbLeftButton, 0, 10, 10
      
      btnZoomOut_Click
End Sub

Private Sub mnuviewfont_Click()
      btnFont_Click
End Sub

Private Sub mnuviewzoomin_Click()
'      btnZoomIn_MouseDown vbLeftButton, 0, 10, 10
      btnZoomIn_Click
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
      
      If agEditor.tag = "" Then Exit Sub
      
      Set litCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.tag))
      
      If litCurrentFile Is Nothing Then
            cboPath = SnipFileName(agEditor.tag)
            Set litCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.tag))
            If litCurrentFile Is Nothing Then
                  MsgBox "It seems that your file was deleted by another application." & _
                        "  If you wish to keep it, save at once!"
                  Exit Sub
            End If
      End If
      litCurrentFile.Selected = True
      If Not gfFullScreenMode Then litCurrentFile.EnsureVisible
End Sub

Private Sub mnuFileExit_Click()
      Unload Me
End Sub

Private Sub mnuFileHistory_Click(Index As Integer)
      Dim sEx As String
      
      If FileSize(mnuFileHistory(Index).tag) > BIGFILESIZE Then
            EditorLoadFile mnuFileHistory(Index).tag, Properties
      Else
             sEx = gFSO.getextensionname(mnuFileHistory(Index).tag)
            EditorLoadFile mnuFileHistory(Index).tag, FileTypeFromExtension(sEx)
      End If
      
      mnuFileCurrentDirectory_Click
End Sub


Private Sub mnuFileNext_Click()
      btnNextFile_Click
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
      
      If gBrowserData.ERROR Or sParentDir = "" Then
            cboPath = sParentDir
      Else
            cboPath = sParentDir & gBrowserData.Filter
      End If
End Sub

Private Sub mnuFilePrev_Click()
      btnPrevFile_Click
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
      
      sOldPath = agEditor.tag
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
                        agEditor.tag = sNewPath
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
                  agEditor.tag = sNewPath
                  btnRefresh_Click
                  btnCurrentDirectory_Click
            End If
            On Error GoTo 0
      
      End If
End Sub
Private Sub mnuFileSaveAs_Click()
      Dim sDefaultPath As String, sFileName As String
      Dim vDate As Variant
      
      If Not agEditor.Visible Then
            Caption = "ERROR: can only save in editor mode."
            Exit Sub
      ElseIf chkReadOnly.Value = vbChecked Then
            Caption = "ERROR: can't save in Read Only mode."
            Exit Sub
      End If

      vDate = Date
      msPhlegmDate = year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
     
      ' here we decide on a default file name to suggest to the user,
      ' based on a whether the editor.tag is empty, and whether the file browser is at a valid folder.
      If agEditor.tag <> "" Then
            sDefaultPath = agEditor.tag  ' It means this is not a new file we're saving.  Default to old name.
            
      ElseIf gBrowserData.ValidPath Then
            sDefaultPath = gBrowserData.Dir & msPhlegmDate & ".txt"  ' New file, good directory in browser.
      Else
            sDefaultPath = CurDir & "\" & msPhlegmDate & ".txt"  ' New file, no good directory present.
      End If
      
      sFileName = InputBox("File name:", "Save", sDefaultPath)
      If sFileName <> "" Then SaveFile sFileName
End Sub

Private Sub mnuHelpAbout_Click()
      frmAbout.Show
'      MsgBox App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpReadme_Click()
      EditorLoadFile CurDir & "\README.md"
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
      If gBrowserData.BookmarkMode Or gBrowserData.HistoryMode Then
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

Private Sub lblDivider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      If lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "" Then
            
            lblDivider.tag = "Resizing"
      End If
End Sub

Private Sub lblDivider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim iOffset As Integer

      If lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "Resizing" Then
            Dim prevLeft As Long
            prevLeft = lblDivider.Left
            With picEditor
                  iOffset = BrowserResizeHorizontal(X + lblDivider.Left)
                  .Move .Left + iOffset, .Top, .Width - iOffset, .Height
                  agEditor.Move 0, 0, picEditor.Width, picEditor.Height
            End With
            If X <> 0 Then
                RearrangeControls
            End If
      Else
            lblDivider.MousePointer = vbSizeWE
      End If
End Sub

Private Sub lblDivider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If lblDivider.MousePointer = vbSizeWE Then
            lblDivider.MousePointer = 0
            lblDivider.tag = ""
'            agEditor.Width = picBrowser.Width + 160
'            agEditor.Left = frmMain.Width - agEditor.Width - 150
      End If

End Sub


Private Sub mnuListHideFileBrowser_Click()
      mnuViewFilebrowser_Click
End Sub

Private Sub mnuListOpen_Click()
      BrowserExecuteItem lvwBrowser.SelectedItem
End Sub

Private Sub mnuListOpenDefault_Click()
      Dim sPath As String
      
      ' opens the file in whatever program windows chooses for it.
      If lvwBrowser.ListItems.Count > 0 Then
            If gBrowserData.BookmarkMode Or gBrowserData.HistoryMode Then
                  sPath = lvwBrowser.SelectedItem.Text
            Else
                  sPath = gBrowserData.Dir & lvwBrowser.SelectedItem.Text
            End If
            ShellExecute 0, "open", sPath, 0, "", SW_RESTORE
      End If
End Sub

Private Sub mnuListProperties_Click()
      ' SImply calls the Explorer file properties dialog.  Hope this works.
      
      If gBrowserData.BookmarkMode Or gBrowserData.HistoryMode Then
            ShowFileProperties lvwBrowser.SelectedItem
      Else
            ShowFileProperties gBrowserData.Dir & lvwBrowser.SelectedItem
      End If
End Sub

Private Sub mnuListRename_Click()
      ' I've decided to make history unchangeable.  It could have worked the other way,
      ' but it's one of those features that would make you more scared than impressed.
      
      ' Bookmarks are changeable, but it's rewriting the name of the link, not the name of the file.

      If gBrowserData.HistoryMode Or lvwBrowser.Visible = False Then Exit Sub
      
      lvwBrowser.StartLabelEdit
      
End Sub

'   Show only files of extension sEx.
'
Private Sub mnuListShowOnly_Click()
      Dim sEx As String
      
      If gBrowserData.BookmarkMode Or gBrowserData.HistoryMode Then Exit Sub
      
      sEx = gFSO.getextensionname(lvwBrowser.SelectedItem)
      If sEx <> "" Then sEx = "." & sEx
      cboPath = gBrowserData.Dir & sEx
End Sub


Private Sub mnuNext_Click()
      btnNextFile_Click
End Sub

Private Sub mnuPlus_Click()
      mnuViewToolbar_Click
End Sub

Private Sub mnuPrev_Click()
      btnPrevFile_Click
End Sub

Private Sub mnuQuery_Click()
      Debug.Print "mnuQuery_Click"
End Sub

Private Sub mnuQueryClose_Click()
      If Not mfHideFind Then
            mfHideFind = True
            picQuery.Visible = False
            RearrangeControls
      End If
End Sub


Private Sub mnuQueryMatchCase_Click()
      mnuQueryMatchCase.Checked = Not mnuQueryMatchCase.Checked
End Sub

Private Sub mnuQueryReplace_Click()
      ' This is where the mode actually switches
      If mfReplaceMode Then
            mfReplaceMode = False
            txtReplace.Visible = False
            txtFind.SetFocus
            
'            btnFindNext.Move 480, 270, 975, 300
'            btnFindPrev.Move 1440, 270, 855, 300
'            chkFindOptions.Move 0, 270, 495, 300
'            btnReplace.Move 2280, 270, 615, 300
            btnFindNext.Move 480, 270, 1095, 300
            btnFindPrev.Move 1560, 270, 1095, 300
            chkFindOptions.Move 2640, 0, 375, 285
            btnReplace.Move 2640, 270, 375, 300
            
            mnuQueryReplace.Caption = "Show &Replace"
            btnReplace.Enabled = True
      Else
            mfReplaceMode = True
            txtReplace.Visible = True
            If txtReplace = "" Then txtReplace = txtFind
            txtReplace.SetFocus
            txtReplace.SelStart = 0
            txtReplace.SelLength = Len(txtReplace)
            txtReplace_Change

            btnReplace.Move 2640, 295, 372, 285
            btnFindNext.Move 2640, 0, 372, 310
            btnFindPrev.Move 3000, 0, 372, 310
            chkFindOptions.Move 3000, 295, 372, 285
            chkFindOptions.ZOrder
            'txtReplace.Move 0, 290, 2415, 288
            mnuQueryReplace.Caption = "Hide &Replace"
      End If
End Sub

Private Sub mnuQueryWholeWord_Click()
      mnuQueryWholeWord.Checked = Not mnuQueryWholeWord.Checked
End Sub


Private Sub mnueditfind_Click()
      ' Ctrl+F puts the selected text into the query box, but does not proceed with a find until you hit the button.
      
      If giEditorMode = eMode.Picture Or giEditorMode = Properties Then Exit Sub  ' no search/replace within pictures.
      
      If Not mnuViewToolbar.Checked Then mnuViewToolbar_Click
      If mfHideFind Or Not picQuery.Visible Then
            mfHideFind = False
            picQuery.Visible = True
            RearrangeControls
      End If
      On Error Resume Next
      If ActiveControl.Name = "agEditor" And agEditor.SelectedText <> "" Then _
            txtFind = Trim(agEditor.SelectedText)
      On Error GoTo 0
      txtFind.SetFocus
End Sub

Private Sub mnueditfindBackwards_Click()
      btnFindPrev_Click
End Sub

Private Sub mnueditfindNext_Click()
      btnFindNext_Click
End Sub

Private Sub mnuViewOptions_Click()
'      frmOptions.Show
End Sub

Private Sub mnuViewReadOnly_Click()
      chkReadOnly.Value = Abs(chkReadOnly.Value - 1)
End Sub


Private Sub mnuEditReplace_Click()
      If giEditorMode = eMode.Picture Or giEditorMode = Properties Then Exit Sub
      
      mfHideFind = False
      If mnuViewToolbar.Checked = False Then mnuViewToolbar_Click
      picQuery.Visible = True
      If mfReplaceMode And ActiveControl.Name <> "txtReplace" And txtReplace = "" Then
            txtReplace.SetFocus
      ElseIf mfReplaceMode And ActiveControl.Name <> "txtreplace" And Not btnReplace.Enabled Then
            txtReplace.SetFocus
      Else
            btnReplace_Click
      End If
End Sub

Private Sub mnuWriteDelete_Click()
      agEditor.InsertContents SF_TEXT, ""
End Sub

Private Sub mnuWriteFind_Click()
      mnueditfind_Click
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
                  Image1.Move 0, 0, gImageData.DefaultWidth, gImageData.DefaultHeight
            Case 103, 55   ' 7 and Keypad 7
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
                  
            Case vbKeyHome
                  Image1.Top = 0
            Case vbKeyEnd
                  Image1.Top = picEditor.Height - Image1.Height
                  
            Case vbKeyPageUp
                  If Image1.Top < -picEditor.Height Then
                        Image1.Top = Image1.Top + picEditor.Height
                  ElseIf Image1.Top < 0 Then
                        Image1.Top = 0
                  End If
                  
            Case vbKeyPageDown
                  If Image1.Top + Image1.Height > picEditor.Height * 2 Then
                        Image1.Top = Image1.Top - picEditor.Height
                  ElseIf Image1.Top + Image1.Height > picEditor.Height Then
                        Image1.Top = picEditor.Height - Image1.Height
                  End If
            
            Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                  If Shift = 0 Then BrowserExecuteNext
            Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                  If Shift = 0 Then BrowserExecutePrev
      End Select
End Sub

Private Sub picEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error Resume Next
      If GetForegroundWindow = frmMain.hwnd And Not (ActiveControl.Name = "picEditor") Then
            picEditor.SetFocus
      End If
      On Error GoTo 0
End Sub

Private Sub picEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not gImageData.Zoomed And Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            BrowserExecuteNext
      ElseIf Not gImageData.Zoomed And Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            BrowserExecutePrev
      End If
      
      gImageData.Zoomed = False
      gImageData.Moved = False
End Sub

Private Sub sliZoom_Change()
      ImageSetZoom (sliZoom.Value)
End Sub

Private Sub sliZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error Resume Next
      If GetForegroundWindow = frmMain.hwnd And Not (ActiveControl.Name = "sliZoom") Then
            sliZoom.SetFocus
      End If
      On Error GoTo 0
End Sub

Private Sub sliZoom_Scroll()
      ImageSetZoom (sliZoom.Value)
End Sub

Private Sub txtFind_Change()
      If mfReplaceMode And txtFind = agEditor.SelectedText And txtFind <> "" Then
            btnReplace.Enabled = True
            If ActiveControl.Name = "txtReplace" Then btnReplace.Default = True
      ElseIf mfReplaceMode Then
            btnReplace.Enabled = False
      End If
      
      ' Reset Find results when the find box is changed, even slightly, even if it's never used in a search.
      miFindResult = 0
      miTotalResults = 0
      lblFindResult = ""
End Sub

'Private Sub txtFind_Change()
'      Dim pos As Integer
'      Dim quickkey As String
'      Dim NewQuery As URLQueryType
'
'      pos = InStr(0, txtFind, " ", )
'      NewQuery.key = Left(txtFind, pos)
'      NewQuery.URL = Right(txtFind, pos)
'End Sub

Private Sub txtFind_GotFocus()
      txtFind.SelStart = 0
      txtFind.SelLength = Len(txtFind)
      btnFindNext.Default = True
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
'      If KeyCode = vbKeyReturn And Shift = 0 Then
'            If txtFind <> "" Then btnFindNext_Click
      If KeyCode = vbKeyReturn And Shift = vbShiftMask Then
            If txtFind <> "" Then btnFindPrev_Click
      ElseIf KeyCode = vbKeyShift Then
            btnFindPrev.Default = True
      End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyShift Then
            btnFindNext.Default = True
      End If
End Sub


Private Sub txtFind_LostFocus()
      btnFindPrev.Default = False
      btnFindNext.Default = False
End Sub


Private Sub txtFind_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
      txtFind = Data.GetData(vbCFText)
      txtFind_KeyDown vbKeyReturn, 0
End Sub

Private Sub txtFind_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
      txtFind.SelStart = 0
      txtFind.SelLength = Len(txtFind)
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
            With mStats
                .Y = lLineIndex + 1
                
                ' We want mStats.i to count CRs and LFs both, since agEditor.CharacterCount does that.
                .i = lMin
                SendMessage agEditor.RichEdithWnd, EM_EXGETSEL, 0, chrSelection
                .X = lMin - lCharIndex + 1
                .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
            End With
        
            FillStats
            staTusBar1.Panels(eStat.SelText) = lMax - lMin
      End If
      
      If mfReplaceMode And txtFind = agEditor.SelectedText And txtFind <> "" Then
            btnReplace.Enabled = True
            If ActiveControl.Name = "txtReplace" Then btnReplace.Default = True
      ElseIf mfReplaceMode Then
            btnReplace.Enabled = False
      End If
      
      ' Reset Find result count whenever the selection changes
      ' (...changes from something other than inside a Find)
      If Not mfFinding Then
            miFindResult = 0
            miTotalResults = 0
            lblFindResult = ""
      End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

      Select Case KeyCode
            Case 220 '  Backslash.  Making Alt+\ into a spare Tab key for the right side of the keyboard.
                  If Shift And vbAltMask Then
                        SendKeys "{TAB}"
                  End If
                  
            Case vbKeyReturn           ' File Properties window from Explorer!
                  If Shift And vbAltMask Then
                        If ActiveControl.Name = "lvwBrowser" And lvwBrowser.ListItems.Count > 0 Then
                              ShowFileProperties gBrowserData.Dir & lvwBrowser.SelectedItem
                        Else
                              ShowFileProperties agEditor.tag
                        End If
                  End If
                  
            Case vbKeyF
                  If Shift = vbCtrlMask + vbShiftMask Then ' TODO: this is still writing letters to the editor.
                        btnFont_Click
                  End If
            
            Case vbKeyF11
                  If Shift = 0 Then btnFullScreen_Click
            
            Case 221 ' Right Bracket "]"
                  If Shift = vbCtrlMask Then BrowserExecuteNext
            
            Case 219 ' Left Bracket "["
                  If Shift = vbCtrlMask Then BrowserExecutePrev
                  
            Case 188 ' Comma (",")  ...also "<"
                  If Shift = vbCtrlMask + vbShiftMask Then
                        mnuviewzoomout_Click
                  End If
                  
            Case 190 ' Period (".") ...also ">"
                  If Shift = vbCtrlMask + vbShiftMask Then
                        mnuviewzoomin_Click
                        
                  ElseIf Shift = vbAltMask And chkFindOptions.Value = vbUnchecked Then
                        'Alt+period  opens popup menu for find options
                        chkFindOptions.SetFocus
                        chkFindOptions.Value = vbChecked
                  ElseIf chkFindOptions.Value = vbChecked Then
                        ' Same button closes find options menu, if already opened
                        chkFindOptions.Value = vbUnchecked
                  End If
            
            Case vbKeyEscape  ' Popup menu doesn't wanna die by itself; escape closes it.
                                                ' Sure wish there were a way to test if a menu is open!
                  If Shift = 0 And chkFindOptions.Value = vbChecked Then
                        chkFindOptions.Value = vbUnchecked
                  ElseIf Shift = 0 And (ActiveControl.Name = "txtFind" Or ActiveControl.Name = "txtReplace") Then
                        ' Get rid of the find, on Esc button from within the find.
                        btnCloseFind_Click
                        If agEditor.Visible Then agEditor.SetFocus
                  End If
      End Select
End Sub



Private Sub Form_Load()
      Dim vDate As Variant
      Dim sCommandFile As String

      InitializeMenus
            
      Set gFSO = CreateObject("Scripting.FileSystemObject") ' Just so I'll never have to do this again.
      Set gImageData.OutPic = Image1
      
      gBrowserData.ListEmpty = True
      
      giEditorMode = Text

'      miImageZoom = 100
      
      'Debug.Print "command line sayeth: [" & Command() & "]"
      sCommandFile = Trim(Command())
      If Left(sCommandFile, 1) = Chr(34) Then sCommandFile = Mid(sCommandFile, 2, Len(sCommandFile) - 2)
      If sCommandFile <> "" And Not (sCommandFile Like "*:\*") Then sCommandFile = CurDir & "\" & sCommandFile
      agEditor.tag = sCommandFile
      
      msPhlegmKey = "Software\" & App.title & "\" & msSettingsVersion
      
      vDate = Date
      msPhlegmDate = year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
      
      GetWindowSettings
      mStats.imax = CharacterCount(agEditor)
      FillStats
      staTusBar1.Panels(eStat.Modified) = ""

      If Not Debugging Then
            gpOldLvwProc = SetWindowLong(lvwBrowser.hwnd, GWL_WNDPROC, _
                  AddressOf ListViewProc)
      End If
      
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
      ctrlhWnd = WindowFromPoint(poiCursor.X, poiCursor.Y)
      
      On Error Resume Next
      If Screen.ActiveControl.Container.Name = "picQuery" Then
            txtFind.SetFocus
      End If
      On Error GoTo 0
      
'      If Not mfSkipMouseEventCrap And ctrlhWnd = agEditor.RichEdithWnd Then
'            mouse_event MOUSEEVENTF_LEFTDOWN, poiCursor.x, poiCursor.y, 0, 0
'            mouse_event MOUSEEVENTF_LEFTUP, poiCursor.x, poiCursor.y, 0, 0
'      ElseIf ctrlhWnd = lvwBrowser.hwnd Then
'            FuckIHateThis = True
'            mouse_event MOUSEEVENTF_LEFTDOWN, poiCursor.x, poiCursor.y, 0, 0
'            mouse_event MOUSEEVENTF_RIGHTUP, poiCursor.x, poiCursor.y, 0, 0
'      ElseIf ctrlhWnd = txtFind.hwnd Then
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
      Const ColumnSizeInc = 50
      
      'debug.print "KEYDOWN " & KeyCode & " " & Shift
      
      Select Case KeyCode
                        
            Case vbKeyC
                  If Shift = vbCtrlMask Then mnuListCopyPath_Click
                  
            Case vbKeyLeft
                  If Shift = vbAltMask Then ' Alt+Left = go back in the recent paths list
                        PathBack
                  ElseIf (Shift And vbCtrlMask) Then
                        ' Ctrl+left = scroll left.  No additional coding needed.
                  ElseIf Shift = vbShiftMask Then
                        ' Shift+left is going to size column width
                        If lvwBrowser.ColumnHeaders.Item(1).Width >= ColumnSizeInc Then
                              lvwBrowser.ColumnHeaders.Item(1).Width = _
                                    lvwBrowser.ColumnHeaders.Item(1).Width - ColumnSizeInc
                        End If
                  Else
                        mnuFileParentDirectory_Click   ' Ordinary left arrow...
                  End If
                                                 
                                                 
            Case vbKeyF13 ' F13, but contains code for it and for right arrow.
                              ' See ListViewProc for details.
                              
                  ' Right = open a folder or a drive, but leave a file alone.
                  '     ...and don't fucking scroll anywhere.
                  
                  If Shift = vbShiftMask Then
                        ' Oh, and shift+right is going to increase column width
                        lvwBrowser.ColumnHeaders.Item("Name").Width = _
                              lvwBrowser.ColumnHeaders.Item("Name").Width + ColumnSizeInc
                              
                  ElseIf lvwBrowser.ListItems.Count > 0 Then
                        With lvwBrowser.SelectedItem
                              If .Icon = eMode.Directory Or .Icon = eMode.Drive Or .Icon = eMode.Cdrom Or _
                                                .Icon = eMode.Floppy Or .Icon = eMode.Network And Shift = 0 Then
                                    BrowserExecuteItem lvwBrowser.SelectedItem
                              End If
                        End With
                  End If
            
            Case vbKeyInsert
                  ' I'm making Insert be a reverse sort order.  It's right up there next to the
                  ' navigational keys and I'm always wanting a reverse right near them and
                  ' insert wasn't serving any purpose.
                  
                  btnSort_Click
            
            Case vbKeyRight
            
                  ' Alt+Right = go forward in the recent paths list
                  
                  If Shift = vbAltMask Then
                        PathForward
                  ' Ctrl+right = scroll right.  (Happens automatically.)
                  End If
                  
            Case vbKeyN
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.Item("Name")
            Case vbKeyT
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.Item("Type")
            Case vbKeyZ
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.Item("Size")
            Case vbKeyM
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.Item("Modified")
            
            Case vbKeyDelete
                  If Shift = 0 Then BrowserDeleteSelected
                  
            Case 219 ' Left Bracket [
                  If Shift = 0 Then BrowserExecutePrev
            
            Case 221 ' Right Bracket ]
                  If Shift = 0 Then BrowserExecuteNext
                  
            Case 220 ' Backslash \
                  If Shift = 0 Then SendKeys "{TAB}"
            
            Case vbKeyBack
                  If Shift = 0 Then PathBack
            
            Case 93 ' That keyboard button that usually means right click.
                  Dim iItemX As Integer, iItemY As Integer
                  
                  iItemX = picBrowser.Left + lvwBrowser.Left + lvwBrowser.SelectedItem.Left + lvwBrowser.SelectedItem.Width
                  iItemY = picBrowser.Top + lvwBrowser.Top + lvwBrowser.SelectedItem.Top + lvwBrowser.SelectedItem.Height
                  Me.PopupMenu mnuList, , iItemX, iItemY, mnuListOpen
            
            Case vbKeyReturn
                  If Shift = 0 Then
                        BrowserExecuteItem lvwBrowser.SelectedItem
                  ElseIf Shift = vbCtrlMask + vbShiftMask Then
                        mnuListOpenDefault_Click
                  End If
      End Select
End Sub

Private Sub lvwBrowser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'debug.print "lvwBrowser_MOUSEUP " & Button & " " & Shift
'      If FuckIHateThis Then
'            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'            FuckIHateThis = True
'      End If
      
      Dim litHoverItem As ListItem
      
      Set litHoverItem = lvwBrowser.HitTest(X, Y)  ' To see if we're over an item.
      
      If (Button = vbRightButton And Shift = 0) Then
            If litHoverItem Is Nothing Then
                  ListMenuDisable
            Else
                  ListMenuEnable litHoverItem
            End If
            Me.PopupMenu mnuList
      
      ElseIf Button = vbLeftButton And Shift = 0 Then
            
            If Not (litHoverItem Is Nothing) Then
                  ' Open the file/folder on an ordinary left click.
                  BrowserExecuteItem litHoverItem
            Else
                  ' Clicking on empty space deselects the selected item.
                  If Not gBrowserData.ListEmpty Then lvwBrowser.SelectedItem.Selected = False
            End If
      
      ElseIf Button = vbMiddleButton And Shift = 0 Then
            If litHoverItem Is Nothing Then PathBack
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
      
      agEditor.Text = ""
      agEditor.tag = ""
      EditorSetMode Text
      frmMain.Caption = "(New File)"
      chkReadOnly.Value = vbUnchecked
End Sub

Private Sub mnuFileSave_Click()
      Dim fSuccess As Boolean
      Dim dteSaveTime As Date

      If Not agEditor.Visible Then
            Caption = "ERROR: can only save in editor mode."
            Exit Sub
      ElseIf chkReadOnly.Value = vbChecked Then
            Caption = "ERROR: can't save in Read Only mode."
            Exit Sub
      End If
      
      If agEditor.tag = "" Then  ' If they try to save a nameless New File
            mnuFileSaveAs_Click
      Else
            SaveFile agEditor.tag
      End If
End Sub


Public Function SaveFile(ByVal sFileName As String)
      Dim fSuccess, fNewFile As Boolean
      Dim dteSaveTime As Date

      If Len(sFileName) > 100 Or agEditor.Text = "" Or gTextEncoding = eTextEncoding.UNICODE Then
            Dim ts
            On Error GoTo File_Error
            If Not FileExists(sFileName) Then
                  fNewFile = True
                  Set ts = gFSO.CreateTextFile(sFileName, eOverwrite.Yes, gTextEncoding)
            Else
                  Set ts = gFSO.OpenTextFile(sFileName, eIoMode.ForWriting, eCreate.No, gTextEncoding)
            End If
            If agEditor.Text = "" And Not fNewFile Then
                  ts.Write ("temporary text to make sure the file counts as modified")
                  ts.Close
                  Set ts = gFSO.OpenTextFile(sFileName, eIoMode.ForWriting)
                  ' TODO: titlebar will show a false positive that it wrote to file in this ONE scenario
                  '     * already existing file
                  '     * we do not have permission to write to it
                  '     * we are trying to save a blank file of exactly 0 bytes
                  '     ...this is just too niche to care about anymore
            End If
            ts.Write (agEditor.Text)
            ts.Close
            On Error GoTo 0
            fSuccess = True
      Else
            fSuccess = agEditor.SaveToFile(sFileName, SF_TEXT)
      End If

      If fSuccess Then
            Dim bytes As Long
            bytes = agEditor.CharacterCount
            staTusBar1.Panels(eStat.encoding) = "ASCII"
            If gTextEncoding = eTextEncoding.UNICODE Then
                  bytes = bytes * 2 + 2
                  staTusBar1.Panels(eStat.encoding) = "UNICODE"
            End If
            staTusBar1.Panels(eStat.Modified) = ""
            agEditor.tag = sFileName
            Caption = sFileName & "  (" & Format(bytes, "#,#0") & " bytes saved on " _
                  & FileModifiedTime(sFileName) & ")"
            btnRefresh_Click
            btnCurrentDirectory_Click
      Else
            frmMain.Caption = "ERROR: cannot save to " & sFileName
      End If
      Exit Function
      
File_Error:
      frmMain.Caption = "ERROR: cannot save to " & sFileName
End Function
Private Sub mnuViewFilebrowser_Click()
    chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub agEditor_Change()

      If Not mfEditorLoading Then staTusBar1.Panels(eStat.Modified) = "Modified"
      
      If staTusBar1.Visible Then
            With mStats
                .imax = CharacterCount(agEditor)
                .ymax = agEditor.LineCount
            End With
            
            FillStats
      End If
      
      ' Reset Find result count when the document changes.
      miFindResult = 0
      miTotalResults = 0
      lblFindResult = ""
End Sub

Private Sub agEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If (Button = vbRightButton And Shift = 0) Then
          Me.PopupMenu mnuWrite
      End If
End Sub

Private Sub FillStats()

      staTusBar1.Panels(eStat.Stats) = "Char: " & Format(mStats.i, "#,#0") & "/" & Format(mStats.imax, "#,#0") _
            & "  Ln: " & Format(mStats.Y, "#,#0") & "/" & Format(mStats.ymax, "#,#0") & "  Col: " & mStats.X _
            & "/" & mStats.xmax
End Sub


Private Sub RearrangeControls()

      ' TODO: clean up these godawful variable names!

      ' Put the various controls where they need to be.
      '   agEditor, lvwBrowser
      ' Made to go on a window resize or when showing or hiding a control
      
      Dim iEdHeight As Integer, iEdWidth As Integer, iEdTop As Integer, iEdLeft As Integer
      Dim iToolbarFullWidth As Integer
      Dim iBrowserTop, iBrowserHeight As Integer
      Dim lineindex As Long, charindex As Long, lMin As Long, lMax As Long
      Dim fValidWindowSize As Boolean, iRedoResizeX As Integer, iRedoResizeY As Integer
      Dim iPicBoxMarginsY As Integer, iFormMarginsX As Integer, iFormMarginsY As Integer
      Dim sHadFocus As String
      
      Const topmargin = 100
      Const leftmargin = 0
      Const rightmargin = 150
      Const midspace = 100
      Const bottommargin = 30
      Const toolbarWidth = 4905
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      fValidWindowSize = True ' ...until proven guilty.
      iRedoResizeY = frmMain.Height
      iRedoResizeX = frmMain.Width
      
      If Not (ActiveControl Is Nothing) Then  ' activecontrol is nothing if image1 is up front...
            sHadFocus = ActiveControl.Name                               ' images cannot take focus.
            picEditor.Visible = False ' MUCH faster if you turn him off while thinking (unless he's empty).
      End If
      
      ' Calculate control positions...
      
      iEdLeft = leftmargin
      If mnuViewFilebrowser.Checked Then iEdLeft = iEdLeft + picBrowser.Width
      
      iEdWidth = frmMain.ScaleWidth - iEdLeft
      
      If mnuViewToolbar.Checked Then
            iBrowserTop = picToolBar.Height
            If Not mfHideFind And picQuery.Visible Then
                  iToolbarFullWidth = toolbarWidth + picQuery.Width
            Else
                  iToolbarFullWidth = toolbarWidth
            End If
      Else
            iToolbarFullWidth = 0
            iBrowserTop = 0
      End If
      
      iEdTop = 0
      If mnuViewToolbar.Checked And (Not mnuViewFilebrowser.Checked Or picBrowser.Width < iToolbarFullWidth) Then
            iEdTop = topmargin + picToolBar.Height
      End If
      
      iEdHeight = frmMain.ScaleHeight - iEdTop - bottommargin
      iBrowserHeight = frmMain.ScaleHeight - iBrowserTop - bottommargin
      If mnuViewStatusBar.Checked Then
            iEdHeight = iEdHeight - staTusBar1.Height
            iBrowserHeight = iBrowserHeight - staTusBar1.Height
      End If
      
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
      
      picBrowser.Move 0, iBrowserTop, picBrowser.Width, iBrowserHeight
      picEditor.Move iEdLeft, iEdTop, iEdWidth, iEdHeight
      agEditor.Move 0, 0, iEdWidth, iEdHeight
      lvwBrowser.Height = iBrowserHeight - lvwBrowser.Top + topmargin


'      If mnuViewToolbar.Checked Then btnToolbarClose.Left = frmMain.ScaleWidth - btnToolbarClose.Width - 50
            
      If giEditorMode = eMode.Text Then
            ' a few things in the statusbar could change in a window resize:
            '   x, xmax, y, ymax
            ' and some shouldn't change:
            '   i, imax,   (we're not adding or deleting characters or moving the cursor)
            '   sellength
            
            agEditor.GetSelection lMin, lMax
            lineindex = agEditor.CurrentLine
            charindex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lineindex, 0)
            
            If staTusBar1.Visible Then
                  With mStats
                      .X = lMin - charindex + 1
                      .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal charindex, 0) + 1
                      .Y = lineindex + 1
                      .ymax = agEditor.LineCount
                  End With
                  FillStats
            End If
      End If
      staTusBar1.Panels(eStat.Tips).Width = frmMain.Width
      
      picEditor.Visible = True
      'If sHadFocus = "agEditor" Then agEditor.SetFocus
End Sub

Private Sub mnuViewStatusBar_Click()
      staTusBar1.Visible = Not staTusBar1.Visible
      mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
      RearrangeControls
End Sub

Private Sub InitializeMenus()
'      Dim tempinfo As MENUITEMINFO
'      Dim hMenu As Long, retval As Long
'
'      hMenu = GetMenu(hwnd)
'      hMenu = GetSubMenu(hMenu, 2)
'      retval = ModifyMenu(hMenu, 0, MF_STRING + MF_BYPOSITION, 2, "&Penis" + vbTab + "Ctrl+P")
      
'      mnuviewzoomin.Caption = "&Increase Font Size" & vbTab & "Alt+="
      mnuEditUndo.Caption = "Undo" & vbTab & "Ctrl+Z"
      mnuEditRedo.Caption = "Redo" & vbTab & "Ctrl+Y"
      mnuViewFont.Caption = "Font..." & vbTab & "Shift+Ctrl+F"
      mnuViewZoomIn.Caption = mnuViewZoomIn.Caption & vbTab & "Shift+Ctrl+" & QUOT & ">" & QUOT
      mnuViewZoomOut.Caption = mnuViewZoomOut.Caption & vbTab & "Shift+Ctrl+" & QUOT & "<" & QUOT
      
      mnuFileNext.Caption = mnuFileNext.Caption & vbTab & "Ctrl+]"
      mnuFilePrev.Caption = mnuFilePrev.Caption & vbTab & "Ctrl+["
      
      mnuWriteFind.Caption = mnuWriteFind.Caption & vbTab & "Ctrl+F"
      mnuWriteCut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnuWriteCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuEditCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuWritePaste.Caption = "&Paste" & vbTab & "Ctrl+V"
      
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
      Dim sNextDrive As String, iDriveIcon As eMode
      Dim lLength As Long
      Dim iIndex As Integer, iTempKey As Integer
      Dim litCurrentItem As ListItem
      
            
      lLength = GetLogicalDriveStrings(100, sDrivesFixed)
      sDriveString = Left(sDrivesFixed, lLength)
      sDriveArray = Split(sDriveString, Chr(0)) ' "(x,x, , )" is an error.  don't put in more commas unless
      lvwBrowser.ListItems.Clear          ' they lead to something.
      lvwBrowser.tag = ""
      
      iTempKey = lvwBrowser.SortKey
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      
      iIndex = LBound(sDriveArray)
      sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      
      Do While (sNextDrive <> "") And (sNextDrive <> Chr(0))
            
            Select Case gFSO.getdrive(sNextDrive).drivetype
                  Case 1: iDriveIcon = Floppy
                  Case 2: iDriveIcon = Drive
                  Case 3: iDriveIcon = Network
                  Case 4: iDriveIcon = Cdrom
            End Select
            Set litCurrentItem = lvwBrowser.ListItems.Add( _
                  1, , sNextDrive, iDriveIcon, iDriveIcon)
            litCurrentItem.ListSubItems.Add , , 0
            
            iIndex = iIndex + 1
            sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      Loop
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = iTempKey
      BrowserGetDrives = iIndex - 1
      
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " drives"
End Function
Private Sub mnuViewToolbar_Click()
      If mnuViewToolbar.Checked Then
            mnuViewToolbar.Checked = False
            picToolBar.Visible = False
            picQuery.Visible = False
            mnuPlus.Caption = "+"
            mnuNext.Visible = True
            mnuPrev.Visible = True
      Else
            mnuViewToolbar.Checked = True
            picToolBar.Visible = True
            If Not mfHideFind Then picQuery.Visible = True
            mnuPlus.Caption = "="
            mnuNext.Visible = False
            mnuPrev.Visible = False
      End If
      RearrangeControls
End Sub

Private Sub mnuViewWordWrap_Click()
      chkWordWrap.Value = Abs(chkWordWrap.Value - 1)
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

Private Function EditorLoadFile(ByVal sFileName As String, Optional ByVal iMode As eMode = Text) As Boolean
      
      Dim fLoadSuccess As Boolean, sCaption As String

      If mfEditorLoading Then agEditor.Text = ""
      mfEditorLoading = True
      
      If Trim(sFileName) = "" Then  ' Blank means start a new file.  For when the registry settings come up empty.
            mnuFileNew_Click
      
      ElseIf Not FileExists(sFileName) Then
            frmMain.Caption = "ERROR: file does not exist."
            agEditor.tag = ""
            
      Else ' Normal file load.
            
            EditorSetMode iMode
            
            Select Case iMode
                  
                  Case eMode.Text, eMode.other, eMode.rtf
                        ' pass along the boolean return value, if anyone wants it.
                        If Not gfFullScreenMode And FileSize(sFileName) > 0 Then
                              Dim encoding As Integer
                              encoding = IsUnicodeFile(sFileName)
                              
                              If encoding = eTextEncoding.ERROR Then
                                    frmMain.Caption = "Could not load file: " + sFileName
                                    mfEditorLoading = False
                                    EditorLoadFile = False
                                    Exit Function
                              ElseIf Len(sFileName) > 100 Or encoding = eTextEncoding.UNICODE Then
                                    Dim f, ts
                                    Set f = gFSO.getfile(sFileName)
                                    Set ts = f.OpenAsTextStream(eIoMode.ForReading, encoding)
                                    If ts.atendofstream() Then
                                          agEditor.Text = ""
                                    Else
                                          agEditor.Text = ts.readall()
                                    End If
                                    ts.Close
                                    fLoadSuccess = True
                              Else
                                    fLoadSuccess = agEditor.LoadFromFile(sFileName, SF_TEXT)
                              End If
                              gTextEncoding = encoding
                              If encoding = eTextEncoding.UNICODE Then
                                    staTusBar1.Panels(eStat.encoding) = "UNICODE"
                              Else
                                    staTusBar1.Panels(eStat.encoding) = "ASCII"
                              End If
                        Else
                              agEditor.Text = ""
                              gTextEncoding = eTextEncoding.ASCII
                              staTusBar1.Panels(eStat.encoding) = "ASCII"
                              fLoadSuccess = True
                        End If
                  
                        sCaption = sFileName & "  (" & Format(FileSize(sFileName), "#,#0") & " bytes saved on " _
                              & FileModifiedTime(sFileName) & ")"
'                  Case rtf
'                        fLoadSuccess = agEditor.LoadFromFile(sFileName, SF_RTF)
                        
                  Case eMode.Picture
                        Dim DefaultWidth, DefaultHeight
                        fLoadSuccess = True
                        On Error Resume Next
                        gImageData.OutPic.Picture = LoadPicture(sFileName)
                        Const twipConversion = 0.567
                        DefaultWidth = gImageData.OutPic.Picture.Width * twipConversion
                        DefaultHeight = gImageData.OutPic.Picture.Height * twipConversion
                        If Width >= 65536 Then
                            DefaultWidth = 65535
                        End If
                        If DefaultHeight >= 65536 Then
                            DefaultHeight = 65535
                        End If
                        gImageData.DefaultWidth = DefaultWidth
                        gImageData.DefaultHeight = DefaultHeight
                        ImageSetZoom (sliZoom.Value)
                        sCaption = sFileName & "  (" & sliZoom.Value & "%)"
                        
                        If Err > 0 Then
                              Caption = "ERROR: " & sFileName & ", picture couldn't load"
                              fLoadSuccess = False
                        End If
                        On Error GoTo 0
                  
                  Case eMode.Properties
                        GetFileProperties sFileName
                        sCaption = sFileName
                        fLoadSuccess = True
            End Select
                  
            If fLoadSuccess Or FileSize(sFileName) = 0 Then  ' Success!
                  agEditor.tag = sFileName
                  frmMain.Caption = sCaption
                  If gfFullScreenMode Then
                        frmFullScreen.lblFileNameZoom = sCaption & "  "
                  End If
                  staTusBar1.Panels(eStat.Modified) = ""
                  agEditor.SetSelection 0, 0
                  AddToHistory sFileName
            
            Else  ' Miscellaneous Failure!  agEditor returns no clues as to the problem.
                  frmMain.Caption = "Could not load file.  command() = " & Chr(34) & Command() & Chr(34) _
                        & "; File = " & Chr(34) & sFileName & Chr(34)
                  agEditor.tag = ""
            End If
      End If
      
      EditorLoadFile = fLoadSuccess
      mfEditorLoading = False
End Function

Public Sub ImageSetZoom(iZoom As Integer)
      gImageData.OutPic.Stretch = True
      gImageData.OutPic.Move gImageData.OutPic.Left, gImageData.OutPic.Top, _
            gImageData.DefaultWidth * CSng(iZoom) / 100#, gImageData.DefaultHeight * CSng(iZoom) / 100#
'      miImageZoom = iZoom
      Caption = agEditor.tag & "  (" & iZoom & "%)"
End Sub

Private Sub SaveWindowSettings()
      Dim lMin As Long, lMax As Long, lKey As Long, lRetVal As Long
      Dim lNewOrUsed As Long, lValueSize As Long
      Dim iIndex As Integer
      
'      Dim wnpPlacement As WINDOWPLACEMENT'
'      Dim rectRestored As RECT
      Dim fntTemp As New StdFont
'      Dim poiTemp As POINTAPI
      
      
      With mWindowPrefs
            .WNP.Length = LenB(.WNP)
            GetWindowPlacement hwnd, .WNP
            If .WNP.showCmd = SW_MINIMIZE Then
                  .WNP.showCmd = SW_RESTORE
            ElseIf .WNP.showCmd = SW_SHOWMINIMIZED Then  '  <-- It'll be this one, not SW_MINIMIZE.
                  .WNP.showCmd = SW_SHOWNORMAL                ' Including the other for paranoia.
            End If
            
            .BrowserWidth = picBrowser.Width
            .SortKey = lvwBrowser.SortKey
            On Error GoTo 0
            .NameColumn = lvwBrowser.ColumnHeaders.Item("Name").Position
            .TypeColumn = lvwBrowser.ColumnHeaders.Item("Type").Position
            .SizeColumn = lvwBrowser.ColumnHeaders.Item("Size").Position
            .ModifiedColumn = lvwBrowser.ColumnHeaders.Item("Modified").Position
            On Error Resume Next
            
            .ShowFileBrowser = picBrowser.Visible
            .ShowStatusBar = staTusBar1.Visible
            .ShowToolBar = picToolBar.Visible
            .ShowFind = Not mfHideFind
            .SortMethod = lvwBrowser.SortOrder
            .AutoLoadFile = agEditor.tag
            .cboPath = cboPath
            .BookmarkCount = mnuBookmark.UBound
            .HistoryCount = mnuFileHistory.UBound
      End With
      
      agEditor.GetSelection lMin, lMax
      
      With mEditorPrefs
            .FirstVisibleLine = agEditor.FirstVisibleLine
            .SelEnd = lMax
            .SelStart = lMin
            .WordWrap = chkWordWrap.Value
            ' If we were set to readonly while looking at pictures, I'll assume the setting wasn't
            ' REALLy that important, at the time.  So, not saving it in that case.
            If giEditorMode <> Picture And chkReadOnly.Value = vbChecked Then
                  .ReadOnly = vbChecked
            Else
                  .ReadOnly = vbUnchecked
            End If
            
            Set fntTemp = GetRealStdFont(agEditor, .TextColor)
            ' Here, we'll store the color as a system color, if it happens to match the button text.
            If .TextColor = TranslateColor(vbWindowText) Then .TextColor = vbWindowText
            .FontBold = fntTemp.Bold
            .FontItalic = fntTemp.Italic
            .FontName = fntTemp.Name
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
      
      lValueSize = LenB(mWindowPrefs)
      lRetVal = RegSetValueExAny(lKey, "Settings", 0, REG_NONE, _
                  ByVal mWindowPrefs, lValueSize)
      If lRetVal <> 0 Then MsgBox "RegSetValueEx Failed.  settings: " & _
                  LenB(mWindowPrefs) & " " & lKey, , App.title
      
      ' Store the File Settings.
      
      lValueSize = LenB(mEditorPrefs)
      lRetVal = RegSetValueExAny(lKey, "agEditor", 0, REG_NONE, _
                  ByVal mEditorPrefs, lValueSize)
      If lRetVal <> 0 Then MsgBox "RegSetValueEx Failed.  mEditorPrefs: " & _
                  LenB(mEditorPrefs) & " " & lKey, , App.title
      
      ' Store Bookmarks.
      
      For iIndex = 1 To mnuBookmark.UBound
            lValueSize = LenB(mnuBookmark(iIndex).tag)
            lRetVal = RegSetValueExString(lKey, "Bookmark" & CStr(iIndex), 0, REG_SZ, _
                        ByVal mnuBookmark(iIndex).tag, lValueSize)
      Next iIndex
      
      For iIndex = mnuBookmark.UBound + 1 To mWindowPrefs.BookmarkCount
            RegDeleteValue lKey, "Bookmark" & CStr(iIndex)
      Next iIndex
      
      ' Store History.
      
      For iIndex = 1 To mnuFileHistory.UBound
            lValueSize = LenB(mnuFileHistory(iIndex).tag)
            lRetVal = RegSetValueExString(lKey, "History" & CStr(iIndex), 0, REG_SZ, _
                  ByVal mnuFileHistory(iIndex).tag, lValueSize)
      Next iIndex
      
      For iIndex = mnuFileHistory.UBound + 1 To mWindowPrefs.HistoryCount
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
      Dim sTemp As String * 255
      Dim fntTemp As New StdFont
      Dim iBookm As Integer, iHistIndex As Integer
      Dim sEx As String
      
      Dim udtWindowPlacement As WINDOWPLACEMENT
      Dim rectRestored As RECT
      
      lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, msPhlegmKey, 0, KEY_QUERY_VALUE, lKey)
      
      lValueSize = LenB(mWindowPrefs)
      lRetVal = RegQueryValueExAny(lKey, "Settings", 0, lDataType, ByVal mWindowPrefs, lValueSize)
      If lRetVal = 0 Then
            With mWindowPrefs
                  mfSkipFormResize = True
                  BrowserResizeHorizontal .BrowserWidth
                  
                  .WNP.Length = LenB(.WNP)
                  SetWindowPlacement hwnd, .WNP
                  
                  lvwBrowser.SortOrder = .SortMethod
                  If agEditor.tag = "" Then agEditor.tag = Trim(CstringToVBstring(.AutoLoadFile))
                  
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
      
                  lvwBrowser.ColumnHeaders.Item("Name").Position = .NameColumn
                  lvwBrowser.ColumnHeaders.Item("Type").Position = .TypeColumn
                  lvwBrowser.ColumnHeaders.Item("Size").Position = .SizeColumn
                  lvwBrowser.ColumnHeaders.Item("Modified").Position = .ModifiedColumn
                  lvwBrowser.SortKey = .SortKey
                  
                  cboPath = Trim(CstringToVBstring(.cboPath))
      
                  chkFileBrowser.Value = -CInt(.ShowFileBrowser)
                  chkFileBrowser_Click
                  'picBrowser.Visible = .ShowFileBrowser
                 ' mnuViewFilebrowser.Checked = .ShowFileBrowser
                  
                  staTusBar1.Visible = .ShowStatusBar
                  mnuViewStatusBar.Checked = .ShowStatusBar
                  picToolBar.Visible = .ShowToolBar
                  mnuViewToolbar.Checked = .ShowToolBar
                  If Not .ShowToolBar Then
                        mnuPlus.Caption = "+"
                        mnuNext.Visible = True
                        mnuPrev.Visible = True
                  End If
                  
                  If .ShowToolBar Then picQuery.Visible = .ShowFind
                  mfHideFind = Not .ShowFind
            
                  mfSkipFormResize = False
                  RearrangeControls
            End With
      Else
            cboPath = ""
      End If
      
'      If Trim(Command()) = "" Then ' no command line argument was given
            lValueSize = LenB(mEditorPrefs)
            lRetVal = RegQueryValueExAny(lKey, "agEditor", 0, lDataType, ByVal mEditorPrefs, lValueSize)
            If lRetVal = 0 Then
                  With mEditorPrefs
                        chkWordWrap.Value = .WordWrap
                        chkWordWrap_Click
                        chkReadOnly.Value = .ReadOnly
                        chkReadOnly_Click
                  
                        fntTemp.Name = Trim(CstringToVBstring(.FontName))
                        fntTemp.Size = .FontSize
                        fntTemp.Bold = .FontBold
                        fntTemp.Italic = .FontItalic
                        fntTemp.Strikethrough = .FontStrikethrough
                        fntTemp.Underline = .FontUnderline
                        If Len(Trim(.FontName)) > 11 Then
                              btnFont.Caption = Left(Trim(.FontName), 10) & "..."
                        Else
                              btnFont.Caption = Trim(.FontName)
                        End If
                        lblFontSize = Round(.FontSize, 1)
                        SetRealStdFont agEditor, fntTemp, .TextColor
                        
                        ' It's important to set the above prior to loading a file.
                        ' Otherwise agEditor's display routines are called again and again for an entire file,
                        ' rather than for a blank editor.
                        
                        If FileSize(agEditor.tag) > BIGFILESIZE Then
                              EditorLoadFile agEditor.tag, Properties
                        Else
                              sEx = gFSO.getextensionname(agEditor.tag)
                              EditorLoadFile agEditor.tag, FileTypeFromExtension(sEx)
                        End If
                        
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

Private Sub txtReplace_Change()
      If mfReplaceMode And txtFind = agEditor.SelectedText And txtFind <> "" Then
            btnReplace.Enabled = True
      ElseIf mfReplaceMode Then
            btnReplace.Enabled = False
      End If
End Sub


Private Sub txtReplace_GotFocus()
      If btnReplace.Enabled Then
            btnReplace.Default = True
      Else 'If btnReplace.Enabled Then
            btnFindNext.Default = True
      End If
End Sub


Private Sub txtReplace_KeyDown(KeyCode As Integer, Shift As Integer)
'      If KeyCode = vbKeyReturn And Shift = 0 Then
'            If btnReplace.Default Then btnReplace_Click
'      End If
End Sub


Private Sub txtReplace_LostFocus()
      btnReplace.Default = False
      btnFindNext.Default = False
End Sub


