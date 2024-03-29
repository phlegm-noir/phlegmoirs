VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DD32A320-6E5E-44C8-BCE6-5908CA400231}#1.0#0"; "agRichEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "(New File)"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11760
   ForeColor       =   &H80000005&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   6960
      Left            =   3000
      ScaleHeight     =   6960
      ScaleWidth      =   8535
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   840
      Width           =   8535
      Begin TabDlg.SSTab sstProperties 
         Height          =   6375
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   11245
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
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
         TabPicture(0)   =   "frmMain.frx":0CCA
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
            TabIndex        =   46
            Top             =   480
            Width           =   6135
            Begin VB.CommandButton btnOpenDefault 
               Caption         =   "&Open"
               Height          =   375
               Left            =   4440
               TabIndex        =   58
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
               TabIndex        =   57
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
               TabIndex        =   47
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
               TabIndex        =   49
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
               TabIndex        =   51
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
               TabIndex        =   53
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
               TabIndex        =   55
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
               TabIndex        =   48
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
               TabIndex        =   50
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
               TabIndex        =   52
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
               TabIndex        =   54
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
               TabIndex        =   56
               Top             =   2400
               Width           =   1125
            End
         End
         Begin VB.Frame fraID3 
            Caption         =   "ID3 tag info"
            Height          =   2415
            Left            =   240
            TabIndex        =   59
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
               TabIndex        =   67
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
               TabIndex        =   65
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
               TabIndex        =   63
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
               TabIndex        =   61
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
               TabIndex        =   64
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
               TabIndex        =   62
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
               TabIndex        =   66
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
               TabIndex        =   60
               Top             =   360
               Width           =   735
            End
         End
      End
      Begin agRichEditBox.agRichEdit agEditor 
         Height          =   5535
         Left            =   5520
         TabIndex        =   44
         Top             =   -240
         Visible         =   0   'False
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
         ViewMode        =   1
         TextLimit       =   9999999
         AutoURLDetect   =   0   'False
         TextOnly        =   -1  'True
         ScrollBars      =   0
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   0
         MousePointer    =   15  'Size All
         Stretch         =   -1  'True
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":124A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":169C
            Key             =   "textfile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AEE
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FBC
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
      ScaleWidth      =   5100
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   600
      Width           =   5100
      Begin MSComctlLib.ImageList ilsFileIcons2 
         Left            =   1875
         Top             =   3435
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   8388863
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":42D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4628
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":497A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4CCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":501E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5370
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":56C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5A14
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5D66
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":60B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":640A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":675C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6AAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6E00
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton btnScrollToTop 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   2380
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":7152
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Scroll To Top"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnSyncContents 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   2040
         MaskColor       =   &H80000001&
         Picture         =   "frmMain.frx":7608
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Jump to the directory containing your open file... (Ctrl+F5)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnDeleteSelected 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1700
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":794A
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Delete File (Del)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnRefresh 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1360
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":7C8C
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Refresh Files (F5)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnSort 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   1020
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":7FCE
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Reverse the sort order (Ctrl+H)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnFolderUp 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":8310
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Go up a directory (Left arrow key or Ctrl+F6)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin VB.CommandButton btnPathForward 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   340
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":8652
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Go forward a directory (Alt+Right)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
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
         ItemData        =   "frmMain.frx":8ACC
         Left            =   0
         List            =   "frmMain.frx":8ACE
         TabIndex        =   32
         Text            =   "*"
         ToolTipText     =   "Type a directory into here, or select one below.  You can even specify a file extension.  Example:   c:\windows\*.dll"
         Top             =   100
         Width           =   2295
      End
      Begin VB.CommandButton btnPathBack 
         Appearance      =   0  'Flat
         Height          =   340
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":8AD0
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Go back a directory (Alt+Left)"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   340
      End
      Begin MSComctlLib.ListView lvwBrowser 
         Height          =   4335
         Left            =   0
         TabIndex        =   40
         Tag             =   "c:\test\"
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7646
         SortKey         =   1
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsFileIcons2"
         SmallIcons      =   "ilsFileIcons2"
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
         MouseIcon       =   "frmMain.frx":8F4A
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
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "Size"
            Text            =   "Si[z]e"
            Object.Width           =   2293
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
      Begin VB.Label lblDivider 
         BackStyle       =   0  'Transparent
         Height          =   25005
         Left            =   2280
         TabIndex        =   41
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
      FontName        =   "MS Sans Serif"
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      ClipControls    =   0   'False
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   11700
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11760
      Begin VB.PictureBox picQuery 
         ClipControls    =   0   'False
         Height          =   600
         Left            =   4800
         ScaleHeight     =   540
         ScaleWidth      =   4035
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   -25
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton btnReplace 
            Appearance      =   0  'Flat
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
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMain.frx":90AC
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Replace (Ctrl+R)"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   375
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
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "More search options (Alt+period)"
            Top             =   0
            Width           =   375
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
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMain.frx":93EE
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Find Previous (Shift+F3)"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   1095
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
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMain.frx":9730
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Find Next (F3)"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtFind 
            Height          =   288
            Left            =   480
            MaxLength       =   50
            OLEDropMode     =   1  'Manual
            TabIndex        =   11
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
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Close Find Dialog"
            Top             =   0
            Width           =   175
         End
         Begin VB.TextBox txtReplace 
            Height          =   288
            Left            =   480
            MaxLength       =   50
            OLEDropMode     =   1  'Manual
            TabIndex        =   19
            ToolTipText     =   "Replace"
            Top             =   290
            Visible         =   0   'False
            Width           =   2175
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
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
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
            TabIndex        =   13
            Top             =   60
            Width           =   465
         End
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
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":9A72
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Next file down (Ctrl+""]"")"
         Top             =   0
         UseMaskColor    =   -1  'True
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
         MaskColor       =   &H80000005&
         Picture         =   "frmMain.frx":A174
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Next file up (Ctrl+""["")"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnZoomIn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3150
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":A876
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   260
         UseMaskColor    =   -1  'True
         Width           =   470
      End
      Begin VB.CommandButton btnZoomDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2700
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":ABB8
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Reset Zoom"
         Top             =   260
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   460
      End
      Begin VB.CommandButton btnFitImage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2250
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":AEFA
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Fit Image To Window"
         Top             =   260
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   460
      End
      Begin VB.CommandButton btnZoomOut 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1800
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":B23C
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   260
         UseMaskColor    =   -1  'True
         Width           =   460
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
         TabIndex        =   23
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
         Picture         =   "frmMain.frx":B57E
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnFont 
         Caption         =   "MS Sans Serif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Set Font (Shift+Ctrl+F)"
         Top             =   0
         Width           =   1815
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         Picture         =   "frmMain.frx":B9C0
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Slider sliZoom 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Adjust picture zoom"
         Top             =   0
         Width           =   1836
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   393216
         LargeChange     =   100
         SmallChange     =   10
         Max             =   500
         SelStart        =   100
         TickFrequency   =   100
         Value           =   100
      End
      Begin VB.CommandButton btnEdit 
         Appearance      =   0  'Flat
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
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":BE02
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Edit This File"
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Toggle Word Wrap (Ctrl+W)"
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnFullScreen 
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
         Height          =   570
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":C144
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Full Screen (F11)"
         Top             =   0
         UseMaskColor    =   -1  'True
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
         MaskColor       =   &H00000000&
         Picture         =   "frmMain.frx":C486
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Save File (Ctrl+S)"
         Top             =   0
         UseMaskColor    =   -1  'True
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
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":CB88
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "New File (Ctrl+N)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         DownPicture     =   "frmMain.frx":D28A
         Height          =   570
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":D98C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Show/Hide the File Browser (F8)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label lblFontSize 
         Alignment       =   2  'Center
         Caption         =   "8"
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
         Left            =   2240
         TabIndex        =   27
         Top             =   320
         Width           =   960
      End
   End
   Begin MSComctlLib.StatusBar staTusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   68
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
         Caption         =   "&Open (File Browser)"
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
      Begin VB.Menu mnuFileNext 
         Caption         =   "Open Next &File"
      End
      Begin VB.Menu mnuFilePrev 
         Caption         =   "Open &Previous File"
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
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
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
      Begin VB.Menu mnuViewFitImage 
         Caption         =   "&Always Fit Images"
         Checked         =   -1  'True
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
         Visible         =   0   'False
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

Option Compare Binary
Option Explicit

' *************************************************************
' General/Settings variables
' *************************************************************
Private mbSkipFormResize As Boolean
Private mbEditorLoading As Boolean

' *************************************************************
' File Browser related variables
' *************************************************************
Private mbBrowserDoubleClick As Boolean
Private mbBrowserItemClicked As Boolean
Private miBrowserMouseButton As Integer
Private miBrowserShift As Integer

' *************************************************************
' cboPath related variables
' *************************************************************
Private miPathRecent As Integer

' *************************************************************
' Find related variables
' *************************************************************
Private mbHideFind As Boolean
Private mbReplaceMode As Boolean
Private miFindResult As Integer
Private mlFirstResultPos As Long
Private mbFinding As Boolean
Private miTotalResults As Integer

Private Sub AddToBookmarks(ByVal sNewBookmark As String)
      Dim iIndex As Integer
      
      If mnuBookmark.UBound >= MAX_BOOKMARKS Then
            MsgBox "You've reached the " & MAX_BOOKMARKS & " bookmark limit. " & _
                  "Manage your bookmarks to make room.", , "Bookmark Limit"
            Exit Sub
      End If
      
      sNewBookmark = CstringToVBstring(sNewBookmark)
      If sNewBookmark = "" Then Exit Sub
     
      iIndex = mnuBookmark.UBound + 1
      Load mnuBookmark(iIndex)
      With mnuBookmark(iIndex)
            .tag = sNewBookmark  ' exact path here, for safe keeping
            .Caption = GetNumberedCaption(sNewBookmark, iIndex)
            .Visible = True
      End With
End Sub

Private Function AddToHistorySimply(ByVal sNewHistory As String) As String
      On Error GoTo SIMPLY_ERROR
      Dim iIndex As Integer

      sNewHistory = CstringToVBstring(sNewHistory)
      AddToHistorySimply = sNewHistory
      If sNewHistory = "" Then Exit Function
     
      iIndex = mnuFileHistory.UBound + 1
      Load mnuFileHistory(iIndex)
      With mnuFileHistory(iIndex)
            .tag = sNewHistory  ' exact path here, for safe keeping
            .Caption = GetNumberedCaption(sNewHistory, iIndex)
            .Visible = True
      End With
      Exit Function
SIMPLY_ERROR:
      DebugLog "      SIMPLY AN ERROR. Error: " & Err.Description, 2
      DebugLog "            New history: " & sNewHistory, 2
End Function

Private Sub AddToHistorySmartly(ByVal sNewHistory As String)
      Dim iIndex As Integer
      Dim sPrevTag As String, sTempTag As String
      Dim bFoundSame As Boolean, bHistoryGrew As Boolean

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
            bHistoryGrew = True
      End If
      
      ' What it SHOULD do:
      '   Put current file at the top.
      '   Start shifting the rest down.
      '   If current file was already in History, remove that one.
      '   Stop shifting.  Don't shift anything below that one.
      sPrevTag = mnuFileHistory(1).tag
      mnuFileHistory(1).tag = sNewHistory
      mnuFileHistory(1).Caption = "&1   " & mnuFileHistory(1).tag
      
      For iIndex = 2 To mnuFileHistory.UBound
            With mnuFileHistory(iIndex)
                  If .tag = sNewHistory Then
                        .tag = sPrevTag
                        bFoundSame = True
                  Else
                        sTempTag = .tag
                        .tag = sPrevTag
                        sPrevTag = sTempTag
                  End If
                  .Caption = GetNumberedCaption(.tag, iIndex)
                  
                  If bFoundSame Then Exit For
            End With
      Next iIndex
      
      If bFoundSame And bHistoryGrew Then Unload mnuFileHistory(mnuFileHistory.UBound)
      If gtBrowserData.HistoryMode Then RefreshAll
End Sub

Private Sub agEditor_Change()

      If Not mbEditorLoading And geEditorMode = eViewMode.TextView And Not (agEditor.tag = "" And agEditor.Text = "") Then
            staTusBar1.Panels(eStat.Modified) = "Modified"
      End If
      
      If staTusBar1.Visible Then
            With gtStats
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

Private Sub agEditor_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                  If Shift = 0 And chkReadOnly.Value = vbChecked Then BrowserExecuteNext
            Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                  If Shift = 0 And chkReadOnly.Value = vbChecked Then BrowserExecuteNext True
            Case vbKeyM
                  If Shift = vbCtrlMask Then
                        mnuBookmarksManage_Click
                  End If
            Case vbKeyF
                  If Shift = vbCtrlMask + vbShiftMask Then
                        btnFont_Click
                  End If
      End Select
End Sub

Private Sub ageditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If FOCUS_FOLLOWS_MOUSE Then
            On Error Resume Next
            If GetForegroundWindow = frmMain.hWnd And Not (ActiveControl.Name = "agEditor") And _
                  Not ActiveControl.Name = "txtFind" And Not ActiveControl.Name = "txtReplace" Then
                  agEditor.SetFocus
            End If
            On Error GoTo 0
      End If
      
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

Private Sub agEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If (Button = vbRightButton And Shift = 0) Then
            Me.PopupMenu mnuWrite
      End If
End Sub

Private Sub agEditor_ProgressStatus(ByVal lAmount As Long, ByVal lTotal As Long)
      ' Debug.Print "PROGRESS: "; lAmount & " " & lTotal

      ' TODO: if a second file is told to load, it cancels this one but won't remove it from the editor first.
      
      DoEvents
End Sub

Private Sub agEditor_SelectionChange(ByVal lMin As Long, ByVal lMax As Long, ByVal oFoundSame As agricheditbox.ERECSelectionTypeConstants)
      ' Update a few items on the status bar.
      
      Dim lLineIndex As Long, lCharIndex As Long
      Dim tSelection As CHARRANGE
      
      lLineIndex = agEditor.CurrentLine
      lCharIndex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lLineIndex, 0)
      
      If staTusBar1.Visible Then
            With gtStats
                  .Y = lLineIndex + 1
                
                  ' We want gtStats.i to count CRs and LFs both, since agEditor.CharacterCount does that.
                  .i = lMin
                  SendMessage agEditor.RichEdithWnd, EM_EXGETSEL, 0, tSelection
                  .X = lMin - lCharIndex + 1
                  .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
            End With
        
            FillStats
            staTusBar1.Panels(eStat.SelText) = lMax - lMin
      End If
      
      If mbReplaceMode And StrComp(txtFind, agEditor.SelectedText, GetFindCompareMode()) = 0 And txtFind <> "" Then
            If Not chkReadOnly Then btnReplace.Enabled = True
            If ActiveControl.Name = "txtReplace" Then btnReplace.Default = True
      ElseIf mbReplaceMode Then
            btnReplace.Enabled = False
      End If
      
      ' Reset Find result count whenever the selection changes
      ' (...changes from something other than inside a Find)
      If Not mbFinding Then
            miFindResult = 0
            miTotalResults = 0
            lblFindResult = ""
      End If
End Sub

Private Sub AutosizeColumns()
      If AUTOSIZE_COLUMNS And gtBrowserData.DoneLoading Then
            SendMessage lvwBrowser.hWnd, LVM_SETCOLUMNWIDTH, ByVal 1, LVSCW_AUTOSIZE
            SendMessage lvwBrowser.hWnd, LVM_SETCOLUMNWIDTH, ByVal 2, LVSCW_AUTOSIZE
            SendMessage lvwBrowser.hWnd, LVM_SETCOLUMNWIDTH, ByVal 3, LVSCW_AUTOSIZE
            If lvwBrowser.ColumnHeaders.item("Type").Width <= COLUMN_TOO_SMALL Then
                  lvwBrowser.ColumnHeaders.item("Type").Width = COLUMN_TOO_SMALL
            End If
            If lvwBrowser.ColumnHeaders.item("Size").Width <= COLUMN_TOO_SMALL Then
                  lvwBrowser.ColumnHeaders.item("Size").Width = COLUMN_TOO_SMALL
            End If
            If lvwBrowser.ColumnHeaders.item("Modified").Width <= COLUMN_TOO_SMALL Then
                  lvwBrowser.ColumnHeaders.item("Modified").Width = COLUMN_TOO_SMALL
            End If
      End If
End Sub

Private Sub BookmarkSaveChanges()
      Dim iIndex As Integer
      
      For iIndex = 1 To lvwBrowser.ListItems.Count
            With mnuBookmark(iIndex)
                  .tag = lvwBrowser.ListItems(iIndex)
                  .Caption = GetNumberedCaption(.tag, iIndex)
            End With
      Next iIndex
      
      For iIndex = iIndex To mnuBookmark.UBound
            Unload mnuBookmark(iIndex)
      Next iIndex
      SaveSettingsToRegistry
End Sub

Private Function BrowserAutoSelectListItem(ByRef rtBdata As TBrowserData)
      Dim oCurrentItem As ListItem
      
      If rtBdata.ListEmpty Or rtBdata.BookmarkMode Then Exit Function
      
      If rtBdata.PartialFileName <> "" Then
            ' Auto-select first filename to match partialfilename, if given.
            Set oCurrentItem = lvwBrowser.FindItem(rtBdata.PartialFileName, , , lvwPartial)
            If Not (oCurrentItem Is Nothing) Then oCurrentItem.Selected = True
      
            
      ElseIf rtBdata.GoingToParent Then
            ' Auto-select the directory we just moved out of, if doing a ParentDirectory.
            Set oCurrentItem = lvwBrowser.FindItem(goFso.GetBaseName(rtBdata.DirPrev))
            If Not (oCurrentItem Is Nothing) Then oCurrentItem.Selected = True
            
      
      ElseIf rtBdata.DirUnchanged Then
            ' Auto-select the item previously selected, for a refresh.
            Set oCurrentItem = lvwBrowser.FindItem(rtBdata.SelTextPrev)

            If (oCurrentItem Is Nothing) Then
                  lvwBrowser.ListItems(1).Selected = True
            Else
                  oCurrentItem.Selected = True
            End If
            
      Else ' Otherwise, auto-select the first item.
            lvwBrowser.ListItems(1).Selected = True
      End If
                  
      DoEvents ' Just doesn't seem to work without DoEvents first.
      If Not gbFullScreenMode And Not (lvwBrowser.SelectedItem Is Nothing) Then
            lvwBrowser.SelectedItem.EnsureVisible
      End If
End Function

Private Sub BrowserDeleteSelected()
      Dim sBookKey As String, iRetVal As Integer
      Dim sTheDamned As String
      
      ' No deletion of history.  If you'd like to delete a file you see in the history,
      ' do it some other way like by opening 10 more unique files.
      
      If lvwBrowser.ListItems.Count = 0 Or gtBrowserData.HistoryMode Then Exit Sub
      
      sTheDamned = gtBrowserData.Dir & lvwBrowser.SelectedItem
      
      If gtBrowserData.BookmarkMode Then
            sBookKey = lvwBrowser.SelectedItem.Key
            lvwBrowser.ListItems.Remove sBookKey
            BookmarkSaveChanges
            Exit Sub
            
      ElseIf gtBrowserData.DrivesMode Then
            frmMain.Caption = "I WILL NOT DELETE YOUR DISK. FIND SOMEONE ELSE."
            Exit Sub
      
      ElseIf Not FileExists(sTheDamned) Then
            frmMain.Caption = "Can't delete what isn't there: " & sTheDamned
            Exit Sub
      End If
      
      On Error GoTo DELETION_ERROR
            
      Dim oAttrs
      oAttrs = GetAttr(sTheDamned)
      
      If oAttrs And vbDirectory Then
            frmMain.Caption = "This program would rather not be held responsible for mass deletions. Please use another."
            Exit Sub
            ' RmDir sTheDamned
            ' frmMain.Caption = "Folder deleted successfully: " & sTheDamned
            ' RefreshAll
      End If

      iRetVal = RecycleFile(sTheDamned)
      If iRetVal <> 0 Then
            frmMain.Caption = "ERROR deleting file. Return code: " & iRetVal
            DebugLog Caption
      Else
            If sTheDamned = agEditor.tag Then
                  agEditor.tag = ""
                  mnuFileNew_Click
            End If
            frmMain.Caption = "File deleted successfully: " & sTheDamned
            RefreshAll
      End If
      Exit Sub

DELETION_ERROR:
      frmMain.Caption = "ERROR deleting file: " & Err.Description
      DebugLog Caption

End Sub

Private Sub BrowserExecuteItem(ByVal oItem As MSComctlLib.ListItem)
      If (lvwBrowser.ListItems.Count = 0) Then Exit Sub
      
      Dim sItemName As String
      sItemName = gtBrowserData.Dir & oItem.Text
      
      Select Case oItem.Icon
      
            Case eIconType.Directory, eIconType.Drive, eIconType.Floppy, eIconType.Cdrom, eIconType.Network
                  ' Open the folder, or go up a folder.
                  If oItem.Text = ".." Then
                        btnFolderUp_Click
                  Else
                        cboPath = sItemName & "\"
                  End If
            
            Case eIconType.Bookmark
                  LoadFile oItem.Text, GetViewMode(oItem.Text, oItem.Icon)
                  
            Case Else
                  LoadFile sItemName, GetViewMode(sItemName, oItem.Icon)
      End Select
End Sub

'   BrowserExecuteNext
'   Select the next item after the selection, and open it.
'
Public Sub BrowserExecuteNext(Optional ByVal Reverse As Boolean = False)
      DoEvents
      Dim iIndex As Integer, oNext As ListItem
      Dim iInc As Integer, iLimit As Integer
      Dim eMode As eViewMode

      ' Open the item next to the open file, not next to whatever thing is selected.
      ' It is possible to select something else via arrow keys or right-click + cancel.
      ' So we start with a sync.
      If (agEditor.tag <> "") And (Not gtBrowserData.BookmarkMode) And (Not gtBrowserData.HistoryMode) Then
            btnSyncContents_Click
      End If
      
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      If Reverse Then
            iInc = -1
            iLimit = 1
      Else
            iInc = 1
            iLimit = lvwBrowser.ListItems.Count
      End If
      iIndex = lvwBrowser.SelectedItem.Index
      
      ' Fullscreen mode is for image view only.
      ' So want the next item to be the next image, skipping over the rest.
      If gbFullScreenMode Then
            Do While iIndex <> iLimit
                  iIndex = iIndex + iInc
                  Set oNext = lvwBrowser.ListItems(iIndex)
                  eMode = GetViewMode(cboPath.Text & oNext.Text, oNext.Icon)
                  If eMode = eViewMode.PictureView Then
                        oNext.Selected = True
                        BrowserExecuteItem oNext
                        Exit Do
                  End If
            Loop
      Else
            If iIndex <> iLimit Then
                  Set oNext = lvwBrowser.ListItems(iIndex + iInc)
                  If oNext.Icon <> eIconType.Directory Then
                        oNext.EnsureVisible
                        oNext.Selected = True
                        BrowserExecuteItem oNext
                  End If
            End If
      End If
End Sub

Private Sub BrowserGetBookmarks()
      Dim iIndex As Integer
      Dim oCurrentItem As ListItem
      
      lvwBrowser.Visible = False
      lvwBrowser.ListItems.Clear
      lvwBrowser.tag = "(Bookmarks)"
      ' I'm adding the index as a Key, to avoid using real indeces.
      ' (So that I can use API functions that desynchronize listitem indexing.)
      ' Edit: I'm not really doing that. Using bookmarks as a test case on whether that might be doable.
      For iIndex = 1 To mnuBookmark.UBound
            Set oCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuBookmark(iIndex).tag, _
                  eIconType.Bookmark, eIconType.Bookmark)
            oCurrentItem.ListSubItems.Add 1, , goFso.getextensionname(mnuBookmark(iIndex).tag)
      Next iIndex
      gtBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      AutosizeColumns
      lvwBrowser.Visible = True
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " bookmarks"
End Sub

Private Function BrowserGetDrives() As Integer
      ' Find all logical drives and display them in lvwBrowser
      ' Returns the number of logical drives found.
      
      Dim sDrivesFixed As String * 255
      Dim sDriveString As String
      Dim sDriveArray() As String
      Dim sNextDrive As String, eDriveIcon As eIconType
      Dim lLength As Long
      Dim iIndex As Integer, iTempKey As Integer
      Dim oCurrentItem As ListItem
      
      lvwBrowser.Visible = False
      
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
            
            Select Case goFso.getdrive(sNextDrive).drivetype
                  Case 1: eDriveIcon = Floppy
                  Case 2: eDriveIcon = Drive
                  Case 3: eDriveIcon = Network
                  Case 4: eDriveIcon = Cdrom
            End Select
            Set oCurrentItem = lvwBrowser.ListItems.Add(1, , sNextDrive, eDriveIcon, eDriveIcon)
            
            iIndex = iIndex + 1
            sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      Loop
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = iTempKey
      BrowserGetDrives = iIndex - 1
      AutosizeColumns
      lvwBrowser.Visible = True
      
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " drives"
End Function

Private Sub BrowserGetFilesAndFolders(ByRef rtBdata As TBrowserData)
      DebugLog "Gonna load some files and folders at: " & rtBdata.Dir
      Dim eIcon As eIconType, iTempKey As Integer
      Dim oTotalBytes
      Dim oCurrentItem As ListItem
      Dim lNextFile As Long, sFileName As String, sEx As String
      Dim tWfd As WIN32_FIND_DATA
      Dim bHadFocus As Boolean
      Dim sErrorMsg As String
      
      On Error Resume Next
      bHadFocus = (ActiveControl.Name = "lvwBrowser")
      On Error GoTo BROWSER_LOAD_FILES_ERROR
      
      lvwBrowser.tag = rtBdata.Dir
      
      lvwBrowser.Visible = False
      lvwBrowser.ListItems.Clear
      iTempKey = lvwBrowser.SortKey
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      If rtBdata.Filter = "" Then rtBdata.Filter = "*"
      lNextFile = FindFirstFile(rtBdata.Dir & rtBdata.Filter, tWfd)
      oTotalBytes = 0
      
      Do
            On Error Resume Next
            
            ' Divide the file types for icon selection
            
            sFileName = CstringToVBstring(tWfd.cFileName) ' Lots of junk past the null character.
            sEx = goFso.getextensionname(sFileName)
            
            If (tWfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                  eIcon = eIconType.Directory
            Else
                  eIcon = GetIconType(sEx)
            End If

            If Err > 0 Then
                  eIcon = eIconType.IconError
                  DebugLog "Icon error for file: " & sFileName & ": " & Err & ": " & Err.Description, 2
            End If
            
            Dim oSizeBig ' this can be bigger than a long integer
            oSizeBig = 0
            oSizeBig = goFso.getfile(rtBdata.Dir + sFileName).Size ' size=0 if error
            
            ' Add that file!
            
            On Error GoTo BROWSER_LOAD_FILES_ERROR
            If sFileName <> "." And sFileName <> "" Then ' what would be the point in providing a "." folder?
                  Set oCurrentItem = lvwBrowser.ListItems.Add(, , sFileName, eIcon, eIcon)
                  
                  ' here, let's keep an invisible second column for sorting by directory later
                  If eIcon = eIconType.Directory Then
                        'oCurrentItem.ListSubItems.Add iFileTypeHeader, "Type", ""
                  Else
                        oCurrentItem.ListSubItems.Add , "Type", sEx
                        oCurrentItem.ListSubItems.Add , "Size", Format(oSizeBig, "#,#0")
                        oCurrentItem.ListSubItems.Add , "Modified", FormatNonLocalFileTime(tWfd.ftLastWriteTime)
                        oCurrentItem.ListSubItems.Add , "SortSize", Format(oSizeBig, "00000000000000000")
                        oTotalBytes = oTotalBytes + oSizeBig
                  End If
            End If
      
      Loop While FindNextFile(lNextFile, tWfd) <> 0
           
      
      If rtBdata.Filter = "*" Then rtBdata.Filter = ""
      rtBdata.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = iTempKey
      
      AutosizeColumns
      lvwBrowser.Visible = True
      If bHadFocus Then lvwBrowser.SetFocus
      
      staTusBar1.Panels(eStat.BrowserStats).Text = FormatBytes(oTotalBytes, 1) & " in " & _
            (lvwBrowser.ListItems.Count - 1) & " objects"  ' -1 for the ".." folder
      Exit Sub

BROWSER_LOAD_FILES_ERROR:
      lvwBrowser.Visible = True
      sErrorMsg = "BrowserGetFilesAndFolders error: " & Err.Description
      DebugLog sErrorMsg, 2
      MsgBox sErrorMsg
      Exit Sub
End Sub

Private Sub BrowserGetHistory()
      Dim iIndex As Integer
      Dim oCurrentItem As ListItem
      
      lvwBrowser.Visible = False
      lvwBrowser.ListItems.Clear
      lvwBrowser.tag = "(History)"
      lvwBrowser.Sorted = False
      
      For iIndex = 1 To mnuFileHistory.UBound
            Set oCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuFileHistory(iIndex).tag, _
                  eIconType.Bookmark, eIconType.Bookmark)
            oCurrentItem.ListSubItems.Add 1, , goFso.getextensionname(mnuFileHistory(iIndex).tag)
      Next iIndex
      
      gtBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
      If Not gtBrowserData.ListEmpty Then AutosizeColumns
      lvwBrowser.Visible = True
      staTusBar1.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " most recent files"
End Sub

Private Function BrowserResizeHorizontal(ByVal iSupposedWidth As Integer) As Integer
      ' This is like a miniature RearrangeControls() for just picBrowser and everything within,
      ' and it happens to only affect their horizontal components.
      
      ' The return value is the difference (in twips) that picBrowser has grown.  Can be negative.
      
      Dim iOffset As Integer, iRealWidth As Integer, iScrollButtonX As Integer
      Const RIGHT_MARGIN = 117
      
      If iSupposedWidth < 1000 Then
            iRealWidth = 1000
      ElseIf picBrowser.Left + iSupposedWidth + 1500 > frmMain.ScaleWidth Then
            iRealWidth = frmMain.ScaleWidth - picBrowser.Left - 1500
      Else
            iRealWidth = iSupposedWidth
      End If
       
      iOffset = iRealWidth - picBrowser.Width
      
      picBrowser.Width = iRealWidth
      lvwBrowser.Width = iRealWidth - RIGHT_MARGIN
      lblDivider.Left = lvwBrowser.Width
      cboPath.Width = iRealWidth - RIGHT_MARGIN
      
      iScrollButtonX = lvwBrowser.Left + lvwBrowser.Width - btnScrollToTop.Width - 30
      If btnSyncContents.Left + btnSyncContents.Width <= iScrollButtonX Then
            btnScrollToTop.Left = iScrollButtonX
      Else
            btnScrollToTop.Left = btnSyncContents.Left + btnSyncContents.Width
      End If
      
      BrowserResizeHorizontal = iOffset
End Function

Private Sub btnCloseFind_Click()
      mnuQueryClose_Click
End Sub

Private Sub btnDeleteSelected_Click()
      BrowserDeleteSelected
End Sub

Private Sub btndeleteselected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnDeleteSelected.ToolTipText
End Sub

Private Sub btnEdit_Click()
      mnuViewReadOnly_Click
End Sub

Private Sub btnfileback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFileBack.ToolTipText
End Sub

Private Sub btnfileforward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFileForward.ToolTipText
End Sub

Private Sub btnFindNext_Click()
      If geEditorMode = eViewMode.PictureView Or geEditorMode = eViewMode.PropertiesView Then Exit Sub
      
      If txtFind = "" Then txtFind = agEditor.SelectedText
      
      Dim lFoundMin As Long, lFoundMax As Long, lStartMin As Long, lStartMax As Long
      Dim lFindRetval As Long
      
      agEditor.GetSelection lStartMin, lStartMax
      
      lFindRetval = EditorFindText(txtFind, Forward, lStartMax, _
            agEditor.CharacterCount, lFoundMin, lFoundMax)
      
      If lFindRetval = -1 Then
            ' Nothing found downward.  Search from beginning.
            lFindRetval = EditorFindText(txtFind, Forward, 0, _
                  lStartMax, lFoundMin, lFoundMax)
      End If
            
      If lFindRetval > -1 Then
            ' Found something!
            mbFinding = True ' make sure the find count doesn't reset when we highlight a find result!
            agEditor.SetSelection lFoundMin, lFoundMax
            mbFinding = False
            
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
End Sub

Private Sub btnFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFindNext.ToolTipText
End Sub

Private Sub btnFindPrev_Click()
      ' So I've decided that it's possible to have negative numbers of find results.
      ' This is what happens when you click "Find Previous",
      ' and there wasn't a previous find, but there is a match.
      ' We can't just call it #N, where N is the total number of matches in the document,
      ' because we haven't searched the entire document!  That would take too long.
      ' So instead, it's #-1.
      
      ' No searching text within a picture or a properties tab.
      If geEditorMode = eViewMode.PictureView Or geEditorMode = PropertiesView Then Exit Sub
      
      If txtFind = "" Then txtFind = agEditor.SelectedText
      
      Dim lFoundMin As Long, lFoundMax As Long, lStartMin As Long, lStartMax As Long
      Dim lFindRetval As Long
      
      agEditor.GetSelection lStartMin, lStartMax
      
      lFindRetval = EditorFindText(txtFind, back, lStartMin, 0, lFoundMin, lFoundMax)
      
      If lFindRetval = -1 Then
            ' Nothing found upward.  Search from end of file.
            lFindRetval = EditorFindText(txtFind, back, agEditor.CharacterCount, _
                  lStartMin, lFoundMin, lFoundMax)
      End If
            
      If lFindRetval > -1 Then
            ' Found something!
            mbFinding = True ' make sure the find count doesn't reset when we highlight a find result!
            agEditor.SetSelection lFoundMin, lFoundMax
            mbFinding = False
            
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

Private Sub btnfindprev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFindPrev.ToolTipText
End Sub

Private Sub btnFitImage_Click()
      ImageZoomFit gtImageData.OutPic.Picture, agEditor.tag
End Sub

Private Sub btnFolderUp_Click()
      ' When we go up a dir, preserve the existing filter except in a drives list.
      Dim sParentDir As String
      
      If gtBrowserData.DrivesMode Or gtBrowserData.BookmarkMode Then Exit Sub
      
      sParentDir = ParentDirectoryOf(gtBrowserData.Dir)
      
      If gtBrowserData.ERROR Or sParentDir = "" Then
            cboPath = sParentDir
      Else
            cboPath = sParentDir & gtBrowserData.Filter
      End If
End Sub

Private Sub btnfolderup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFolderUp.ToolTipText
End Sub

Private Sub btnFont_Click()
      Dim oTempFont As New StdFont ' StdFont is a Class
      Dim lRetVal As Long, lTextColor As Long
      Const CF_SCREEN_FONTS As Long = &H1
      Const CF_SCALABLE_ONLY As Long = &H20000
      Const CF_EFFECTS As Long = &H100
      
      Set oTempFont = GetRealStdFont(agEditor.RichEdithWnd, lTextColor)
      
      'make the dialog choices begin with what the agEditor shows
      With dlgFont
            .Flags = CF_SCREEN_FONTS + cdlCFApply + CF_EFFECTS ' btw, Apply doesn't work
            .FontName = oTempFont.Name
            .FontBold = oTempFont.Bold
            .FontUnderline = oTempFont.Underline
            .FontSize = oTempFont.Size  ' one uses Single, the other Currency
            .FontStrikethru = oTempFont.Strikethrough
            .Color = lTextColor
      End With

      On Error Resume Next 'trap the error. if they hit cancel, do nothing and exit
      dlgFont.ShowFont
      If Err.Number = cdlCancel Then Exit Sub
      On Error GoTo 0
      
      With oTempFont
            .Name = dlgFont.FontName
            ' If you set a font name, you set a charset (automatically). Same for weight.
            ' agRichEdit's SetFont method does not pass the charset on to the rich edit control.
            
            ' It probably uses a CHARFORMAT2, and neglects to give its dwMask property the CFM_CHARSET flag.
            ' So that even if it did set the bCharset property to the stdfont.charset value,
            ' it would not have been seen.
            
            ' And it assumes charset = 0, which is true for most fonts.
            ' That's why it wouldn't work (until now) with symbol fonts, which have charset = 2.
            
            .Bold = dlgFont.FontBold
            .Italic = dlgFont.FontItalic
            .Underline = dlgFont.FontUnderline
            .Strikethrough = dlgFont.FontStrikethru
            .Size = CCur(dlgFont.FontSize)
            ' Weight is set automatically. (It seems that) 400 = plain, 700 = bold.
      End With
      'agEditor.SetFont oTempFont, , , , ercSetFormatAll <-- the simple call that doesn't work looks like this
      lRetVal = SetRealStdFont(agEditor.RichEdithWnd, oTempFont, dlgFont.Color)
      
      btnFont.Caption = GetRealStdFont(agEditor.RichEdithWnd).Name
      If Len(btnFont.Caption) > 18 Then
            btnFont.Caption = Left(Trim(btnFont.Caption), 17) & "..."
      End If
      lblFontSize = Round(GetRealStdFont(agEditor.RichEdithWnd).Size, 0)
      
      SaveSettingsToRegistry ' losing your font setting is so annoying; save them NOW!
End Sub

Private Sub btnfont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnFont.ToolTipText
End Sub

Private Sub btnFullScreen_Click()
      Hide
      frmFullScreen.Show
End Sub

Private Sub btnNewFile_Click()
      mnuFileNew_Click
End Sub

Private Sub btnnewfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnNewFile.ToolTipText
End Sub

Private Sub btnNextFile_Click()
      BrowserExecuteNext
End Sub

Private Sub btnnextfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnNextFile.ToolTipText
End Sub

Private Sub btnPathBack_Click()
      PathBack
End Sub

Private Sub btnpathback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPathBack.ToolTipText
End Sub

Private Sub btnPathForward_Click()
      PathForward
End Sub

Private Sub btnpathforward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPathForward.ToolTipText
End Sub

Private Sub btnPrevFile_Click()
      BrowserExecuteNext True
End Sub

Private Sub btnprevfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnPrevFile.ToolTipText
End Sub

Private Sub btnRefresh_Click()
      RefreshAll
      If agEditor.tag = "" Then
            frmMain.Caption = "(New File)"
      Else
            frmMain.Caption = agEditor.tag & "  (" & Format(GetFileSize(agEditor.tag), "#,#0") & " bytes saved on " _
                  & FileModifiedTime(agEditor.tag) & ")"
      End If
End Sub

Private Sub btnrefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnRefresh.ToolTipText
End Sub

Private Sub btnReplace_Click()
      If Not mbReplaceMode Then
            ' The replace button puts us in replace mode if we aren't already there.
            mnuQueryReplace_Click

      ElseIf btnReplace.Enabled And Not chkReadOnly Then
            ' Otherwise, if we were already in replace mode, it replaces (if legal).
            agEditor.InsertContents SF_TEXT, txtReplace
            btnFindNext_Click
      End If
End Sub

Private Sub btnReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnReplace.ToolTipText
End Sub

Private Sub btnSave_Click()
      mnuFileSave_Click
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnSave.ToolTipText
End Sub

Private Sub btnScrollToTop_Click()
      If lvwBrowser.ListItems.Count > 0 Then lvwBrowser.ListItems(1).EnsureVisible
End Sub

Private Sub btnScrolltotop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnScrollToTop.ToolTipText
End Sub

Private Sub btnSort_Click()
      Dim iTempKey As Integer
      
      ' List remains sorted at all times.  Only the order can be reversed.
      
      If gtBrowserData.HistoryMode Then Exit Sub
      
      With lvwBrowser
            .Sorted = True
            iTempKey = .SortKey
            .SortKey = 0
            .SortOrder = Abs(.SortOrder - 1)
            .SortKey = iTempKey
      End With
            
      If gtBrowserData.BookmarkMode Then BookmarkSaveChanges
End Sub

Private Sub btnSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnSort.ToolTipText
End Sub

Private Sub btnSyncContents_Click()
      
      ' What this really does is:
      '     1. go to directory containing open file
      '     2. select open file from list

      Dim oCurrentFile As ListItem
      
      If agEditor.tag = "" Then Exit Sub
      
      Set oCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.tag))
      
      If oCurrentFile Is Nothing Then
            cboPath = SnipFileName(agEditor.tag)
            Set oCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.tag))
            If oCurrentFile Is Nothing Then
                  MsgBox "It seems that your file was deleted by another application." & _
                        "  If you wish to keep it, save at once!"
                  Exit Sub
            End If
      End If
      oCurrentFile.Selected = True
      If Not gbFullScreenMode Then oCurrentFile.EnsureVisible
End Sub

Private Sub btnSyncContents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = btnSyncContents.ToolTipText
End Sub

Private Sub btnToolbarClose_Click()
      mnuViewToolbar_Click
End Sub

Private Sub btnZoomDefault_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ImageSetZoom 100
      Image1.Move 0, 0, gtImageData.DefaultWidth, gtImageData.DefaultHeight
End Sub

 
Private Sub btnZoomIn_Click()
      Select Case geEditorMode
      
            Case eViewMode.PictureView
                  ' Go to the next zoom divisible by the zoom increment
                  If sliZoom.Value < 100 Then
                        ImageZoomIn 25
                  Else
                        ImageZoomIn sliZoom.LargeChange
                  End If
                  
            Case Else  ' Increase the Font Size
                  Dim iFontSize As Integer
                  iFontSize = CInt(GetRealFontSize(agEditor.RichEdithWnd))
                  iFontSize = SetRealFontSize(agEditor.RichEdithWnd, GetNextFontSize(iFontSize))
                  lblFontSize = iFontSize
      End Select
End Sub

Private Sub btnZoomOut_Click()
      Select Case geEditorMode
            
            Case eViewMode.PictureView
                  ' Go to the next lowest zoom % divisible by the zoom increment
                  If sliZoom.Value <= 100 Then
                        ImageZoomOut 25
                  Else
                        ImageZoomOut sliZoom.LargeChange
                  End If
            
            Case Else   ' Decrease the Font Size
                  Dim iFontSize As Integer
                  iFontSize = CInt(GetRealFontSize(agEditor.RichEdithWnd))
                  iFontSize = SetRealFontSize(agEditor.RichEdithWnd, GetPrevFontSize(iFontSize))
                  lblFontSize = iFontSize
      End Select
End Sub

Private Sub cboPath_Change()
      
      ParsePath cboPath, gtBrowserData
      
      If gtBrowserData.BookmarkMode Then
            BrowserGetBookmarks
            PathAddRecent "(Bookmarks)"
      
      ElseIf gtBrowserData.HistoryMode Then
            BrowserGetHistory
            PathAddRecent "(History)"
      
      ElseIf gtBrowserData.DrivesMode Then
            BrowserGetDrives
            PathAddRecent ""
            
      ElseIf Not (gtBrowserData.DirUnchanged And gtBrowserData.FilterUnchanged) Then
            BrowserGetFilesAndFolders gtBrowserData
            ' Add to recent paths only if filtration was fruitful.
            If Not gtBrowserData.ListEmpty Then PathAddRecent gtBrowserData.Dir & gtBrowserData.Filter
      End If

      BrowserAutoSelectListItem gtBrowserData
End Sub

Private Sub cboPath_Click()
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

Private Sub cboPath_GotFocus()
      ' When focus is obtained, put the cursor right where we would have moved it anyway:
      ' At the end of the path, before the extension if one exists.
      
      If cboPath <> "(Bookmarks)" And cboPath <> "(History)" Then
            
            Dim iExtensionLength As Integer
            
            iExtensionLength = Len(goFso.getextensionname(cboPath))
            If iExtensionLength > 0 Then iExtensionLength = iExtensionLength + 1 ' include the dot
            cboPath.SelStart = Len(cboPath) - iExtensionLength
      End If
End Sub

Private Sub cboPath_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
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

Private Sub chkFileBrowser_Click()
      picBrowser.Visible = chkFileBrowser.Value
      mnuViewFilebrowser.Checked = chkFileBrowser.Value
      staTusBar1.Panels(eStat.BrowserStats).Visible = chkFileBrowser.Value
      
      RearrangeControls
End Sub

Private Sub chkFileBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = chkFileBrowser.ToolTipText
End Sub

Private Sub chkFindOptions_Click()
      If chkFindOptions.Value = vbChecked Then
            PopupMenu mnuQuery, vbPopupMenuRightAlign, AbsoluteRight(chkFindOptions), _
                  AbsoluteBottom(chkFindOptions)
      End If
End Sub

Private Sub chkFindOptions_LostFocus()
      chkFindOptions.Value = vbUnchecked
End Sub

Private Sub chkFindOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = chkFindOptions.ToolTipText
End Sub

Private Sub chkReadOnly_Click()
      
      mnuViewReadOnly.Checked = chkReadOnly.Value
      agEditor.ReadOnly = chkReadOnly.Value
      If chkReadOnly.Value = vbChecked Then
            agEditor.BackColor = &H8000000F
            btnEdit.Visible = True
            If mbReplaceMode Then mnuQueryReplace_Click
            btnReplace.Enabled = False
      Else
            btnEdit.Visible = False
            btnReplace.Enabled = True
            agEditor.BackColor = &H80000005
      End If
End Sub

Private Sub chkWordWrap_Click()
      
      Dim lLineIndex As Long, lCharIndex As Long, lMin As Long, lMax As Long
      
      mnuViewWordWrap.Checked = chkWordWrap.Value
      agEditor.ViewMode = chkWordWrap.Value
      
      ' a few things in the statusbar could change in a word wrap:
      '   x, xmax, y, ymax
      ' and some shouldn't change:
      '   i, imax,   (we're not adding or deleting characters or moving the cursor)
      '   sellength
      
      If agEditor.CharacterCount = 0 Then Exit Sub
      agEditor.GetSelection lMin, lMax
      lLineIndex = agEditor.CurrentLine
      lCharIndex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lLineIndex, 0)
      
      If staTusBar1.Visible Then
            With gtStats
                  .X = lMin - lCharIndex + 1
                  .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
                  .Y = lLineIndex + 1
                  .ymax = agEditor.LineCount
            End With
            FillStats
      End If
End Sub

' EditorFindText
'   Finds the search string sFindMe in agEditor between values of lRangeStart and lRangeEnd.
'   This function DOES NOT HIGHLIGHT ANYTHING OR MOVE THE CURSOR.
'
'  rlFoundMin and rlFoundMax receive the start and end positions of the found string.
'  Returns -1 if nothing found, returns rlFoundMin if successful.
'
'  The way EM_FINDTEXTEX works is that it goes from lRangeStart to lRangeEnd in the
'  specified direction.  That means the start position has to come first.  NOT the lower of the values first.

Private Function EditorFindText( _
      ByVal sFindme As String, _
      ByVal eDir As eDirection, _
      ByVal lRangeStart As Long, _
      ByVal lRangeEnd As Long, _
      ByRef rlFoundMin As Long, _
      ByRef rlFoundMax As Long) As Long
      
      Const FR_MATCHCASE As Long = &H4
      Const FR_WHOLEWORD As Long = &H2
      Const FR_DOWN As Long = &H1
      ' Const EM_FINDTEXT As Long = (WM_USER + 56)
      Const EM_FINDTEXTEX As Long = (WM_USER + 79)

      Dim lFindOptions As Long
      Dim tFindData As FINDTEXTEX
      
      If eDir = Forward Then lFindOptions = FR_DOWN ' fr_down = go from lStartMin to end of editor.
      If mnuQueryWholeWord.Checked Then lFindOptions = lFindOptions + FR_WHOLEWORD
      If mnuQueryMatchCase.Checked Then lFindOptions = lFindOptions + FR_MATCHCASE
      
      tFindData.chrg.cpMin = lRangeStart
      tFindData.chrg.cpMax = lRangeEnd
      tFindData.lpstrText = sFindme & Chr(0) ' it wants a C string
      
      EditorFindText = SendMessage(agEditor.RichEdithWnd, EM_FINDTEXTEX, ByVal lFindOptions, tFindData)
      
      rlFoundMin = tFindData.chrgText.cpMin
      rlFoundMax = tFindData.chrgText.cpMax
End Function

Private Function LoadEditorFile(ByVal sFileName As String, Optional ByVal iMode As eViewMode) As Boolean
                  
      If Not gbFullScreenMode And GetFileSize(sFileName) > 0 Then
            Dim iEncoding As Integer
            iEncoding = IsUnicodeFile(sFileName)
            
            If iEncoding = eTextEncoding.ERROR Then
                  SetCaption "Could not load file: " + sFileName
                  agEditor.tag = ""
                  giTextEncoding = eTextEncoding.ASCII
                  staTusBar1.Panels(eStat.encoding) = "ASCII"
                  mbEditorLoading = False
                  Exit Function
            ElseIf Len(sFileName) > 100 Or iEncoding = eTextEncoding.UNICODE Then
                  Dim oFile, oStream
                  Set oFile = goFso.getfile(sFileName)
                  Set oStream = oFile.OpenAsTextStream(eIoMode.ForReading, iEncoding)
                  If oStream.atendofstream() Then
                        agEditor.Text = ""
                  Else
                        agEditor.Text = oStream.readall()
                  End If
                  oStream.Close
                  LoadEditorFile = True
            Else
                  LoadEditorFile = agEditor.LoadFromFile(sFileName, SF_TEXT)
            End If
            giTextEncoding = iEncoding
            If iEncoding = eTextEncoding.UNICODE Then
                  staTusBar1.Panels(eStat.encoding) = "UNICODE"
            Else
                  staTusBar1.Panels(eStat.encoding) = "ASCII"
            End If
      Else
            agEditor.Text = ""
            giTextEncoding = eTextEncoding.ASCII
            staTusBar1.Panels(eStat.encoding) = "ASCII"
            LoadEditorFile = True
      End If

      SetCaption sFileName & "  (" & Format(GetFileSize(sFileName), "#,#0") & " bytes saved on " _
            & FileModifiedTime(sFileName) & ")"

End Function

Private Function LoadPictureFile(ByVal sFileName As String, Optional ByVal iMode As eViewMode) As Boolean
       
      Dim cRenderer As New stdPicEx2
      LoadPictureFile = True
      On Error Resume Next
      Dim oPic As IPictureDisp
      Set gtImageData.OutPic.Picture = Nothing
      Set oPic = cRenderer.LoadPictureEx(sFileName, mgtAutoSelect)
      gtImageData.DefaultWidth = ScaleX(oPic.Width, vbHimetric, vbTwips)
      gtImageData.DefaultHeight = ScaleY(oPic.Height, vbHimetric, vbTwips)
      
      If geImageSizingMode = eImageSizingMode.Default100 Then
            gtImageData.OutPic.Picture = oPic
            ImageSetZoom sliZoom.Value, sFileName
      
      ElseIf geImageSizingMode = eImageSizingMode.AlwaysFit Then
            ImageZoomFit oPic, sFileName
      End If
      
      If Err > 0 Then
            SetCaption "ERROR: " & sFileName & ", picture couldn't load"
            LoadPictureFile = False
      End If
      On Error GoTo 0
End Function

Private Function LoadFile(ByVal sFileName As String, Optional ByVal iMode As eViewMode) As Boolean
      
      If mbEditorLoading Then agEditor.Text = ""
      
      If Trim(sFileName) = "" Then ' Blank means start a new file.
            mnuFileNew_Click
            Exit Function
      ElseIf Not FileExists(sFileName) Then
            SetCaption "ERROR: file does not exist."
            agEditor.tag = ""
            Exit Function
      End If
      
      EditorSetMode iMode
      
      Select Case iMode
            Case eViewMode.TextView
                  mbEditorLoading = True
                  LoadFile = LoadEditorFile(sFileName, iMode)
                  
            Case eViewMode.PictureView
                  mbEditorLoading = True
                  LoadFile = LoadPictureFile(sFileName, iMode)
            
            Case eViewMode.PropertiesView
                  mbEditorLoading = True
                  LoadPropertiesView sFileName
                  SetCaption sFileName
                  LoadFile = True
      End Select
            
      If LoadFile Or GetFileSize(sFileName) = 0 Then  ' Success!
            agEditor.tag = sFileName
            staTusBar1.Panels(eStat.Modified) = ""
            agEditor.SetSelection 0, 0
            AddToHistorySmartly sFileName
      
      Else  ' Miscellaneous Failure!  agEditor returns no clues as to the problem.
            SetCaption "Could not load file. Command() = """ & Command() & """; File = """ & sFileName & """"
            agEditor.tag = ""
            giTextEncoding = eTextEncoding.ASCII
            staTusBar1.Panels(eStat.encoding) = "ASCII"
      End If
      
      mbEditorLoading = False
End Function

Private Sub SetCaption(sCaption As String)
      frmMain.Caption = sCaption
      If gbFullScreenMode Then
            frmFullScreen.lblFileNameZoom = sCaption & "  "
      End If
End Sub

Private Sub EditorSetMode(iMode As eViewMode)

      ' When we change the sort of data to display (text, picture, more to be determined),
      ' there are some things that have to be set, hidden, etc.
      
      If iMode = geEditorMode Then Exit Sub
      
      Select Case iMode
            Case eViewMode.TextView
      
                  geEditorMode = iMode
                  agEditor.Visible = True
                  Image1.Visible = False
                  Image1.Picture = Nothing
                  sstProperties.Visible = False
                  btnFont.Visible = True
                  chkWordWrap.Visible = False
                  btnFullScreen.Visible = False
                  If chkReadOnly Then btnEdit.Visible = True
                  If Not mbHideFind And mnuViewToolbar.Checked Then picQuery.Visible = True
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3000, 260, 615, 320
                  btnZoomOut.Move 1800, 260, 615, 320
                  btnZoomDefault.Visible = False
                  btnFitImage.Visible = False
                  
                  staTusBar1.Panels(eStat.encoding).Visible = True
                  staTusBar1.Panels(eStat.Modified).Visible = True
                  staTusBar1.Panels(eStat.Stats).Visible = True
                  staTusBar1.Panels(eStat.SelText).Visible = True
      
            Case eViewMode.PictureView
                  
                  geEditorMode = eViewMode.PictureView
                  agEditor.Visible = False
                  agEditor.Text = ""
                  Image1.Visible = True
                  sstProperties.Visible = False
                  btnFont.Visible = False
                  chkWordWrap.Visible = False
                  btnFullScreen.Visible = True
                  btnEdit.Visible = False
                  If Not mbHideFind Then picQuery.Visible = False
                  
                  '4 buttons x-pos: 1800, 2250, 2700, 3150
                  sliZoom.Visible = True
                  btnZoomIn.Move 3150, 260, 470, 320
                  btnZoomOut.Move 1800, 260, 460, 320
                  btnZoomDefault.Visible = True
                  btnFitImage.Visible = True
                  
                  If glOldpicEditorProc = 0 Then
                        glOldpicEditorProc = SetWindowLong(picEditor.hWnd, GWL_WNDPROC, _
                              AddressOf TrackMouseWheel)
                  End If
                  
                  staTusBar1.Panels(eStat.encoding).Visible = False
                  staTusBar1.Panels(eStat.Modified).Visible = False
                  staTusBar1.Panels(eStat.Stats).Visible = False
                  staTusBar1.Panels(eStat.SelText).Visible = False
                  
            Case eViewMode.PropertiesView
            
                  geEditorMode = eViewMode.PropertiesView
                  agEditor.Visible = False
                  agEditor.Text = ""
                  Image1.Visible = False
                  Image1.Picture = Nothing
                  sstProperties.Visible = True
                  
                  btnFont.Visible = True
                  chkWordWrap.Visible = False
                  btnFullScreen.Visible = False
                  btnEdit.Visible = False
                  If Not mbHideFind Then picQuery.Visible = False
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3000, 260, 615, 320
                  btnZoomOut.Move 1800, 260, 615, 320
                  btnZoomDefault.Visible = False
                  btnFitImage.Visible = False
                  
                  staTusBar1.Panels(eStat.encoding).Visible = False
                  staTusBar1.Panels(eStat.Modified).Visible = False
                  staTusBar1.Panels(eStat.Stats).Visible = False
                  staTusBar1.Panels(eStat.SelText).Visible = False
            
            Case Else
                  DebugLog "How did we get to the ERROR ViewMode? agEditor.tag: """ + agEditor.tag + """"
      End Select
      
      RearrangeControls
End Sub

Private Sub FillStats()

      staTusBar1.Panels(eStat.Stats) = "Char: " & Format(gtStats.i, "#,#0") & "/" & Format(gtStats.imax, "#,#0") _
            & "  Ln: " & Format(gtStats.Y, "#,#0") & "/" & Format(gtStats.ymax, "#,#0") & "  Col: " & gtStats.X _
            & "/" & gtStats.xmax
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
                              ShowFileProperties gtBrowserData.Dir & lvwBrowser.SelectedItem
                        Else
                              ShowFileProperties agEditor.tag
                        End If
                  End If
            
            
            Case vbKeyF
                  If ActiveControl.Name <> "agEditor" And Shift = vbCtrlMask + vbShiftMask Then
                        btnFont_Click
                  End If
                  
            Case vbKeyH
                  If Shift = vbCtrlMask And chkFileBrowser Then btnSort_Click
            
            Case vbKeyF11
                  If Shift = 0 Then btnFullScreen_Click
            
            Case 221 ' Right Bracket "]"
                  If Shift = vbCtrlMask Then BrowserExecuteNext
            
            Case 219 ' Left Bracket "["
                  If Shift = vbCtrlMask Then BrowserExecuteNext True
                  
            Case 188 ' Comma (",")  ...also "<"
                  If Shift = vbCtrlMask + vbShiftMask Then
                        mnuviewzoomout_Click
                  End If
                  
            Case 190 ' Period (".") ...also ">"
                  If Shift = vbCtrlMask + vbShiftMask Then
                        mnuviewzoomin_Click
                        
                  ElseIf Shift = vbAltMask And chkFindOptions.Visible And chkFindOptions.Value = vbUnchecked Then
                        'Alt+period  opens popup menu for find options
                        chkFindOptions.SetFocus
                        chkFindOptions.Value = vbChecked
                  ElseIf chkFindOptions.Value = vbChecked Then
                        ' Same button closes find options menu, if already opened
                        chkFindOptions.Value = vbUnchecked
                  End If
                  
            Case vbKeyF5
                  If Shift = vbCtrlMask And chkFileBrowser Then btnSyncContents_Click
            
            Case vbKeyEscape
                  ' Popup menu doesn't wanna die by itself; escape closes it.
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
      Dim sLoadErrorMsg As String
      
      On Error GoTo LOAD_ERROR
      DebugLog "----------------------------------------", 2
      DebugLog "LOADING PHLEGMOIRS v" & App.Major & "." & App.Minor & "." & App.Revision, 2
      If (Trim(Command()) <> "") Then
            DebugLog "Command(): " & Command(), 2
      End If
      DebugLog "", 2

      DebugLog "Initializing menus..."
      InitializeMenus
            
      Set goFso = CreateObject("Scripting.FileSystemObject") ' Just one of these will do.
      Set gtImageData.OutPic = Image1
      Set gtImageData.SurroundingBox = picEditor
      gtBrowserData.ListEmpty = True
      geEditorMode = Text
      
      gsCommandFile = Trim(Command())
      
      If Left(gsCommandFile, 1) = Chr(34) Then
            gsCommandFile = Mid(gsCommandFile, 2, Len(gsCommandFile) - 2)
      End If
      
      If gsCommandFile <> "" And Not (gsCommandFile Like "*:\*") Then
            gsCommandFile = CurDir & "\" & gsCommandFile
      End If
      
      agEditor.tag = gsCommandFile
      
      gsPhlegmKey = "Software\" & App.title & "\" & REGISTRY_VERSION
      
      vDate = Date
      gsPhlegmDate = year(vDate) & "-" & Format(Month(vDate), "0#") & "-" & Format(Day(vDate), "0#")
      
      DoEvents
      LoadRegistrySettings
      gtBrowserData.DoneLoading = True ' Why does it finish loading *here*? Proved this by debugging.

      
      gtStats.imax = CharacterCount(agEditor)
      FillStats
      staTusBar1.Panels(eStat.Modified) = ""

      If Not DEBUGGING Then
            glOldLvwProc = SetWindowLong(lvwBrowser.hWnd, GWL_WNDPROC, AddressOf ListViewProc)
      End If
      
      If agEditor.Visible Then
            agEditor.SetFocus
      ElseIf lvwBrowser.Visible Then
            lvwBrowser.SetFocus
      End If
      
      If agEditor.tag = "" Then AutosizeColumns ' in this specific case, the normal time to do this is too soon
      
      btnSyncContents_Click
      Exit Sub

LOAD_ERROR:
      sLoadErrorMsg = "Load error. Err: " & Err.Description
      DebugLog sLoadErrorMsg, 2
      MsgBox sLoadErrorMsg
      Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' If we open a popupmenu, and then right click off into space,
      '   the mousedown event is called for the form (not for the control we are
      '   hovering over nor the menu itself.)
      ' Our form doesn't need it.  We'll have him pass it to the control it's over.

      Dim lCtrlHwnd As Long
      Dim tCursor As POINTAPI

      If Button <> vbRightButton Or Shift <> 0 Then Exit Sub

      GetCursorPos tCursor
      lCtrlHwnd = WindowFromPoint(tCursor.X, tCursor.Y)
      
      On Error Resume Next
      If Screen.ActiveControl.Container.Name = "picQuery" Then
            txtFind.SetFocus
      End If
      On Error GoTo 0
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      staTusBar1.Panels(eStat.Tips).Text = ""
End Sub

Private Sub Form_Resize()
      If mbSkipFormResize Then
            ' Beep
      Else
            RearrangeControls
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
      SaveSettingsToRegistry
      Unload frmAbout
      Unload frmFullScreen
      If Not DEBUGGING Then
            SetWindowLong lvwBrowser.hWnd, GWL_WNDPROC, glOldLvwProc
            glOldLvwProc = 0
      End If
      
      If glOldpicEditorProc <> 0 Then
            SetWindowLong picEditor.hWnd, GWL_WNDPROC, glOldpicEditorProc
            glOldpicEditorProc = 0
      End If
      DebugLog "", 2
      DebugLog "UNLOADING PHLEGMOIRS", 2
      Set goFso = Nothing
End Sub

Private Sub GatherBrowserPrefs(ByRef rtPrefs As TBrowserPrefs)
      With rtPrefs
            .AutoLoadPath = cboPath
            .SortMethod = lvwBrowser.SortOrder
            .SortKey = lvwBrowser.SortKey
            On Error GoTo 0
            .NameColumnIndex = lvwBrowser.ColumnHeaders.item("Name").Position
            .TypeColumnIndex = lvwBrowser.ColumnHeaders.item("Type").Position
            .SizeColumnIndex = lvwBrowser.ColumnHeaders.item("Size").Position
            .ModifiedColumnIndex = lvwBrowser.ColumnHeaders.item("Modified").Position
            .NameColumnWidth = lvwBrowser.ColumnHeaders.item("Name").Width
            .TypeColumnWidth = lvwBrowser.ColumnHeaders.item("Type").Width
            .SizeColumnWidth = lvwBrowser.ColumnHeaders.item("Size").Width
            .ModifiedColumnWidth = lvwBrowser.ColumnHeaders.item("Modified").Width
            On Error Resume Next
      End With
End Sub

Private Sub GatherEditorPrefs(ByRef rtPrefs As TEditorPrefs)
      Dim lMin As Long, lMax As Long
      Dim oTempFont As New StdFont
      
      agEditor.GetSelection lMin, lMax

      With rtPrefs
            .AutoLoadFile = agEditor.tag
            .FirstVisibleLine = agEditor.FirstVisibleLine
            .SelEnd = lMax
            .SelStart = lMin
            .WordWrap = chkWordWrap.Value
            ' If we were set to readonly while looking at pictures, I'll assume the setting wasn't
            ' REALLy that important, at the time.  So, not saving it in that case.
            If geEditorMode <> Picture And chkReadOnly.Value = vbChecked Then
                  .ReadOnly = vbChecked
            Else
                  .ReadOnly = vbUnchecked
            End If
            
            Set oTempFont = GetRealStdFont(agEditor.RichEdithWnd, .TextColor)
            ' Here, we'll store the color as a system color, if it happens to match the button text.
            If .TextColor = TranslateColor(vbWindowText) Then .TextColor = vbWindowText
            .FontBold = oTempFont.Bold
            .FontItalic = oTempFont.Italic
            .FontName = oTempFont.Name
            .FontSize = oTempFont.Size
            .FontStrikethrough = oTempFont.Strikethrough
            .FontUnderline = oTempFont.Underline
            
            SendMessage agEditor.RichEdithWnd, EM_GETSCROLLPOS, 0, .ScrollPos
      End With
End Sub

Private Sub GatherHistoryAndBookmarks(ByRef rtPrefs As TAllPrefs)
      Dim iIndex As Integer
      With rtPrefs
            .BookmarkCount = mnuBookmark.UBound
            For iIndex = 1 To .BookmarkCount
                  .Bookmarks(iIndex) = mnuBookmark(iIndex).tag
            Next iIndex
            
            .HistoryCount = mnuFileHistory.UBound
            For iIndex = 1 To .HistoryCount
                  .History(iIndex) = mnuFileHistory(iIndex).tag
            Next iIndex
            
            .PathHistoryCount = cboPath.ListCount
            If .PathHistoryCount > MAX_HISTORY Then .PathHistoryCount = MAX_HISTORY
            For iIndex = 1 To .PathHistoryCount
                  .PathHistory(iIndex) = cboPath.List(iIndex)
            Next iIndex
      End With
End Sub

Private Sub GatherWindowPrefs(ByRef rtPrefs As TWindowPrefs)
      With rtPrefs
            .WNP.length = LenB(.WNP)
            GetWindowPlacement hWnd, .WNP
            If .WNP.showCmd = SW_MINIMIZE Then
                  .WNP.showCmd = SW_RESTORE
            ElseIf .WNP.showCmd = SW_SHOWMINIMIZED Then  '  <-- It'll be this one, not SW_MINIMIZE.
                  .WNP.showCmd = SW_SHOWNORMAL                ' Including the other for paranoia.
            End If
            .BrowserWidth = picBrowser.Width
            .ShowFileBrowser = picBrowser.Visible
            .ShowStatusBar = staTusBar1.Visible
            .ShowToolBar = picToolBar.Visible
            .ShowFind = Not mbHideFind
            .ImageZoom = sliZoom.Value
            .ImageSizingMode = geImageSizingMode
      End With
End Sub

Private Function GetFindCompareMode() As Integer
      If mnuQueryMatchCase.Checked Then
            GetFindCompareMode = vbBinaryCompare
      Else
            GetFindCompareMode = vbTextCompare
      End If
End Function

Private Sub Image1_DblClick()
      ' This needs to (effectively) call an Image1_mousedown... but with what parameters???
      Dim tPrev As POINTAPI
      Dim tPicBoxRect As RECT
      
      GetCursorPos tPrev
      GetWindowRect picEditor.hWnd, tPicBoxRect
      
      gtImageData.PrevX = (tPrev.X - tPicBoxRect.Left) * Screen.TwipsPerPixelX - Image1.Left
      gtImageData.PrevY = (tPrev.Y - tPicBoxRect.Top) * Screen.TwipsPerPixelY - Image1.Top
      gtImageData.Dragging = True
      picEditor.SetFocus
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      gtImageData.PrevX = X
      gtImageData.PrevY = Y
      If Button = vbLeftButton Then
            gtImageData.Dragging = True
            picEditor.SetFocus
      End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If FOCUS_FOLLOWS_MOUSE Then
            On Error Resume Next
            If GetForegroundWindow = frmMain.hWnd And Not (ActiveControl.Name = "picEditor") Then
                  picEditor.SetFocus
            End If
            On Error GoTo 0
      End If
            
      If gtImageData.Dragging Then
            Image1.Move Image1.Left + X - gtImageData.PrevX, Image1.Top + Y - gtImageData.PrevY, Image1.Width, Image1.Height
            If X <> gtImageData.PrevX Or Y <> gtImageData.PrevY Then gtImageData.Moved = True
      End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' Mouse button lifted?  Stop the drag!
      gtImageData.Dragging = False
      
      picEditor_MouseUp Button, Shift, X, Y
End Sub

Public Sub ImageSetZoom(ByVal iZoom As Integer, Optional ByVal sFileName As String)
      If sFileName = "" Then sFileName = agEditor.tag
      If iZoom > sliZoom.Max Then
            iZoom = sliZoom.Max
      ElseIf iZoom < 0 Then
            iZoom = 0
      End If
      
      gtImageData.OutPic.Move gtImageData.OutPic.Left, gtImageData.OutPic.Top, _
            gtImageData.DefaultWidth * CSng(iZoom) / 100#, gtImageData.DefaultHeight * CSng(iZoom) / 100#
      SetCaption sFileName & "  (" & iZoom & "%)"
      sliZoom.Value = iZoom
End Sub

Public Sub ImageZoomFit(ByRef roPic As IPictureDisp, ByVal sFileName As String)
      If roPic = 0 Then Exit Sub
      Dim dRatio As Double, dAreaRatio As Double
      Dim lTop As Long, lLeft As Long, lHeight As Long, lWidth As Long
      dRatio = CDbl(gtImageData.DefaultWidth) / gtImageData.DefaultHeight
      dAreaRatio = CDbl(gtImageData.SurroundingBox.ScaleWidth) / gtImageData.SurroundingBox.ScaleHeight
      
      If dRatio > dAreaRatio Then
            ' pic is more wide than long, compared to its space
            lLeft = 0
            lWidth = gtImageData.SurroundingBox.ScaleWidth
            lHeight = lWidth / dRatio
            lTop = gtImageData.SurroundingBox.ScaleHeight / 2 - lHeight / 2
      Else
            lTop = 0
            lHeight = gtImageData.SurroundingBox.ScaleHeight
            lWidth = lHeight * dRatio
            lLeft = gtImageData.SurroundingBox.ScaleWidth / 2 - lWidth / 2
      End If
      gtImageData.OutPic.Move lLeft, lTop, lWidth, lHeight
      gtImageData.OutPic.Picture = roPic
      sliZoom.Value = Format(CDbl(lHeight) / gtImageData.DefaultHeight * 100, "#,#0")
      SetCaption sFileName & "  (" & sliZoom.Value & "%)"
End Sub

Public Sub ImageZoomIn(iStep As Integer)
      ' goes up to the next zoom divisible by iStep
      If sliZoom.Value >= sliZoom.Max Then Exit Sub
      ImageSetZoom sliZoom.Value + (iStep - (sliZoom.Value Mod iStep))
End Sub

Public Sub ImageZoomOut(iStep As Integer)
      ' Sets zoom to the next lowest integer divisibly by iStep.
      
      If sliZoom.Value <= 0 Then Exit Sub
      
      If sliZoom.Value Mod iStep = 0 Then
            ImageSetZoom sliZoom.Value - iStep
      Else
            ImageSetZoom sliZoom.Value - (sliZoom.Value Mod iStep)
      End If
End Sub

Private Sub InitializeMenus()
      mnuEditUndo.Caption = "Undo" & vbTab & "Ctrl+Z"
      mnuEditRedo.Caption = "Redo" & vbTab & "Ctrl+Y"
      mnuViewFont.Caption = "Font..." & vbTab & "Shift+Ctrl+F"
      mnuViewZoomIn.Caption = mnuViewZoomIn.Caption & vbTab & "Shift+Ctrl+"">"""
      mnuViewZoomOut.Caption = mnuViewZoomOut.Caption & vbTab & "Shift+Ctrl+""<"""
      
      mnuFileNext.Caption = mnuFileNext.Caption & vbTab & "Ctrl+]"
      mnuFilePrev.Caption = mnuFilePrev.Caption & vbTab & "Ctrl+["
      
      mnuWriteFind.Caption = mnuWriteFind.Caption & vbTab & "Ctrl+F"
      mnuWriteCut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnueditcut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnuWriteCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuEditCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuWritePaste.Caption = "&Paste" & vbTab & "Ctrl+V"
      mnuEditPaste.Caption = "&Paste" & vbTab & "Ctrl+V"
      
      mnuListDelete.Caption = mnuListDelete.Caption & vbTab & "Del"
      mnuListProperties.Caption = "&Properties" & vbTab & "Alt+Enter"
      mnuListCancel.Caption = "&Cancel" & vbTab & "Esc"
End Sub

Private Sub lblDivider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      If lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "" Then
            
            lblDivider.tag = "Resizing"
      End If
End Sub

Private Sub lblDivider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim iOffset As Integer

      If lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "Resizing" Then
            Dim lPrevLeft As Long
            lPrevLeft = lblDivider.Left
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
            SaveSettingsToRegistry
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

Private Sub ListMenuEnable(oHoverItem As ListItem)
      ' This will be called when a listitem is clicked, and it will enable or disable parts
      ' of the right click menu, based on the sort of listitem is passed to it.
      
      mnuListOpenDefault.Enabled = True
      mnuListOpen.Enabled = True
      mnuListOpenDefault.Caption = "Open With Default Program..." & vbTab & "Shift+Ctrl+Enter"
      mnuListCopyPath.Enabled = True
      mnuListProperties.Enabled = True
      
      If gtBrowserData.BookmarkMode Then
            mnuListShowOnly.Enabled = False
            mnuListDelete.Caption = "&Delete Bookmark" & vbTab & "Del"
      
      ElseIf gtBrowserData.HistoryMode Then
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
      
      If oHoverItem.Icon = eIconType.Directory Or oHoverItem.Icon = eIconType.Drive Then
            mnuListOpenDefault.Caption = "Explore..." & vbTab & "Shift+Ctrl+Enter"
            mnuListDelete = False
            If oHoverItem.Text = ".." Or oHoverItem.Icon = eIconType.Drive Then mnuListRename = False
      End If

End Sub

Private Sub LoadBrowserPrefs(ByRef rtPrefs As TBrowserPrefs)
      Dim sCboPath As String
      With rtPrefs
            lvwBrowser.SortOrder = .SortMethod
            lvwBrowser.SortKey = .SortKey
            If ALLOW_REARRANGE_COLUMNS Then
                  lvwBrowser.AllowColumnReorder = True
                  lvwBrowser.ColumnHeaders.item("Name").Position = .NameColumnIndex
                  lvwBrowser.ColumnHeaders.item("Type").Position = .TypeColumnIndex
                  lvwBrowser.ColumnHeaders.item("Size").Position = .SizeColumnIndex
                  lvwBrowser.ColumnHeaders.item("Modified").Position = .ModifiedColumnIndex
            End If
            lvwBrowser.ColumnHeaders.item("Name").Width = .NameColumnWidth
            If Not AUTOSIZE_COLUMNS Then
                  lvwBrowser.ColumnHeaders.item("Type").Width = .TypeColumnWidth
                  lvwBrowser.ColumnHeaders.item("Size").Width = .SizeColumnWidth
                  lvwBrowser.ColumnHeaders.item("Modified").Width = .ModifiedColumnWidth
            End If
            sCboPath = Trim(CstringToVBstring(.AutoLoadPath))
            If agEditor.tag = "" Then
                  DebugLog "We don't have a file to load, so load the most recent directory...", 2
                  cboPath = sCboPath
                  DebugLog "File browser path set to: " & cboPath, 2
            ElseIf agEditor.tag <> "" And sCboPath <> "" Then
                  ' if we're not gonna load it, at least make it the most recent path history item
                  PathAddRecent sCboPath
            End If
            DoEvents
      End With
End Sub

Private Sub LoadEditorPrefs(ByRef rtPrefs As TEditorPrefs)
      Dim oTempFont As New StdFont
            
      On Error GoTo EDITOR_PREFS_ERROR
      DebugLog "Found editor settings. Applying them..."
      With rtPrefs
            chkWordWrap.Value = .WordWrap
            chkWordWrap_Click

            chkReadOnly.Value = .ReadOnly
            chkReadOnly_Click

            oTempFont.Name = Trim(CstringToVBstring(.FontName))
            oTempFont.Size = .FontSize
            oTempFont.Bold = .FontBold
            oTempFont.Italic = .FontItalic
            oTempFont.Strikethrough = .FontStrikethrough
            oTempFont.Underline = .FontUnderline
            If Len(Trim(.FontName)) > 28 Then
                  btnFont.Caption = Left(Trim(.FontName), 17) & "..."
            Else
                  btnFont.Caption = Trim(.FontName)
            End If
            SetRealStdFont agEditor.RichEdithWnd, oTempFont, .TextColor
            lblFontSize = Round(.FontSize, 0)
      End With
      Exit Sub
      
EDITOR_PREFS_ERROR:
      frmMain.Caption = "ERROR: Could not load editor prefs. Err: " & Err.Description
      DebugLog frmMain.Caption, 2
      MsgBox frmMain.Caption
End Sub

Private Sub LoadHistoryAndBookmarks(ByRef rtPrefs As TAllPrefs)
      Dim iBookm As Integer, iHistIndex As Integer
      With rtPrefs
            DebugLog "Loading bookmarks..."
            For iBookm = 1 To .BookmarkCount
                  AddToBookmarks Trim(CstringToVBstring(.Bookmarks(iBookm)))
            Next iBookm
            DebugLog "Loaded " & .BookmarkCount & " bookmarks."
            
            DebugLog "Loading file history..."
            For iHistIndex = 1 To .HistoryCount
                  AddToHistorySimply Trim(CstringToVBstring(.History(iHistIndex)))
            Next iHistIndex
            DebugLog "Loaded " & .HistoryCount & " historical file records."
            
            DebugLog "Loading path history..."
            For iHistIndex = .PathHistoryCount To 1 Step -1
                  PathAddRecent Trim(CstringToVBstring(.PathHistory(iHistIndex)))
            Next iHistIndex
            DebugLog "Loaded " & .HistoryCount & " historical path records."
      End With
End Sub

Private Sub LoadPropertiesView(ByVal sFileName As String)
      Dim tWfd As WIN32_FIND_DATA
      Dim lFile As Long
      Dim sEx As String
      
      lFile = FindFirstFile(sFileName, tWfd)
      fraProperties.Caption = tWfd.cFileName
      lblPropValue(2) = Format(goFso.getfile(sFileName).Size, "#,#0")
      lblPropValue(4) = FormatNonLocalFileTime(tWfd.ftLastWriteTime)
      lblPropValue(3) = FormatNonLocalFileTime(tWfd.ftCreationTime)
      lblPropValue(5) = FormatNonLocalFileTime(tWfd.ftLastAccessTime)
      FindClose lFile

      sEx = goFso.getextensionname(sFileName)
      If sEx = "mp3" Then
            Dim tMp3Info As MP3TagInfo
            
            GetMP3Info sFileName, tMp3Info
            With tMp3Info
                  lblPropValue(6) = tMp3Info.title
                  lblPropValue(7) = tMp3Info.artist
                  lblPropValue(8) = tMp3Info.album
                  lblPropValue(9) = tMp3Info.year
            End With
      Else
            With tMp3Info
                  lblPropValue(6) = ""
                  lblPropValue(7) = ""
                  lblPropValue(8) = ""
                  lblPropValue(9) = ""
            End With
      End If
      ' getAllProperties sFileName
End Sub

Private Sub LoadRegistrySettings()
      DebugLog "Retrieving registry settings, version " & REGISTRY_VERSION & "..."
      On Error GoTo SETTINGS_ERROR
      
      Dim lRetVal As Long, lKey As Long
      Dim lDataType As Long ' receiving only
      Dim lValueSize As Long ' in/out
      Dim bFileLoaded As Boolean
      Dim tPrefs As TAllPrefs
      
      mbSkipFormResize = True
      
      lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, gsPhlegmKey, 0, KEY_QUERY_VALUE, lKey)
      lValueSize = LenB(tPrefs)
      lRetVal = RegQueryValueExAny(lKey, "Settings", 0, lDataType, ByVal tPrefs, lValueSize)
      
      If lRetVal = 0 Then
            If agEditor.tag = "" Then agEditor.tag = Trim(CstringToVBstring(tPrefs.EditorPrefs.AutoLoadFile))
            LoadWindowPrefs tPrefs.WindowPrefs
            LoadHistoryAndBookmarks tPrefs
            LoadBrowserPrefs tPrefs.BrowserPrefs
            LoadEditorPrefs tPrefs.EditorPrefs
      Else
            DebugLog "Did not find any previous settings."
            cboPath = ""
            BrowserResizeHorizontal picBrowser.Width
      End If
      
      ' It's important to set the above prior to loading a file.
      ' Otherwise agEditor's display routines are called again and again for an entire file,
      ' rather than for a blank editor.
      
      DebugLog "Attempting to auto-load file: " & agEditor.tag & "..."
      bFileLoaded = LoadFile(agEditor.tag, GetViewMode(agEditor.tag, GetIconType(goFso.getextensionname(agEditor.tag))))
      If bFileLoaded Then
            DebugLog "File loaded."
      Else
            DebugLog "File was NOT loaded."
      End If

      If lRetVal = 0 And bFileLoaded Then
            With tPrefs.EditorPrefs
                  ' If the file has been changed so that selection and scroll positions are meaningless,
                  ' just skip them...
                  On Error Resume Next
                  If Trim(gsCommandFile) = "" Then
                        agEditor.SetSelection .SelStart, .SelEnd
                        SendMessage agEditor.RichEdithWnd, EM_SETSCROLLPOS, 0, .ScrollPos
                  End If
                  On Error GoTo 0
            End With
      End If
      mbSkipFormResize = False
      RegCloseKey lKey
      DebugLog "All settings complete."
      Exit Sub
      
SETTINGS_ERROR:
      frmMain.Caption = "ERROR: Could not load settings. Err: " & Err.Description
      DebugLog frmMain.Caption, 2
      MsgBox frmMain.Caption
      mbSkipFormResize = False
End Sub

Private Sub LoadWindowPrefs(ByRef rtPrefs As TWindowPrefs)
      With rtPrefs
            BrowserResizeHorizontal .BrowserWidth

            .WNP.length = LenB(.WNP)
            SetWindowPlacement hWnd, .WNP

            chkFileBrowser.Value = -CInt(.ShowFileBrowser)
            chkFileBrowser_Click
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
            mbHideFind = Not .ShowFind
            ImageSetZoom .ImageZoom
            geImageSizingMode = .ImageSizingMode
            If geImageSizingMode = AlwaysFit Then
                  mnuViewFitImage.Checked = True
            Else
                  mnuViewFitImage.Checked = False
            End If
            DebugLog "Rearranging controls..."
            RearrangeControls
            DebugLog "Rearranged controls."
      End With
End Sub

'
'  lvwBrowser_AfterLabelEdit (in other words, "rename")
'
'     It is even allowable to rename an open file without saving as a new file or deleting anything.
'
'     Unsaved progress will not be tampered with, but NOR WILL IT BE SAVED, until you save it.
'
Private Sub lvwBrowser_AfterLabelEdit(Cancel As Integer, NewString As String)

      Dim sFolder As String, sOldPath As String
      
      sFolder = gtBrowserData.Dir
      
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
      sOldPath = gtBrowserData.Dir & lvwBrowser.SelectedItem
      
      Cancel = RenameFileWithChecks(sOldPath, sFolder & NewString)
End Sub

Private Sub lvwBrowser_Click()
      miBrowserMouseButton = 0  ' These probably an overcaution --
      miBrowserShift = 0                  ' They are reset in the next MouseDown anyway.
End Sub

Private Sub lvwBrowser_ColumnClick(ByVal oColumnHeader As MSComctlLib.ColumnHeader)
      If gtBrowserData.HistoryMode Then Exit Sub
      
      Dim iNewKey As Integer
      
      With lvwBrowser
            ' This overhead maneuver can't be used without major, major, major overhaul...
'            If oColumnHeader.key = "Size" Then
'                  lRetVal = SendMessage(.hwnd, LVM_SORTITEMSEX, ByVal .SortOrder, _
'                        AddressOf CompareLong)

            If oColumnHeader.Key = "Size" Then
                  iNewKey = 4  ' Doing the switch... 5th column stores size invisibly, with leading zeroes for text sorting.
            Else
                  iNewKey = oColumnHeader.Index - 1
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
      
      If gtBrowserData.BookmarkMode Then BookmarkSaveChanges

End Sub

Private Sub lvwBrowser_DblClick()
      mbBrowserDoubleClick = True
      If Not mbBrowserItemClicked Then
            btnFolderUp_Click
      End If
End Sub

Private Sub lvwBrowser_ItemClick(ByVal oItem As MSComctlLib.ListItem)
      mbBrowserItemClicked = True
      ListMenuEnable lvwBrowser.SelectedItem
End Sub

Private Sub lvwBrowser_KeyDown(KeyCode As Integer, Shift As Integer)
      ' Left = up folder.  Right = open folder.
      ' Trying to copy the functionality of explorer somehow, but without a visible tree.
      
      Const COLUMN_SIZE_INC = 50
      
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
                        If lvwBrowser.ColumnHeaders.item(1).Width >= COLUMN_SIZE_INC Then
                              lvwBrowser.ColumnHeaders.item(1).Width = _
                                    lvwBrowser.ColumnHeaders.item(1).Width - COLUMN_SIZE_INC
                        End If
                  Else
                        btnFolderUp_Click   ' Ordinary left arrow...
                  End If
            
            Case vbKeyF2
                  mnuListRename_Click
                                                 
            Case vbKeyF13 ' F13, but contains code for it and for right arrow. See ListViewProc for details.
                              
                  ' Right = open a folder or a drive, but leave a file alone.
                  '     ...and we take pains to disarm the listview's urge to scroll right on right arrow
                  
                  If Shift = vbShiftMask Then
                        ' Oh, and shift+right is going to increase column width
                        lvwBrowser.ColumnHeaders.item("Name").Width = _
                              lvwBrowser.ColumnHeaders.item("Name").Width + COLUMN_SIZE_INC
                              
                  ElseIf lvwBrowser.ListItems.Count > 0 Then
                        With lvwBrowser.SelectedItem
                              If .Icon = eIconType.Directory Or .Icon = eIconType.Drive Or .Icon = eIconType.Cdrom Or _
                                    .Icon = eIconType.Floppy Or .Icon = eIconType.Network And Shift = 0 Then
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
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.item("Name")
            Case vbKeyT
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.item("Type")
            Case vbKeyZ
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.item("Size")
            Case vbKeyM
                  lvwBrowser_ColumnClick lvwBrowser.ColumnHeaders.item("Modified")
            
            Case vbKeyDelete
                  If Shift = 0 Then BrowserDeleteSelected
                  
            Case 219 ' Left Bracket [
                  If Shift = 0 Then BrowserExecuteNext True
            
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

Private Sub lvwBrowser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lvwBrowser_MouseMove Button, Shift, X, Y
      mbBrowserItemClicked = False
      miBrowserMouseButton = Button
      miBrowserShift = Shift
End Sub

Private Sub lvwBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      Dim oHoverItem As ListItem
      
      ' Can't prevent people from making this helper column visible... but it'll be gone pretty quick
      If lvwBrowser.ColumnHeaders(5).Width > 0 Then lvwBrowser.ColumnHeaders(5).Width = 0
      
      If FOCUS_FOLLOWS_MOUSE Then
            ' Autofocus on the file browser.
            ' But we don't do that from within cboPath, because it would be very annoying to
            ' have your typing of a directory interrupted by stray movement of the mouse.
            On Error Resume Next
            If GetForegroundWindow = frmMain.hWnd And Not (ActiveControl.Name = "lvwBrowser") _
                  And Not (ActiveControl.Name = "cboPath") Then
                  lvwBrowser.SetFocus
            End If
            On Error GoTo 0
      End If
      
      ' See if we're over an item.
      Set oHoverItem = lvwBrowser.HitTest(X, Y)
      
      ' Show file names in statusbar on mouseover.
      If Not (oHoverItem Is Nothing) Then
            staTusBar1.Panels(eStat.Tips).Text = oHoverItem.Text
            lvwBrowser.MousePointer = ccCustom
            
            If Button = vbLeftButton Or Button = vbRightButton Then
                  oHoverItem.Selected = True
            End If
      Else
            staTusBar1.Panels(eStat.Tips).Text = ""
            lvwBrowser.MousePointer = ccDefault
      End If
End Sub

Private Sub lvwBrowser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim oHoverItem As ListItem
      
      Set oHoverItem = lvwBrowser.HitTest(X, Y)  ' To see if we're over an item.
      
      If (Button = vbRightButton And Shift = 0) Then
            If oHoverItem Is Nothing Then
                  ListMenuDisable
            Else
                  ListMenuEnable oHoverItem
            End If
            Me.PopupMenu mnuList
      
      ElseIf Button = vbLeftButton And Shift = 0 Then
            
            If Not (oHoverItem Is Nothing) Then
                  ' Open the file/folder on an ordinary left click.
                  BrowserExecuteItem oHoverItem
            Else
                  ' Clicking on empty space deselects the selected item.
                  If Not gtBrowserData.ListEmpty Then lvwBrowser.SelectedItem.Selected = False
            End If
      
      ElseIf Button = vbMiddleButton And Shift = 0 Then
            If oHoverItem Is Nothing Then PathBack
      End If
      
      ' For use in the click event, so we know what was clicked.  OBSOLETE.
      miBrowserMouseButton = Button
      miBrowserShift = Shift
End Sub

Private Sub mnuBookmarksAdd_Click()
      Dim iBookm As Integer
      
      For iBookm = 1 To mnuBookmark.UBound
            If mnuBookmark(iBookm).tag = agEditor.tag Then
                  Exit Sub
            End If
      Next iBookm
      
      AddToBookmarks agEditor.tag
      SaveSettingsToRegistry
      
      If gtBrowserData.BookmarkMode Then RefreshAll
End Sub

Private Sub mnuBookmarksManage_Click()
      If mnuViewFilebrowser.Checked = False Then mnuViewFilebrowser_Click
      cboPath = "(Bookmarks)"
End Sub

Private Sub mnuBookmark_Click(Index As Integer)
      LoadFile mnuBookmark(Index).tag, GetViewMode(mnuBookmark(Index).tag, eIconType.Bookmark)
      
      btnSyncContents_Click
End Sub

Private Sub mnuBrowserRefresh_Click()
      btnRefresh_Click
End Sub

Private Sub mnuEditCopy_Click()
      agEditor.Copy
End Sub

Private Sub mnuEditCut_Click()
      agEditor.Cut
End Sub

Private Sub mnuEditFindBackwards_Click()
      btnFindPrev_Click
End Sub

Private Sub mnuEditFindNext_Click()
      btnFindNext_Click
End Sub

Private Sub mnuEditFind_Click()
      ' Ctrl+F puts the selected text into the query box, but does not proceed with a find until you hit the button.
      
      If geEditorMode = eViewMode.PictureView Or geEditorMode = eViewMode.PropertiesView Then Exit Sub  ' no search/replace within pictures.
      
      If Not mnuViewToolbar.Checked Then mnuViewToolbar_Click
      If mbHideFind Or Not picQuery.Visible Then
            mbHideFind = False
            picQuery.Visible = True
            RearrangeControls
      End If
      On Error Resume Next
      If agEditor.SelectedText <> "" Then
            txtFind = Trim(agEditor.SelectedText)
            txtReplace = ""
      End If
      On Error GoTo 0
      txtFind.SetFocus
End Sub

Private Sub mnuEditPaste_Click()
      agEditor.Paste
End Sub

Private Sub mnuEditRedo_Click()
      agEditor.Redo
End Sub

Private Sub mnuEditReplace_Click()
      If geEditorMode = eViewMode.PictureView Or geEditorMode = eViewMode.PropertiesView Then Exit Sub
      
      mnuEditFind_Click
      
      If mbReplaceMode And ActiveControl.Name <> "txtReplace" And txtReplace = "" Then
            txtReplace.SetFocus
      ElseIf mbReplaceMode And ActiveControl.Name <> "txtreplace" And Not btnReplace.Enabled Then
            txtReplace.SetFocus
      Else
            btnReplace_Click
      End If
End Sub

Private Sub mnuEditUndo_Click()
      agEditor.Undo
End Sub

Private Sub mnuEdit_Click()
      mnuEditUndo.Enabled = agEditor.CanUndo
      mnuEditRedo.Enabled = agEditor.CanRedo
      mnueditcut.Enabled = True
      mnuEditCopy.Enabled = True
      mnuEditPaste.Enabled = True
      mnuEditFind.Enabled = True
      mnuEditReplace.Enabled = True
      mnuEditFindNext.Enabled = True
      mnuEditFindBackwards.Enabled = True
      
      If geEditorMode <> eViewMode.TextView Then
            mnuEditUndo.Enabled = False
            mnuEditRedo.Enabled = False
            mnueditcut.Enabled = False
            mnuEditCopy.Enabled = False
            mnuEditPaste.Enabled = False
            mnuEditFind.Enabled = False
            mnuEditReplace.Enabled = False
            mnuEditFindNext.Enabled = False
            mnuEditFindBackwards.Enabled = False
      End If
      
      If chkReadOnly Then
            mnuEditUndo.Enabled = False
            mnuEditRedo.Enabled = False
            mnueditcut.Enabled = False
            mnuEditPaste.Enabled = False
            mnuEditReplace.Enabled = False
      End If
      
      If ActiveControl.Name <> "agEditor" Then
            mnuEditUndo.Enabled = False
            mnuEditRedo.Enabled = False
            mnueditcut.Enabled = False
            mnuEditCopy.Enabled = False
            mnuEditPaste.Enabled = False
      End If
      
      If agEditor.SelectedText = "" Then
            mnueditcut.Enabled = False
            mnuEditCopy.Enabled = False
            If txtFind = "" Then
                  mnuEditFindNext.Enabled = False
                  mnuEditFindBackwards.Enabled = False
            End If
      End If
End Sub

Private Sub mnuFileExit_Click()
      Unload Me
End Sub

Private Sub mnuFileHistory_Click(Index As Integer)
      LoadFile mnuFileHistory(Index).tag, GetViewMode(mnuFileHistory(Index).tag, eIconType.Bookmark)
      
      btnSyncContents_Click
End Sub

Private Sub mnuFileNew_Click()
      agEditor.Text = ""
      agEditor.tag = ""
      giTextEncoding = eTextEncoding.ASCII
      EditorSetMode eViewMode.TextView
      frmMain.Caption = "(New File)"
      staTusBar1.Panels(eStat.encoding) = "ASCII"
      chkReadOnly.Value = vbUnchecked
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

Private Sub mnuFilePrev_Click()
      btnPrevFile_Click
End Sub

Private Sub mnuFileSaveAs_Click()
      Dim sDefaultPath As String, sFileName As String
      Dim vDate As Variant
      
      If Not agEditor.Visible Then
            frmMain.Caption = "ERROR: can only save in editor mode: " & sFileName
            DebugLog Caption
            Exit Sub
      ElseIf chkReadOnly.Value = vbChecked Then
            frmMain.Caption = "ERROR: can't save in Read Only mode: " & sFileName
            DebugLog Caption
            Exit Sub
      End If

      vDate = Date
      gsPhlegmDate = year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
     
      ' here we decide on a default file name to suggest to the user,
      ' based on a whether the editor.tag is empty, and whether the file browser is at a valid folder.
      If agEditor.tag <> "" Then
            sDefaultPath = agEditor.tag  ' It means this is not a new file we're saving.  Default to old name.
            
      ElseIf gtBrowserData.ValidPath Then
            sDefaultPath = gtBrowserData.Dir & gsPhlegmDate & ".txt"  ' New file, good directory in browser.
      Else
            sDefaultPath = CurDir & "\" & gsPhlegmDate & ".txt"  ' New file, no good directory present.
      End If
      
      While FileExists(sDefaultPath)
            Dim sEx As String
            sEx = goFso.getextensionname(sDefaultPath)
            If sEx <> "" Then
                  Dim oRegex
                  Set oRegex = CreateObject("VBScript.RegExp")
                  oRegex.Global = True
                  oRegex.Pattern = "\." + sEx + "$"
                  sDefaultPath = oRegex.Replace(sDefaultPath, "_." + sEx)
            Else
                  sDefaultPath = sDefaultPath + "_"
            End If
      Wend
      
      sFileName = InputBox("File name:", "Save", sDefaultPath)
      If sFileName <> "" Then SaveFile sFileName
End Sub

Private Sub mnuFileSave_Click()
      If Not agEditor.Visible Then
            frmMain.Caption = "ERROR: can only save in editor mode."
            Exit Sub
      ElseIf chkReadOnly.Value = vbChecked Then
            frmMain.Caption = "ERROR: can't save in Read Only mode."
            Exit Sub
      End If
      
      If agEditor.tag = "" Then  ' If they try to save a nameless New File
            mnuFileSaveAs_Click
      Else
            SaveFile agEditor.tag
      End If
End Sub

Private Sub mnuFile_Click()
      mnuFileSave.Enabled = True
      mnuFileSaveAs.Enabled = True
      If geEditorMode <> eViewMode.TextView Or chkReadOnly Then
            mnuFileSave.Enabled = False
            mnuFileSaveAs.Enabled = False
      End If
End Sub

Private Sub mnuHelpAbout_Click()
      frmAbout.Show vbModal
      ' MsgBox App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpReadme_Click()
      ShellExecute 0, "open", "https://github.com/phlegm-noir/phlegmoirs/blob/main/README.md", 0, 0, 1
End Sub

Private Sub mnuListCancel_Click()
      SendKeys "{ESC}"
End Sub

Private Sub mnuListCopyPath_Click()
      Clipboard.Clear
      If gtBrowserData.BookmarkMode Or gtBrowserData.HistoryMode Then
            Clipboard.SetText lvwBrowser.SelectedItem
      Else
            Clipboard.SetText gtBrowserData.Dir & lvwBrowser.SelectedItem
      End If
End Sub

Private Sub mnuListDelete_Click()
      BrowserDeleteSelected
End Sub

Private Sub mnuListHideFileBrowser_Click()
      mnuViewFilebrowser_Click
End Sub

Private Sub mnuListOpenDefault_Click()
      Dim sPath As String
      
      ' opens the file in whatever program windows chooses for it.
      If lvwBrowser.ListItems.Count > 0 Then
            If gtBrowserData.BookmarkMode Or gtBrowserData.HistoryMode Then
                  sPath = lvwBrowser.SelectedItem.Text
            Else
                  sPath = gtBrowserData.Dir & lvwBrowser.SelectedItem.Text
            End If
            ShellExecute 0, "open", sPath, 0, "", SW_RESTORE
      End If
End Sub

Private Sub mnuListOpen_Click()
      BrowserExecuteItem lvwBrowser.SelectedItem
End Sub

Private Sub mnuListProperties_Click()
      ' SImply calls the Explorer file properties dialog.  Hope this works.
      
      If gtBrowserData.BookmarkMode Or gtBrowserData.HistoryMode Then
            ShowFileProperties lvwBrowser.SelectedItem
      Else
            ShowFileProperties gtBrowserData.Dir & lvwBrowser.SelectedItem
      End If
End Sub

Private Sub mnuListRename_Click()
      ' I've decided to make history unchangeable.  It could have worked the other way,
      ' but it's one of those features that would make you more scared than impressed.
      
      ' Bookmarks are changeable, but it's rewriting the name of the link, not the name of the file.

      If gtBrowserData.HistoryMode Or lvwBrowser.Visible = False Then Exit Sub
      
      lvwBrowser.StartLabelEdit
      
End Sub

'   Show only files of extension sEx.
'
Private Sub mnuListShowOnly_Click()
      Dim sEx As String
      
      If gtBrowserData.BookmarkMode Or gtBrowserData.HistoryMode Then Exit Sub
      
      sEx = goFso.getextensionname(lvwBrowser.SelectedItem)
      If sEx <> "" Then sEx = "." & sEx
      cboPath = gtBrowserData.Dir & sEx
End Sub

Private Sub mnuList_Click()
      
      ' This is the popup menu for lvwBrowser.  Click fires whenever the menu is popped up.
      
      ' Most menu items are enabled/disabled in lvwBrowser_ItemClick.
      ' Here, we un-set some of them if the user has clicked somewhere that is not a list item.
      
      ' Events happen in this order: lvwBrowser_MouseDown, lvwBrowser_ItemClick, mnuList_Click.
      
      ' mbBrowserItemClicked is set to False on the MouseDown, and True on the ItemClick.
      ' So if it gets here as False, that means ItemClick did not happen on this mouse event.
      
'      If Not mbBrowserItemClicked Then
'            mnuListOpenDefault.Enabled = False
'            mnuListDelete.Enabled = False
'            mnuListRename.Enabled = False
'            mnuListCopyPath.Enabled = False
'            mnuListShowOnly.Enabled = False
'            mnuListProperties.Enabled = False
'      End If
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

Private Sub mnuQueryClose_Click()
      If Not mbHideFind Then
            mbHideFind = True
            picQuery.Visible = False
            RearrangeControls
      End If
End Sub

Private Sub mnuQueryMatchCase_Click()
      mnuQueryMatchCase.Checked = Not mnuQueryMatchCase.Checked
End Sub

Private Sub mnuQueryReplace_Click()
      ' This is where the mode actually switches
      If mbReplaceMode Then
            mbReplaceMode = False
            txtReplace.Visible = False
            txtFind.SetFocus
            
            btnFindNext.Move 480, 270, 1095, 300
            btnFindPrev.Move 1560, 270, 1095, 300
            chkFindOptions.Move 2640, 0, 375, 285
            btnReplace.Move 2640, 270, 375, 300
            
            mnuQueryReplace.Caption = "Show &Replace"
            btnReplace.Enabled = True
      Else
            mbReplaceMode = True
            txtReplace.Visible = True
            txtReplace.SetFocus
            txtReplace.SelStart = 0
            txtReplace.SelLength = Len(txtReplace)
            txtReplace_Change

            btnReplace.Move 2640, 295, 372, 285
            btnFindNext.Move 2640, 0, 372, 310
            btnFindPrev.Move 3000, 0, 372, 310
            chkFindOptions.Move 3000, 295, 372, 285
            chkFindOptions.ZOrder
            mnuQueryReplace.Caption = "Hide &Replace"
      End If
End Sub

Private Sub mnuQueryWholeWord_Click()
      mnuQueryWholeWord.Checked = Not mnuQueryWholeWord.Checked
End Sub

Private Sub mnuQuery_Click()
      mnuQueryReplace.Enabled = True
      If chkReadOnly Then mnuQueryReplace.Enabled = False
End Sub

Private Sub mnuViewFilebrowser_Click()
      chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub mnuViewFitImage_Click()
      mnuViewFitImage.Checked = Not mnuViewFitImage.Checked
      If mnuViewFitImage.Checked Then
            geImageSizingMode = eImageSizingMode.AlwaysFit
      Else
            geImageSizingMode = eImageSizingMode.Default100
      End If
End Sub

Private Sub mnuviewfont_Click()
      btnFont_Click
End Sub

Private Sub mnuViewHistory_Click()
      If mnuViewFilebrowser.Checked = False Then mnuViewFilebrowser_Click
      cboPath = "(History)"
End Sub

Private Sub mnuViewReadOnly_Click()
      chkReadOnly.Value = Abs(chkReadOnly.Value - 1)
End Sub

Private Sub mnuViewStatusBar_Click()
      staTusBar1.Visible = Not staTusBar1.Visible
      mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
      RearrangeControls
End Sub

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
            If Not mbHideFind Then picQuery.Visible = True
            mnuPlus.Caption = "="
            mnuNext.Visible = False
            mnuPrev.Visible = False
      End If
      RearrangeControls
End Sub

Private Sub mnuViewWordWrap_Click()
      chkWordWrap.Value = Abs(chkWordWrap.Value - 1)
End Sub

Private Sub mnuviewzoomin_Click()
'      btnZoomIn_MouseDown vbLeftButton, 0, 10, 10
      btnZoomIn_Click
End Sub

Private Sub mnuviewzoomout_Click()
'      btnZoomOut_MouseDown vbLeftButton, 0, 10, 10
      
      btnZoomOut_Click
End Sub

Private Sub mnuView_Click()
      mnuViewFont.Enabled = True
      mnuViewZoomIn.Enabled = True
      mnuViewZoomOut.Enabled = True
      mnuViewReadOnly.Enabled = True
      mnuViewWordWrap.Enabled = True
      
      If geEditorMode = eViewMode.PropertiesView Then
            mnuViewZoomIn.Enabled = False
            mnuViewZoomOut.Enabled = False
      End If
      
      If geEditorMode <> eViewMode.TextView Then
            mnuViewFont.Enabled = False
            mnuViewReadOnly.Enabled = False
            mnuViewWordWrap.Enabled = False
      End If
End Sub

Private Sub mnuWindowSaveSettings_Click()
      SaveSettingsToRegistry
End Sub

Private Sub mnuWriteCopy_Click()
      agEditor.Copy
End Sub

Private Sub mnuWriteCut_Click()
      agEditor.Cut
End Sub

Private Sub mnuWriteDelete_Click()
      agEditor.InsertContents SF_TEXT, ""
End Sub

Private Sub mnuWriteFind_Click()
      mnuEditFind_Click
End Sub

Private Sub mnuWritePaste_Click()
      agEditor.Paste
End Sub

Private Sub mnuWrite_Click()
      mnuWriteDelete.Enabled = True
      mnuWriteCut.Enabled = True
      mnuWriteCopy.Enabled = True
      mnuWritePaste.Enabled = True
      
      If chkReadOnly Then
            mnuWriteDelete.Enabled = False
            mnuWriteCut.Enabled = False
            mnuWritePaste.Enabled = False
      End If
      
      If agEditor.SelectedText = "" Then
            mnuWriteDelete.Enabled = False
            mnuWriteCut.Enabled = False
            mnuWriteCopy.Enabled = False
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
'   ParsePath translates input string sInput into referenced data structure rtBdata.
'   rtBdata holds the working directory, filter, previous directory, mode,
'   ...and much, much more!
'
Private Sub ParsePath(ByVal sInput As String, ByRef rtBdata As TBrowserData)
      ' (Bookmarks)      (that means bookmark mode, of course!)
      ' (History)           (History mode)
      '                            (a blank is intrepreted as "root" / drives list mode)
      ' c:\temp\  (just a plain old directory)
      ' c:\temp\.txt  (wildcard implied)
      ' c:\temp\READM*  (contains wildcard(s) after the directory, will filter the list)
      ' c:\temp\READMYLIPS  (no wildcard, won't filter but will move selection to a matching filename)
      
      Dim sFileName As String
      sInput = Trim(sInput)
      
      With rtBdata
      
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
                        If Not goFso.FolderExists(.Dir) Then .ValidPath = False
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
                        .Filter = "*." & goFso.getextensionname(sFileName)
                        
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

Private Sub picEditor_KeyDown(KeyCode As Integer, Shift As Integer)
      
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
                  ImageSetZoom 100
                  Image1.Move 0, 0, gtImageData.DefaultWidth, gtImageData.DefaultHeight
            Case 103, 55   ' 7 and Keypad 7
                  ImageSetZoom sliZoom.Value / 2
            Case 104, 56   ' 8 and Keypad 8
                  ImageSetZoom sliZoom.Value * 2
            Case vbKeyDown
                  Image1.Top = Image1.Top + MOVE_INCREMENT
            Case vbKeyUp
                  Image1.Top = Image1.Top - MOVE_INCREMENT
            Case vbKeyLeft
                  Image1.Left = Image1.Left - MOVE_INCREMENT
            Case vbKeyRight
                  Image1.Left = Image1.Left + MOVE_INCREMENT
                  
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
                  If Shift = 0 Then BrowserExecuteNext True
      End Select
End Sub

Private Sub picEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If FOCUS_FOLLOWS_MOUSE Then
            On Error Resume Next
            If GetForegroundWindow = frmMain.hWnd And Not (ActiveControl.Name = "picEditor") Then
                  picEditor.SetFocus
            End If
            On Error GoTo 0
      End If
End Sub

Private Sub picEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If geEditorMode = eViewMode.PictureView And Not gtImageData.Moved And Not gtImageData.Zoomed And _
            Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            BrowserExecuteNext
      ElseIf geEditorMode = eViewMode.PictureView And Not gtImageData.Moved And Not gtImageData.Zoomed And _
            Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            BrowserExecuteNext True
      ElseIf geEditorMode = eViewMode.PictureView And Not gtImageData.Moved And Not gtImageData.Zoomed And _
            Button = vbMiddleButton Then
            
            btnFullScreen_Click
      End If
      
      gtImageData.Zoomed = False
      gtImageData.Moved = False
End Sub

Private Sub RearrangeControls()

      ' Put the various controls where they need to be.
      '   agEditor, lvwBrowser
      ' Made to go on a window resize or when showing or hiding a control
      
      Dim iEdHeight As Integer, iEdWidth As Integer, iEdTop As Integer, iEdLeft As Integer
      Dim iToolbarFullWidth As Integer
      Dim iBrowserTop, iBrowserHeight As Integer
      Dim lLineIndex As Long, lCharIndex As Long, lMin As Long, lMax As Long
      Dim bValidWindowSize As Boolean, iRedoResizeX As Integer, iRedoResizeY As Integer
      Dim iPicBoxMarginsY As Integer, iFormMarginsX As Integer, iFormMarginsY As Integer
      Dim sHadFocus As String
      
      Const TOP_MARGIN = 100
      Const LEFT_MARGIN = 0
      Const BOTTOM_MARGIN = 30
      Const TOOLBAR_WIDTH = 4905
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      bValidWindowSize = True ' ...until proven guilty.
      iRedoResizeY = frmMain.Height
      iRedoResizeX = frmMain.Width
      
      If Not (ActiveControl Is Nothing) Then  ' activecontrol is nothing if image1 is up front...
            sHadFocus = ActiveControl.Name                               ' images cannot take focus.
            picEditor.Visible = False ' MUCH faster if you turn him off while thinking (unless he's empty).
      End If
      
      ' Calculate control positions...
      
      iEdLeft = LEFT_MARGIN
      If mnuViewFilebrowser.Checked Then iEdLeft = iEdLeft + picBrowser.Width
      
      iEdWidth = frmMain.ScaleWidth - iEdLeft
      
      If mnuViewToolbar.Checked Then
            iBrowserTop = picToolBar.Height
            If Not mbHideFind And picQuery.Visible Then
                  iToolbarFullWidth = TOOLBAR_WIDTH + picQuery.Width
            Else
                  iToolbarFullWidth = TOOLBAR_WIDTH
            End If
      Else
            iToolbarFullWidth = 0
            iBrowserTop = 0
      End If
      
      iEdTop = 0
      If mnuViewToolbar.Checked And (Not mnuViewFilebrowser.Checked Or picBrowser.Width < iToolbarFullWidth) Then
            iEdTop = TOP_MARGIN + picToolBar.Height
      End If
      
      iEdHeight = frmMain.ScaleHeight - iEdTop - BOTTOM_MARGIN
      iBrowserHeight = frmMain.ScaleHeight - iBrowserTop - BOTTOM_MARGIN
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
            bValidWindowSize = False
            iFormMarginsX = frmMain.Width - frmMain.ScaleWidth
            iRedoResizeX = iEdLeft + iFormMarginsX + 1510
      End If
      
      If iEdHeight < 1500 Then
            bValidWindowSize = False
            iFormMarginsY = frmMain.Height - frmMain.ScaleHeight
            iPicBoxMarginsY = picBrowser.Height - picBrowser.ScaleHeight
            iRedoResizeY = iEdTop + lvwBrowser.Top + iPicBoxMarginsY + iFormMarginsY + 1510
      End If
      
      If Not bValidWindowSize Then
            frmMain.Move Left, Top, iRedoResizeX, iRedoResizeY
            Exit Sub
      End If
      
      ' It's all good.  Move the controls now!
      
      picBrowser.Move 0, iBrowserTop, picBrowser.Width, iBrowserHeight
      picEditor.Move iEdLeft, iEdTop, iEdWidth, iEdHeight
      agEditor.Move 0, 0, iEdWidth, iEdHeight
      lvwBrowser.Height = iBrowserHeight - lvwBrowser.Top + TOP_MARGIN


      If geEditorMode = eViewMode.TextView Then
            ' a few things in the statusbar could change in a window resize:
            '   x, xmax, y, ymax
            ' and some shouldn't change:
            '   i, imax,   (we're not adding or deleting characters or moving the cursor)
            '   sellength
            
            agEditor.GetSelection lMin, lMax
            lLineIndex = agEditor.CurrentLine
            lCharIndex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lLineIndex, 0)
            
            If staTusBar1.Visible Then
                  With gtStats
                        .X = lMin - lCharIndex + 1
                        .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
                        .Y = lLineIndex + 1
                        .ymax = agEditor.LineCount
                  End With
                  FillStats
            End If
      ElseIf geEditorMode = eViewMode.PictureView And geImageSizingMode = eImageSizingMode.AlwaysFit Then
            btnFitImage_Click
      End If
      staTusBar1.Panels(eStat.Tips).Width = frmMain.Width
      
      picEditor.Visible = True
End Sub

Private Sub RefreshAll()
      With gtBrowserData
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
                  BrowserGetFilesAndFolders gtBrowserData
            End If
      End With
      
      BrowserAutoSelectListItem gtBrowserData
End Sub

Private Function RenameFile(sOldPath As String, sNewPath As String, sIfOpenFile As String, sIfOtherFile As String) As Boolean
      On Error Resume Next
      Dim bCancel As Boolean
      Name sOldPath As sNewPath
      If Err > 0 Then
            frmMain.Caption = Err.Number & ": " & Err.Description
            DebugLog frmMain.Caption
            bCancel = True
      ElseIf sOldPath = agEditor.tag Then
            frmMain.Caption = sIfOpenFile
            agEditor.tag = sNewPath
      Else
            frmMain.Caption = sIfOtherFile
      End If
      On Error GoTo 0
      RenameFile = bCancel
End Function

Private Function RenameFileWithChecks(sOldPath As String, sNewPath As String) As Boolean
      Dim bCancel As Boolean
      
      If Not FileExists(sOldPath) Then
            frmMain.Caption = "Can't rename what's not there: " & sOldPath
            RefreshAll
            bCancel = True
      
      ElseIf StrComp(sOldPath, sNewPath, vbBinaryCompare) = 0 Then
            
            bCancel = True  ' No change whatsoever.
      
      ElseIf StrComp(sOldPath, sNewPath, vbTextCompare) = 0 Then
            
            bCancel = RenameFile(sOldPath, sNewPath, _
                  "Adjusted the capitalization of open file to: " & sNewPath, _
                  "Renamed.  Even though all you changed was the capitalization. Freak.")
      
      ElseIf FileExists(sNewPath) Then
            frmMain.Caption = "This name sucks: " & Chr(34) & sNewPath & Chr(34) & ".  Change it."
            bCancel = True
      Else
            bCancel = RenameFile(sOldPath, sNewPath, _
                  "Renamed open file: " & sNewPath, "Rename successful: " & sNewPath)
      End If
      
      If Not bCancel Then
            lvwBrowser.SelectedItem.Text = goFso.GetFileName(sNewPath)
            'RefreshAll
            btnSyncContents_Click
      End If
      RenameFileWithChecks = bCancel
End Function

Public Function SaveFile(ByVal sFileName As String)
      Dim bSuccess As Boolean, bNewFile As Boolean
      
      If Len(sFileName) > 100 Or agEditor.Text = "" Or giTextEncoding = eTextEncoding.UNICODE Then
            Dim oStream
            On Error GoTo File_Error
            If Not FileExists(sFileName) Then
                  bNewFile = True
                  Set oStream = goFso.CreateTextFile(sFileName, eOverwrite.Yes, giTextEncoding)
            Else
                  Set oStream = goFso.OpenTextFile(sFileName, eIoMode.ForWriting, eCreate.No, giTextEncoding)
            End If
            If agEditor.Text = "" And Not bNewFile Then
                  oStream.Write ("temporary text to make sure the file counts as modified")
                  oStream.Close
                  Set oStream = goFso.OpenTextFile(sFileName, eIoMode.ForWriting)
                  ' TODO: titlebar will show a false positive that it wrote to file in this ONE scenario
                  '     * already existing file
                  '     * we do not have permission to write to it
                  '     * we are trying to save a blank file of exactly 0 bytes
                  '     ...this is just too niche to care about anymore
            End If
            oStream.Write (agEditor.Text)
            oStream.Close
            On Error GoTo 0
            bSuccess = True
      Else
            bSuccess = agEditor.SaveToFile(sFileName, SF_TEXT)
      End If

      If bSuccess Then
            Dim lBytes As Long
            lBytes = agEditor.CharacterCount
            staTusBar1.Panels(eStat.encoding) = "ASCII"
            If giTextEncoding = eTextEncoding.UNICODE Then
                  lBytes = lBytes * 2 + 2
                  staTusBar1.Panels(eStat.encoding) = "UNICODE"
            End If
            staTusBar1.Panels(eStat.Modified) = ""
            agEditor.tag = sFileName
            frmMain.Caption = sFileName & "  (" & Format(lBytes, "#,#0") & " bytes saved on " _
                  & FileModifiedTime(sFileName) & ")"
            RefreshAll
            btnSyncContents_Click
            AddToHistorySmartly sFileName
      Else
            frmMain.Caption = "ERROR: cannot save to " & sFileName
      End If
      SaveSettingsToRegistry
      Exit Function
      
File_Error:
      frmMain.Caption = "ERROR: cannot save to " & sFileName
End Function

Private Sub SaveSettingsToRegistry()
      Dim lKey As Long, lRetVal As Long
      Dim lNewOrUsed As Long, lValueSize As Long
      Dim sErrorMsg As String
      Dim tPrefs As TAllPrefs
      
      GatherWindowPrefs tPrefs.WindowPrefs
      GatherBrowserPrefs tPrefs.BrowserPrefs
      GatherEditorPrefs tPrefs.EditorPrefs
      GatherHistoryAndBookmarks tPrefs
            
      ' Create storage key.
      
      lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, gsPhlegmKey, 0, "", 0, _
            KEY_ALL_ACCESS, ByVal 0, lKey, lNewOrUsed)
      If lRetVal <> 0 Then MsgBox "RegCreateKey Failed: " & lKey
      
      ' Store the Settings.
      
      lValueSize = LenB(tPrefs)
      lRetVal = RegSetValueExAny(lKey, "Settings", 0, REG_NONE, ByVal tPrefs, lValueSize)
      If lRetVal <> 0 Then
            sErrorMsg = "RegSetValueEx Failed. settings: " & LenB(tPrefs) & " " & lKey
            DebugLog sErrorMsg, 2
            MsgBox sErrorMsg, , App.title
      End If

      lRetVal = RegCloseKey(lKey)
      Debug.Print "Settings saved at: " & Now
End Sub

Private Sub ShowFileProperties(ByVal sPath As String)
      ' SImply calls the Explorer file properties dialog.  Hope this works.
      
      Dim tExInfo As SHELLEXECUTEINFO
            
      tExInfo.cbSize = LenB(tExInfo)
      tExInfo.lpFile = sPath
      tExInfo.lpVerb = "properties"
      tExInfo.fMask = SEE_MASK_INVOKEIDLIST
      
      ShellExecuteEx tExInfo
End Sub

Private Sub sliZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If FOCUS_FOLLOWS_MOUSE Then
            On Error Resume Next
            If GetForegroundWindow = frmMain.hWnd And Not (ActiveControl.Name = "sliZoom") Then
                  sliZoom.SetFocus
            End If
            On Error GoTo 0
      End If
End Sub

Private Sub sliZoom_Scroll()
      ImageSetZoom (sliZoom.Value)
End Sub

Private Sub txtFind_Change()
      If mbReplaceMode And StrComp(txtFind, agEditor.SelectedText, GetFindCompareMode()) = 0 And txtFind <> "" Then
            If Not chkReadOnly Then btnReplace.Enabled = True
            If ActiveControl.Name = "txtReplace" Then btnReplace.Default = True
      ElseIf mbReplaceMode Then
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

Private Sub txtReplace_Change()
      If mbReplaceMode And StrComp(txtFind, agEditor.SelectedText, GetFindCompareMode()) = 0 And txtFind <> "" Then
            If Not chkReadOnly Then btnReplace.Enabled = True
      ElseIf mbReplaceMode Then
            btnReplace.Enabled = False
      End If
End Sub

Private Sub txtReplace_GotFocus()
      If btnReplace.Enabled Then
            btnReplace.Default = True
      Else
            btnFindNext.Default = True
      End If
End Sub

Private Sub txtReplace_LostFocus()
      btnReplace.Default = False
      btnFindNext.Default = False
End Sub

Public Sub WheelInput(iWheelTurn As Integer, iVirtKeys As Integer, lx As Long, ly As Long)
      ' This is called from modPhlegmoirs.TrackMouseWheel
      ' It acts on picEditor while in picture mode.
      
      Dim lWheelMoveIncrement As Long
      ' lWheelMoveIncrement will be the positive distance that the wheel moves a picture.
      lWheelMoveIncrement = -MOVE_INCREMENT * 3 * Abs(CLng(iWheelTurn)) * sliZoom.Value / 100
      
      With gtImageData.OutPic
            ' Wheel scroll up = move picture down = make Top value HIGHER
            ' ...but not to rise above zero.
            If iVirtKeys = 0 And iWheelTurn > 0 Then
                  If gtImageData.OutPic.Height > gtImageData.SurroundingBox.ScaleHeight Then
                        If .Top < -lWheelMoveIncrement Then
                              .Top = .Top + lWheelMoveIncrement
                        ElseIf .Top < 0 Then
                              .Top = 0
                        End If
                  Else
                        BrowserExecuteNext True
                  End If
            
            ElseIf iVirtKeys = 0 And iWheelTurn < 0 Then
                  ' Wheel scroll down = move picture up = make Top value LOWER.
                  ' ...the bottom value not to fall below the bottom value of its container.
                  
                  If gtImageData.OutPic.Height > gtImageData.SurroundingBox.ScaleHeight Then
                        If .Top + .Height > .Container.Height + lWheelMoveIncrement Then
                              .Top = .Top - lWheelMoveIncrement
                        ElseIf .Top + .Height > .Container.Height Then
                              .Top = .Container.Height - .Height
                        End If
                  Else
                        BrowserExecuteNext
                  End If
                  
            ElseIf iVirtKeys = MK_LBUTTON Then ' Right mouse button + wheel scroll
                  ' Move picture right/left
                  .Left = .Left - iWheelTurn * MOVE_INCREMENT * 3
                  gtImageData.Moved = True
                  
            ElseIf iVirtKeys = MK_MBUTTON Then ' Hold down the wheel while spinning it (if you can even do that)
                  
                  ' Picture Zoom, large increment
                  Dim iPresses As Integer
                  ' So we'll be lazy and just press the appropriate zoom button once for each mouse turn.
                  For iPresses = 1 To Abs(iWheelTurn)
                        If iWheelTurn > 0 Then
                              btnZoomIn_Click
                                          
                        ElseIf iWheelTurn < 0 Then
                              btnZoomOut_Click
                        End If
                  Next iPresses
                  gtImageData.Zoomed = True
                  If gbFullScreenMode Then frmFullScreen.lblFileNameZoom = Caption & "  "
                  
            ElseIf iVirtKeys = MK_RBUTTON Then ' Left mouse button + wheel scroll
                  
                  ' Picture zoom, small increment
                  ImageSetZoom sliZoom.Value + iWheelTurn * sliZoom.SmallChange
                  gtImageData.Zoomed = True
                  If gbFullScreenMode Then frmFullScreen.lblFileNameZoom = Caption & "  "
            End If
      End With
End Sub

