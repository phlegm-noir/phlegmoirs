VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMane 
   Caption         =   "phlegmoirs"
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   12285
   Icon            =   "frmMane.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin phlegmoirs.RechEdit Editor 
      Height          =   8070
      Left            =   4215
      TabIndex        =   16
      Top             =   615
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   14235
   End
   Begin phlegmoirs.PhlegmoFiler Filer 
      Height          =   8070
      Left            =   0
      TabIndex        =   15
      Top             =   615
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   14235
   End
   Begin MSComctlLib.StatusBar staTusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   8685
      Width           =   12285
      _ExtentX        =   21669
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
   Begin VB.PictureBox picToolBar 
      ClipControls    =   0   'False
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   4770
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4830
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
         Picture         =   "frmMane.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMane.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMane.frx":1ACE
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMane.frx":1E10
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMane.frx":2152
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmMane.frx":2494
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   260
         UseMaskColor    =   -1  'True
         Width           =   460
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
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Set Font (Shift+Ctrl+F)"
         Top             =   0
         Width           =   1815
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
         Picture         =   "frmMane.frx":27D6
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMane.frx":2ED8
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "New File (Ctrl+N)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         DownPicture     =   "frmMane.frx":35DA
         Height          =   570
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMane.frx":3CDC
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Show/Hide the File Browser (F8)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
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
         Picture         =   "frmMane.frx":43DE
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Full Screen (F11)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
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
         Picture         =   "frmMane.frx":4720
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Edit This File"
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.Slider sliZoom 
         Height          =   330
         Left            =   1800
         TabIndex        =   6
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
         TabIndex        =   14
         Top             =   320
         Width           =   960
      End
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
End
Attribute VB_Name = "frmMane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MIN_EDITOR_WIDTH = 3000
Const MIN_HEIGHT = 3300
Const TOP_MARGIN = 105
Private mlFormMarginsHoriz As Long
Private mePrevWindowState As FormWindowStateConstants
Private mtPrevWindowPos As POINTAPI
Private mtPrevWindowSize As POINTAPI

Private Sub chkFileBrowser_Click()
      Filer.Visible = (vbChecked = chkFileBrowser.Value)
      DoEvents
      mnuViewFilebrowser.Checked = Filer.Visible
      
      FormAdjustWidthIfTooSmall
      Form_Resize
      SizeLimiterHook Me.hWnd, GetFilerWidth() + MIN_EDITOR_WIDTH, MIN_HEIGHT
End Sub

Private Sub Filer_ResizeHorizontal(ByVal lWidth As Long)
      Form_Resize
End Sub

Private Sub Filer_SeriousResize(ByVal lWidth As Long)
      Form_Resize
      SizeLimiterHook Me.hWnd, lWidth + MIN_EDITOR_WIDTH
End Sub

Private Sub Form_Load()
      mlFormMarginsHoriz = Me.Width - Me.ScaleWidth
      SizeLimiterHook Me.hWnd, GetFilerWidth() + MIN_EDITOR_WIDTH, MIN_HEIGHT
End Sub

Private Sub Form_Resize()
      If (mePrevWindowState <> vbNormal) And (Me.WindowState = vbNormal) Then
            FormAdjustWidthIfTooSmall
      End If
      mePrevWindowState = Me.WindowState
      mtPrevWindowPos.X = Me.Left
      mtPrevWindowPos.Y = Me.Top
      mtPrevWindowSize.X = Me.Width
      mtPrevWindowSize.Y = Me.Height
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      Dim lFilerHeight, lEditorTop As Long
      lFilerHeight = Me.ScaleHeight - GetToolbarHeight - GetStatusbarHeight
      lEditorTop = Ternary(GetFilerWidth > picToolBar.Width + 105, 0, GetToolbarHeight)
      
      Editor.Move GetFilerWidth, lEditorTop, Me.Width - GetFilerWidth - mlFormMarginsHoriz, Me.ScaleHeight - GetStatusbarHeight - lEditorTop
      If Filer.Visible Then
            Filer.Move Filer.Left, GetToolbarHeight, Filer.Width, lFilerHeight
      End If
      Filer.SetMaxWidth (Me.Width - MIN_EDITOR_WIDTH - mlFormMarginsHoriz)
End Sub

Private Sub FormAdjustWidthIfTooSmall()
      If GetFilerWidth() + MIN_EDITOR_WIDTH > frmMane.ScaleWidth Then
            frmMane.Width = Filer.Width + MIN_EDITOR_WIDTH
      End If
End Sub

Private Function GetFilerWidth() As Long
      GetFilerWidth = Ternary(chkFileBrowser, Filer.Width + Filer.Left, 0)
End Function

Private Function GetStatusbarHeight() As Long
      GetStatusbarHeight = Ternary(mnuViewStatusBar.Checked, staTusBar1.Height, 0)
End Function

Private Function GetToolbarHeight() As Long
      GetToolbarHeight = Ternary(mnuViewToolbar.Checked, picToolBar.Height + picToolBar.Top + TOP_MARGIN, 0)
End Function

Private Sub Form_Unload(Cancel As Integer)
      SizeLimiterUnhook Me.hWnd
      Unload frmAbout
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

Private Sub mnuViewFilebrowser_Click()
      chkFileBrowser.Value = Ternary(mnuViewFilebrowser.Checked, vbUnchecked, vbChecked)
End Sub

Private Sub mnuViewStatusBar_Click()
      mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
      staTusBar1.Visible = mnuViewStatusBar.Checked
      Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
      mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
      picToolBar.Visible = mnuViewToolbar.Checked
      mnuPlus.Caption = Ternary(mnuViewToolbar.Checked, "=", "+")
      mnuNext.Visible = Not mnuViewToolbar.Checked
      mnuPrev.Visible = Not mnuViewToolbar.Checked
      Form_Resize
End Sub
