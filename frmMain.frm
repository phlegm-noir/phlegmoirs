VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.3#0"; "VBCCR17.OCX"
Begin VB.Form frmMain 
   Caption         =   "phlegmoirs"
   ClientHeight    =   10005
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin phlegmoirs.PhlegmoFinder Finder 
      Height          =   600
      Left            =   4920
      TabIndex        =   18
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1058
   End
   Begin phlegmoirs.MadProps Props 
      Height          =   5655
      Left            =   4800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9975
   End
   Begin phlegmoirs.PhlegmoFoto Foto 
      Height          =   5895
      Left            =   5640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10398
   End
   Begin phlegmoirs.RetchEdit Editor 
      Height          =   8070
      Left            =   4215
      TabIndex        =   15
      Top             =   615
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   14235
   End
   Begin phlegmoirs.PhlegmoFiler Filer 
      Height          =   8070
      Left            =   0
      TabIndex        =   14
      Top             =   615
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   14235
   End
   Begin VBCCR17.StatusBar staTusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      Top             =   9645
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitPanels      =   "frmMain.frx":0CCA
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
         Picture         =   "frmMain.frx":159A
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMain.frx":1C9C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMain.frx":239E
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmMain.frx":26E0
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMain.frx":2A22
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMain.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   260
         UseMaskColor    =   -1  'True
         Width           =   460
      End
      Begin VB.CommandButton btnFont 
         Caption         =   "Lucida Sans Unicode"
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
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Set Font (Shift+Ctrl+F)"
         Top             =   0
         Width           =   1815
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
         Picture         =   "frmMain.frx":30A6
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMain.frx":33E8
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Edit This File"
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
         Picture         =   "frmMain.frx":372A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmMain.frx":3E2C
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
         DownPicture     =   "frmMain.frx":452E
         Height          =   570
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":4C30
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Show/Hide the File Browser (F8)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VBCCR17.Slider sliZoom 
         Height          =   300
         Left            =   1800
         TabIndex        =   19
         Top             =   0
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         Max             =   500
         Value           =   100
         TickFrequency   =   100
         SmallChange     =   10
         LargeChange     =   100
         TickStyle       =   3
         TipSide         =   1
      End
      Begin VB.Label lblFontSize 
         Alignment       =   2  'Center
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2240
         TabIndex        =   11
         Top             =   300
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MIN_EDITOR_WIDTH = 3000
Const MIN_HEIGHT = 3300
Const TOP_MARGIN = 105
Const FIND_LEFT_MARGIN = 105

Private mlFormMarginsHoriz As Long
Private mePrevWindowState As FormWindowStateConstants
Private mtPrevWindowPos As POINTAPI
Private mtPrevWindowSize As POINTAPI

Private msFileName As String
Private mvView As Variant  ' holds the active usercontrol on the right side of the screen
Private meMode As eViewMode
Private mbHideFind As Boolean

Private mtCharStats
Private msCommandFile As String
Private mbFullScreenMode As Boolean
Private meImageSizingMode As eImageSizingMode

Private Sub btnFont_Click()
      Dim sName As String
      Dim iSize As Integer
      Editor.OpenFontDialog sName, iSize
      If iSize > 0 Then
            If Len(sName) > 18 Then
                  btnFont.Caption = Left(Trim(sName), 17) & "..."
            Else
                  btnFont.Caption = sName
            End If
            lblFontSize = iSize
      End If
End Sub

Private Sub btnNextFile_Click()
      If meMode <= 3 Then SetMode meMode + 1
End Sub

Private Sub btnPrevFile_Click()
      If meMode >= 1 Then SetMode meMode - 1
End Sub

Private Sub btnZoomIn_Click()
      lblFontSize = Editor.NextFontSize
End Sub

Private Sub btnZoomOut_Click()
      lblFontSize = Editor.PrevFontSize
End Sub

Private Sub chkFileBrowser_Click()
      Filer.Visible = (vbChecked = chkFileBrowser.Value)
      DoEvents
      mnuViewFilebrowser.Checked = Filer.Visible
      
      FormAdjustWidthIfTooSmall
      Form_Resize
      SizeLimiterHook Me.hwnd, GetFilerWidth() + MIN_EDITOR_WIDTH, MIN_HEIGHT
End Sub

Private Sub Filer_HoverItem(ByVal sItemText As String)
      staTusBar.Panels(eStat.StatTips).Text = sItemText
End Sub

Private Sub Filer_OpenFile(ByVal sFileName As String)
      Editor.LoadFile sFileName
End Sub

Private Sub Filer_ResizeHorizontal(ByVal lWidth As Long)
      Form_Resize
End Sub

Private Sub Filer_SeriousResize(ByVal lWidth As Long)
      Form_Resize
      SizeLimiterHook Me.hwnd, lWidth + MIN_EDITOR_WIDTH
End Sub

Private Sub Filer_StatsUpdate(ByVal sFilerStats As String)
      staTusBar.Panels(eStat.BrowserStats).Text = sFilerStats
End Sub

Private Sub Finder_Closing()
      Finder.Visible = False
      Form_Resize
End Sub

Private Sub Form_Load()
      DebugLog ""
      DebugLog App.title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " starting at " & Now
      Set mvView = Editor
      mlFormMarginsHoriz = Me.Width - Me.ScaleWidth
      SizeLimiterHook Me.hwnd, GetFilerWidth() + MIN_EDITOR_WIDTH, MIN_HEIGHT
      SetMode (eViewMode.TextView)
End Sub

Private Sub Form_Resize()
      If (mePrevWindowState <> vbNormal) And (Me.WindowState = vbNormal) Then
            FormAdjustWidthIfTooSmall
      End If
      mePrevWindowState = Me.WindowState
      mtPrevWindowPos.x = Me.Left
      mtPrevWindowPos.y = Me.Top
      mtPrevWindowSize.x = Me.Width
      mtPrevWindowSize.y = Me.Height
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      Dim lFilerHeight, lViewTop As Long
      lFilerHeight = Me.ScaleHeight - GetToolbarHeight - GetStatusbarHeight
      lViewTop = Ternary(GetFilerWidth > GetToolbarWidth + FIND_LEFT_MARGIN, 0, GetToolbarHeight)
      
      mvView.Move GetFilerWidth, lViewTop, Me.Width - GetFilerWidth - mlFormMarginsHoriz, Me.ScaleHeight - GetStatusbarHeight - lViewTop
      If Filer.Visible Then
            Filer.Move Filer.Left, GetToolbarHeight, Filer.Width, lFilerHeight
      End If
      Filer.SetMaxWidth (Me.Width - MIN_EDITOR_WIDTH - mlFormMarginsHoriz)
      
      ' Workaround because the statusbar refuses to update its maximized width AND refuses to let you set a width when aligned-bottom
      If Me.WindowState = vbMaximized Then
            staTusBar.Align = vbAlignNone
            staTusBar.Width = Me.ScaleWidth
            staTusBar.Align = vbAlignBottom
      End If
      staTusBar.Panels(eStat.StatTips).Width = Max(0, staTusBar.Width - staTusBar.Panels(eStat.StatTips).Left)
      DebugLog "Statusbar width: " & staTusBar.Width & "; Form scalewidth: " & Me.ScaleWidth
End Sub

Private Sub FormAdjustWidthIfTooSmall()
      If GetFilerWidth() + MIN_EDITOR_WIDTH > Me.ScaleWidth Then
            Me.Width = Filer.Width + MIN_EDITOR_WIDTH
      End If
End Sub

Private Function GetFilerWidth() As Long
      GetFilerWidth = Ternary(chkFileBrowser, Filer.Width + Filer.Left, 0)
End Function

Private Function GetStatusbarHeight() As Long
      GetStatusbarHeight = Ternary(mnuViewStatusBar.Checked, staTusBar.Height, 0)
End Function

Private Function GetToolbarHeight() As Long
      GetToolbarHeight = Ternary(mnuViewToolbar.Checked, picToolBar.Height + picToolBar.Top + TOP_MARGIN, 0)
End Function

Private Function GetToolbarWidth() As Long
      GetToolbarWidth = Ternary(Finder.Visible, Finder.Left + Finder.Width - picToolBar.Left, picToolBar.Width)
End Function

Private Sub Form_Unload(Cancel As Integer)
      SizeLimiterUnhook Me.hwnd
      Unload frmAbout
End Sub

Private Sub mnuBrowserRefresh_Click()
      Editor.ForceRefresh
End Sub

Private Sub mnuFileSave_Click()
      Editor.Save
End Sub

Private Sub mnuHelpAbout_Click()
      frmAbout.Show vbModal
End Sub

Private Sub mnuHelpReadme_Click()
      ShellExecute 0, "open", "https://github.com/phlegm-noir/phlegmoirs/blob/main/README.md", 0, 0, 1
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
      staTusBar.Visible = mnuViewStatusBar.Checked
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

Private Sub SetMode(eMode As eViewMode)
      
      meMode = eMode
      Select Case eMode
            Case eViewMode.TextView
      
                  Editor.Visible = True
                  Foto.Visible = False
                  Props.Visible = False
                  Set mvView = Editor
'                  Image1.Picture = LoadPicture

                  btnFont.Visible = True
                  btnFullScreen.Visible = False
                  If mnuViewReadOnly.Checked Then btnEdit.Visible = True
                  If Not mbHideFind And mnuViewToolbar.Checked Then Finder.Visible = True
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3000, 260, 615, 320
                  btnZoomOut.Move 1800, 260, 615, 320
                  btnZoomDefault.Visible = False
                  btnFitImage.Visible = False
                  
                  staTusBar.Panels(eStat.Encoding).Visible = True
                  staTusBar.Panels(eStat.Modified).Visible = True
                  staTusBar.Panels(eStat.CharStats).Visible = True
                  staTusBar.Panels(eStat.SelText).Visible = True
      
            Case eViewMode.PictureView
                  Editor.Visible = False
                  Foto.Visible = True
                  Props.Visible = False
                  Set mvView = Foto
                  
                  btnFont.Visible = False
                  btnFullScreen.Visible = True
                  btnEdit.Visible = False
                  If Not mbHideFind Then Finder.Visible = False
                  
                  sliZoom.Visible = True
                  btnZoomIn.Move 3150, 260, 470, 320
                  btnZoomOut.Move 1800, 260, 460, 320
                  btnZoomDefault.Visible = True
                  btnFitImage.Visible = True
                  
'                  If glOldpicEditorProc = 0 Then
'                        glOldpicEditorProc = SetWindowLong(picEditor.hWnd, GWL_WNDPROC, _
'                              AddressOf TrackMouseWheel)
'                  End If

                  staTusBar.Panels(eStat.Encoding).Visible = False
                  staTusBar.Panels(eStat.Modified).Visible = False
                  staTusBar.Panels(eStat.CharStats).Visible = False
                  staTusBar.Panels(eStat.SelText).Visible = False
                  
            Case eViewMode.PropertiesView
            
                  Editor.Visible = False
'                  agEditor.Text = ""
                  Foto.Visible = False
'                  Image1.Picture = LoadPicture
                  Props.Visible = True
                  Set mvView = Props
                  
                  btnFont.Visible = True
                  btnFullScreen.Visible = False
                  btnEdit.Visible = False
                  If Not mbHideFind Then Finder.Visible = False
                  
                  sliZoom.Visible = False
                  btnZoomIn.Move 3000, 260, 615, 320
                  btnZoomOut.Move 1800, 260, 615, 320
                  btnZoomDefault.Visible = False
                  btnFitImage.Visible = False
                  
                  staTusBar.Panels(eStat.Encoding).Visible = False
                  staTusBar.Panels(eStat.Modified).Visible = False
                  staTusBar.Panels(eStat.CharStats).Visible = False
                  staTusBar.Panels(eStat.SelText).Visible = False
            
            Case Else
                  DebugLog "How did we get to the ERROR ViewMode? Filename: """ + msFileName + """", 2
      End Select
      
      Form_Resize
End Sub


