VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DD32A320-6E5E-44C8-BCE6-5908CA400231}#1.0#0"; "AGRICHEDIT.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "phlegmoirs"
   ClientHeight    =   6465
   ClientLeft      =   135
   ClientTop       =   540
   ClientWidth     =   10080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   10080
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0452
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0CF6
            Key             =   "textfile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1148
            Key             =   "otherfile"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFileBrowser 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      Height          =   5352
      Left            =   60
      MousePointer    =   1  'Arrow
      ScaleHeight     =   5355
      ScaleWidth      =   2415
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   712
      Width           =   2412
      Begin VB.CommandButton btnDeleteSelected 
         Appearance      =   0  'Flat
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   264
         Left            =   1320
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Delete"
         Top             =   300
         Width           =   264
      End
      Begin VB.CommandButton btnPathForward 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   264
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":2292
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Go forward a directory (Alt+Right)"
         Top             =   300
         Width           =   264
      End
      Begin VB.ComboBox cboPath 
         Height          =   288
         ItemData        =   "Main.frx":23DC
         Left            =   0
         List            =   "Main.frx":23DE
         TabIndex        =   9
         Text            =   "*"
         Top             =   0
         Width           =   2292
      End
      Begin MSComctlLib.ListView lvwBrowser 
         Height          =   4335
         Left            =   0
         TabIndex        =   10
         Tag             =   "c:\test\"
         Top             =   600
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
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsFileIcons"
         SmallIcons      =   "ilsFileIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
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
      Begin VB.CommandButton btnSort 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   792
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":23E0
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Reverse the sort order"
         Top             =   300
         Width           =   264
      End
      Begin VB.CommandButton btnCurrentDirectory 
         Appearance      =   0  'Flat
         Caption         =   "<>"
         Height          =   264
         Left            =   1584
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Jump to the directory containing your open file..."
         Top             =   300
         Width           =   264
      End
      Begin VB.CommandButton btnFolderUp 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   528
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Go up a directory (Left arrow key)"
         Top             =   300
         Width           =   264
      End
      Begin VB.CommandButton btnRefresh 
         Appearance      =   0  'Flat
         Caption         =   "R"
         Height          =   264
         Left            =   1056
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Refresh Files"
         Top             =   300
         Width           =   264
      End
      Begin VB.CommandButton btnPathBack 
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   0
         MaskColor       =   &H80000000&
         Picture         =   "Main.frx":286C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Go back a directory (Alt+Left)"
         Top             =   300
         Width           =   264
      End
      Begin VB.Label lblDivider 
         BackStyle       =   0  'Transparent
         Height          =   25005
         Left            =   2295
         TabIndex        =   16
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   7920
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   12
   End
   Begin VB.PictureBox picToolBox 
      ClipControls    =   0   'False
      Height          =   612
      Left            =   -130
      ScaleHeight     =   555
      ScaleWidth      =   8475
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8532
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
         Height          =   580
         Left            =   1560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   732
      End
      Begin VB.TextBox txtQueryBox 
         Height          =   372
         Left            =   3480
         MaxLength       =   50
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Text            =   "query box"
         Top             =   120
         Width           =   3732
      End
      Begin VB.CheckBox chkDummy 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   732
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   120
         Picture         =   "Main.frx":29B6
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   732
      End
   End
   Begin MSComctlLib.StatusBar staTusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6195
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
            TextSave        =   "Char: 0/00000  Ln: 0/000  Col: 0/000"
            Key             =   "statStats"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Modified"
            TextSave        =   "Modified"
            Key             =   "statModified"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Ins"
            TextSave        =   "Ins"
            Key             =   "statInsert"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Seltext: 000"
            TextSave        =   "Seltext: 000"
            Key             =   "statSelText"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Tips and things"
            TextSave        =   "Tips and things"
            Key             =   "statTips"
         EndProperty
      EndProperty
   End
   Begin agRichEditBox.agRichEdit agEditor 
      Height          =   5352
      Left            =   2572
      TabIndex        =   4
      Top             =   712
      Width           =   5858
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
      DisableNoScroll =   -1  'True
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
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
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
         Shortcut        =   ^U
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
         Caption         =   "&Dictionary"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuViewThesaurus 
         Caption         =   "&Thesaurus"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWordWrap 
         Caption         =   "&Word Wrap"
         Checked         =   -1  'True
         Shortcut        =   ^W
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
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuListDelete 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu mnuListDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListCopyPath 
         Caption         =   "&Copy Full File Name"
      End
      Begin VB.Menu mnuListOpenDefault 
         Caption         =   "&Open In Default Application"
      End
      Begin VB.Menu mnuListShowOnly 
         Caption         =   "&Show only this file type"
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

Const Debugging = False
Const mfSkipMouseEventCrap = True
Const msSettingsVersion = "0.8.2"

Dim msPhlegmDate As String
Dim mudtStats As TStatType ' TODO: seriously, rename this to something good.
'Dim FuckIHateThis As Boolean
Dim mfValidCboPath As Boolean
Dim mfStartLabelEditFromButton As Boolean
Dim mfSkipFormResize As Boolean

Dim mfBrowserItemClicked As Boolean
Dim mfBrowserButtonPressed As Boolean
Dim miBrowserMouseButton As Integer
Dim miBrowserShift As Integer

Dim miPathRecent As Integer

Dim msPhlegmKey As String

Dim msTest As String * 100
Dim msTestArray() As String

'Dim EditorAccelTable() As ACCEL
'Dim ControlInfoData As CONTROLINFO
'Dim ctrlInfo1 As CTRLINFO

Dim mudtSettings As TWindowPrefs
Dim mudtCurrentFileSettings As TEditorPrefs

Enum EFileIcon
      Directory = 1
      Drive = 3
      Text = 4
      Other = 5
      Error = 7
      Bookmark = 8
End Enum

Enum EStat  ' TODO: Seriously, rename this to something good.
      Stats = 1
      Modified = 2
      Insert = 3
      SelText = 4
      LastSaved = 5
      Tips = 6
End Enum


Const isaidSHOWMEALLTHEFILESdammit = vbDirectory + vbHidden + vbReadOnly + _
            vbSystem + vbArchive


Private Sub AddToBookmarks(ByVal sNewBookmark As String)
      Dim iIndex As Integer

      If sNewBookmark = "" Then Exit Sub
     
      iIndex = mnuBookmark.UBound + 1
      Load mnuBookmark(iIndex)
      With mnuBookmark(iIndex)
            .Tag = sNewBookmark  ' exact path here, for safe keeping
            .Caption = "&" & iIndex & "   " & sNewBookmark ' here, to make it look all nice
            .Visible = True
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

Private Sub BrowserGetFilesAndFolders(ByVal BrowseDir As String)
      Dim iIcon As Integer
      Dim litCurrentItem As ListItem
      Dim hNextFile As Long, sFileName As String
      Dim WFD As WIN32_FIND_DATA
      Dim fHadFocus As Boolean, fDirUnchanged As Boolean
      Dim sOldSelectedItem As String
      'Dim sngStartTime As Single
      
      
      On Error Resume Next    ' there won't be an active control during form_load, so skip this part.
      fHadFocus = (frmMain.ActiveControl.name = "lvwBrowser")
      On Error GoTo 0
      
       If lvwBrowser.ListItems.Count > 0 Then
            fDirUnchanged = (BrowseDir = lvwBrowser.Tag)
            sOldSelectedItem = lvwBrowser.SelectedItem.Text
      End If
      
      lvwBrowser.Visible = False  ' a nice idea, but we don't want to lose focus while under.  OR DO WE ?
      lvwBrowser.ListItems.Clear
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
     ' sNextFile = Dir(BrowseDir, isaidSHOWMEALLTHEFILESdammit)
     ' Do While sNextFile <> ""
      
      
      'sngStartTime = Timer
      hNextFile = FindFirstFile(BrowseDir & "*", WFD)
      Do
            On Error Resume Next
            
            ' Divide the file types up slightly for icon selection
            sFileName = CstringToVBstring(WFD.cFileName)
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                  iIcon = EFileIcon.Directory
            
            ElseIf Right(sFileName, 4) = ".txt" Then
                  iIcon = EFileIcon.Text
            
            Else
                  iIcon = EFileIcon.Other
                  
            End If

            If Err > 0 Then
                  iIcon = EFileIcon.Error
                  Debug.Print Err & ": " & Err.Description
            End If
            On Error GoTo 0
            
            If sFileName <> "." Then ' just what is the point in providing them with a "." folder?
                  Set litCurrentItem = lvwBrowser.ListItems.Add(, , sFileName, iIcon, iIcon)
                  
                  ' here, let's keep an invisible second column for sorting by directory later
                  If iIcon = EFileIcon.Directory Then
                        litCurrentItem.ListSubItems.Add , , 0
                  Else
                        litCurrentItem.ListSubItems.Add , , 1
                  End If
            End If
      
      Loop While FindNextFile(hNextFile, WFD) <> 0
           
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = 1
      lvwBrowser.Visible = True
      If fHadFocus Then lvwBrowser.SetFocus
      
      ' Auto-select the item previously selected, for a refresh.  Otherwise, the first item.
      
      If lvwBrowser.ListItems.Count > 0 Then
            If fDirUnchanged Then Set litCurrentItem = lvwBrowser.FindItem(sOldSelectedItem)

            If (litCurrentItem Is Nothing) Or fDirUnchanged = False Then
                  lvwBrowser.ListItems(1).Selected = True
            Else
                  litCurrentItem.Selected = True
                  lvwBrowser.SelectedItem.EnsureVisible
            End If
      End If
      'Debug.Print Timer - sngStartTime
End Sub

Private Function BrowserResizeHorizontal(ByVal iSupposedWidth As Integer) As Integer
      ' This is like a miniature RearrangeControls() for just picFileBrowser and everything within,
      ' and it happens to only affect their horizontal components.
      
      ' The return value is the difference (in twips) that picFileBrowser has grown.  Can be negative.
      
      Dim iOffset As Integer, iRealWidth As Integer
      
      If iSupposedWidth < 1000 Then
            iRealWidth = 1000
      ElseIf picFileBrowser.Left + iSupposedWidth + 1500 > frmMain.ScaleWidth Then
            iRealWidth = frmMain.ScaleWidth - picFileBrowser.Left - 1500
      Else
            iRealWidth = iSupposedWidth
      End If
      
      iOffset = iRealWidth - picFileBrowser.Width
      
      picFileBrowser.Width = iRealWidth
      lvwBrowser.Width = lvwBrowser.Width + iOffset
      lblDivider.Left = lvwBrowser.Width
      lvwBrowser.ColumnHeaders(1).Width = lvwBrowser.Width - 100
      cboPath.Width = cboPath.Width + iOffset
      
      BrowserResizeHorizontal = iOffset
End Function

Private Function ParentDirectoryOf(ByVal sPath As String)
      Dim iSlash As Integer
      
      If sPath = "\" Then
            ParentDirectoryOf = ""
      Else
            iSlash = InStrRev(sPath, "\", Len(sPath) - 1)
            ParentDirectoryOf = Left(sPath, iSlash)
      End If
End Function

Private Sub PathAddRecent(ByVal sPath As String)
      ' Supplement recent paths list, unless we are currently scrolling through them.
      ' Top of the List = lowest of the ListIndeces = forward(recent)most of the paths.
            
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
      
      lvwBrowser.ListItems.Clear
      lvwBrowser.Tag = "(Bookmarks)"
      For iIndex = 1 To mnuBookmark.UBound
            Set litCurrentItem = lvwBrowser.ListItems.Add(, , mnuBookmark(iIndex).Tag, _
                  EFileIcon.Bookmark, EFileIcon.Bookmark)
            litCurrentItem.ListSubItems.Add 1, , 1
      Next iIndex
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

Private Sub agEditor_ProgressStatus(ByVal lAmount As Long, ByVal lTotal As Long)
'      Debug.Print "PROGRESS: "; lAmount & " " & lTotal

      ' TODO: this!
      
End Sub

Private Sub btnCurrentDirectory_Click()
      mnuFileCurrentDirectory_Click
End Sub

Private Sub btnDeleteSelected_Click()
      BrowserDeleteSelected
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


Private Sub btnPathBack_Click()
      PathBack
End Sub

Private Sub btnPathForward_Click()
      PathForward
End Sub


Private Sub btnRefresh_Click()
    
      
      If cboPath.ListCount > 0 And cboPath.Text <> cboPath.List(miPathRecent) Then
            cboPath.ListIndex = miPathRecent
      Else
            cboPath_Change
      End If
End Sub

Private Sub btnSort_Click()
      
      ' List remains sorted at all times.  Only the order can be reversed.
      
      With lvwBrowser
            .SortKey = 0
            .SortOrder = Abs(.SortOrder - 1)
            .SortKey = 1
      End With
      
      If lvwBrowser.Tag = "(Bookmarks)" Then BookmarkSaveChanges
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
      ' TODO: Fix this.  Button must sync with menu, registry, and actual Browser visibility.
      
      picFileBrowser.Visible = chkFileBrowser.Value
      mnuViewFilebrowser.Checked = chkFileBrowser.Value
      mnuBrowser.Enabled = chkFileBrowser.Value
      
      RearrangeControls
      'agEditor.SetFocus
End Sub

Private Sub cboPath_Change()
      ' Type a directory into cboPath.  Valid paths will be loaded as you type.
      '     (actually, anything ending in "\" will be loaded)
      '     (TODO: also, you can specify an attribute)
      
      Dim sDirRetval As String
      Dim iIndex As Integer
      Dim litCurrentItem As ListItem
      
      'debug.print "cboPath_CHANGE.   text: " & cboPath.Text
      
      mfValidCboPath = True
      
      If cboPath = "" Then ' TODO: decide what mfValidCboPath should be here
      
            ' Path is at pseudo-"My Computer" location.  Display all drives in lvwBrowser.
            
            lvwBrowser.Tag = ""
            BrowserGetDrives
            PathAddRecent ""
            Exit Sub
      
      ElseIf cboPath = "(Bookmarks)" Then ' TODO: and here
      
            ' Manage bookmarks (load them into the browser with little paperclip icons)
      
            lvwBrowser.Tag = "(Bookmarks)"
            BrowserGetBookmarks
            PathAddRecent "(Bookmarks)"
            Exit Sub
      
      ElseIf Not (cboPath Like "*:\*") Then
            mfValidCboPath = False
      End If
      
      ' Monitor cboPath at each Change, see if a valid directory appeared.
            
      If mfValidCboPath Then
            On Error Resume Next
            sDirRetval = Dir(cboPath.Text, vbDirectory)
            If Err.Number > 0 Then
                  sDirRetval = ""
                  MsgBox "ERROR: some weird thing went wrong with Dir() because it is stupid."
            End If
            On Error GoTo 0
      End If
      
      If sDirRetval = "" Or mfValidCboPath = False Then     ' Found a bad directory
            mfValidCboPath = False
                                                            
            If Right(cboPath, 1) = "\" Then
                  lvwBrowser.ListItems.Clear
                  lvwBrowser.Tag = "(Error)"
            End If
            
      Else              ' Found a good directory!
            mfValidCboPath = True
            
            PathAddRecent cboPath.Text
            
            lvwBrowser.Tag = SnipFileName(GetFullPathName(cboPath))
            BrowserGetFilesAndFolders lvwBrowser.Tag
      End If
End Sub


Private Sub cboPath_GotFocus()
      With cboPath
            If .Text <> "(Bookmarks)" Then .SelStart = Len(.Text)
      End With
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

Private Sub Form_Unload(Cancel As Integer)
      If Not Debugging Then
            SetWindowLong lvwBrowser.hwnd, GWL_WNDPROC, gpOldLvwBrowserProc
            gpOldLvwBrowserProc = 0
      End If
      
      SaveWindowSettings
End Sub

Private Sub lvwBrowser_AfterLabelEdit(Cancel As Integer, NewString As String)

      ' TODO: finish this, and make it work for directories.
      ' Trap errors, particularly, Name can't create a directory.
      ' If there's a slash in the name, look for it without lvwBrowser.tag, and then refresh.
      
      Dim sFolder As String, sOldPath As String
      
      sFolder = lvwBrowser.Tag
      
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
      sOldPath = lvwBrowser.Tag & lvwBrowser.SelectedItem
      
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
End Sub

Private Sub lvwBrowser_BeforeLabelEdit(Cancel As Integer)
      'debug.print "lvwBrowser_Before " & Cancel
End Sub

Private Sub lvwBrowser_Click()
      'debug.print "lvwBrowser_CLICK"
      If mfBrowserItemClicked = True And _
            miBrowserMouseButton = vbLeftButton And _
            miBrowserShift = 0 Then
            
            BrowserExecuteItem lvwBrowser.SelectedItem
      End If
      
      miBrowserMouseButton = 0  ' These probably an overcaution --
      miBrowserShift = 0                  ' They are reset in the next MouseDown anyway.
End Sub

Private Sub BrowserExecuteItem(ByVal Item As MSComctlLib.ListItem, Optional ByVal sOp As String)
      If (lvwBrowser.ListItems.Count = 0) Then Exit Sub
      
      Select Case Item.Icon
      
            Case EFileIcon.Text, EFileIcon.Other
                  ' Open the file.
                  EditorLoadFile lvwBrowser.Tag & Item.Text
                  
            Case EFileIcon.Directory, EFileIcon.Drive
                  ' Open the folder, or go up a folder.
                  If Item.Text = ".." Then
                        mnuFileParentDirectory_Click
                  Else
                        cboPath = lvwBrowser.Tag & Item.Text & "\"
                  End If
            
            Case EFileIcon.Bookmark
                  ' Open the bookmarked file.
                  EditorLoadFile Item.Text
      End Select
End Sub

Private Sub lvwBrowser_DblClick()
      'debug.print "lvwBrowser_DBLCLICK"
End Sub

Private Sub lvwBrowser_ItemClick(ByVal Item As MSComctlLib.ListItem)
      ' ItemClick fires every time the selection changes, or a selection is clicked.
      'Debug.Print "itemclick " & Item.Index
            
      mfBrowserItemClicked = True
      
      mnuListOpenDefault.Enabled = True
      mnuListOpenDefault.Caption = "Open With Default Program..."
      mnuListDelete.Enabled = True
      mnuListRename.Enabled = True
      mnuListCopyPath.Enabled = True
      
      If Item.Icon = EFileIcon.Directory Or Item.Icon = EFileIcon.Drive Then
            mnuListOpenDefault.Caption = "Explore..."
            mnuListDelete = False
            If Item.Text = ".." Or Item.Icon = EFileIcon.Drive Then mnuListRename = False
      End If
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
      
      ' TODO: decide how Shift is going to affect all this.
      ' With a one-click open-file, we'd like shift-click to leave it alone, at the very least.
      
      mfBrowserItemClicked = False
      miBrowserMouseButton = Button
      miBrowserShift = Shift
End Sub

Private Sub lvwBrowser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'      Debug.Print "lvwBROWSER MOUSEMOVE, X: " & x

'      If picFileBrowser.MousePointer = vbSizeWE Then
'            picFileBrowser.Tag = ""
'            picFileBrowser.MousePointer = 0
'      End If
      'staTusBar1.Panels(EStat.LastSaved) = ppppppp
End Sub

Private Sub mnuBookmark_Click(Index As Integer)
      EditorLoadFile mnuBookmark(Index).Tag
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
      
      If lvwBrowser.Tag = "(Bookmarks)" Then btnRefresh_Click
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
            mnuBrowserOpenDefault.Caption = "Open With Default Program..."
            mnuBrowserDelete.Enabled = False
            mnuBrowserRename.Enabled = False
            Exit Sub
      Else
            mnuBrowserOpenDefault.Enabled = True
            mnuBrowserOpenDefault.Caption = "Open With Default Program..."
            mnuBrowserDelete.Enabled = True
            mnuBrowserRename.Enabled = True
      End If
      
      With lvwBrowser.SelectedItem
            If .Icon = EFileIcon.Directory Or .Icon = EFileIcon.Drive Then
                  mnuBrowserOpenDefault.Caption = "Explore Selected..."
                  mnuBrowserDelete = False
                  If .Text = ".." Or .Icon = EFileIcon.Drive Then mnuBrowserRename = False
            End If
      End With

End Sub

Private Sub mnuBrowserDelete_Click()
      BrowserDeleteSelected
End Sub

Private Sub BrowserDeleteSelected()
      Dim iBookm As Integer
      Dim sTheDamned As String
      
      If lvwBrowser.ListItems.Count = 0 Then Exit Sub
      
      sTheDamned = lvwBrowser.Tag & lvwBrowser.SelectedItem
      
      With lvwBrowser
            If .Tag = "(Bookmarks)" Then
                  
                  iBookm = .SelectedItem.Index      ' TODO: FIIIIIIIIXXXXXXXXX
                  .ListItems.Remove iBookm
                  
                  BookmarkSaveChanges
                  
            ElseIf .Tag = "" Then
                  Caption = "I WILL NOT DELETE YOUR DISK.  FIND SOMEONE ELSE."
            ElseIf Not FileExists(sTheDamned) Then
                  Caption = "Can't delete what isn't there: " & sTheDamned
            ElseIf sTheDamned = agEditor.Tag Then
                  Caption = "Can't delete your open file.  Sorry."
            ElseIf GetAttr(sTheDamned) And vbDirectory Then
                  Caption = "This program would rather not be held responsible for mass deletions.  Please use another."

'                  RmDir sTheDamned
'                  Caption = "Folder deleted successfully: " & sTheDamned
'                  btnRefresh_Click
            Else
                  On Error Resume Next
                  Kill sTheDamned
                  If Err > 0 Then
                        Caption = Err.Number & ": " & Err.Description
                  Else
                        Caption = "File deleted successfully: " & sTheDamned
                        ' TODO: this is no better a refresh than cboPath_change.  fix.
                        btnRefresh_Click
                  End If
                  On Error GoTo 0
            End If
      End With
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

Private Sub mnuEditFont_Click()
      btnFont_Click
End Sub

Private Sub mnuFileCurrentDirectory_Click()
      Dim litCurrentFile As ListItem
      
      cboPath = SnipFileName(agEditor.Tag)
      Set litCurrentFile = lvwBrowser.FindItem(SnipPath(agEditor.Tag))
      litCurrentFile.Selected = True
      litCurrentFile.EnsureVisible
End Sub

Private Sub mnuFileOpen_Click()
      If mnuViewFilebrowser.Checked = False Then
            mnuViewFilebrowser.Checked = True
            mnuViewFilebrowser_Click
      End If
      lvwBrowser.SetFocus
End Sub

Private Sub mnuFileParentDirectory_Click()
      cboPath = ParentDirectoryOf(cboPath)
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
            End If
            On Error GoTo 0
      
      End If
End Sub
Private Sub mnuFileSaveAs_Click()
      Dim sDefaultPath As String, sFileName As String
      Dim fSuccess As Boolean
      
      If lvwBrowser.Tag = "(Bookmarks)" Or lvwBrowser.Tag = "" Or _
                  lvwBrowser.Tag = "(Error)" Then
                  
            sDefaultPath = CurDir & "\"
            
      Else   ' If none of the special cases, the tag should be a valid path
            sDefaultPath = lvwBrowser.Tag
      End If
      
      sFileName = InputBox("File name:", "Save", sDefaultPath & msPhlegmDate & ".txt")
      fSuccess = agEditor.SaveToFile(sFileName, SF_TEXT)

      If fSuccess Then
            staTusBar1.Panels(EStat.Modified) = ""
            agEditor.Tag = sFileName
            frmMain.Caption = sFileName
            btnRefresh_Click
      Else
            frmMain.Caption = "ERROR: cannot save to " & sFileName
      End If
End Sub

Private Sub mnuList_Click()
      
      ' This is the popup menu for lvwBrowser.  Click fires whenever the menu is popped up.
      
      ' Most menu items are enabled/disabled in lvwBrowser_ItemClick.
      ' Here, we un-set some of them if the user has clicked somewhere that is not a list item.
      
      ' Events happen in this order: lvwBrowser_MouseDown, lvwBrowser_ItemClick, mnuList_Click.
      
      ' mfBrowserItemClicked is set to False on the MouseDown, and True on the ItemClick.
      ' So if it gets here as False, that means ItemClick did not happen on this mouse event.
      
      If Not mfBrowserItemClicked Then
            mnuListOpenDefault.Enabled = False
            mnuListDelete.Enabled = False
            mnuListRename.Enabled = False
            mnuListCopyPath.Enabled = False
      End If
End Sub

Private Sub mnuListCopyPath_Click()
      Clipboard.Clear
      If lvwBrowser.Tag = "(Bookmarks)" Then
            Clipboard.SetText lvwBrowser.SelectedItem
      Else
            Clipboard.SetText lvwBrowser.Tag & lvwBrowser.SelectedItem
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
            With agEditor
'                  .Visible = False
                  iOffset = BrowserResizeHorizontal(x + lblDivider.Left)
                  .Move .Left + iOffset, .Top, .Width - iOffset, .Height
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
'            agEditor.Width = picFileBrowser.Width + 160
'            agEditor.Left = frmMain.Width - agEditor.Width - 150
      End If

End Sub


Private Sub mnuListOpenDefault_Click()
      Dim sPath As String
      
      With lvwBrowser
            If .ListItems.Count > 0 Then
                  If .Tag = "(Bookmarks)" Then
                        sPath = .SelectedItem.Text
                  Else
                        sPath = .Tag & .SelectedItem.Text
                  End If
                  ShellExecute 0, "open", sPath, 0, "", SW_RESTORE
            End If
      End With
End Sub

Private Sub mnuListRename_Click()
      lvwBrowser.StartLabelEdit
      
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

Private Sub txtQueryBox_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
          retval = ShellExecute(0, "open", _
                  "http://dictionary.reference.com/search?q=" & txtQueryBox, 0, "", 8)
      End If
End Sub

Private Sub txtQueryBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
      txtQueryBox = Data.GetData(vbCFText)
      txtQueryBox_KeyPress vbKeyReturn
End Sub

Private Sub txtQueryBox_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
      txtQueryBox.SelStart = 0
      txtQueryBox.SelLength = Len(txtQueryBox)
End Sub

Private Sub agEditor_KeyDown(KeyCode As Integer, Shift As Integer)
'      gpOldProc = SetWindowLong(agEditor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)
      Select Case KeyCode
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
      
      lLineIndex = agEditor.CurrentLine
      lCharIndex = SendMessage(agEditor.RichEdithWnd, EM_LINEINDEX, ByVal lLineIndex, 0)
      
      If staTusBar1.Visible Then
            With mudtStats
                .i = lMin + 1
                .y = lLineIndex + 1
                .x = lMin - lCharIndex + 1
                .xmax = SendMessage(agEditor.RichEdithWnd, EM_LINELENGTH, ByVal lCharIndex, 0) + 1
            End With
        
            FillStats
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
End Sub



Private Sub Form_Load()
      Dim vDate As Variant

      InitializeMenus
            
      Debug.Print "command line sayeth: [" & Command() & "]"
      agEditor.Tag = Trim(Command())
      
      msPhlegmKey = "Software\" & App.Title & "\" & msSettingsVersion
      
      vDate = Date
      msPhlegmDate = Year(vDate) & "-" & Format(Month(vDate), "0#") & _
            "-" & Format(Day(vDate), "0#")
      
      GetWindowSettings
      mudtStats.imax = agEditor.CharacterCount
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
                                                 
                                                 
            Case 221 ' Right bracket, but contains code for it and for right arrow.
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
                  If Shift = vbShiftMask + vbCtrlMask Then BrowserDeleteSelected
                  
      End Select
End Sub

Private Sub lvwBrowser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      'debug.print "lvwBrowser_MOUSEUP " & Button & " " & Shift
'      If FuckIHateThis Then
'            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'            FuckIHateThis = True
'      End If
      
      If (Button = vbRightButton And Shift = 0) Then
            Me.PopupMenu mnuList
      End If
      
      ' For use in the click event, so we know what was clicked.
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
      agEditor.Tag = ""
      frmMain.Caption = "(New File)"
End Sub

Private Sub mnuFileSave_Click()
      Dim fSuccess As Boolean
      
      fSuccess = agEditor.SaveToFile(agEditor.Tag, SF_TEXT)

      If fSuccess Then
            staTusBar1.Panels(EStat.Modified) = ""
            Caption = agEditor.Tag
      Else
            frmMain.Caption = "ERROR: cannot save to " & agEditor.Tag
      End If
End Sub

Private Sub mnuViewDictionary_Click()
      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
      If txtQueryBox.Visible Then
            txtQueryBox.SetFocus
      Else
            txtQueryBox_GotFocus
      End If
      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
      txtQueryBox.SetFocus
      If agEditor.SelectedText <> "" Then txtQueryBox_KeyPress vbKeyReturn
End Sub

Private Sub mnuViewFilebrowser_Click()
    chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub agEditor_Change()

      If staTusBar1.Panels(EStat.Modified) = "" Then
          staTusBar1.Panels(EStat.Modified) = "Modified"
      End If
      
      If staTusBar1.Visible Then
            With mudtStats
                .imax = agEditor.CharacterCount
                .ymax = agEditor.LineCount
            End With
            
            FillStats
      End If
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
      Const leftmargin = 60
      Const rightmargin = 150
      Const midspace = 100
      
      If Me.WindowState = vbMinimized Then Exit Sub
      
      fValidWindowSize = True ' ...until proven guilty.
      iRedoResizeY = frmMain.Height
      iRedoResizeX = frmMain.Width
      
      sHadFocus = ActiveControl.name
      agEditor.Visible = False ' MUCH faster if you turn him off while thinking
      
      ' Calculate control positions...
      
      iEdTop = 0
      If mnuViewToolbar.Checked Then iEdTop = iEdTop + picToolBox.Height + midspace
      
      iEdHeight = frmMain.ScaleHeight - iEdTop - topmargin
      If mnuViewStatusBar.Checked Then iEdHeight = iEdHeight - staTusBar1.Height
      
      iEdLeft = leftmargin
      If mnuViewFilebrowser.Checked Then iEdLeft = iEdLeft + picFileBrowser.Width
      
      iEdWidth = frmMain.ScaleWidth - iEdLeft
      
      
      ' Check to see if we've gone out of bounds...
      
            ' TODO: iEdWidth comes back around a second time as 1499
      If iEdWidth < 1500 And WindowState = vbMaximized Then
            BrowserResizeHorizontal picFileBrowser.Width
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
            iPicBoxMarginsY = picFileBrowser.Height - picFileBrowser.ScaleHeight
            iRedoResizeY = iEdTop + lvwBrowser.Top + iPicBoxMarginsY + iFormMarginsY + 1510
      End If
      
      If Not fValidWindowSize Then
            Move Left, Top, iRedoResizeX, iRedoResizeY
            Exit Sub
      End If
      
      ' It's all good.  Move the controls now!
      
      agEditor.Move iEdLeft, iEdTop, iEdWidth, iEdHeight
      With picFileBrowser
            .Move .Left, iEdTop, .Width, iEdHeight
      End With
      lvwBrowser.Height = iEdHeight - lvwBrowser.Top
      picToolBox.Width = frmMain.Width
            
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
      
      agEditor.Visible = True
      If sHadFocus = "agEditor" Then agEditor.SetFocus
End Sub

Private Sub mnuViewStatusBar_Click()
      staTusBar1.Visible = Not staTusBar1.Visible
      mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
      RearrangeControls
End Sub

Private Sub mnuViewThesaurus_Click()
      If agEditor.SelectedText <> "" Then txtQueryBox = agEditor.SelectedText
      If txtQueryBox.Visible Then
            txtQueryBox.SetFocus
      Else
            txtQueryBox_GotFocus
      End If
      If agEditor.SelectedText <> "" Then
            retval = ShellExecute(0, "open", _
                    "http://thesaurus.reference.com/search?q=" & txtQueryBox, 0, "", 8)
            Me.SetFocus
      End If
End Sub

Private Sub InitializeMenus()
'      Dim tempinfo As MENUITEMINFO
'      Dim hMenu As Long, retval As Long
'
'      hMenu = GetMenu(hwnd)
'      hMenu = GetSubMenu(hMenu, 2)
'      retval = ModifyMenu(hMenu, 0, MF_STRING + MF_BYPOSITION, 2, "&Penis" + vbTab + "Ctrl+P")
      
      mnuEditIncFont.Caption = "&Increase Font Size" & vbTab & "Alt+="
      
      
      mnuWriteCut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnuWriteCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuWritePaste.Caption = "&Paste" & vbTab & "Ctrl+V"
      
      mnuWindowMinimize.Caption = mnuWindowMinimize.Caption & vbTab & "Alt+F10"
      mnuWindowMaximize.Caption = mnuWindowMaximize.Caption & vbTab & "Alt+F12"
      mnuWindowRestore.Caption = mnuWindowRestore.Caption & vbTab & "Alt+F11"
      
      mnuListDelete.Caption = mnuListDelete.Caption & vbTab & "Shift+Ctrl+Del"

End Sub

Private Function BrowserGetDrives() As Integer
      ' Find all logical drives and display them in lvwBrowser
      ' Returns the number of logical drives found.
      
      Dim sFix As String * 255
      Dim sVar As String
      Dim sArray() As String
      Dim lLength As Long
      Dim iIndex As Integer
      Dim litCurrentItem As ListItem
      
      
      lLength = GetLogicalDriveStrings(100, sFix)
      sVar = Left(sFix, lLength)
      sArray = Split(sVar, Chr(0)) ' "(x,x, , )" is an error.  don't put in more commas unless
      lvwBrowser.ListItems.Clear          ' they lead to something.
      lvwBrowser.Tag = ""
      
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      
      iIndex = LBound(sArray)
      sVar = TrimTrailingSlash(sArray(iIndex))
      
      Do While (sVar <> "") And (sVar <> Chr(0))
            
            Set litCurrentItem = lvwBrowser.ListItems.Add( _
                  1, , sVar, EFileIcon.Drive, EFileIcon.Drive)
            litCurrentItem.ListSubItems.Add , , 0
            
            iIndex = iIndex + 1
            sVar = TrimTrailingSlash(sArray(iIndex))
      Loop
      
      lvwBrowser.Sorted = True
      lvwBrowser.SortKey = 1
      BrowserGetDrives = iIndex - 1
End Function

Private Sub mnuViewToolbar_Click()
      picToolBox.Visible = Not picToolBox.Visible
      mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
      RearrangeControls
End Sub

Private Sub mnuViewWordWrap_Click()
      agEditor.ViewMode = Abs(agEditor.ViewMode - 1)
      mnuViewWordWrap.Checked = Not mnuViewWordWrap.Checked
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

Private Function EditorLoadFile(sFileName As String) As Boolean
      
      ' pass along the boolean return value, if anyone wants it.
      EditorLoadFile = agEditor.LoadFromFile(sFileName, SF_TEXT)
      
      If (EditorLoadFile = True) Then    ' Success!
            agEditor.Tag = sFileName
            frmMain.Caption = sFileName
            staTusBar1.Panels(EStat.Modified) = ""
            agEditor.SetSelection 0, 0
            
      Else  ' Failure!
            frmMain.Caption = "ERROR.  command() = " & Command() & " Tag: " & sFileName
            agEditor.Tag = ""  ' if anyone asks the Tag about a loaded file, we won't let him
      End If                        ' give them false hope.
End Function

Private Sub SaveWindowSettings()
      Dim lMin As Long, lMax As Long, lKey As Long, lRetval As Long
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
            
            .BrowserWidth = picFileBrowser.Width
            .ShowFileBrowser = picFileBrowser.Visible
            .ShowStatusBar = staTusBar1.Visible
            .ShowToolBar = picToolBox.Visible
            .SortMethod = lvwBrowser.SortOrder
            .AutoLoadFile = agEditor.Tag
            .cboPath = cboPath
            .BookmarkCount = mnuBookmark.UBound
      End With
      
      agEditor.GetSelection lMin, lMax
      
      With mudtCurrentFileSettings
            .FirstVisibleLine = agEditor.FirstVisibleLine
            .SelEnd = lMax
            .SelStart = lMin
            .WordWrap = agEditor.ViewMode
            
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
      
            
      
      
      lRetval = RegCreateKeyEx(HKEY_CURRENT_USER, msPhlegmKey, 0, "", 0, _
                  KEY_ALL_ACCESS, ByVal 0, lKey, lNewOrUsed)
      If lRetval <> 0 Then MsgBox "RegCreateKey Failed: " & lKey
      
      lValueSize = LenB(mudtSettings)
      lRetval = RegSetValueExAny(lKey, "Settings", 0, REG_NONE, _
                  ByVal mudtSettings, lValueSize)
      If lRetval <> 0 Then MsgBox "RegSetValueEx Failed.  settings: " & _
                  LenB(mudtSettings) & " " & lKey, , App.Title
      
      lValueSize = LenB(mudtCurrentFileSettings)
      lRetval = RegSetValueExAny(lKey, "agEditor", 0, REG_NONE, _
                  ByVal mudtCurrentFileSettings, lValueSize)
      If lRetval <> 0 Then MsgBox "RegSetValueEx Failed.  mudtCurrentFileSettings: " & _
                  LenB(mudtCurrentFileSettings) & " " & lKey, , App.Title
      
      
      For iIndex = 1 To mnuBookmark.UBound
            lValueSize = LenB(mnuBookmark(iIndex).Tag)
            lRetval = RegSetValueExString(lKey, "Bookmark" & CStr(iIndex), 0, REG_SZ, _
                        ByVal mnuBookmark(iIndex).Tag, lValueSize)
      Next iIndex
      
      For iIndex = mnuBookmark.UBound + 1 To mudtSettings.BookmarkCount
            RegDeleteValue lKey, "Bookmark" & CStr(iIndex)
      Next iIndex
      
      ' TODO: Gotta remember to delete bookmarks in the regisy that were
      ' deleted in the program!

      lRetval = RegCloseKey(lKey)
End Sub

Private Sub GetWindowSettings()
      Dim lRetval As Long, lKey As Long
      Dim lDataType As Long ' receiving only
      Dim lValueSize As Long ' in/out
      Dim poiFirstLine As POINTAPI
      Dim sTemp As String * 255
      Dim fntTemp As New StdFont
      Dim iBookm As Integer
      
      Dim udtWindowPlacement As WINDOWPLACEMENT
      Dim rectRestored As RECT
      Dim poiTemp As POINTAPI
      
      lRetval = RegOpenKeyEx(HKEY_CURRENT_USER, msPhlegmKey, 0, KEY_QUERY_VALUE, lKey)
      
      lValueSize = LenB(mudtSettings)
      lRetval = RegQueryValueExAny(lKey, "Settings", 0, lDataType, ByVal mudtSettings, lValueSize)
      If lRetval = 0 Then
            With mudtSettings
                  mfSkipFormResize = True
                  BrowserResizeHorizontal .BrowserWidth
                  
                  .WNP.Length = LenB(.WNP)
                  SetWindowPlacement hwnd, .WNP
                  
                  lvwBrowser.SortOrder = .SortMethod
                  If agEditor.Tag = "" Then agEditor.Tag = Trim(CstringToVBstring(.AutoLoadFile))
                  
                  For iBookm = 1 To mudtSettings.BookmarkCount ' TODO: THIS BEFORE SETTINGS... SOMEHOW...
                        lValueSize = 255 * 2
                        lRetval = RegQueryValueExString(lKey, "Bookmark" & CStr(iBookm), 0, lDataType, _
                                    ByVal sTemp, lValueSize)
                        AddToBookmarks Left(sTemp, lValueSize - 1) ' size included the null
                  Next iBookm
      
                  cboPath = Trim(CstringToVBstring(.cboPath))
      
                  chkFileBrowser.Value = -CInt(.ShowFileBrowser)
                  chkFileBrowser_Click
                  'picFileBrowser.Visible = .ShowFileBrowser
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
      
      EditorLoadFile agEditor.Tag
      
      If Trim(Command()) = "" Then
            lValueSize = LenB(mudtCurrentFileSettings)
            lRetval = RegQueryValueExAny(lKey, "agEditor", 0, lDataType, ByVal mudtCurrentFileSettings, lValueSize)
            If lRetval = 0 Then
                  With mudtCurrentFileSettings
                        agEditor.ViewMode = .WordWrap
                        
                        fntTemp.Bold = .FontBold
                        fntTemp.Charset = .FontCharset
                        fntTemp.Italic = .FontItalic
                        fntTemp.name = Trim(CstringToVBstring(.FontName))
                        fntTemp.Size = .FontSize
                        fntTemp.Strikethrough = .FontStrikethrough
                        fntTemp.Underline = .FontUnderline
                        agEditor.SetFont fntTemp, , , , ercSetFormatAll
                        
                        ' If the file has been changed so that selection and scroll positions are meaningless,
                        ' just skip them...
                        
                        On Error Resume Next
                        agEditor.SetSelection .SelStart, .SelEnd
                        SendMessage agEditor.RichEdithWnd, EM_SETSCROLLPOS, 0, .ScrollPos
                        On Error GoTo 0
                  End With
            End If
      End If
      
      RegCloseKey lKey
End Sub

