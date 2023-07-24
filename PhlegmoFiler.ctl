VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.3#0"; "VBCCR17.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PhlegmoFiler 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   5235
   Begin MSComctlLib.ListView lvwBrowser 
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   780
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   4419
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   1138
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2302
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "SortedSize"
         Text            =   "SortedSize"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "IsFolder"
         Text            =   "IsFolder"
         Object.Width           =   0
      EndProperty
   End
   Begin VBCCR17.ComboBoxW cboPath 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Type a directory into here, or select one below.  You can even specify a file extension.  Example:   c:\windows\*.dll"
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "PhlegmoFiler.ctx":0000
      ScrollTrack     =   0   'False
   End
   Begin VB.CommandButton btnScrollToTop 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2730
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":0036
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Scroll To Top"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   390
   End
   Begin VB.CommandButton btnSyncContents 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2340
      MaskColor       =   &H80000001&
      Picture         =   "PhlegmoFiler.ctx":04EC
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Jump to the directory containing your open file... (Ctrl+F5)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnDeleteSelected 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1950
      MaskColor       =   &H00000000&
      Picture         =   "PhlegmoFiler.ctx":082E
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Delete File (Del)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnRefresh 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1560
      MaskColor       =   &H80000005&
      Picture         =   "PhlegmoFiler.ctx":0B70
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Refresh Files (F5)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnSort 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1170
      MaskColor       =   &H80000005&
      Picture         =   "PhlegmoFiler.ctx":0EB2
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Reverse the sort order (Ctrl+H)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnFolderUp 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   780
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":11F4
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Go up a directory (Left arrow key or Ctrl+F6)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnPathForward 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   390
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":1536
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Go forward a directory (Alt+Right)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnPathBack 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":19B0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Go back a directory (Alt+Left)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin MSComctlLib.ImageList ilsFileIcons2 
      Left            =   3360
      Top             =   480
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
            Picture         =   "PhlegmoFiler.ctx":1E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":217C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":2820
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":2B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":2EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3216
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3568
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":38BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":42B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":4602
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":4954
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDivider 
      Height          =   5670
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "PhlegmoFiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AUTOSIZE_COLUMNS As Boolean = False ' Will autosize type, size, and modified while retaining filename width
Private Const COLUMN_TOO_SMALL = 600 ' We won't keep the auto-size if it's beneath a certain point

' This is not fully implemented, especially any possible conflicts with AUTOSIZE_COLUMNS
Private Const ALLOW_REARRANGE_COLUMNS As Boolean = False

Private Const INITIAL_WIDTH = 6000
Private Const RIGHT_MARGIN = 105
Private Const BOTTOM_MARGIN = 15
Private Const MIN_WIDTH = 995

Private mlInitialPointerX As Long
Private mlPrevPointerX As Long
Private mlUserCtlMaxWidth As Long
Private miInitializings As Integer ' Did we initialize? Multiple times for some reason? Just checking

Private msDir As String
Private meMode As eFilerMode
Private moFso As Object

Private msDirPrev As String
Private msFilter As String
Private msFilterPrev As String
Private miLastPathIndex As Integer
Private msPartialFileName As String
Private msSelTextPrev As String
Private mbDirUnchanged As Boolean
Private mbFilterUnchanged As Boolean
Private mbGoingToParent As Boolean
Private mbItemClicked As Boolean
Private mbDoneLoading As Boolean

Private msRealSortKey As String ' always end on a sort of IsFolder to keep folders separate; need to track "real" order
Private miRealSortOrder As Integer

Public Event ResizeHorizontal(ByVal lWidth As Long)

' Sent less frequently, use this to hard-limit the form's min-width via API calls
Public Event SeriousResize(ByVal lWidth As Long)

Private Function AutoSelectListItem()
      Dim oCurrentItem As ListItem
      
      If lvwBrowser.ListItems.Count = 0 Or meMode = Bookmarks Or meMode = History Then Exit Function
      
      If msPartialFileName <> "" Then
            ' Auto-select first filename to match partialfilename, if given.
            Set oCurrentItem = lvwBrowser.FindItem(msPartialFileName, , lvwPartial)
            If Not (oCurrentItem Is Nothing) Then oCurrentItem.Selected = True
            
      ElseIf mbGoingToParent Then
            ' Auto-select the directory we just moved out of, if doing a ParentDirectory.
            Set oCurrentItem = lvwBrowser.FindItem(moFso.GetBaseName(msDirPrev))
            If Not (oCurrentItem Is Nothing) Then oCurrentItem.Selected = True
            
      
      ElseIf mbDirUnchanged Then
            ' Auto-select the item previously selected, for a refresh.
            Set oCurrentItem = lvwBrowser.FindItem(msSelTextPrev)

            If (oCurrentItem Is Nothing) Then
                  lvwBrowser.ListItems(1).Selected = True
            Else
                  oCurrentItem.Selected = True
            End If
            
      Else ' Otherwise, auto-select the first item.
            lvwBrowser.ListItems(1).Selected = True
      End If
                  
      DoEvents ' Just doesn't seem to work without DoEvents first.
      If Not (lvwBrowser.SelectedItem Is Nothing) Then
            lvwBrowser.SelectedItem.EnsureVisible
      End If
End Function

Private Sub AutosizeColumns()
      With lvwBrowser
            If AUTOSIZE_COLUMNS And mbDoneLoading Then
                  Dim iColumns(1 To 3) As Integer, iColumn As Variant
                  DebugLog "Autosizing column indeces: " & .ColumnHeaders("Type").Index & ", " _
                        & .ColumnHeaders("Size").Index & ", " & .ColumnHeaders("Modified").Index
                  
                  iColumns(1) = .ColumnHeaders("Type").Index - 1
                  iColumns(2) = .ColumnHeaders("Size").Index - 1
                  iColumns(3) = .ColumnHeaders("Modified").Index - 1
                  For Each iColumn In iColumns
                        SendMessage lvwBrowser.hwnd, LVM_SETCOLUMNWIDTH, ByVal CLng(iColumn), LVSCW_AUTOSIZE
                        If .ColumnHeaders(iColumn + 1).Width <= COLUMN_TOO_SMALL Then
                              .ColumnHeaders(iColumn + 1).Width = COLUMN_TOO_SMALL
                        End If
                  Next iColumn
            End If
            ' These should have never changed... but who knows?
            .ColumnHeaders("SortedSize").Width = 0
            .ColumnHeaders("IsFolder").Width = 0
      End With
End Sub

Private Sub btnFolderUp_Click()
      ' When we go up a dir, preserve the existing filter except in a drives list.
      
      If meMode <> eFilerMode.Files Then Exit Sub
      
      Dim sParentDir As String
      sParentDir = ParentDirectoryOf(msDir)
      
      If sParentDir = "" Then
            cboPath = sParentDir
      Else
            cboPath = sParentDir & msFilter
      End If
End Sub

Private Sub btnSort_Click()
      With lvwBrowser
            .SortOrder = Abs(.SortOrder - 1)
            .SortKey = .ColumnHeaders("Name").Index - 1
            .SortKey = .ColumnHeaders(msRealSortKey).Index - 1
            .SortKey = .ColumnHeaders("IsFolder").Index - 1
      End With
End Sub

Private Sub cboPath_Change()
      
      ParsePath cboPath
      
      Select Case meMode
            
            Case Bookmarks
                  LoadBookmarks
                  PathAddRecent "(Bookmarks)"
            
            Case History
                  LoadHistory
                  PathAddRecent "(History)"
      
            Case Drives
                  If Not (mbDirUnchanged And mbFilterUnchanged) Then
                        LoadDrives
                        PrintComboBoxW cboPath
                        PathAddRecent ""
                        PrintComboBoxW cboPath
                  End If
            
            Case Else
                  If Not (mbDirUnchanged And mbFilterUnchanged) Then
                        LoadFilesAndFolders
                        ' Add to recent paths only if filtration was fruitful.
                        If lvwBrowser.ListItems.Count > 0 Then
                              PrintComboBoxW cboPath
                              PathAddRecent msDir & msFilter
                              PrintComboBoxW cboPath
                        End If
                  End If
      End Select
      
      AutoSelectListItem
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
      
      miLastPathIndex = cboPath.ListIndex
     
      cboPath_Change
End Sub

Private Sub cboPath_GotFocus()
      ' When focus is obtained, put the cursor right where we would have moved it anyway:
      ' At the end of the path, before the extension if one exists.
      
      If cboPath <> "(Bookmarks)" And cboPath <> "(History)" Then
            
            Dim iExtensionLength As Integer
            
            iExtensionLength = Len(moFso.GetExtensionName(cboPath))
            If iExtensionLength > 0 Then iExtensionLength = iExtensionLength + 1 ' include the dot
            cboPath.SelStart = Len(cboPath) - iExtensionLength
      End If
End Sub

Private Sub LoadBookmarks() ' TODO
'      Dim iIndex As Integer
'      Dim oCurrentItem As ListItem
'
'      lvwBrowser.Visible = False
'      lvwBrowser.ListItems.Clear
'      msDir = "(Bookmarks)"
'      ' I'm adding the index as a Key, to avoid using real indeces.
'      ' (So that I can use API functions that desynchronize listitem indexing.)
'      ' Edit: I'm not really doing that. Using bookmarks as a test case on whether that might be doable.
'      For iIndex = 1 To mnuBookmark.UBound
'            Set oCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuBookmark(iIndex).tag, _
'                  eIconType.Bookmark, eIconType.Bookmark)
'            oCurrentItem.ListSubItems.Add 1, , goFso.GetExtensionName(mnuBookmark(iIndex).tag)
'      Next iIndex
'      gtBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
'      AutosizeColumns
'      lvwBrowser.Visible = True
'      staTusBar.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " bookmarks"
End Sub

Private Sub LoadDrives()
      ' Find all logical drives and display them in the file browser
      
      DebugLog "Gonna load some drives"
      Dim sDrivesFixed As String * 255
      Dim sDriveString As String
      Dim sDriveArray() As String
      Dim sNextDrive As String, eDriveIcon As eIconType
      Dim lLength As Long
      Dim iIndex As Integer, iTempKey As Integer
      Dim oCurrentItem As ListItem
      Dim oDrive
      
      On Error GoTo LOAD_DRIVES_ERROR
      lLength = GetLogicalDriveStrings(100, sDrivesFixed)
      sDriveString = Left(sDrivesFixed, lLength)
      DebugLog "Drives found = " & sDriveString, 2
      sDriveArray = Split(sDriveString, Chr(0))
      
      lvwBrowser.ListItems.Clear
      msDir = ""
      iIndex = LBound(sDriveArray)
      sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      
      On Error Resume Next
      Do While (sNextDrive <> "") And (sNextDrive <> Chr(0))
            
            Set oDrive = moFso.GetDrive(sNextDrive)
            If Err > 0 Then
                  DebugLog "Problem with drive " & sNextDrive & " " & Err.Description, 2
                  Err = 0
            Else
                  Select Case oDrive.DriveType
                        Case 1: eDriveIcon = Floppy
                        Case 2: eDriveIcon = Drive
                        Case 3: eDriveIcon = Network
                        Case 4: eDriveIcon = Cdrom
                        Case Else: eDriveIcon = IconError
                  End Select
                  Set oCurrentItem = lvwBrowser.ListItems.Add(1, , sNextDrive, , eDriveIcon)
                  oCurrentItem.ListSubItems.Add , "Type", ""
                  oCurrentItem.ListSubItems.Add , "Size", ""
                  oCurrentItem.ListSubItems.Add , "Modified", ""
                  oCurrentItem.ListSubItems.Add , "IsFolder", 0 ' Yes
                  oCurrentItem.ListSubItems.Add , "SortedSize", 0
            End If
            iIndex = iIndex + 1
            sNextDrive = TrimTrailingSlash(sDriveArray(iIndex))
      Loop
      AutosizeColumns
      lvwBrowser.SortKey = lvwBrowser.ColumnHeaders(msRealSortKey).Index - 1
      lvwBrowser.Sorted = True
      
      ' staTusBar.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " drives"
      Exit Sub

LOAD_DRIVES_ERROR:
      Dim sErrorMsg As String
      sErrorMsg = "LoadDrives error: " & Err.Description
      DebugLog sErrorMsg, 2
      MsgBox sErrorMsg
End Sub

Private Sub LoadFilesAndFolders()
      DebugLog "Gonna load some files and folders at: " & msDir
      
      Dim eIcon As eIconType, iTempKey As Integer
      Dim oCurrentItem As ListItem
      Dim lNextFile As Long, sFileName As String, sEx As String
      Dim tWfd As WIN32_FIND_DATA
      Dim bHadFocus As Boolean
      Dim oTotalBytes, oSizeBig, oSizeBig2 ' these can be bigger than a long integer
      
      ' bHadFocus = (ActiveControl.Name = "lvwBrowser")
      On Error GoTo LOAD_FILES_ERROR
      
      lvwBrowser.Visible = False
      Screen.MousePointer = vbHourglass
      
      lvwBrowser.ListItems.Clear
      lvwBrowser.SortKey = 0
      lvwBrowser.Sorted = False ' Sorting each element would have to slow things down, wouldn't it?
      
      If msFilter = "" Then msFilter = "*"
      lNextFile = FindFirstFile(msDir & msFilter, tWfd)
      oTotalBytes = 0
      
      Do
            On Error Resume Next
            
            sFileName = CstringToVBstring(tWfd.cFileName) ' Lots of junk past the null character.
            sEx = moFso.GetExtensionName(sFileName)
            
            If (tWfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                  eIcon = eIconType.Directory
            Else
                  eIcon = GetIconType(sEx)
            End If

            If Err > 0 Then
                  eIcon = eIconType.IconError
                  DebugLog "Icon error for file: " & sFileName & ": " & Err & ": " & Err.Description, 2
            End If
            
            oSizeBig = GetBigFileSize(tWfd, sFileName, msDir, moFso)
            
            On Error GoTo LOAD_FILES_ERROR
            If sFileName <> "." And sFileName <> "" Then ' what would be the point in providing a "." folder?
                  Set oCurrentItem = lvwBrowser.ListItems.Add(, , sFileName, , eIcon)
                  
                  ' here, let's keep an invisible second column for sorting by directory later
                  If eIcon = eIconType.Directory Then
                        oCurrentItem.ListSubItems.Add , "Type", ""
                        oCurrentItem.ListSubItems.Add , "Size", ""
                        oCurrentItem.ListSubItems.Add , "Modified", ""
                        oCurrentItem.ListSubItems.Add , "IsFolder", 0 ' Yes
                  Else
                        oCurrentItem.ListSubItems.Add , "Type", sEx
                        oCurrentItem.ListSubItems.Add , "Size", Format(oSizeBig, "#,#0")
                        oCurrentItem.ListSubItems.Add , "Modified", FormatNonLocalFileTime(tWfd.ftLastWriteTime)
                        oCurrentItem.ListSubItems.Add , "SortedSize", Format(oSizeBig, "00000000000000000")
                        oCurrentItem.ListSubItems.Add , "IsFolder", 1 ' No
                        oTotalBytes = oTotalBytes + oSizeBig
                  End If
            End If
      
      Loop While FindNextFile(lNextFile, tWfd) <> 0
           
      
      If msFilter = "*" Then msFilter = ""
      If msRealSortKey = "" Then msRealSortKey = "Name"
      lvwBrowser.SortKey = lvwBrowser.ColumnHeaders(msRealSortKey).Index - 1
      lvwBrowser.SortKey = lvwBrowser.ColumnHeaders("IsFolder").Index - 1
      lvwBrowser.Sorted = True

      AutosizeColumns

'      If bHadFocus Then lvwBrowser.SetFocus
      
'      staTusBar.Panels(eStat.BrowserStats).Text = FormatBytes(oTotalBytes, 1) & " in " & _
'            (lvwBrowser.ListItems.Count - 1) & " objects"  ' -1 for the ".." folder
      
      lvwBrowser.Visible = True
      Screen.MousePointer = vbDefault
      
      Exit Sub

LOAD_FILES_ERROR:
      Dim sErrorMsg As String
      lvwBrowser.Visible = True
      Screen.MousePointer = vbDefault
      sErrorMsg = "LoadFilesAndFolders error: " & Err.Description
      DebugLog sErrorMsg, 2
      MsgBox sErrorMsg
End Sub

Private Sub LoadHistory() ' TODO
'      Dim iIndex As Integer
'      Dim oCurrentItem As ListItem
'
'      lvwBrowser.Visible = False
'      lvwBrowser.ListItems.Clear
'      msDir = "(History)"
'      lvwBrowser.Sorted = False
'
'      For iIndex = 1 To mnuFileHistory.UBound
'            Set oCurrentItem = lvwBrowser.ListItems.Add(, "b" & CInt(iIndex), mnuFileHistory(iIndex).tag, _
'                  eIconType.Bookmark, eIconType.Bookmark)
'            oCurrentItem.ListSubItems.Add 1, , goFso.GetExtensionName(mnuFileHistory(iIndex).tag)
'      Next iIndex
'
'      gtBrowserData.ListEmpty = (lvwBrowser.ListItems.Count = 0)
'      If Not gtBrowserData.ListEmpty Then AutosizeColumns
'      lvwBrowser.Visible = True
'      staTusBar.Panels(eStat.BrowserStats).Text = lvwBrowser.ListItems.Count & " most recent files"
End Sub

'   Much can be learned that is locked within cboPath.
'
Private Sub ParsePath(ByVal sInput As String)
      ' (Bookmarks)      (Manage Bookmarks mode)
      ' (History)           (History mode)
      '                            (a blank is intrepreted as "root" / drives list mode)
      ' c:\temp\  (plain directory)
      ' c:\temp\.txt  (wildcard implied)
      ' c:\temp\READM*  (contains wildcard(s) after the directory, will filter the list)
      ' c:\temp\READMYLIPS  (no wildcard, won't filter but will move selection to a matching filename)
      
      Dim bValidPath As Boolean
      Dim sFileName As String
      sInput = Trim(sInput)
            
      msDirPrev = msDir
      msFilterPrev = msFilter
      
      If sInput = "(Bookmarks)" Then
            meMode = Bookmarks
            msDir = "(Bookmarks)"  ' Just so that (msDir = X) never accidentally returns true.
            msFilter = ""
            msPartialFileName = ""
            bValidPath = False
      
      ElseIf sInput = "(History)" Then
            meMode = History
            msDir = "(History)"
            msFilter = ""
            msPartialFileName = ""
            bValidPath = False
      
      Else
            If Not (sInput Like "*:\*") Then  ' Drives mode, root of the file system.
                  bValidPath = False
                  meMode = Drives
                  msPartialFileName = sInput
                  msDir = ""
            Else                                            ' Ordinary (folder) mode.
                  bValidPath = True
                  meMode = Files
                  msDir = SnipFileName(sInput)
                  If Not moFso.FolderExists(msDir) Then bValidPath = False
            End If
            mbDirUnchanged = (msDir = msDirPrev)
            mbGoingToParent = (msDir = ParentDirectoryOf(msDirPrev)) And Not mbDirUnchanged
      End If
      
      sFileName = SnipPath(sInput)
      
      If bValidPath Then
            msPartialFileName = ""
            If Right(sInput, 1) = "\" Then  ' c:\temp\   (just a plain old directory)
                  msFilter = ""
            ElseIf sFileName Like ".*" And Not (sFileName Like "*.") Then  ' c:\temp\.txt  (wildcard implied)
                  msFilter = "*." & moFso.GetExtensionName(sFileName)
                  
            ElseIf sFileName Like "*[?*]*" Then  ' c:\temp\READM*   (contains wildcard(s) after the directory)
                  msFilter = sFileName
                  
            ElseIf lvwBrowser.ListItems.Count > 0 Then  ' c:\temp\READMYLIPS   (some trailing characters, but no wildcard)
                  msFilter = ""
                  msPartialFileName = sFileName
            End If
      End If
      mbFilterUnchanged = (msFilter = msFilterPrev)
      
      If lvwBrowser.ListItems.Count > 0 Then msSelTextPrev = lvwBrowser.SelectedItem.Text
End Sub

Private Sub PathAddRecent(ByVal sPath As String)
      ' Supplement recent paths list, unless we are currently scrolling through them.
      ' Top of the List = Lowest of the ListIndeces = Forward(recent)most of the paths.
            
      Dim iIndex As Integer
      
      With cboPath
            
            If .ListIndex = -1 Then  ' (not scrolling through them)
                  
                  ' Delete forward history, if any, and insert current path.
                  
                  For iIndex = 0 To miLastPathIndex - 1
                        .RemoveItem 0
                  Next iIndex
                  
                  Dim iSelStart As Integer, sText As String
                  sText = cboPath.Text
                  iSelStart = cboPath.SelStart

                  If .ListCount = 0 Then
                        .AddItem sPath
                        cboPath.Text = sText
                        cboPath.SelStart = iSelStart
                        cboPath.SelLength = 0
                  
                  ' May contain repeats, but we don't want any consecutive repeats.
                  ElseIf .List(0) <> sPath Then
                        .AddItem sPath, 0
                        cboPath.Text = sText
                        cboPath.SelStart = iSelStart
                        cboPath.SelLength = 0
                  End If
                  
                  miLastPathIndex = 0
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

' The Filer calls the shots on where the dividing line goes.
' When it enlarges or shrinks itself, the other controls have to adapt. Not vice-versa.
'
' But the form will place an upper limit on the Filer's width.
' The idea is that anything bigger would be forcing the form to expand, and we shouldn't do that.
'
' The form has to respect the MINIMUM width of other controls.
' With a FIXED form width, that implies a MAXIMUM Filer width at any given time.
'
Public Sub SetMaxWidth(ByVal lMaxWidth As Long)
      mlUserCtlMaxWidth = lMaxWidth
      DebugLog "Filer max width has been set to: " & mlUserCtlMaxWidth
End Sub

Private Sub lblDivider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
      If lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "" Then
            lblDivider.tag = "Resizing"
            mlInitialPointerX = GetCursorPosX() * Screen.TwipsPerPixelX
            mlPrevPointerX = mlInitialPointerX
      End If
End Sub

Private Sub lblDivider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim lDeltaX As Long
      
      lDeltaX = GetCursorPosX() * Screen.TwipsPerPixelX - mlPrevPointerX
      If Abs(lDeltaX) > 15 And (lblDivider.Left + lDeltaX > MIN_WIDTH) And (lblDivider.Left + lDeltaX < mlUserCtlMaxWidth) _
                  And lblDivider.MousePointer = vbSizeWE And lblDivider.tag = "Resizing" Then
            lblDivider.tag = "Busy"
            
            mlPrevPointerX = mlPrevPointerX + lDeltaX
            RearrangeControls lblDivider.Left + lDeltaX
            If lblDivider.tag = "Busy" Then lblDivider.tag = "Resizing"
      Else
            lblDivider.MousePointer = vbSizeWE
      End If
End Sub

Private Sub lblDivider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      If lblDivider.MousePointer = vbSizeWE Then
            ' SaveSettingsToRegistry
            RaiseEvent SeriousResize(Width)
      End If
      lblDivider.MousePointer = 0
      lblDivider.tag = ""
      DebugLog "We are no longer resizing the Filer"
End Sub

Private Sub RearrangeControls(Optional ByVal lSupposedWidth As Long = -1)
      Dim lRightWall As Long, lToolbarRightEdge As Long
      
      If lSupposedWidth = -1 Then lSupposedWidth = Width - RIGHT_MARGIN
      Bound lSupposedWidth, MIN_WIDTH, mlUserCtlMaxWidth
      
      DebugLog "Setting Filer to supposed width: " & lSupposedWidth
      lvwBrowser.Width = lSupposedWidth
      cboPath.Width = lSupposedWidth
      lblDivider.Left = lSupposedWidth
      If miInitializings > 0 Then
            Width = lSupposedWidth + RIGHT_MARGIN
            RaiseEvent ResizeHorizontal(Width)
      Else
            DebugLog "Too bad we're not initialized though...", 2
      End If
      
      lToolbarRightEdge = btnSyncContents.Left + btnSyncContents.Width
      lRightWall = lvwBrowser.Left + lvwBrowser.Width - btnScrollToTop.Width - 30
      btnScrollToTop.Left = Max(lRightWall, lToolbarRightEdge)
End Sub

Private Sub lvwBrowser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      If meMode = History Then Exit Sub
      
      Dim lNewKey As Long
      Dim sNewKey As String
      
      With lvwBrowser
            sNewKey = ColumnHeader.Text
            If ColumnHeader.Text = "Size" Then sNewKey = "SortedSize"
            lNewKey = .ColumnHeaders(sNewKey).Index - 1
            
            .Sorted = True
            If msRealSortKey = sNewKey Then
                  .SortOrder = Abs(.SortOrder - 1)
                  .SortKey = .ColumnHeaders("Name").Index - 1
            End If
            msRealSortKey = sNewKey
            .SortKey = lNewKey
            .SortKey = .ColumnHeaders("IsFolder").Index - 1
            
            DebugLog "Sorted: " & .Sorted & "; Key: " & msRealSortKey & "; Order: " & .SortOrder _
                  & "; Column: " & ColumnHeader.Text & " (" & ColumnHeader.Index & ")"
      End With
      
      ' If meMode = Bookmarks Then BookmarkSaveChanges
End Sub

Private Sub UserControl_Initialize()
      ListViewSpyHook lvwBrowser.hwnd
      
      Set moFso = CreateObject("Scripting.FileSystemObject")
      
      miInitializings = miInitializings + 1
      
      mbDoneLoading = True
End Sub

Private Sub UserControl_Resize()
      lvwBrowser.Height = ScaleHeight - lvwBrowser.Top - BOTTOM_MARGIN
      lblDivider.Height = ScaleHeight
End Sub

' Initialization is too early for some tasks.
' If a usercontrol raises an event before the Main form loads, the event is ignored completely.
Private Sub UserControl_Show()
      mlUserCtlMaxWidth = 10000
      RearrangeControls INITIAL_WIDTH - RIGHT_MARGIN
      DoEvents
      cboPath_Change
End Sub
