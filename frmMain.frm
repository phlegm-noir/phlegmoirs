VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "VBALEDIT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "phlegmoirs"
   ClientHeight    =   6528
   ClientLeft      =   132
   ClientTop       =   432
   ClientWidth     =   8532
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6528
   ScaleWidth      =   8532
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnFolderUp 
      Height          =   240
      Left            =   1800
      MaskColor       =   &H80000000&
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1575
      Width           =   264
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   1680
      Top             =   480
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   12
   End
   Begin VB.ComboBox comboPath 
      Height          =   288
      ItemData        =   "frmMain.frx":047A
      Left            =   60
      List            =   "frmMain.frx":047C
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Text            =   "c:\my documents\mx_bin\"
      Top             =   1200
      Width           =   2292
   End
   Begin vbalEdit.vbalRichEdit Editor 
      Height          =   2532
      Left            =   3000
      TabIndex        =   5
      Tag             =   "c:\my documents\mx_bin\trash.txt"
      Top             =   1560
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   4466
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
   Begin VB.PictureBox picToolBox 
      Height          =   612
      Left            =   -120
      ScaleHeight     =   564
      ScaleWidth      =   8484
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8532
      Begin VB.CommandButton btnFont 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   732
      End
      Begin VB.TextBox DicBox 
         Height          =   372
         Left            =   3480
         MaxLength       =   50
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Text            =   "Look it up!"
         Top             =   120
         Width           =   3732
      End
      Begin VB.CheckBox Check1 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   732
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   120
         Picture         =   "frmMain.frx":047E
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   732
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   264
      Left            =   0
      TabIndex        =   1
      Top             =   6264
      Width           =   8532
      _ExtentX        =   15050
      _ExtentY        =   466
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
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4332
      Left            =   60
      TabIndex        =   0
      Tag             =   "c:\my documents\mx_bin\"
      Top             =   1560
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   7641
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "filename"
         Text            =   "Um"
         Object.Width           =   3351
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
      Begin VB.Menu mnuFileDiv1 
         Caption         =   "-"
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
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditIncFont 
         Caption         =   "&Increase Font Size\tAlt+F6"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFilebrowser 
         Caption         =   "File &Browser"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewDictionary 
         Caption         =   "&Dictionary"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuViewThesaurus 
         Caption         =   "&Thesaurus"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListTest 
         Caption         =   "&Test"
      End
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "Write"
      Visible         =   0   'False
      Begin VB.Menu mnuWriteCut 
         Caption         =   "Cut\tCtrl+X"
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
Dim DateString As String
Dim Stat1 As StatType
Dim FuckIHateThis As Boolean
'Dim Var1 As Variant
Dim tempfont As New StdFont
'Dim EditorAccelTable() As ACCEL
'Dim ControlInfoData As CONTROLINFO
'Dim ctrlInfo1 As CTRLINFO
'Dim tempOLE As olelib.IOleObject
'Dim tempOLEcontrol As olelib.IOleControl

Const statStats = 1, statModified = 2, statInsert = 3
Const statSelText = 4, statLastSaved = 5, statTips = 6

Private Sub btnFolderUp_Click()
      Dim islash As Integer
      
      islash = InStrRev(comboPath, "\", Len(comboPath) - 1, vbTextCompare)
      comboPath = Left(comboPath, islash)
      
      lvwFiles.SetFocus
End Sub


Private Sub btnFont_Click()
      'Dim tempfont As New StdFont ' New because StdFont is a Class
      
      With dlgFont 'make the dialog choices begin with what the editor shows
            .Flags = cdlCFBoth + cdlCFApply ' and allow for all font types.
            .FontName = Editor.GetFont.Name                    ' btw, Apply doesn't work
            .FontBold = Editor.GetFont.Bold
            .FontUnderline = Editor.GetFont.Underline
            .FontSize = Editor.GetFont.Size
      End With

      On Error Resume Next 'trap the error. if they hit cancel, do nothing and exit
      dlgFont.ShowFont
      If Err.Number = cdlCancel Then Exit Sub
      Err.Clear
      On Error GoTo 0  'btw, I think this has the effect of err.Clear
      
      With tempfont
            .Name = dlgFont.FontName
            .Bold = dlgFont.FontBold
            .Italic = dlgFont.FontItalic
            .Underline = dlgFont.FontUnderline
            .Size = dlgFont.FontSize
      End With
      Editor.SetFont tempfont, , , , ercSetFormatAll
      Me.Caption = Editor.GetFont.Name & " " & Editor.GetFont.Charset & " " & Editor.GetFont.Size
End Sub


Private Sub chkFileBrowser_Click()
    lvwFiles.Visible = Not lvwFiles.Visible
    comboPath.Visible = Not comboPath.Visible
    btnFolderUp.Visible = Not btnFolderUp.Visible
    RearrangeControls
    Editor.SetFocus
End Sub

Private Sub comboPath_Change()
      ' Type a directory into comboPath.  Valid paths will be loaded as you type.
      '     (actually, anything ending in "\" will be loaded)
      '     (also, you can specify an attribute)
      
      Dim s As String
      p = Dir(comboPath.Text)
      If p <> "" Then
            lvwFiles.Tag = TrimPath(comboPath)
            FillFileBrowser p
      ElseIf Right(comboPath, 1) = "\" Then  ' when an empty directory is found, we
            lvwFiles.ListItems.Clear                ' want to know that it is empty
            lvwFiles.Tag = TrimPath(comboPath)
      End If
End Sub

Private Sub comboPath_GotFocus()
      comboPath.SelStart = Len(comboPath)
End Sub

Private Sub comboPath_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim islash As Integer
      If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
            islash = InStrRev(comboPath, "\", , vbTextCompare)
            comboPath = Left(comboPath, islash)
      ElseIf KeyCode = vbKeyReturn Then
            
      End If
End Sub

Private Sub DicBox_GotFocus()
    DicBox.SelStart = 0
    DicBox.SelLength = Len(DicBox)
End Sub

Private Sub DicBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        retval = ShellExecute(0, "open", _
                "http://dictionary.reference.com/search?q=" & DicBox, 0, "", 8)
    End If
End Sub

Private Sub DicBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DicBox = Data.GetData(vbCFText)
    DicBox_KeyPress vbKeyReturn
End Sub

Private Sub DicBox_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    DicBox.SelStart = 0
    DicBox.SelLength = Len(DicBox)
End Sub

Private Sub Editor_KeyDown(KeyCode As Integer, Shift As Integer)
'      pOldProc = SetWindowLong(Editor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)
      Select Case KeyCode
'            Case vbKeyAdd
'                  If Shift = vbAltMask Then
'                        SendMessage Editor.RichEdithWnd, EM_SETFONTSIZE, ByVal 1, 0
''                        Set tempfont = Editor.GetFont
''                        tempfont.Size = tempfont.Size + 1
''                        Editor.SetFont tempfont, , , , ercSetFormatAll
'                        Me.Caption = Editor.GetFont.Name & " " & Editor.GetFont.Charset & " " & Editor.GetFont.Size
'                  End If
'
'            Case vbKeySubtract
'                  If Shift = vbAltMask Then
'                        Set tempfont = Editor.GetFont
'                        tempfont.Size = tempfont.Size - 1
'                        Editor.SetFont tempfont, , , , ercSetFormatAll
'                        Me.Caption = Editor.GetFont.Name & " " & Editor.GetFont.Charset & " " & Editor.GetFont.Size
'                  End If
                  
      End Select
'      retval = SetWindowLong(Editor.RichEdithWnd, GWL_WNDPROC, pOldProc)
'      pOldProc = 0
End Sub

Private Sub Editor_SelectionChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As vbalEdit.ERECSelectionTypeConstants)
    Dim lineindex As Long, charindex As Long

    lineindex = Editor.CurrentLine
    charindex = SendMessage(Editor.RichEdithWnd, EM_LINEINDEX, ByVal lineindex, 0)

    With Stat1
        .i = lMin + 1
        .Y = lineindex + 1
        .X = lMin - charindex + 1
        .xmax = SendMessage(Editor.RichEdithWnd, EM_LINELENGTH, ByVal charindex, 0) + 1
    End With

    FillStats

    StatusBar1.Panels(statSelText) = lMax - lMin

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim ctrlhWnd As Long
      Dim CursorPos As POINTAPI

      'pOldProc = SetWindowLong(Editor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)
      Select Case KeyCode
            Case vbKeyTab
                  If Shift = vbCtrlMask And Me.ActiveControl.Name = "Editor" Then
                        lvwFiles.SetFocus
                  ElseIf Me.ActiveControl.Name = "Editor" Then
                        
                  Else
                  
                  End If
            
            Case 93
                  GetCursorPos CursorPos
                  ctrlhWnd = WindowFromPoint(CursorPos.X, CursorPos.Y)
                  StatusBar1.Panels(statLastSaved) = "c: " & ctrlhWnd
                  
            Case vbKeyB
                  Me.Caption = Me.ActiveControl.Name
                  
            Case Else
                  StatusBar1.Panels(statLastSaved) = KeyCode
      End Select
'      retval = SetWindowLong(Editor.RichEdithWnd, GWL_WNDPROC, pOldProc)
'      pOldProc = 0
End Sub



Private Sub Form_Load()
      Dim fn As Variant
      Dim i As Integer
      Dim d As Variant
'      Dim tempinfo As MENUITEMINFO
'      Dim hMenu As Long, retval As Long
'
'      hMenu = GetMenu(hwnd)
'      hMenu = GetSubMenu(hMenu, 2)
'      retval = ModifyMenu(hMenu, 0, MF_STRING + MF_BYPOSITION, 2, "&Penis" + vbTab + "Ctrl+P")
      
      mnuFileOpen.Caption = "&Open" & vbTab & "Ctrl+F9"
      mnuEditIncFont.Caption = "&Increase Font Size" & vbTab & "Alt+="
      
      
      mnuWriteCut.Caption = "Cu&t" & vbTab & "Ctrl+X"
      mnuWriteCopy.Caption = "&Copy" & vbTab & "Ctrl+C"
      mnuWritePaste.Caption = "&Paste" & vbTab & "Ctrl+V"

'      GetMenuItemInfo hMenu, 0, True, tempinfo 'position 1 should be
'      tempinfo.dwTypeData = "&Test\tCtrl+F9"                'mnuListTest
'      SetMenuItemInfo hMenu, 0, True, tempinfo
      
      d = Date
      DateString = Year(d) & "-" & Format(Month(d), "0#") & "-" & Format(Day(d), "0#")
      
      fn = Dir(comboPath)
      While fn <> ""
          lvwFiles.ListItems.Add 1, , fn
          fn = Dir
      Wend
      
      'retval = Editor.LoadFromFile(Editor.Tag, SF_TEXT)
      Stat1.imax = Editor.CharacterCount
      FillStats
      StatusBar1.Panels(statModified) = ""

'      pOldProc = SetWindowLong(Editor.RichEdithWnd, GWL_WNDPROC, AddressOf WindowProc)

'      Set tempOLE = RichEdit1
      
      
      
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' If we open a popupmenu, and then right click off into space,
      '   the mousedown event is called for the form (not for the control we are
      '   hovering over nor the menu itself.)
      ' Our form doesn't need it.  We'll have him pass it to the control it's over.

      Dim ctrlhWnd As Long
      Dim retval As Long
      Dim CursorPos As POINTAPI
      
      If Button <> vbRightButton Or Shift <> 0 Then Exit Sub

      GetCursorPos CursorPos
      ctrlhWnd = WindowFromPoint(CursorPos.X, CursorPos.Y)
      If ctrlhWnd = Editor.RichEdithWnd Then
            mouse_event MOUSEEVENTF_LEFTDOWN, CursorPos.X, CursorPos.Y, 0, 0
            mouse_event MOUSEEVENTF_LEFTUP, CursorPos.X, CursorPos.Y, 0, 0
      ElseIf ctrlhWnd = lvwFiles.hwnd Then
            FuckIHateThis = True
            mouse_event MOUSEEVENTF_LEFTDOWN, CursorPos.X, CursorPos.Y, 0, 0
            mouse_event MOUSEEVENTF_RIGHTUP, CursorPos.X, CursorPos.Y, 0, 0
      ElseIf ctrlhWnd = DicBox.hwnd Then
            mouse_event MOUSEEVENTF_LEFTDOWN, CursorPos.X, CursorPos.Y, 0, 0
            'mouse_event MOUSEEVENTF_LEFTUP, CursorPos.X, CursorPos.Y, 0, 0
      Else
            SendMessage frmMain.hwnd, WM_CANCELMODE, 0, 0
      End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      StatusBar1.Panels(statLastSaved) = X & " " & Y

End Sub

Private Sub Form_Resize()
    RearrangeControls
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'this merely reverses the sort order

    lvwFiles.SortOrder = Abs(lvwFiles.SortOrder - 1)
End Sub

Private Sub lvwFiles_DblClick()
      Editor.Tag = lvwFiles.Tag & lvwFiles.SelectedItem.Text
      Editor.LoadFromFile Editor.Tag, SF_TEXT
      StatusBar1.Panels(statModified) = ""
End Sub


Private Sub lvwFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If FuckIHateThis Then
            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
            FuckIHateThis = True
      End If
      If (Button = vbRightButton And Shift = 0) Then
          Me.PopupMenu mnuList
      End If
End Sub

Private Sub mnuFileNew_Click()
      Debug.Print "HAHAHAHAAHAHHA" & vbKeyReturn
End Sub

Private Sub mnuFileSave_Click()
    Editor.SaveToFile Editor.Tag, SF_TEXT
    StatusBar1.Panels(statModified) = ""
    'Editor.Refresh
End Sub

Private Sub mnuViewDictionary_Click()
    If Editor.SelectedText <> "" Then DicBox = Editor.SelectedText
    DicBox.SetFocus
    If Editor.SelectedText <> "" Then DicBox_KeyPress vbKeyReturn
End Sub

Private Sub mnuViewFilebrowser_Click()
    chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub Editor_Change()

    If StatusBar1.Panels(statModified) = "" Then
        StatusBar1.Panels(statModified) = "Modified"
    End If

    With Stat1
        .imax = Editor.CharacterCount
        .ymax = Editor.LineCount
    End With

    FillStats

End Sub

Private Sub Editor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If (Button = vbRightButton And Shift = 0) Then
          Me.PopupMenu mnuWrite
      End If
End Sub

Private Sub FillStats()

    StatusBar1.Panels(statStats) = "Char: " & Stat1.i & "/" & Stat1.imax _
        & "  Ln: " & Stat1.Y & "/" & Stat1.ymax & "  Col: " & Stat1.X _
        & "/" & Stat1.xmax
End Sub


Private Sub RearrangeControls()

    ' Put the various controls where they need to be.
    '   Editor, lvwFiles
    ' Made to go on a window resize or when showing or hiding a control

    Dim h As Integer, w As Integer, top1 As Integer, left1 As Integer
    Dim lineindex As Long, charindex As Long, lMin As Long, lMax As Long
    Const topmargin = 800
    Const leftmargin = 60
    Const rightmargin = 150
    Const midspace = 100

    Editor.Visible = False ' MUCH faster if you turn him off while thinking

    top1 = 0
    If picToolBox.Visible Then top1 = top1 + picToolBox.Height + midspace
    h = frmMain.Height - top1 - topmargin
    If StatusBar1.Visible Then h = h - StatusBar1.Height

    left1 = leftmargin
    If lvwFiles.Visible Then left1 = left1 + lvwFiles.Width + midspace
    w = frmMain.Width - left1 - rightmargin


    Editor.Top = top1
    Editor.Left = left1
    If h > 0 Then Editor.Height = h
    If w > 0 Then Editor.Width = w

    h = h - comboPath.Height
    If h > 0 Then lvwFiles.Height = h
    comboPath.Top = top1
    lvwFiles.Top = top1 + comboPath.Height
    btnFolderUp.Top = lvwFiles.Top + 10

    Editor.Visible = True

    ' a few things in the statusbar could change in a window resize:
    '   x, xmax, y, ymax
    ' and some shouldn't change:
    '   i, imax,   (we're not adding or deleting characters or moving the cursor)
    '   sellength

    Editor.GetSelection lMin, lMax
    lineindex = Editor.CurrentLine
    charindex = SendMessage(Editor.RichEdithWnd, EM_LINEINDEX, ByVal lineindex, 0)

    With Stat1
        .X = lMin - charindex + 1
        .xmax = SendMessage(Editor.RichEdithWnd, EM_LINELENGTH, ByVal charindex, 0) + 1
        .Y = lineindex + 1
        .ymax = Editor.LineCount
    End With
    FillStats
End Sub

Private Sub mnuViewThesaurus_Click()
    If Editor.SelectedText <> "" Then DicBox = Editor.SelectedText
    DicBox.SetFocus
    If Editor.SelectedText <> "" Then
        retval = ShellExecute(0, "open", _
                "http://thesaurus.reference.com/search?q=" & DicBox, 0, "", 8)
        Me.SetFocus
    End If
End Sub

Private Sub FillFileBrowser(ByVal s As String)
    lvwFiles.ListItems.Clear
    Do
        lvwFiles.ListItems.Add 1, , s
        s = Dir
    Loop Until s = ""
End Sub

Private Sub mnuWriteCopy_Click()
      Editor.Copy
End Sub

Private Sub mnuWriteCut_Click()
      Editor.Cut
End Sub

Private Sub mnuWritePaste_Click()
      Editor.Paste
End Sub

Private Function TrimPath(ByVal pathname As String)
      Dim islash As Integer
      islash = InStrRev(pathname, "\", , vbTextCompare)
      TrimPath = Left(pathname, islash)
End Function
