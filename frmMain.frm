VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DD32A320-6E5E-44C8-BCE6-5908CA400231}#1.0#0"; "AGRICHEDIT.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "phlegmoirs"
   ClientHeight    =   6528
   ClientLeft      =   132
   ClientTop       =   432
   ClientWidth     =   8532
   LinkTopic       =   "Form1"
   ScaleHeight     =   6528
   ScaleWidth      =   8532
   StartUpPosition =   2  'CenterScreen
   Begin agRichEditBox.agRichEdit agRichEdit1 
      Height          =   2412
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   4255
      Version         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ViewMode        =   0
      Border          =   0   'False
      AutoURLDetect   =   0   'False
      ScrollBars      =   0
   End
   Begin VB.PictureBox picToolBox 
      Height          =   612
      Left            =   -120
      ScaleHeight     =   564
      ScaleWidth      =   8484
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8532
      Begin VB.CheckBox Check1 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   732
      End
      Begin VB.CheckBox chkFileBrowser 
         CausesValidation=   0   'False
         Height          =   580
         Left            =   120
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   732
      End
   End
   Begin RichTextLib.RichTextBox rtxtWrite 
      Height          =   4932
      Left            =   2520
      TabIndex        =   2
      Tag             =   "c:\my documents\phlegmoir1.txt"
      Top             =   1080
      Width           =   5652
      _ExtentX        =   9970
      _ExtentY        =   8700
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      OLEDragMode     =   0
      TextRTF         =   $"frmMain.frx":0DF0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Top             =   1080
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
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFilebrowser 
         Caption         =   "File &Browser"
         Shortcut        =   {F5}
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
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuWriteCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuWritePaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
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
Dim CursorXY As POINTAPI

Const statStats = 1, statModified = 2, statInsert = 3
Const statSelText = 4, statLastSaved = 5, statTips = 6

Private Sub Check1_Click()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub chkFileBrowser_Click()
    lvwFiles.Visible = Not lvwFiles.Visible
    RearrangeControls
    rtxtWrite.SetFocus
End Sub

Private Sub Form_Activate()
    'RearrangeControls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'StatusBar1.Panels(2) = StatusBar1.Panels(2) + 1
    Select Case KeyCode
    
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case "a"
'            frmMain.WindowState = vbMinimized
'    End Select
'    StatusBar1.Panels(2) = KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'StatusBar1.Panels(2) = StatusBar1.Panels(2) - 1
    Select Case KeyCode
    
    End Select
    
End Sub

Private Sub Form_Load()
'    Dim fs, f As Variant
'    Dim mxname As String
    Dim fn As Variant
    Dim i As Integer
    Dim d As Variant
    
    d = Date
    DateString = Year(d) & "-" & Format(Month(d), "0#") & "-" & Format(Day(d), "0#")
    
    fn = Dir(lvwFiles.Tag)
    While fn <> ""
        lvwFiles.ListItems.Add 1, , fn
        fn = Dir
    Wend
    
    
'    Const ForReading = 1, ForWriting = 2, ForAppending = 3
'    mxname = "C:\my documents\mx_bin\2004-09-01 my editor ideas.txt"
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.OpenTextFile(mxname, ForReading)
'    txtWrite = f.readall
'    f.Close
    
    rtxtWrite.LoadFile (rtxtWrite.Tag)
    StatusBar1.Panels(statModified) = ""
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton And Shift = 0 Then
        retval = GetCursorPos(CursorXY)
        mouse_event MOUSEEVENTF_LEFTDOWN, CursorXY.x, CursorXY.y, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, CursorXY.x, CursorXY.y, 0, 0
    End If
End Sub

Private Sub Form_Resize()
    RearrangeControls
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'this merely reverses the sort order
    ' (data type of sortorders are not technically boolean,
    '   which is stupid)
    
    lvwFiles.SortOrder = Abs(lvwFiles.SortOrder - 1)
End Sub

Private Sub lvwFiles_DblClick()
    rtxtWrite.Tag = lvwFiles.Tag & lvwFiles.SelectedItem.Text
    rtxtWrite.LoadFile rtxtWrite.Tag, rtfText
    StatusBar1.Panels(statModified) = ""
End Sub

Private Sub lvwFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then Me.PopupMenu mnuList
End Sub

Private Sub mnuFileSave_Click()
    rtxtWrite.SaveFile rtxtWrite.Tag, rtfText
    StatusBar1.Panels(statModified) = ""
    'rtxtWrite.Refresh
End Sub

Private Sub mnuViewFilebrowser_Click()
    chkFileBrowser = Abs(chkFileBrowser.Value - 1)
End Sub

Private Sub rtxtWrite_Change()
    
    If StatusBar1.Panels(statModified) = "" Then
        StatusBar1.Panels(statModified) = "Modified"
    End If
    
    With Stat1
        .imax = Len(rtxtWrite.Text) + 1
        .ymax = SendMessage(rtxtWrite.hwnd, EM_GETLINECOUNT, 0, 0)
    End With
    
    FillStats

End Sub


Private Sub rtxtWrite_KeyDown(KeyCode As Integer, Shift As Integer)
    StatusBar1.Panels(statLastSaved) = KeyCode
End Sub

Private Sub rtxtWrite_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'hmmm...
    If Button = vbRightButton And Shift = 0 Then
        retval = GetCursorPos(CursorXY)
        mouse_event MOUSEEVENTF_LEFTDOWN, CursorXY.x, CursorXY.y, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, CursorXY.x, CursorXY.y, 0, 0
    End If
End Sub

Private Sub rtxtWrite_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbRightButton And Shift = 0) Then
        Me.PopupMenu mnuWrite
    End If
End Sub

Private Sub FillStats()
    
    StatusBar1.Panels(statStats) = "Char: " & Stat1.i & "/" & Stat1.imax _
        & "  Ln: " & Stat1.y & "/" & Stat1.ymax & "  Col: " & Stat1.x _
        & "/" & Stat1.xmax
End Sub

Private Sub rtxtWrite_SelChange()
    Dim rtxtlength As Long
    Dim charindex As Long, lineindex As Long

    ' EM_LINEINDEX gets the index, from the rtxt's beginning,
    '   of the first character on a specific line.

    lineindex = rtxtWrite.GetLineFromChar(rtxtWrite.SelStart)
    charindex = SendMessage(rtxtWrite.hwnd, EM_LINEINDEX, ByVal lineindex, 0)

    With Stat1
        .i = rtxtWrite.SelStart + 1
        .y = lineindex + 1
        .x = rtxtWrite.SelStart - charindex + 1
        .xmax = SendMessage(rtxtWrite.hwnd, EM_LINELENGTH, ByVal charindex, 0) + 1
    End With

    FillStats

    StatusBar1.Panels(statSelText) = rtxtWrite.SelLength
End Sub

Private Sub RearrangeControls()

    ' Put the various controls where they need to be.
    ' Made to go on a window resize or when showing or hiding a control

    Dim h As Integer, w As Integer, top1 As Integer, left1 As Integer
    Dim lineindex As Long, charindex As Long
    Const topmargin = 800
    Const leftmargin = 60
    Const rightmargin = 150
    Const midspace = 100

    rtxtWrite.Visible = False ' MUCH faster if you turn him off while thinking
    
    top1 = 0
    If picToolBox.Visible Then top1 = top1 + picToolBox.Height + midspace
    h = frmMain.Height - top1 - topmargin
    If StatusBar1.Visible Then h = h - StatusBar1.Height
    
    left1 = leftmargin
    If lvwFiles.Visible Then left1 = left1 + lvwFiles.Width + midspace
    w = frmMain.Width - left1 - rightmargin
    
    
    rtxtWrite.Top = top1
    rtxtWrite.Left = left1
    If h > 0 Then rtxtWrite.Height = h
    If w > 0 Then rtxtWrite.Width = w
    
    If h > 0 Then lvwFiles.Height = h
    lvwFiles.Top = top1
    lvwFiles.Left = leftmargin
    
    rtxtWrite.Visible = True
    
    lineindex = rtxtWrite.GetLineFromChar(rtxtWrite.SelStart)
    charindex = SendMessage(rtxtWrite.hwnd, EM_LINEINDEX, ByVal lineindex, 0)
    
    With Stat1
        .x = rtxtWrite.SelStart - charindex + 1
        .xmax = SendMessage(rtxtWrite.hwnd, EM_LINELENGTH, ByVal charindex, 0) + 1
        .y = lineindex + 1
        .ymax = SendMessage(rtxtWrite.hwnd, EM_GETLINECOUNT, 0, 0)
    End With
    FillStats
End Sub
