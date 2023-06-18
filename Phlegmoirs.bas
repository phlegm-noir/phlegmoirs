Attribute VB_Name = "modPhlegmoirs"


' *************************************************************
' My custom types
' *************************************************************

Option Compare Text
Option Explicit

Public Type TStatType
    i As Long
    imax As Long
    Y As Long
    ymax As Long
    X As Long
    xmax As Long
End Type

Public Type TEditorPrefs
      WordWrap As Integer
      ReadOnly As Integer ' new
      SelStart As Long
      SelEnd As Long
      FirstVisibleLine As Long
      FontName As String * 255
      FontSize As Currency
      FontBold As Boolean
      FontItalic As Boolean
      FontUnderline As Boolean
      FontStrikethrough As Boolean
      TextColor As Long 'new
      ScrollPos As POINTAPI
      
      ' Reserved, not used:
      BackColor As Long 'new
End Type

Public Type TWindowPrefs
      WNP As WINDOWPLACEMENT
      BrowserWidth As Long
      SortMethod As Integer
      SortKey As Integer
      NameColumn As Integer  ' These store the column positions, and are negative if hidden
      TypeColumn As Integer
      SizeColumn As Integer
      ModifiedColumn As Integer
      
      ShowFileBrowser As Boolean
      ShowStatusBar As Boolean
      ShowToolBar As Boolean
      ShowFind As Boolean 'new
      BookmarkCount As Integer
      HistoryCount As Integer
      AutoLoadFile As String * 255
      cboPath As String * 255
      
      ' Reserved, not used:
      FullScreen As Boolean 'new
End Type

Public Type TBrowserData
      Dir As String
      DirPrev As String
      InputPrev As String
      Filter As String
      FilterPrev As String
      PartialFileName As String
      SelTextPrev As String
      DirUnchanged As Boolean
      FilterUnchanged As Boolean
      GoingToParent As Boolean
      BookmarkMode As Boolean
      DrivesMode As Boolean
      HistoryMode As Boolean
      ValidPath As Boolean
      Error As Boolean
      ListEmpty As Boolean
      
      ItemClicked As Boolean
      ButtonPressed As Boolean
      MouseButton As Integer
      Shift As Integer
      RecentPath As Integer
End Type

Public Type TImageData
      OutPic As Object
      DefaultHeight As Integer
      DefaultWidth As Integer
      PrevX As Single
      PrevY As Single
      
      Dragging As Boolean
      Moved As Boolean
      Zoomed As Boolean
End Type

' *************************************************************
' Global enumerations
' *************************************************************

' eMode is meant to encompass browser modes, editor modes, and file types.
Enum eMode
      Directory = 1
      Properties = 2
      Drive = 3
      Text = 4
      other = 5
      Picture = 6
      Error = 7
      Bookmark = 8
      Floppy = 9
      Network = 10
      Cdrom = 11
      rtf = 12
End Enum

Enum eStat
      BrowserStats = 1
      Stats = 2
      Modified = 3
      SelText = 4
      Tips = 5
End Enum

Enum eDirection
      Forward = 1
      back = -1
End Enum

Enum eQuery
      Find = 0
End Enum

      
' *************************************************************
' Global Variables
' *************************************************************

Public gpOldLvwProc As Long, gpOldpicBrowserProc As Long, gpOldpicEditorProc As Long
Public gpOldfrmFullScreenProc As Long
Public objtest As Object
Public gFSO As Object  ' nothing I hate more than declaring one of these in each function, for one tiny little purpose.
Public gBrowserData As TBrowserData
Public gImageData As TImageData
Public giEditorMode As eMode  ' use eMode
Public gfFullScreenMode As Integer

Public Const MoveIncrement = -512
Public Const glMAX_LONG_INTEGER = &H7FFFFFFF   '   2147483647
Public QUOT As String ' = Chr(34) can't do this in the declaration, I see!


' *************************************************************
' Window Procedures
' *************************************************************
    
'Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
'            ByVal wParam As Long, ByVal lParam As Long) As Long
'
'      If uMsg <> 15 And uMsg <> 32 And uMsg <> 132 And uMsg <> 512 And _
'            uMsg <> 275 Then
'            ' WM_PAINT, WM_SETCURSOR, WM_SETICON, WM_MOUSEMOVE
'            ' WM_TIMER
'            Debug.Print Hex(uMsg) & vbTab & "(" & uMsg & ")" & _
'                        vbTab & wParam & vbTab & lParam
'      End If
'
'      Select Case uMsg
'            Case 3
'
'      End Select
'      WindowProc = CallWindowProc(gpOldProc, hwnd, uMsg, wParam, lParam)
'End Function

Function FormatBytes(ByVal curBytes As Currency, iPrecision As Integer) As String
      ' This function takes a quantity of bytes as a currency value ('cause it's 64-bit),
      ' formats it to read likeathis:
      
      ' 45.2 MB
      ' 300.2 KB
      ' 666 Bytes
      ' 99.4444 GB
      ' 20000 TB  (not supporting terabytes at the moment, sorry folks)
      
      ' ...with iPrecision digits after the demical.
      
      If curBytes < 1024@ Then
            FormatBytes = CStr(curBytes) & " b"
      ElseIf curBytes < 1048576@ Then
            FormatBytes = CStr(Round(curBytes / 1024@, iPrecision)) & " KB"
      ElseIf curBytes < 1073741824@ Then
            FormatBytes = CStr(Round(curBytes / 1048576@, iPrecision)) & " MB"
      ElseIf curBytes < 1099511627776@ Then
            FormatBytes = CStr(Round(curBytes / 1073741824@, iPrecision)) & " GB"
      Else
            FormatBytes = "Size Unknown"
      End If
End Function

Public Function FormatNonLocalFileTime(NLFT As FILETIME) As String
      ' example date string:   2005-03-15 6:14:21
      
      Dim localTime As FILETIME
      Dim sysTime As SYSTEMTIME
      Dim hFile As Long
      
      FileTimeToLocalFileTime NLFT, localTime
      FileTimeToSystemTime localTime, sysTime
      With sysTime
            FormatNonLocalFileTime = .wYear & "-" & Format(.wMonth, "00") & "-" & Format(.wDay, "00") _
                  & ", " & Format(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00")
      End With
End Function

 Function GetRealStdFont(ByRef editor As agRichEdit, Optional ByRef lTextColor As Long) As StdFont
      ' OK, I put in a byref value to pass on the text color, which is not included in the StdFont type.
      ' The function returns a StdFont containing the rest of the font data.
      
      Const CFE_BOLD As Long = &H1
      Const CFE_ITALIC As Long = &H2
      Const CFE_STRIKEOUT As Long = &H8
      Const CFE_UNDERLINE As Long = &H4
      
      Const CFM_FACE As Long = &H20000000
      Const CFM_SIZE As Long = &H80000000
      Const CFM_CHARSET As Long = &H8000000
      Const CFM_BOLD As Long = &H1
      Const CFM_COLOR As Long = &H40000000
      Const CFM_ITALIC As Long = &H2
      Const CFM_LINK As Long = &H20
      Const CFM_OFFSET As Long = &H10000000
      Const CFM_STRIKEOUT As Long = &H8
      Const CFM_UNDERLINE As Long = &H4
      Const CFM_WEIGHT As Long = &H400000


      Dim char2 As CHARFORMAT2
      Dim lRetVal As Long
      Dim fntNew As New StdFont
      
      char2.cbSize = LenB(char2)
      ' Tell it which CHARFORMAT2 properties carry relevant data:
      char2.dwMask = CFM_SIZE + CFM_FACE + CFM_BOLD + CFM_COLOR + _
                  CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE
            ' I took out cfm_charset and cfm_weight, because they are set automatically
      lRetVal = SendMessage(editor.RichEdithWnd, EM_GETCHARFORMAT, ByVal 0, char2)
      
      If lRetVal <> 0 Then
            With fntNew
                  .Size = char2.yHeight / 20!
                  .Name = CstringToVBstring(StrConv(char2.szFaceName, vbUnicode))
                  .Bold = char2.dwEffects And CFE_BOLD
                  .Italic = char2.dwEffects And CFE_ITALIC
                  .Strikethrough = char2.dwEffects And CFE_STRIKEOUT
                  .Underline = char2.dwEffects And CFE_UNDERLINE
            End With
            lTextColor = char2.crTextColor
            Set GetRealStdFont = fntNew
      Else
            Set GetRealStdFont = Nothing
      End If
End Function

      ' The problem: right arrow key wants to scroll forward.  I want it to do things
      ' like BrowserExecuteItem (open folder/drive, etc.) and it's REALLY ANNOYING
      ' when the listview does both at once.
      
      ' The solution: The nonexistant F13 key will be
      ' given the same right arrow implementation, and this window procedure will
      ' merely wait for a right arrow (without ctrl), and when it finds one, it will
      ' continue the window procedure as though F13 had been pressed.
      
      ' If somebody wants to scroll right with arrow keys, he may use ctrl+right,
      ' and nothing bad will happen.
      
      ' The one regret: I hope the system dependent "scan codes" in the low-order
      ' word of a WM_KEYDOWN F13 message aren't used for anything.  Because
      ' it's getting scan codes meant for a right arrow.
'
Public Function ListViewProc(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

      Select Case uMsg
            Case WM_KEYDOWN
                  
                  'Debug.Print "LVW_BROWSER: WM_KEYDOWN: " & wParam & " " & lParam
                  
                  If wParam = vbKeyRight And Not IsKeyDown(VK_CONTROL) Then
                                    
                        ListViewProc = CallWindowProc(gpOldLvwProc, hwnd, _
                              WM_KEYDOWN, vbKeyF13, lParam)
                        Exit Function
                  End If
            
            Case WM_MOUSEWHEEL
                  ' This procedure is going to scroll horizontally when you mousewheel
                  ' over the hscrollbar.  JUST LIKE OPERA!!! <3

                  Const LVHT_ABOVE As Long = &H8
                  Const LVHT_BELOW As Long = &H10
                  Const LVHT_NOWHERE As Long = &H1
                  Const LVHT_ONITEMICON As Long = &H2
                  Const LVHT_ONITEMLABEL As Long = &H4
                  Const LVHT_ONITEMSTATEICON As Long = &H8
                  Const LVHT_TORIGHT As Long = &H20
                  Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

                  Const SB_ENDSCROLL As Long = 8
                  Const SB_LEFT As Long = 6
                  Const SB_LINELEFT As Long = 0
                  Const SB_LINERIGHT As Long = 1
                  Const SB_PAGELEFT As Long = 2
                  Const SB_PAGERIGHT As Long = 3
                  Const SB_RIGHT As Long = 7

                  Dim lRetVal As Long
                  Dim lWheelTurns As Long  ' Can be positive or negative (not zero, of course!)
                                                                        ' If you spin the wheel REALLY fast, it sends a single
                                                                        ' WM_MOUSEWHEEL message grouping multiple turns together.
                                                                        ' But it's normally either 1 or -1, sending a message for each turn.
                  Dim HitTestInfo As LVHITTESTINFO
                  Dim recClient As RECT

                  lWheelTurns = MAKEPOINT(wParam).Y / WHEEL_DELTA
                  lRetVal = GetWindowRect(hwnd, recClient)
                  'Debug.Print "from wm_mousewheel:     x: "; MAKEPOINT(lParam).X & " y: " & MAKEPOINT(lParam).Y
                  'Debug.Print "from wm_mousewheel:     x: "; recClient.Left & " y: " & recClient.Top

                  With HitTestInfo
                        .pt.X = MAKEPOINT(lParam).X - recClient.Left
                        .pt.Y = MAKEPOINT(lParam).Y - recClient.Top
                  End With
                  lRetVal = SendMessage(hwnd, LVM_HITTEST, ByVal 0, HitTestInfo)
'                  Debug.Print "from wm_mousewheel:     x: "; HitTestInfo.pt.X & " y: " & HitTestInfo.pt.Y & _
                        "   " & HitTestInfo.flags & "    wheel turns: " & lWheelTurns

                  If (HitTestInfo.flags And LVHT_BELOW) Or (HitTestInfo.flags And LVHT_ABOVE) Then
'                  If Not CBool(HitTestInfo.flags And LVHT_TORIGHT) Then
                        
                  ' For the moment, as you can see above, I'm commenting out the above/below condition.
                  ' I'm thinking it's better to always mousewheel left/right, and scroll manually for updown.
                  ' Since the listview is pretty thin, and the scrollbar is just right there begging to be touched.
                  ' Also, we DO NOT want them attempting left/right buttons to move the scrollbar!
                  ' That'll find them in some hot water, but at least it doesn't open files (just folders/drives).
                                    
                        Dim iTurn As Integer
                        If lWheelTurns > 0 Then
                              For iTurn = 1 To lWheelTurns * 3
                                    lRetVal = SendMessage(hwnd, WM_HSCROLL, ByVal SB_LINELEFT, ByVal 0)
                              Next iTurn
                              Exit Function
                        ElseIf lWheelTurns < 0 Then
                              For iTurn = -1 To lWheelTurns * 3 Step -1
                                    lRetVal = SendMessage(hwnd, WM_HSCROLL, ByVal SB_LINERIGHT, ByVal 0)
                              Next iTurn
                              Exit Function
                        ElseIf lWheelTurns = 0 Then
                              MsgBox "ERROR: wheel turn = 0.  How can you turn the wheel zero turns?"
                        End If
                  End If
      End Select
                  
      ListViewProc = CallWindowProc(gpOldLvwProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function TrackMouseWheel(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

      ' This is for picture manipulation; applies to picEditor.
      
      If uMsg = WM_MOUSEWHEEL Then
            Dim poiWheel As POINTAPI
            Dim iVirtKeys As Integer, iWheelTurn As Integer
            
            poiWheel = MAKEPOINT(wParam)
            iWheelTurn = poiWheel.Y / WHEEL_DELTA
            iVirtKeys = poiWheel.X
            poiWheel = MAKEPOINT(lParam)
            'Debug.Print iWheelTurn & "   " & iVirtKeys & "   " & poiWheel.x & "   " & poiWheel.y & "   " & Hex(wParam)
            frmMain.WheelInput iWheelTurn, iVirtKeys, poiWheel.X, poiWheel.Y
      End If
      TrackMouseWheel = CallWindowProc(gpOldpicEditorProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function TrackMouseWheelFullScreen(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

      ' This is for picture manipulation; applies to frmFullScreen.
      
      If uMsg = WM_MOUSEWHEEL Then
            Dim poiWheel As POINTAPI
            Dim iVirtKeys As Integer, iWheelTurn As Integer
            
            poiWheel = MAKEPOINT(wParam)
            iWheelTurn = poiWheel.Y / WHEEL_DELTA
            iVirtKeys = poiWheel.X
            poiWheel = MAKEPOINT(lParam)
            'Debug.Print iWheelTurn & "   " & iVirtKeys & "   " & poiWheel.x & "   " & poiWheel.y & "   " & Hex(wParam)
            frmMain.WheelInput iWheelTurn, iVirtKeys, poiWheel.X, poiWheel.Y
      End If
      TrackMouseWheelFullScreen = CallWindowProc(gpOldfrmFullScreenProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function TrackMouseLeave(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
            
      Dim objPFB As Object
      
      If uMsg = WM_MOUSELEAVE Then
            Beep
      End If
      TrackMouseLeave = CallWindowProc(gpOldpicBrowserProc, hwnd, uMsg, wParam, lParam)
End Function

'Public Function CompareLong(ByVal lParam1 As Long, ByVal lParam2 As Long, _
'      ByVal lParamSort As Long) As Integer
'      ' This is a callback function to be sent with an LVM_SORTITEMSEX message.
'      ' Listviews like to do text sorting.  That's all they do in visual basic, without help.
'      ' I remember seeing applications in which the programmer obviously wasn't prepared
'      ' for this annoying lack of functionality.  (Napster, anyone?)
'
'      ' lParam is going to mirror lvwBrowser.SortOrder, which is
'      ' 0 for lvwAscending
'      ' 1 for lvwDescending
'
'      If lParam1 < lParam2 Then
'            CompareLong = -1
'      ElseIf lParam1 = lParam2 Then
'            CompareLong = 0
'      Else
'            CompareLong = 1
'      End If
'
'      If lParamSort = lvwDescending Then CompareLong = -CompareLong
'End Function



' *************************************************************
' Useful functions and subs
' *************************************************************

Public Function SnipPath(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipPath = Right(sPath, Len(sPath) - iSlash)
End Function

Public Function SnipFileName(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipFileName = Left(sPath, iSlash)
End Function

Public Function CstringToVBstring(ByVal sCstring As String) As String
      ' Removes first null character and anything following it.
      Dim lngNullPosition As Long
      
      lngNullPosition = InStr(1, sCstring, Chr(0))
      If lngNullPosition = 0 Then
            CstringToVBstring = sCstring
      Else
            CstringToVBstring = Left(sCstring, lngNullPosition - 1)
      End If
End Function

Public Function VBstringToCstring(ByVal sVBstring As String) As Byte()
      ' Just, never you mind this.  Not the inverse of above.
      sVBstring = sVBstring & Chr(0)
      VBstringToCstring = sVBstring
      'VBstringToCstring = StrConv(sVBstring, vbFromUnicode)
End Function

Public Function TrimTrailingSlash(ByVal sPath As String) As String
      If Right(sPath, 1) = "\" Then
            TrimTrailingSlash = Left(sPath, Len(sPath) - 1)
      Else
            TrimTrailingSlash = sPath
      End If
End Function

Public Function IsPathFull(ByVal sInput As String) As Long
      ' returns 0 if not a full path
      ' if a full path, returns position of colon
      ' does NOT check if the path is a VALID path
      IsPathFull = InStrRev(sInput, ":")
End Function

Public Function GetFullPathName(ByVal sInput As String) As String
      ' Just for the record:
      '     Dir believes that path names always end with a "\"
      '     CurDir believes that path names *never* end with a "\"
      
      Dim iColonPosition As Integer
      
      iColonPosition = InStrRev(sInput, ":")
      If iColonPosition = 0 Then
            GetFullPathName = CurDir & "\" & sInput
      Else
            GetFullPathName = sInput
      End If
      
End Function

'Public Function FileExistsVBS(ByVal sPath As String) As Boolean
'      Dim objFS As Object
'
'      Set objFS = CreateObject("Scripting.FileSystemObject")
'      FileExists = objFS.FileExists(sPath)
'End Function


Public Function IsKeyDown(ByVal lVirtKey As Long) As Boolean
      IsKeyDown = GetKeyState(lVirtKey) And &HF0000000
End Function

Public Function FileExists(ByVal sSource As String) As Boolean

      Dim WFD As WIN32_FIND_DATA
      Dim hFile As Long
      
      hFile = FindFirstFile(sSource, WFD)
      FileExists = hFile <> -1 ' invalid handle value
      
      FindClose (hFile)
   
End Function

Public Function FileSize(ByVal sSource As String) As Currency
      
      Dim WFD As WIN32_FIND_DATA
      Dim hFile As Long
      
      hFile = FindFirstFile(sSource, WFD)
      
      If hFile > 0 And WFD.nFileSizeHigh = 0 Then
            FileSize = WFD.nFileSizeLow
      ElseIf hFile > 0 And WFD.nFileSizeHigh > 0 Then
            FileSize = -1 ' TODO: account for the high order word which should be multiplied by the maximum long integer
      Else
            FileSize = -1 ' invalid handle value
      End If
      
      FindClose (hFile)
      
End Function

Public Function FileModifiedTime(ByVal sSource As String) As String
      ' example date string:   2005-03-15 6:14:21
      
      Dim WFD As WIN32_FIND_DATA
      Dim localTime As FILETIME
      Dim sysTime As SYSTEMTIME
      Dim hFile As Long
      
      hFile = FindFirstFile(sSource, WFD)
      
      If hFile > 0 Then
            FileModifiedTime = FormatNonLocalFileTime(WFD.ftLastWriteTime)
      Else
            FileModifiedTime = ""
      End If
      
      FindClose (hFile)
End Function


Public Function RecycleFile(ByVal sPath As String) As Integer
      ' Send a file to the Recycle Bin.
      
      Dim shfFileOperation As SHFILEOPSTRUCT
      
      With shfFileOperation
            .wFunc = FO_DELETE
            .pFrom = sPath
            .fFlags = FOF_ALLOWUNDO
      End With
      RecycleFile = SHFileOperation(shfFileOperation)
End Function

' WARNING: these do not account for frame/picturebox borders.  Use informally.

Public Function AbsoluteLeft(ByRef ctrl As Control) As Long
      Dim iMargins As Long
      
      On Error Resume Next
      AbsoluteLeft = ctrl.Left + AbsoluteLeft(ctrl.Container)
      If Err > 0 Then AbsoluteLeft = ctrl.Left
      On Error GoTo 0
End Function

Public Function AbsoluteTop(ByRef ctrl As Control) As Long
      Dim iMargins As Long
      
      On Error Resume Next
      AbsoluteTop = ctrl.Top + AbsoluteTop(ctrl.Container)
      If Err > 0 Then AbsoluteTop = ctrl.Top
      On Error GoTo 0
End Function

Public Function AbsoluteRight(ByRef ctrl As Control) As Long
      AbsoluteRight = AbsoluteLeft(ctrl) + ctrl.Width
End Function

Public Function AbsoluteBottom(ByRef ctrl As Control) As Long
      AbsoluteBottom = AbsoluteTop(ctrl) + ctrl.Height
End Function

' I can't believe it's this difficult to extract the character count, when counting only one character per
' carriage return.  Cannot fucking believe it.

Public Function CharacterCount(ByRef editor As agRichEdit) As Long
      
      Dim lLastLineLength As Long, lLastLineIndex As Long
      
      lLastLineIndex = SendMessage(editor.RichEdithWnd, EM_LINEINDEX, ByVal (editor.LineCount - 1), ByVal 0)
      
      lLastLineLength = SendMessage(editor.RichEdithWnd, EM_LINELENGTH, ByVal lLastLineIndex + 1, ByVal 0)

      CharacterCount = lLastLineIndex + lLastLineLength
End Function

Public Function GetRealFontName(ByRef editor As agRichEdit) As String
      Const CFM_FACE As Long = &H20000000
      Const CFM_SIZE As Long = &H80000000
      
      Dim char2 As CHARFORMAT2
      Dim lRetVal As Long
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_FACE
      char2.dwEffects = 0
      lRetVal = SendMessage(editor.RichEdithWnd, EM_GETCHARFORMAT, ByVal 0, char2)
      
      ' Make it 16-bit characters, and trim the fat.
      GetRealFontName = CstringToVBstring(StrConv(char2.szFaceName, vbUnicode))
End Function

Public Function GetRealFontSize(ByRef editor As agRichEdit) As Single
      Const CFM_FACE As Long = &H20000000
      Const CFM_SIZE As Long = &H80000000
      
      Dim char2 As CHARFORMAT2
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_SIZE '+ CFM_FACE
      char2.dwEffects = 0
      SendMessage editor.RichEdithWnd, EM_GETCHARFORMAT, ByVal SCF_SELECTION, char2
      GetRealFontSize = char2.yHeight / 20! ' convert twips to printer's points
End Function

Public Function SetRealFontSize(ByRef editor As agRichEdit, ByVal sNewSize As Single) As Single
      Const CFM_FACE As Long = &H20000000
      Const CFM_SIZE As Long = &H80000000
      
      Dim char2 As CHARFORMAT2
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_SIZE
      char2.dwEffects = 0
      char2.yHeight = sNewSize * 20
      SendMessage editor.RichEdithWnd, EM_SETCHARFORMAT, ByVal SCF_ALL, char2
      SetRealFontSize = char2.yHeight / 20!  ' return this, see if it's moved or something weird like that.
End Function

Public Function SetRealStdFont(ByRef editor As agRichEdit, ByRef fnt As StdFont, _
      Optional lTextColor As Long = vbWindowText) As Long

      Const CFE_BOLD As Long = &H1
      Const CFE_ITALIC As Long = &H2
      Const CFE_STRIKEOUT As Long = &H8
      Const CFE_UNDERLINE As Long = &H4
      
      Const CFM_FACE As Long = &H20000000
      Const CFM_SIZE As Long = &H80000000
      Const CFM_CHARSET As Long = &H8000000
      Const CFM_BOLD As Long = &H1
      Const CFM_COLOR As Long = &H40000000
      Const CFM_ITALIC As Long = &H2
      Const CFM_LINK As Long = &H20
      Const CFM_OFFSET As Long = &H10000000
      Const CFM_STRIKEOUT As Long = &H8
      Const CFM_UNDERLINE As Long = &H4
      Const CFM_WEIGHT As Long = &H400000
      
      Dim char2 As CHARFORMAT2
      Dim sFontName As String
      Dim bDyn() As Byte, i As Integer
      
      With char2
            .cbSize = LenB(char2)
            ' Tell it which CHARFORMAT2 properties carry relevant data:
            .dwMask = CFM_SIZE + CFM_FACE + CFM_CHARSET + CFM_BOLD + _
                  CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE + CFM_WEIGHT
            
            .dwMask = .dwMask + CFM_COLOR
            .crTextColor = TranslateColor(lTextColor)
            
            If fnt.Bold Then .dwEffects = .dwEffects + CFE_BOLD
            If fnt.Italic Then .dwEffects = .dwEffects + CFE_ITALIC
            If fnt.Underline Then .dwEffects = .dwEffects + CFE_UNDERLINE
            If fnt.Strikethrough Then .dwEffects = .dwEffects + CFE_STRIKEOUT
            .yHeight = fnt.Size * 20
            .bCharSet = fnt.Charset
            .wWeight = fnt.Weight
            ' the font name takes some string manipulation...
            bDyn = StrConv(fnt.Name & Chr(0), vbFromUnicode)
            For i = LBound(bDyn) To UBound(bDyn)
                  .szFaceName(i) = bDyn(i)
            Next i
      End With
      SetRealStdFont = SendMessage(editor.RichEdithWnd, EM_SETCHARFORMAT, ByVal SCF_ALL, char2)
End Function


' DON'T NEED, COMMENTING OUT.  WORKS BUT NOT FOR POPUP MENUS.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
'' Some pages may also contain other copyrights by the author.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Distribution: You can freely use this code in your own
''               applications, but you may not reproduce
''               or publish this code on any web site,
''               online service, or distribute as source
''               on any media without express permission.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Public Sub SetRadioMenuChecksB(ByRef frm As Form, ByVal mnuBarIndex As Long, ByVal mnuItem As Long)
'
'      Dim hMenu As Long
'      Dim hSubMenu As Long
'      Dim mInfo As MENUITEMINFO
'
'      'get the menu handle
'      hMenu = GetMenu(frm.hWnd)
'
'      'get the submenu handle
'      hSubMenu = GetSubMenu(hMenu, mnuBarIndex)
'
'      'fill a structure to retrieve the current
'      'item menu string by first calling
'      'GetMenuItemInfo passing a null string.
'      'The required size is returned in
'      'mInfo.cch. Add 1 to accommodate the
'      'null that will be added when called.
'      With mInfo
'            .cbSize = Len(mInfo)
'            .fMask = MIIM_TYPE
'            .fType = MFT_STRING
'            .dwTypeData = vbNullString
'            .cch = Len(mInfo.dwTypeData)
'
'            'get the needed buffer size
'            Call GetMenuItemInfo(hSubMenu, mnuItem, MENU_IDENTIFIER, mInfo)
'
'            'set the buffer
'            .dwTypeData = Space$(mInfo.cch + 1)
'            .cch = Len(mInfo.dwTypeData)
'
'      End With
'
'      'and get the data
'      If GetMenuItemInfo(hSubMenu, mnuItem, MENU_IDENTIFIER, mInfo) <> 0 Then
'
'            'copy its attributes, changing
'            'the checkmark to a radio button
'            With mInfo
'                  .cbSize = Len(mInfo)
'                  .fType = MFT_RADIOCHECK
'                  .fMask = MIIM_TYPE
'            End With
'
'            'modify the menu item
'            Call SetMenuItemInfo(hSubMenu, mnuItem, MENU_IDENTIFIER, mInfo)
'
'      End If
'
'End Sub





Public Function GetMP3Info(ByVal sFileName As String, mp3info As MP3TagInfo) As String()
      ' Retrieve the informations contained into the standard ID3 tag
      ' of the specified MP3 file
      ' Return an array of 6 elements with the following meaning:
      '   - index 0: song title
      '   - index 1: artist
      '   - index 2: album
      '   - index 3: year
      '   - index 4: comment
      '   - index 5: genre: this is an integer value --> use any MP3 player,
      '  such as Winamp,
      '       to look for the descriptions

'      Dim infoRet(5) As String
'      Dim mp3info As MP3TagInfo
      
      On Error Resume Next
      ' open the specified file
      Open sFileName For Binary As #1
      ' fill the strct's fileds
      
      With mp3info
            Get #1, FileLen(sFileName) - 127, .tag
            If Not .tag = "TAG" Then
                  Close #1
                  Exit Function
            End If
            Get #1, , .title
            Get #1, , .artist
            Get #1, , .album
            Get #1, , .year
            Get #1, , .comment
            Get #1, , .genre
            Close #1
      End With
      
      ' I'm altering what I copied.  This will pass on the structre instead of returning a string array.

'      ' from struct to array
'      infoRet(0) = Trim$(mp3info.title)
'      infoRet(1) = Trim$(mp3info.artist)
'      infoRet(2) = Trim$(mp3info.album)
'      infoRet(3) = Trim$(mp3info.year)
'      infoRet(4) = Trim$(mp3info.comment)
'      infoRet(5) = CInt(Asc(Trim$(mp3info.genre))) - 1
'      'return the array
'      GetMP3Info = infoRet
End Function


