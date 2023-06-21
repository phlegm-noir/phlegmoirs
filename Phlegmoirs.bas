Attribute VB_Name = "modPhlegmoirs"


' *************************************************************
' My custom types
' *************************************************************

Option Compare Binary
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
      ReadOnly As Integer
      SelStart As Long
      SelEnd As Long
      FirstVisibleLine As Long
      FontName As String * 255
      FontSize As Currency
      FontBold As Boolean
      FontItalic As Boolean
      FontUnderline As Boolean
      FontStrikethrough As Boolean
      TextColor As Long
      ScrollPos As POINTAPI
      
      ' Reserved, not used:
      BackColor As Long
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
      ShowFind As Boolean
      BookmarkCount As Integer
      HistoryCount As Integer
      AutoLoadFile As String * 255
      cboPath As String * 255
      FocusFollowsMouse As Boolean
      
      ' Reserved, not used:
      FullScreen As Boolean
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
      ERROR As Boolean
      ListEmpty As Boolean
      
      ItemClicked As Boolean
      ButtonPressed As Boolean
      MouseButton As Integer
      Shift As Integer
      RecentPath As Integer
End Type

Public Type TImageData
      OutPic As Object
      DefaultHeight As Long
      DefaultWidth As Long
      PrevX As Single
      PrevY As Single
      
      Dragging As Boolean
      Moved As Boolean
      Zoomed As Boolean
End Type

' *************************************************************
' Global enumerations
' *************************************************************

Enum eIconType
      Directory = 1
      Properties = 2
      Drive = 3
      Text = 4
      other = 5
      Picture = 6
      IconError = 7
      Bookmark = 8
      Floppy = 9
      Network = 10
      Cdrom = 11
      rtf = 12
      mp3 = 13
      video = 14
End Enum

Enum eViewMode
      ERROR = -2
      TextView = 1
      PictureView = 2
      PropertiesView = 3
End Enum

Enum eStat
      BrowserStats = 1
      Stats = 2
      encoding = 3
      Modified = 4
      SelText = 5
      Tips = 6
End Enum

Enum eDirection
      Forward = 1
      back = -1
End Enum

Enum eQuery
      Find = 0
End Enum

Enum eTextEncoding
      ASCII = 0
      UNICODE = -1
      ERROR = -2
End Enum

Enum eIoMode
      ForReading = 1
      ForWriting = 2
      ForAppending = 8
End Enum

Enum eCreate
      Yes = True
      No = False
End Enum

Enum eOverwrite
      Yes = True
      No = False
End Enum

      
' *************************************************************
' Global Variables
' *************************************************************

Public gpOldLvwProc As Long, gpOldpicBrowserProc As Long, gpOldpicEditorProc As Long
Public gpOldfrmFullScreenProc As Long
Public objtest As Object
Public gFSO As Object
Public gBrowserData As TBrowserData
Public gImageData As TImageData
Public giEditorMode As eViewMode
Public gTextEncoding As Integer
Public gfFullScreenMode As Integer

Public Const MoveIncrement = -512
Public Const glMAX_LONG_INTEGER = &H7FFFFFFF   '   2147483647

Function FormatBytes(ByVal curBytes, iPrecision As Integer) As String
      ' This function takes a quantity of bytes as a currency value ('cause it's 64-bit),
      ' formats it to read likeathis:
      
      ' 45.2 MB
      ' 300.2 KB
      ' 666 Bytes
      ' 99.4444 GB
      ' 20000 TB  (not supporting terabytes at the moment, sorry folks)
      
      ' ...with iPrecision digits after the demical.
      
      If curBytes < 1024@ Then
            FormatBytes = CStr(curBytes) & " bytes"
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

Public Function ListViewProc(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

      Select Case uMsg
            Case WM_KEYDOWN
                  ' The problem: right arrow key wants to scroll right.  I want it to do things
                  ' like enter the given folder, which would cause the listview to do both at once.
                  
                  ' Solution: this modified windows procedure will translate a right arrow (without ctrl)
                  ' into an F13 keyDown and then continue the window procedure.
                  
                  ' Elsewhere, F13 will be bound to do the desired right arrow key things.
                  
                  ' If somebody wants to scroll right with arrow keys, they can use ctrl+right.
                  
                  If wParam = vbKeyRight And Not IsKeyDown(VK_CONTROL) Then
                        
                        ListViewProc = CallWindowProc(gpOldLvwProc, hwnd, _
                              WM_KEYDOWN, vbKeyF13, lParam)
                        Exit Function
                  End If
            
            Case WM_MOUSEWHEEL
                  ' This procedure is going to scroll horizontally when you mousewheel
                  ' over the hscrollbar. Inspired by Opera browser at the time.

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
                  Dim lWheelTurns As Long  ' Can be positive or negative (not zero)
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
      
      ' Do all the default things too, as defined by the old procedure
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

Public Function GetIconType(sEx As String) As eIconType
      ' This function takes an extension (DO NOT INCLUDE DOT) and returns a mode
      
      Select Case sEx
            Case "bmp", "gif", "jpg", "jpeg", "ico", "cur", "png", "webp"
                  GetIconType = eIconType.Picture
            
            Case "dll", "ocx", "exe", "zip", "msi", "sys", "cab", "7z"
                  GetIconType = eIconType.Properties
            
            Case "mp3", "ogg", "wav", "flac"
                  GetIconType = eIconType.mp3
            
            Case "avi", "mpeg", "mp4", "webm", "flv"
                  GetIconType = eIconType.video
            
            Case "rtf"
                  GetIconType = eIconType.rtf
            
            Case "txt", "log"
                  GetIconType = eIconType.Text
            
            Case Else
                  GetIconType = eIconType.other
      End Select
End Function

Public Function GetViewMode(ByVal sFileName As String, ByVal iMode As eIconType) As eViewMode
      Const PicFileTooBig = 10000000
      Const NonPicFileTooBig = 2097152
      Dim sEx As String
      Dim fileSize
      
      If Not gFSO.FileExists(sFileName) Then
            GetViewMode = eViewMode.ERROR
            Exit Function
      End If
      
      sEx = gFSO.getextensionname(sFileName)
      
      If iMode = eIconType.Bookmark Then
            iMode = GetIconType(sEx)
      End If
      
      Select Case iMode
            Case eIconType.Picture
                  fileSize = GetFileSize(sFileName)
                  
                  If sEx = "png" Or sEx = "webp" Or fileSize > PicFileTooBig Then
                        GetViewMode = eViewMode.PropertiesView
                  Else
                        GetViewMode = eViewMode.PictureView
                  End If
            
            Case eIconType.Text, eIconType.rtf, eIconType.other
                  fileSize = GetFileSize(sFileName)
                  
                  If fileSize > NonPicFileTooBig Then
                        GetViewMode = eViewMode.PropertiesView
                  Else
                        GetViewMode = eViewMode.TextView
                  End If
            
            Case eIconType.Cdrom, eIconType.Directory, eIconType.Drive, eIconType.IconError, eIconType.Floppy, eIconType.Network
                  GetViewMode = eViewMode.ERROR
            
            Case Else
                  GetViewMode = eViewMode.PropertiesView
      End Select
End Function
Public Function IsUnicodeFile(FilePath)
      Dim objStream
      Dim intAsc1Chr, intAsc2Chr
      
      If Not gFSO.FileExists(FilePath) Then
            IsUnicodeFile = eTextEncoding.ERROR
            Exit Function
      ElseIf GetFileSize(FilePath) = 1 Then
            IsUnicodeFile = eTextEncoding.ASCII
            Exit Function
      End If
      
      On Error Resume Next
      Set objStream = gFSO.OpenTextFile(FilePath, eIoMode.ForReading, eCreate.No, eTextEncoding.ASCII)
      intAsc1Chr = AscW(objStream.Read(1))
      intAsc2Chr = AscW(objStream.Read(1))
      objStream.Close
      If Err > 0 Then
            IsUnicodeFile = eTextEncoding.ERROR
            Exit Function
      End If
      On Error GoTo 0
      
      If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
          IsUnicodeFile = eTextEncoding.UNICODE
      Else
          IsUnicodeFile = eTextEncoding.ASCII
      End If
      
      Set objStream = Nothing
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

Public Function GetFileSize(ByVal sSource As String) As Currency
      If sSource = "" Then
            GetFileSize = 0
      Else
            GetFileSize = gFSO.GetFile(sSource).Size
      End If
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

' Extract the character count, when counting only one character per carriage return.

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


Public Function GetMP3Info(ByVal sFileName As String, mp3info As MP3TagInfo) As String()
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
End Function

Public Function getAllProperties(sFileName As String)
      Dim sBaseName, sPathName
      sPathName = gFSO.getParentFolderName(sFileName)
      sBaseName = gFSO.GetFileName(sFileName)
      
      Dim objShell, objFolder
      Set objShell = CreateObject("shell.application")
      Set objFolder = objShell.NameSpace(sPathName)
      
      If (Not objFolder Is Nothing) Then
            Dim objFolderItem
            Set objFolderItem = objFolder.ParseName(sBaseName)
            
            If (Not objFolderItem Is Nothing) Then
                  Dim uninteresting
                  uninteresting = Array("Attributes", "Owner", "Total size", "Computer", "File extension", _
                  "Filename", "Space free", "Shared", "Folder name", "File location", "Folder", "Path", "Type", _
                  "Link status", "Space used", "Sharing status", "UNKNOWN(296)", "Content", "Rating", "Shared with", _
                  "Protected")
                  
                  Dim i
                  For i = 0 To 320
                        Dim columnName, value
                        columnName = objFolder.GetDetailsOf(objFolder.Items, i)
                        value = objFolder.GetDetailsOf(objFolderItem, i)
                        If columnName = "" Then columnName = "UNKNOWN(" + Trim(Str(i)) + ")"
                        
                        If value <> "" And Not ArrContains(uninteresting, columnName) Then
                              Debug.Print Str(i) + ". " + columnName + ": " + vbTab + value
                        End If
                  Next i
            End If
            
            Set objFolderItem = Nothing
      End If
      
      Set objFolder = Nothing
      Set objShell = Nothing
End Function

Public Function ArrContains(arrString, ByVal PassedValue As String) As Boolean
   Dim Index As Integer
    For Index = LBound(arrString) To UBound(arrString)
      If arrString(Index) = PassedValue Then
        ArrContains = True
        Exit Function
      End If
    Next
    ArrContains = False
End Function
