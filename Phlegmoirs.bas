Attribute VB_Name = "modPhlegmoirs"
Option Compare Binary
Option Explicit
      
' *************************************************************
' Global Settings
' *************************************************************

Public Const REGISTRY_VERSION = "0.19.0" ' Not the current build number, but the last time I changed the registry structure.
Public Const LOG_TO_FILE As Boolean = False
Public Const MINIMUM_LOG_LEVEL As Integer = 1
Public Const DEBUGGING As Boolean = False
Public Const MAX_HISTORY As Integer = 10
Public Const MAX_BOOKMARKS As Integer = 30
Public Const FOCUS_FOLLOWS_MOUSE As Boolean = False
Public Const AUTOSIZE_COLUMNS As Boolean = True ' Will autosize type, size, and modified while retaining filename width
Public Const COLUMN_TOO_SMALL = 600 ' We won't keep the auto-size if it's beneath a certain point

' Got registry slots for their order just in case
' This is not fully implemented, especially any possible conflicts with AUTOSIZE_COLUMNS
Public Const ALLOW_REARRANGE_COLUMNS As Boolean = False

' *************************************************************
' My custom types
' *************************************************************

Public Type TStatType
      i As Long
      imax As Long
      Y As Long
      ymax As Long
      X As Long
      xmax As Long
End Type

' These prefs objects are for the registry.
' Not for use during the session.
Public Type TEditorPrefs
      AutoLoadFile As String * 255
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
End Type

Public Type TBrowserPrefs
      AutoLoadPath As String * 255
      SortMethod As Integer
      SortKey As Integer
      NameColumnWidth As Long
      TypeColumnWidth As Long
      SizeColumnWidth As Long
      ModifiedColumnWidth As Long
      NameColumnIndex As Integer
      TypeColumnIndex As Integer
      SizeColumnIndex As Integer
      ModifiedColumnIndex As Integer
End Type

Public Type TWindowPrefs
      WNP As WINDOWPLACEMENT
      BrowserWidth As Long
      ShowFileBrowser As Boolean
      ShowStatusBar As Boolean
      ShowToolBar As Boolean
      ShowFind As Boolean
      ImageZoom As Integer
End Type

Public Type TAllPrefs
      WindowPrefs As TWindowPrefs
      BrowserPrefs As TBrowserPrefs
      EditorPrefs As TEditorPrefs
      HistoryCount As Integer
      History(MAX_HISTORY) As String * 255
      PathHistoryCount As Integer
      PathHistory(MAX_HISTORY) As String * 255
      BookmarkCount As Integer
      Bookmarks(MAX_BOOKMARKS) As String * 255
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

Public gStats As TStatType
Public gsPhlegmKey As String
Public gsPhlegmDate As String

Public gpOldLvwProc As Long, gpOldpicBrowserProc As Long, gpOldpicEditorProc As Long
Public gpOldfrmFullScreenProc As Long
Public objtest As Object
Public gFSO As Object
Public gBrowserData As TBrowserData
Public gImageData As TImageData
Public giEditorMode As eViewMode
Public gTextEncoding As Integer
Public gfFullScreenMode As Integer
Public gCommandFile As String

Public Const MoveIncrement = -512
Public Const glMAX_LONG_INTEGER = &H7FFFFFFF   '   2147483647

Public Function AbsoluteBottom(ByRef ctrl As Control) As Long
Attribute AbsoluteBottom.VB_UserMemId = 1610612760
      AbsoluteBottom = AbsoluteTop(ctrl) + ctrl.Height
End Function

' WARNING: these do not account for frame/picturebox borders.  Use informally.

Public Function AbsoluteLeft(ByRef ctrl As Control) As Long
Attribute AbsoluteLeft.VB_UserMemId = 1610612757
      Dim iMargins As Long
      
      On Error Resume Next
      AbsoluteLeft = ctrl.Left + AbsoluteLeft(ctrl.Container)
      If Err > 0 Then AbsoluteLeft = ctrl.Left
      DebugLog "AbsoluteLeft throws an error, whatever that means..."
      On Error GoTo 0
End Function

Public Function AbsoluteRight(ByRef ctrl As Control) As Long
Attribute AbsoluteRight.VB_UserMemId = 1610612759
      AbsoluteRight = AbsoluteLeft(ctrl) + ctrl.Width
End Function

Public Function AbsoluteTop(ByRef ctrl As Control) As Long
Attribute AbsoluteTop.VB_UserMemId = 1610612758
      Dim iMargins As Long
      
      On Error Resume Next
      AbsoluteTop = ctrl.Top + AbsoluteTop(ctrl.Container)
      If Err > 0 Then AbsoluteTop = ctrl.Top
      DebugLog "AbsoluteTop throws an error, whatever that means..."
      On Error GoTo 0
End Function

Public Function ArrContains(arrString, ByVal PassedValue As String) As Boolean
Attribute ArrContains.VB_UserMemId = 1610612764
   Dim Index As Integer
    For Index = LBound(arrString) To UBound(arrString)
      If arrString(Index) = PassedValue Then
        ArrContains = True
        Exit Function
      End If
    Next
    ArrContains = False
End Function

' Extract the character count, when counting only one character per carriage return.

Public Function CharacterCount(ByRef editor As agRichEdit) As Long
Attribute CharacterCount.VB_UserMemId = 1610612761
      
      Dim lLastLineLength As Long, lLastLineIndex As Long
      
      lLastLineIndex = SendMessage(editor.RichEdithWnd, EM_LINEINDEX, ByVal (editor.LineCount - 1), ByVal 0)
      
      lLastLineLength = SendMessage(editor.RichEdithWnd, EM_LINELENGTH, ByVal lLastLineIndex + 1, ByVal 0)

      CharacterCount = lLastLineIndex + lLastLineLength
End Function

Public Function CstringToVBstring(ByVal sCstring As String) As String
Attribute CstringToVBstring.VB_UserMemId = 1610612743
      ' Removes first null character and anything following it.
      On Error GoTo CONVERSION_ERROR
      Dim lngNullPosition As Long
      
      lngNullPosition = InStr(1, sCstring, Chr(0))
      If lngNullPosition = 0 Then
            CstringToVBstring = sCstring
      Else
            CstringToVBstring = Left(sCstring, lngNullPosition - 1)
      End If
      Exit Function
CONVERSION_ERROR:
      DebugLog "CONVERSION ERROR: " & sCstring, 2
End Function

Public Sub DebugLog(ByVal sMsg As String, Optional ByVal iLogLevel As Integer = 1)
Attribute DebugLog.VB_UserMemId = 1610612745
      Debug.Print sMsg
      If LOG_TO_FILE And iLogLevel >= MINIMUM_LOG_LEVEL Then
            Dim iFile As Integer
            Dim sFile As String
            sFile = "phlegmoirs_err.log"
            iFile = FreeFile
            Open sFile For Append As #iFile
                  Print #iFile, Now & ": " & sMsg
            Close #iFile
      End If
End Sub

Public Function FileExists(ByVal sSource As String) As Boolean
Attribute FileExists.VB_UserMemId = 1610612753

      Dim WFD As WIN32_FIND_DATA
      Dim hFile As Long
      
      hFile = FindFirstFile(sSource, WFD)
      FileExists = hFile <> -1 ' invalid handle value
      
      FindClose (hFile)
   
End Function

Public Function FileModifiedTime(ByVal sSource As String) As String
Attribute FileModifiedTime.VB_UserMemId = 1610612755
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

Function FormatBytes(ByVal curBytes, iPrecision As Integer) As String
      ' This function takes a quantity of bytes as a currency value ('cause it's 64-bit),
      ' formats it to read likeathis:
      
      ' 45.2 MB
      ' 300.2 KB
      ' 666 Bytes
      ' 99.4444 GB
      ' 20000 TB  (not supporting terabytes at the moment, sorry folks)
      
      ' ...with iPrecision digits after the demical.
      
      If curBytes = 1 Then
            FormatBytes = CStr(curBytes) & " byte"
      ElseIf curBytes < 1024@ Then
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

Public Function getAllProperties(sFileName As String)
Attribute getAllProperties.VB_UserMemId = 1610612763
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

Public Function GetFileSize(ByVal sSource As String) As Currency
Attribute GetFileSize.VB_UserMemId = 1610612754
      If sSource = "" Then
            GetFileSize = 0
      Else
            GetFileSize = gFSO.getfile(sSource).Size
      End If
End Function

Public Function GetFullPathName(ByVal sInput As String) As String
Attribute GetFullPathName.VB_UserMemId = 1610612751
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

Public Function GetIconType(sEx As String) As eIconType
Attribute GetIconType.VB_UserMemId = 1610612746
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

Public Function GetMP3Info(ByVal sFileName As String, mp3info As MP3TagInfo) As String()
Attribute GetMP3Info.VB_UserMemId = 1610612762
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

' Figure out the numbering of a menu caption, and which digit to underline.
Public Function GetNumberedCaption(ByVal sFileName As String, ByVal iIndex As Integer) As String
Attribute GetNumberedCaption.VB_UserMemId = 1610612765
      If iIndex < 10 Then
            GetNumberedCaption = "&" & iIndex & "   " & sFileName
      ElseIf iIndex = 10 Then
            GetNumberedCaption = "1&0   " & sFileName
      Else
            GetNumberedCaption = iIndex & "   " & sFileName
      End If
End Function

Public Function GetViewMode(ByVal sFileName As String, ByVal iMode As eIconType) As eViewMode
Attribute GetViewMode.VB_UserMemId = 1610612747
      Const PicFileTooBig = 10000000
      Const NonPicFileTooBig = 2097152
      Dim sEx As String
      Dim fileSize
      
      If Not gFSO.FileExists(sFileName) Then
            GetViewMode = eViewMode.ERROR
            DebugLog "ViewMode: ERROR for file: " & sFileName & "..."
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
                  DebugLog "ViewMode: ERROR for file: " & sFileName
            
            Case Else
                  GetViewMode = eViewMode.PropertiesView
      End Select
End Function

Public Function IsKeyDown(ByVal lVirtKey As Long) As Boolean
Attribute IsKeyDown.VB_UserMemId = 1610612752
      IsKeyDown = GetKeyState(lVirtKey) And &HF0000000
End Function

Public Function IsPathFull(ByVal sInput As String) As Long
Attribute IsPathFull.VB_UserMemId = 1610612750
      ' returns 0 if not a full path
      ' if a full path, returns position of colon
      ' does NOT check if the path is a VALID path
      IsPathFull = InStrRev(sInput, ":")
End Function

Public Function IsUnicodeFile(FilePath)
Attribute IsUnicodeFile.VB_UserMemId = 1610612748
      Dim objStream
      Dim intAsc1Chr, intAsc2Chr
      
      If Not gFSO.FileExists(FilePath) Then
            IsUnicodeFile = eTextEncoding.ERROR
            DebugLog "Text encoding=ERROR for file: " & FilePath
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
            DebugLog "Text encoding=ERROR for file: " & FilePath
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

                  With HitTestInfo
                        .pt.X = MAKEPOINT(lParam).X - recClient.Left
                        .pt.Y = MAKEPOINT(lParam).Y - recClient.Top
                  End With
                  lRetVal = SendMessage(hwnd, LVM_HITTEST, ByVal 0, HitTestInfo)

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
                              DebugLog "ERROR: wheel turn = 0.  How can you turn the wheel zero turns?"
                        End If
                  End If
      End Select
      
      ' Do all the default things too, as defined by the old procedure
      ListViewProc = CallWindowProc(gpOldLvwProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function RecycleFile(ByVal sPath As String) As Integer
Attribute RecycleFile.VB_UserMemId = 1610612756
      ' Send a file to the Recycle Bin.
      
      Dim shfFileOperation As SHFILEOPSTRUCT
      
      With shfFileOperation
            .wFunc = FO_DELETE
            .pFrom = sPath
            .fFlags = FOF_ALLOWUNDO
      End With
      RecycleFile = SHFileOperation(shfFileOperation)
End Function

Public Function SnipFileName(ByVal sPath As String) As String
Attribute SnipFileName.VB_UserMemId = 1610612742
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipFileName = Left(sPath, iSlash)
End Function

' *************************************************************
' Useful functions and subs
' *************************************************************

Public Function SnipPath(ByVal sPath As String) As String
Attribute SnipPath.VB_UserMemId = 1610612741
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipPath = Right(sPath, Len(sPath) - iSlash)
End Function

Public Function TrackMouseLeave(ByVal hwnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute TrackMouseLeave.VB_UserMemId = 1610612740
            
      Dim objPFB As Object
      
      If uMsg = WM_MOUSELEAVE Then
            Beep
      End If
      TrackMouseLeave = CallWindowProc(gpOldpicBrowserProc, hwnd, uMsg, wParam, lParam)
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
            frmMain.WheelInput iWheelTurn, iVirtKeys, poiWheel.X, poiWheel.Y
      End If
      TrackMouseWheel = CallWindowProc(gpOldpicEditorProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function TrimTrailingSlash(ByVal sPath As String) As String
Attribute TrimTrailingSlash.VB_UserMemId = 1610612749
      If Right(sPath, 1) = "\" Then
            TrimTrailingSlash = Left(sPath, Len(sPath) - 1)
      Else
            TrimTrailingSlash = sPath
      End If
End Function

Public Function VBstringToCstring(ByVal sVBstring As String) As Byte()
Attribute VBstringToCstring.VB_UserMemId = 1610612744
      ' Just, never you mind this.  Not the inverse of above.
      sVBstring = sVBstring & Chr(0)
      VBstringToCstring = sVBstring
      'VBstringToCstring = StrConv(sVBstring, vbFromUnicode)
End Function

