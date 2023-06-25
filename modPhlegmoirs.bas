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

Public gtStats As TStatType
Public gsPhlegmKey As String
Public gsPhlegmDate As String

Public glOldLvwProc As Long, glOldpicBrowserProc As Long, glOldpicEditorProc As Long
Public glOldfrmFullScreenProc As Long
Public goFso As Object
Public gtBrowserData As TBrowserData
Public gtImageData As TImageData
Public geEditorMode As eViewMode
Public giTextEncoding As Integer
Public gbFullScreenMode As Boolean
Public gsCommandFile As String

Public Const MOVE_INCREMENT = -512

Public Function AbsoluteBottom(ByRef roCtrl As Control) As Long
      AbsoluteBottom = AbsoluteTop(roCtrl) + roCtrl.Height
End Function

' WARNING: these do not account for frame/picturebox borders.  Use informally.

Public Function AbsoluteLeft(ByRef roCtrl As Control) As Long
      On Error Resume Next
      AbsoluteLeft = roCtrl.Left + AbsoluteLeft(roCtrl.Container)
      If Err > 0 Then AbsoluteLeft = roCtrl.Left
      DebugLog "AbsoluteLeft throws an error, whatever that means..."
      On Error GoTo 0
End Function

Public Function AbsoluteRight(ByRef roCtrl As Control) As Long
      AbsoluteRight = AbsoluteLeft(roCtrl) + roCtrl.Width
End Function

Public Function AbsoluteTop(ByRef roCtrl As Control) As Long
      On Error Resume Next
      AbsoluteTop = roCtrl.Top + AbsoluteTop(roCtrl.Container)
      If Err > 0 Then AbsoluteTop = roCtrl.Top
      DebugLog "AbsoluteTop throws an error, whatever that means..."
      On Error GoTo 0
End Function

Public Function ArrContains(arrString, ByVal sPassedValue As String) As Boolean
      Dim iIndex As Integer
      For iIndex = LBound(arrString) To UBound(arrString)
            If arrString(iIndex) = sPassedValue Then
                  ArrContains = True
                  Exit Function
            End If
      Next
      ArrContains = False
End Function

' Extract the character count, when counting only one character per carriage return.

Public Function CharacterCount(ByRef roEditor As agRichEdit) As Long
      
      Dim lLastLineLength As Long, lLastLineIndex As Long
      
      lLastLineIndex = SendMessage(roEditor.RichEdithWnd, EM_LINEINDEX, ByVal (roEditor.LineCount - 1), ByVal 0)
      
      lLastLineLength = SendMessage(roEditor.RichEdithWnd, EM_LINELENGTH, ByVal lLastLineIndex + 1, ByVal 0)

      CharacterCount = lLastLineIndex + lLastLineLength
End Function

Public Function CstringToVBstring(ByVal sCstring As String) As String
      ' Removes first null character and anything following it.
      On Error GoTo CONVERSION_ERROR
      Dim lNullPos As Long
      
      lNullPos = InStr(1, sCstring, Chr(0))
      If lNullPos = 0 Then
            CstringToVBstring = sCstring
      Else
            CstringToVBstring = Left(sCstring, lNullPos - 1)
      End If
      Exit Function
CONVERSION_ERROR:
      DebugLog "CONVERSION ERROR: " & sCstring, 2
End Function

Public Sub DebugLog(ByVal sMsg As String, Optional ByVal iLogLevel As Integer = 1)
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

      Dim tWfd As WIN32_FIND_DATA
      Dim lHfile As Long
      
      lHfile = FindFirstFile(sSource, tWfd)
      FileExists = lHfile <> -1 ' invalid handle value
      
      FindClose (lHfile)
   
End Function

Public Function FileModifiedTime(ByVal sSource As String) As String
      ' example date string:   2005-03-15 6:14:21
      
      Dim tWfd As WIN32_FIND_DATA
      Dim lHfile As Long
      
      lHfile = FindFirstFile(sSource, tWfd)
      
      If lHfile > 0 Then
            FileModifiedTime = FormatNonLocalFileTime(tWfd.ftLastWriteTime)
      Else
            FileModifiedTime = ""
      End If
      
      FindClose (lHfile)
End Function

Function FormatBytes(ByVal oBytes, iPrecision As Integer) As String
      ' This function takes a quantity of bytes as a currency value ('cause it's 64-bit),
      ' formats it to read likeathis:
      
      ' 45.2 MB
      ' 300.2 KB
      ' 666 Bytes
      ' 99.4444 GB
      ' 20000 TB  (not supporting terabytes at the moment, sorry folks)
      
      ' ...with iPrecision digits after the demical.
      
      If oBytes = 1 Then
            FormatBytes = CStr(oBytes) & " byte"
      ElseIf oBytes < 1024@ Then
            FormatBytes = CStr(oBytes) & " bytes"
      ElseIf oBytes < 1048576@ Then
            FormatBytes = CStr(Round(oBytes / 1024@, iPrecision)) & " KB"
      ElseIf oBytes < 1073741824@ Then
            FormatBytes = CStr(Round(oBytes / 1048576@, iPrecision)) & " MB"
      ElseIf oBytes < 1099511627776@ Then
            FormatBytes = CStr(Round(oBytes / 1073741824@, iPrecision)) & " GB"
      Else
            FormatBytes = "Size Unknown"
      End If
End Function

Public Function FormatNonLocalFileTime(NLFT As FILETIME) As String
      ' example date string:   2005-03-15 6:14:21
      
      Dim tLocalTime As FILETIME
      Dim tSysTime As SYSTEMTIME
      
      FileTimeToLocalFileTime NLFT, tLocalTime
      FileTimeToSystemTime tLocalTime, tSysTime
      With tSysTime
            FormatNonLocalFileTime = .wYear & "-" & Format(.wMonth, "00") & "-" & Format(.wDay, "00") _
                  & ", " & Format(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00")
      End With
End Function

Public Function getAllProperties(sFileName As String)
      Dim oBaseName, oPathName
      oPathName = goFso.getParentFolderName(sFileName)
      oBaseName = goFso.GetFileName(sFileName)
      
      Dim oShell, oFolder
      Set oShell = CreateObject("shell.application")
      Set oFolder = oShell.NameSpace(oPathName)
      
      If (Not oFolder Is Nothing) Then
            Dim oFolderItem
            Set oFolderItem = oFolder.ParseName(oBaseName)
            
            If (Not oFolderItem Is Nothing) Then
                  Dim oUninteresting
                  oUninteresting = Array("Attributes", "Owner", "Total size", "Computer", "File extension", _
                        "Filename", "Space free", "Shared", "Folder name", "File location", "Folder", "Path", "Type", _
                        "Link status", "Space used", "Sharing status", "UNKNOWN(296)", "Content", "Rating", "Shared with", _
                        "Protected")
                  
                  Dim oIndex
                  For oIndex = 0 To 320
                        Dim oColumnName, oValue
                        oColumnName = oFolder.GetDetailsOf(oFolder.Items, oIndex)
                        oValue = oFolder.GetDetailsOf(oFolderItem, oIndex)
                        If oColumnName = "" Then oColumnName = "UNKNOWN(" + Trim(Str(oIndex)) + ")"
                        
                        If oValue <> "" And Not ArrContains(oUninteresting, oColumnName) Then
                              Debug.Print Str(oIndex) + ". " + oColumnName + ": " + vbTab + oValue
                        End If
                  Next oIndex
            End If
            
            Set oFolderItem = Nothing
      End If
      
      Set oFolder = Nothing
      Set oShell = Nothing
End Function

Public Function GetFileSize(ByVal sSource As String) As Currency
      If sSource = "" Then
            GetFileSize = 0
      Else
            GetFileSize = goFso.getfile(sSource).Size
      End If
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

' Figure out the numbering of a menu caption, and which digit to underline.
Public Function GetNumberedCaption(ByVal sFileName As String, ByVal iIndex As Integer) As String
      If iIndex < 10 Then
            GetNumberedCaption = "&" & iIndex & "   " & sFileName
      ElseIf iIndex = 10 Then
            GetNumberedCaption = "1&0   " & sFileName
      Else
            GetNumberedCaption = iIndex & "   " & sFileName
      End If
End Function

Public Function GetViewMode(ByVal sFileName As String, ByVal eMode As eIconType) As eViewMode
      Const PIC_FILE_TOO_BIG = 10000000
      Const NON_PIC_FILE_TOO_BIG = 2097152
      Dim sEx As String
      Dim oFileSize
      
      If Not goFso.FileExists(sFileName) Then
            GetViewMode = eViewMode.ERROR
            DebugLog "ViewMode: ERROR for file: " & sFileName & "..."
            Exit Function
      End If
      
      sEx = goFso.getextensionname(sFileName)
      
      If eMode = eIconType.Bookmark Then
            eMode = GetIconType(sEx)
      End If
      
      Select Case eMode
            Case eIconType.Picture
                  oFileSize = GetFileSize(sFileName)
                  
                  If sEx = "png" Or sEx = "webp" Or oFileSize > PIC_FILE_TOO_BIG Then
                        GetViewMode = eViewMode.PropertiesView
                  Else
                        GetViewMode = eViewMode.PictureView
                  End If
            
            Case eIconType.Text, eIconType.rtf, eIconType.other
                  oFileSize = GetFileSize(sFileName)
                  
                  If oFileSize > NON_PIC_FILE_TOO_BIG Then
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
      IsKeyDown = GetKeyState(lVirtKey) And &HF0000000
End Function

Public Function IsPathFull(ByVal sInput As String) As Long
      ' returns 0 if not a full path
      ' if a full path, returns position of colon
      ' does NOT check if the path is a VALID path
      IsPathFull = InStrRev(sInput, ":")
End Function

Public Function IsUnicodeFile(oFilePath)
      Dim oStream
      Dim oAsc1Chr, oAsc2Chr
      
      If Not goFso.FileExists(oFilePath) Then
            IsUnicodeFile = eTextEncoding.ERROR
            DebugLog "Text encoding=ERROR for file: " & oFilePath
            Exit Function
      ElseIf GetFileSize(oFilePath) = 1 Then
            IsUnicodeFile = eTextEncoding.ASCII
            Exit Function
      End If
      
      On Error Resume Next
      Set oStream = goFso.OpenTextFile(oFilePath, eIoMode.ForReading, eCreate.No, eTextEncoding.ASCII)
      oAsc1Chr = AscW(oStream.Read(1))
      oAsc2Chr = AscW(oStream.Read(1))
      oStream.Close
      If Err > 0 Then
            IsUnicodeFile = eTextEncoding.ERROR
            DebugLog "Text encoding=ERROR for file: " & oFilePath
            Exit Function
      End If
      On Error GoTo 0
      
      If (oAsc1Chr = 255) And (oAsc2Chr = 254) Then
            IsUnicodeFile = eTextEncoding.UNICODE
      Else
            IsUnicodeFile = eTextEncoding.ASCII
      End If
      
      Set oStream = Nothing
End Function

Public Function ListViewProc(ByVal lHwnd As Long, ByVal lMsg As Long, _
      ByVal lWparam As Long, ByVal lLparam As Long) As Long

      Select Case lMsg
            Case WM_KEYDOWN
                  ' The problem: right arrow key wants to scroll right.  I want it to do things
                  ' like enter the given folder, which would cause the listview to do both at once.
                  
                  ' Solution: this modified windows procedure will translate a right arrow (without roCtrl)
                  ' into an F13 keyDown and then continue the window procedure.
                  
                  ' Elsewhere, F13 will be bound to do the desired right arrow key things.
                  
                  ' If somebody wants to scroll right with arrow keys, they can use ctrl+right.
                  
                  If lWparam = vbKeyRight And Not IsKeyDown(VK_CONTROL) Then
                        
                        ListViewProc = CallWindowProc(glOldLvwProc, lHwnd, _
                              WM_KEYDOWN, vbKeyF13, lLparam)
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
                  
                  ' Can be positive or negative (not zero)
                  ' If you spin the wheel REALLY fast, it sends a single
                  ' WM_MOUSEWHEEL message grouping multiple turns together.
                  ' But it's normally either 1 or -1, sending a message for each turn.
                  Dim lWheelTurns As Long
                  
                  Dim tHitTestInfo As LVHITTESTINFO
                  Dim tRecClient As RECT

                  lWheelTurns = MAKEPOINT(lWparam).Y / WHEEL_DELTA
                  lRetVal = GetWindowRect(lHwnd, tRecClient)

                  With tHitTestInfo
                        .pt.X = MAKEPOINT(lLparam).X - tRecClient.Left
                        .pt.Y = MAKEPOINT(lLparam).Y - tRecClient.Top
                  End With
                  lRetVal = SendMessage(lHwnd, LVM_HITTEST, ByVal 0, tHitTestInfo)

                  If (tHitTestInfo.flags And LVHT_BELOW) Or (tHitTestInfo.flags And LVHT_ABOVE) Then
                        Dim iTurn As Integer
                        If lWheelTurns > 0 Then
                              For iTurn = 1 To lWheelTurns * 3
                                    lRetVal = SendMessage(lHwnd, WM_HSCROLL, ByVal SB_LINELEFT, ByVal 0)
                              Next iTurn
                              Exit Function
                        ElseIf lWheelTurns < 0 Then
                              For iTurn = -1 To lWheelTurns * 3 Step -1
                                    lRetVal = SendMessage(lHwnd, WM_HSCROLL, ByVal SB_LINERIGHT, ByVal 0)
                              Next iTurn
                              Exit Function
                        ElseIf lWheelTurns = 0 Then
                              MsgBox "ERROR: wheel turn = 0.  How can you turn the wheel zero turns?"
                              DebugLog "ERROR: wheel turn = 0.  How can you turn the wheel zero turns?"
                        End If
                  End If
      End Select
      
      ' Do all the default things too, as defined by the old procedure
      ListViewProc = CallWindowProc(glOldLvwProc, lHwnd, lMsg, lWparam, lLparam)
End Function

Public Function RecycleFile(ByVal sPath As String) As Integer
      ' Send a file to the Recycle Bin.
      
      Dim tFileOperations As SHFILEOPSTRUCT
      
      With tFileOperations
            .wFunc = FO_DELETE
            .pFrom = sPath
            .fFlags = FOF_ALLOWUNDO
      End With
      RecycleFile = SHFileOperation(tFileOperations)
End Function

Public Function SnipFileName(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipFileName = Left(sPath, iSlash)
End Function

' *************************************************************
' Useful functions and subs
' *************************************************************

Public Function SnipPath(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipPath = Right(sPath, Len(sPath) - iSlash)
End Function

Public Function TrackMouseLeave(ByVal lHwnd As Long, ByVal lMsg As Long, _
      ByVal lWparam As Long, ByVal lLparam As Long) As Long
            
      If lMsg = WM_MOUSELEAVE Then
            Beep
      End If
      TrackMouseLeave = CallWindowProc(glOldpicBrowserProc, lHwnd, lMsg, lWparam, lLparam)
End Function

Public Function TrackMouseWheel(ByVal lHwnd As Long, ByVal lMsg As Long, _
      ByVal lWparam As Long, ByVal lLparam As Long) As Long

      ' This is for picture manipulation; applies to picEditor.
      
      If lMsg = WM_MOUSEWHEEL Then
            Dim tPoiWheel As POINTAPI
            Dim iVirtKeys As Integer, iWheelTurn As Integer
            
            tPoiWheel = MAKEPOINT(lWparam)
            iWheelTurn = tPoiWheel.Y / WHEEL_DELTA
            iVirtKeys = tPoiWheel.X
            tPoiWheel = MAKEPOINT(lLparam)
            frmMain.WheelInput iWheelTurn, iVirtKeys, tPoiWheel.X, tPoiWheel.Y
      End If
      TrackMouseWheel = CallWindowProc(glOldpicEditorProc, lHwnd, lMsg, lWparam, lLparam)
End Function

Public Function TrimTrailingSlash(ByVal sPath As String) As String
      If Right(sPath, 1) = "\" Then
            TrimTrailingSlash = Left(sPath, Len(sPath) - 1)
      Else
            TrimTrailingSlash = sPath
      End If
End Function

Public Function VBstringToCstring(ByVal sVbString As String) As Byte()
      ' Just, never you mind this.  Not the inverse of above.
      sVbString = sVbString & Chr(0)
      VBstringToCstring = sVbString
      'VBstringToCstring = StrConv(sVbString, vbFromUnicode)
End Function

