Attribute VB_Name = "modPhlegmoirs"


' *************************************************************
' My custom types
' *************************************************************

Option Compare Text
Option Explicit

Public Type TStatType
    i As Long
    imax As Long
    y As Long
    ymax As Long
    x As Long
    xmax As Long
End Type

Public Type TEditorPrefs
      WordWrap As Integer
      SelStart As Long
      SelEnd As Long
      FirstVisibleLine As Long
      FontName As String * 255
      FontSize As Currency
      FontBold As Boolean
      FontItalic As Boolean
      FontUnderline As Boolean
      FontStrikethrough As Boolean
      FontCharset As Integer
      ScrollPos As POINTAPI
      
End Type

Public Type TWindowPrefs
      WNP As WINDOWPLACEMENT
      BrowserWidth As Long
      SortMethod As Integer
      ShowFileBrowser As Boolean
      ShowStatusBar As Boolean
      ShowToolBar As Boolean
      BookmarkCount As Integer
      AutoLoadFile As String * 255
      cboPath As String * 255
End Type

Public Type TBrowserData
      Dir As String
      DirPrev As String
      InputPrev As String
      Filter As String
      PartialFileName As String
      SelTextPrev As String
      DirUnchanged As Boolean
      GoingToParent As Boolean
      BookmarkMode As Boolean
      DrivesMode As Boolean
      ValidPath As Boolean
      Error As Boolean
      ListEmpty As Boolean
      
      ItemClicked As Boolean
      ButtonPressed As Boolean
      MouseButton As Integer
      Shift As Integer
      RecentPath As Integer
End Type
      

Public Type TURLQuery
      name As String * 50
      URL As String * 255
      color As Long
      key As String * 5
End Type

Public gpOldLvwBrowserProc As Long, gpOldpicBrowserProc As Long
Public objtest As Object
Public gFSO As Object  ' nothing I hate more than declaring one of these in each function, for one tiny little purpose.
Public gBrowserData As TBrowserData

Public Const glMAX_LONG_INTEGER = &H7FFFFFFF   '   2147483647


' *************************************************************
' My functions and subs
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

Public Function SuppressArrowKeysProc(ByVal hWnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

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
      
      Select Case uMsg
            Case WM_KEYDOWN
                  
                  'Debug.Print "LVW_BROWSER: WM_KEYDOWN: " & wParam & " " & lParam
                  
                  If wParam = vbKeyRight And Not IsKeyDown(VK_CONTROL) Then
                                    
                        SuppressArrowKeysProc = CallWindowProc(gpOldLvwBrowserProc, hWnd, _
                              WM_KEYDOWN, vbKeyF13, lParam)
                        Exit Function
                  End If
      End Select
                  
      SuppressArrowKeysProc = CallWindowProc(gpOldLvwBrowserProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function TrackMouseLeave(ByVal hWnd As Long, ByVal uMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
            
      Dim objPFB As Object
      
      If uMsg = WM_MOUSELEAVE Then
            Beep
      End If
      TrackMouseLeave = CallWindowProc(gpOldpicBrowserProc, hWnd, uMsg, wParam, lParam)
End Function


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
      Dim lngNullPosition As Long
      
      lngNullPosition = InStr(1, sCstring, Chr(0))
      If lngNullPosition = 0 Then
            CstringToVBstring = sCstring
      Else
            CstringToVBstring = Left(sCstring, lngNullPosition - 1)
      End If
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


Public Sub RecycleFile(ByVal sPath As String)
      ' Send a file to the Recycle Bin.
      
      Dim shfFileOperation As SHFILEOPSTRUCT
      
      With shfFileOperation
            .wFunc = FO_DELETE
            .pFrom = sPath
            .fFlags = FOF_ALLOWUNDO
      End With
      SHFileOperation shfFileOperation
End Sub
