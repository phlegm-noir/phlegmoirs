Attribute VB_Name = "APIOther"
' *************************************************************
' Windows API: other
' *************************************************************

Option Explicit
Option Compare Binary

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
      ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByRef wParam As Any, _
      ByRef lParam As Any) As Long

Public Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" ( _
      ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByRef wParam As Any, _
      ByRef lParam As String) As Long
      
Public Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageA" ( _
      ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByRef wParam As Long, _
      ByRef lParam As Long) As Long

Public Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
End Type

Public Type CHARRANGE
      cpMin As Long
      cpMax As Long
End Type

Public Const LF_FACESIZE As Long = 32

Public Type CHARFORMAT2
      cbSize As Long
      dwMask As Long
      dwEffects As Long
      yHeight As Long
      yOffset As Long
      crTextColor As Long
      bCharSet As Byte
      bPitchAndFamily As Byte
      szFaceName(LF_FACESIZE) As Byte
'      szFaceName As String * LF_FACESIZE
      wWeight As Integer
      sSpacing As Integer
      crBackColor As Long
      lcid As Long
      dwReserved As Long
      sStyle As Integer
      wKerning As Integer
      bUnderlineType As Byte
      bAnimation As Byte
      bRevAuthor As Byte
      bReserved1 As Byte
End Type

Type FINDTEXTEX
      chrg As CHARRANGE
      lpstrText As String
      chrgText As CHARRANGE
End Type

Public Type WINDOWPLACEMENT
      Length As Long
      flags As Long
      showCmd As Long
      ptMinPosition As POINTAPI
      ptMaxPosition As POINTAPI
      rcNormalPosition As RECT
End Type

Public Const SW_MINIMIZE As Long = 6
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWNORMAL As Long = 1

Public Declare Function GetWindowPlacement Lib "user32.dll" ( _
      ByVal hwnd As Long, _
      ByRef lpwndpl As WINDOWPLACEMENT) As Long
      

Public Declare Function SetWindowPlacement Lib "user32.dll" ( _
      ByVal hwnd As Long, _
      ByRef lpwndpl As WINDOWPLACEMENT) As Long
      
Public Declare Function GetWindowRect Lib "user32.dll" ( _
      ByVal hwnd As Long, _
      ByRef lpRect As RECT) As Long

Public Declare Function GetClientRect Lib "user32.dll" ( _
      ByVal hwnd As Long, _
      ByRef lpRect As RECT) As Long

Public Const SWP_HIDEWINDOW As Long = &H80
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOCOPYBITS As Long = &H100
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOREDRAW As Long = &H8
Public Const SWP_NOSENDCHANGING As Long = &H400
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_SHOWWINDOW As Long = &H40


Public Declare Function SetWindowPos Lib "user32.dll" ( _
       ByVal hwnd As Long, _
       ByVal hWndInsertAfter As Long, _
       ByVal x As Long, _
       ByVal y As Long, _
       ByVal cx As Long, _
       ByVal cy As Long, _
       ByVal wFlags As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
       
Public Declare Function ShowScrollBar Lib "user32.dll" ( _
      ByVal hwnd As Long, _
      ByVal wBar As Long, _
      ByVal bShow As Long) As Long
      
Public Const SB_HORZ As Long = 0  ' for wBar


Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
      Destination As Any, _
      Source As Any, _
      ByVal Length As Long)

Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
       ByVal hwnd As Long, _
       ByVal lpString As String) As Long

Public Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
       ByVal hwnd As Long, _
       ByVal lpString As String, _
       ByVal hData As Long) As Long


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
      ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long

Public Const SW_SHOW As Long = 5



Public Declare Function WindowFromPoint Lib "user32.dll" ( _
       ByVal xPoint As Long, _
       ByVal yPoint As Long) As Long
      'gets the control (aka, window) under the cursor.  FINALLY.

Public Declare Function APISetFocus Lib "user32.dll" Alias "SetFocus" ( _
       ByVal hwnd As Long) As Long
      ' why not just use a .SetFocus property?  'cause we were supplied an hWnd, that's why.
      
'Public Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" ( _
'       ByVal lpFileName As String, _
'       ByVal nBufferLength As Long, _
'       ByVal lpBuffer As String, _
'       ByVal lpFilePart As String) As Long
      
Public Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" ( _
       ByVal nBufferLength As Long, _
       ByVal lpBuffer As String) As Long
      ' lpBuffer = string of null terminated drive strings terminated by another null


Public Type FILETIME
      dwLowDateTime As Long
      dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
      wYear As Integer
      wMonth As Integer
      wDayOfWeek As Integer
      wDay As Integer
      wHour As Integer
      wMinute As Integer
      wSecond As Integer
      wMilliseconds As Integer
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" ( _
       ByRef lpFileTime As FILETIME, _
       ByRef lpSystemTime As SYSTEMTIME) As Long
       
Public Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" ( _
       ByRef lpFileTime As FILETIME, _
       ByRef lpLocalFileTime As FILETIME) As Long

Public Const MAX_PATH As Long = 260

Public Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime As FILETIME
      nFileSizeHigh As Long
      nFileSizeLow As Long
      dwReserved0 As Long
      dwReserved1 As Long
      cFileName As String * MAX_PATH
      cAlternate As String * 14
End Type

Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10


Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" ( _
       ByVal lpFileName As String, _
       ByRef lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" ( _
       ByVal hFindFile As Long, _
       ByRef lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare Function FindClose Lib "kernel32.dll" ( _
       ByVal hFindFile As Long) As Long


Public Type SHELLEXECUTEINFO
      cbSize As Long
      fMask As Long
      hwnd As Long
      lpVerb As String
      lpFile As String
      lpParameters As String
      lpDirectory As String
      nShow As Long
      hInstApp As Long
      ' fields
      lpIDList As Long
      lpClass As String
      hkeyClass As Long
      dwHotKey As Long
      hIcon As Long
      hProcess As Long
End Type

Public Const SEE_MASK_INVOKEIDLIST As Long = &HC

Public Declare Function ShellExecuteEx Lib "shell32.dll" ( _
       ByRef lpExecInfo As SHELLEXECUTEINFO) As Long


Public Type SHFILEOPSTRUCT
      hwnd As Long
      wFunc As Long
      pFrom As String
      pTo As String
      fFlags As Integer
      fAborted As Long
      hNameMaps As Long
      sProgress As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" ( _
       ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_MOVE As Long = &H1
Public Const FO_RENAME As Long = &H4
Public Const FOF_SILENT As Long = &H4
Public Const FOF_NORECURSION As Long = &H1000
Public Const FOF_NOCONFIRMMKDIR As Long = &H200
Public Const FOF_ALLOWUNDO = &H40

Public Declare Function GetLastError Lib "kernel32.dll" () As Long
Public Const ERROR_CANCELLED As Long = 1223&


Public Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
       ByVal lOleColor As Long, _
       ByVal lHPalette As Long, _
       ByRef lColorRef As Long) As Long
       
Public Type LVHITTESTINFO
      pt As POINTAPI
      flags As Long
      iItem As Long
      iSubItem As Long
 End Type

Private Const LVM_FIRST As Long = &H1000
Public Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Public Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
      ' p1 = byval 0
      ' p2 = LVHITTESTINFO
      ' returns LVHITTESTINFO.iItem, I believe.  Or -1, if not over an item.
      
      
Public Type MP3TagInfo
      tag As String * 3
      title As String * 30
      artist As String * 30
      album As String * 30
      year As String * 4
      comment As String * 30
      genre As String * 1
End Type

' Y = high order word
' X = low order word

Public Function MAKEPOINT(ByVal lParam As Long) As POINTAPI
      ' BE CAREFUL WITH THESE PESKY AMPERSANDS!
      ' &n = octal n
      ' &Hn = hexadecimal n
      ' n& = Long integer n
      ' &HFFFF = -1   (hex, but NOT A LONG!)
      ' &HFFFF& = 65535   (hex, long)
      ' See the difference?
      
      ' For ints that aren't longs:
      ' &H0001 = 1
      ' &H7FFF = 32767
      ' &H8000 = -32767
      ' &H8001 = -32766
      ' ...
      ' &HFFFE = -2
      ' &HFFFF = -1
      
      MAKEPOINT.y = lParam / &H10000    ' remainderless division to get high order word
      MAKEPOINT.x = lParam And &H7FFF&    ' mask to get low order word
End Function

Function TranslateColor(ByVal Clr As OLE_COLOR, _
      Optional ByVal lHPalette As Long = 0) As Long
      
      Const CLR_INVALID As Long = &HFFFF&

      If OleTranslateColor(Clr, lHPalette, TranslateColor) Then
            TranslateColor = CLR_INVALID
      End If
End Function

