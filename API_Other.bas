Attribute VB_Name = "APIOther"
' *************************************************************
' Windows API: other
' *************************************************************

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal MSG As Long, _
        wParam As Any, _
        lParam As Any) As Long
    'used everywhere, mainly sending EM_* messages to the agEditor


Public Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
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


