Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal MSG As Long, _
        wParam As Any, _
        lParam As Any) As Long
    'used everywhere, mainly sending EM_* messages to the Editor

' *************************************************************
' Windows API: message/notification constants
' *************************************************************

Public Const EM_LINEINDEX = &HBB
    'p1 = line number, -1 for currently selected
    'p2 = 0
    ' EM_LINEINDEX gets the index, from the rtxt's beginning,
    '   of the first character on a specific line.

Public Const EM_LINELENGTH = &HC1
    'p1 = character index (not line index), -1 for (see help)
    'p2 = 0

Public Const EM_SETSEL = &HB1
    'p1 = position of first selected char, -1 for no selection
    'p2 = position of last selected char + 1

Public Const EM_GETLINECOUNT = 186

Public Const EM_GETMODIFY As Long = &HB8

Public Const EM_UNDO = &HC7
    'p1=0
    'p2=0

Public Const WM_CANCELMODE As Long = &H1F
    'p1=0
    'p2=0
    'using this to disappear a popup menu

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
      'p1=modifier keys
      'p2=(ycoord * &H10000) + xcoord
Public Const MK_CONTROL = &H8
Public Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_XBUTTON1 = &H20
Public Const MK_XBUTTON2 = &H40
      ' note that these MKs are modifier key constants;
      '     are not sendable messages

Public Const WM_USER As Long = &H400

Public Const EM_STOPGROUPTYPING As Long = (WM_USER + 88)
      'for rich edit 2.0+ only
      ' stops the current group of undo actions, starts a new one
      'p1=0
      'p2=0

Public Const EM_SETFONTSIZE As Long = (WM_USER + 223)
      'for rich edit 3.0+ only
      ' increases or decreases the font size
      'p1=amount to increase by (can be negative)
      'p2=0

' *************************************************************
' Windows API: carets and cursors
' *************************************************************

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function GetCaretPos Lib "user32.dll" ( _
     ByRef lpPoint As POINTAPI) As Long

Public Declare Function HideCaret Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Public Declare Function SetCaretPos Lib "user32.dll" ( _
     ByVal x As Long, _
     ByVal y As Long) As Long
Public Declare Function ShowCaret Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Public Declare Sub mouse_event Lib "user32.dll" ( _
     ByVal dwFlags As Long, _
     ByVal dx As Long, _
     ByVal dy As Long, _
     ByVal cButtons As Long, _
     ByVal dwExtraInfo As Long)

Public Declare Function GetCursorPos Lib "user32.dll" ( _
     ByRef lpPoint As POINTAPI) As Long

Public Declare Function SetCursorPos Lib "user32.dll" ( _
     ByVal x As Long, _
     ByVal y As Long) As Long


' *************************************************************
' Windows API: menu functions
' *************************************************************

Public Declare Function GetMenu Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32.dll" ( _
       ByVal hMenu As Long, _
       ByVal nPos As Long) As Long

Public Declare Function GetMenuItemID Lib "user32.dll" ( _
       ByVal hMenu As Long, _
       ByVal nPos As Long) As Long

Public Declare Function ModifyMenu Lib "user32.dll" Alias "ModifyMenuA" ( _
       ByVal hMenu As Long, _
       ByVal nPosition As Long, _
       ByVal wFlags As Long, _
       ByVal wIDNewItem As Long, _
       ByVal lpString As Any) As Long
Public Const MF_BYPOSITION As Long = &H400& ' for wFlags
Public Const MF_STRING As Long = &H0& ' this one too

Public Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" ( _
       ByVal hMenu As Long, _
       ByVal uItem As Long, _
       ByVal fByPosition As Boolean, _
       ByRef lpMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" ( _
       ByVal hMenu As Long, _
       ByVal uItem As Long, _
       ByVal fByPosition As Boolean, _
       ByRef lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" ( _
       ByVal hMenu As Long, _
       ByVal un As Long, _
       ByVal bool As Boolean, _
       ByRef lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" ( _
       ByVal hMenu As Long, _
       ByVal wFlags As Long, _
       ByVal wIDNewItem As Long, _
       ByVal lpNewItem As Any) As Long


Public Type MENUITEMINFO
      cbSize As Long
      fMask As Long
      fType As Long
      fState As Long
      wID As Long
      hSubMenu As Long
      hbmpChecked As Long
      hbmpUnchecked As Long
      dwItemData As Long
      dwTypeData As String
      cch As Long
End Type



' *************************************************************
' Windows API: keyboard accelerators
' *************************************************************

Public Type ACCEL
      fVirt As Byte
      key As Integer
      cmd As Integer
End Type

Public Type MSG
      hwnd As Long
      message As Long
      wParam As Long
      lParam As Long
      time As Long
      pt As POINTAPI
End Type

Public Declare Function LoadAccelerators Lib "user32.dll" Alias "LoadAcceleratorsA" ( _
       ByVal hInstance As Long, _
       ByVal lpTableName As String) As Long

Public Declare Function CopyAcceleratorTable Lib "user32.dll" Alias "CopyAcceleratorTableA" ( _
       ByVal hAccelSrc As Long, _
       ByRef lpAccelDst As ACCEL, _
       ByVal cAccelEntries As Long) As Long

Public Declare Function DestroyAcceleratorTable Lib "user32.dll" ( _
       ByVal haccel As Long) As Long

Public Declare Function TranslateAccelerator Lib "user32.dll" Alias "TranslateAcceleratorA" ( _
       ByVal hwnd As Long, _
       ByVal hAccTable As Long, _
       ByRef lpMsg As MSG) As Long

Public Declare Function CreateAcceleratorTable Lib "user32.dll" Alias "CreateAcceleratorTableA" ( _
       ByRef lpaccl As ACCEL, _
       ByVal cEntries As Long) As Long


' *************************************************************
' Windows API: other
' *************************************************************


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
      ' why not just use a property?  'cause we were supplied an hWnd, that's why.

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
       ByVal hwnd As Long, _
       ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC As Long = -4  'for nIndex

Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" ( _
       ByVal hwnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long) As Long

Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
       ByVal lpPrevWndFunc As Long, _
       ByVal hwnd As Long, _
       ByVal MSG As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long) As Long



' *************************************************************
' My custom types
' *************************************************************

Public Type StatType
    i As Long
    imax As Long
    y As Long
    ymax As Long
    x As Long
    xmax As Long
End Type

Public pOldProc As Long

Public SuppressWhatever As Boolean

' *************************************************************
' My functions and subs
' *************************************************************
    
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      
      If uMsg <> 15 And uMsg <> 32 And uMsg <> 132 And uMsg <> 512 And _
            uMsg <> 275 Then
            ' WM_PAINT, WM_SETCURSOR, WM_SETICON, WM_MOUSEMOVE
            ' WM_TIMER
            Debug.Print Hex(uMsg) & vbTab & "(" & uMsg & ")" & vbTab & wParam & vbTab & lParam
      End If

      WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
End Function
