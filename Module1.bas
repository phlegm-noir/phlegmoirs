Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd _
        As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Public Const EM_LINEINDEX = &HBB
    'p1 = line number, -1 for currently selected
    'p2 = 0

Public Const EM_LINELENGTH = &HC1
    'p1 = character index (not line index), -1 for (see help)
    'p2 = 0

Public Const EM_SETSEL = &HB1
    'p1 = position of first selected char, -1 for no selection
    'p2 = position of last selected char + 1

Public Const EM_GETLINECOUNT = 186
Public Const EM_GETMODIFY As Long = &HB8


Public Type StatType
    i As Long
    imax As Long
    y As Long
    ymax As Long
    x As Long
    xmax As Long
End Type
    
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function GetCaretPos Lib "user32.dll" ( _
     ByRef lpPoint As POINTAPI) As Long

Public Declare Function HideCaret Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

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


Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

