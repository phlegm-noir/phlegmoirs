Attribute VB_Name = "APICaret"
' *************************************************************
' Windows API: carets and cursors
' *************************************************************

Public Const EM_GETSCROLLPOS As Long = (WM_USER + 221) ' Rich Edit 3.0+
      ' p1 = 0
      ' p2 = POINT
      ' always returns -1

Public Const EM_SETSCROLLPOS As Long = (WM_USER + 222) ' Rich Edit 3.0+
      ' p1 = 0
      ' p2 = POINT
      ' always returns 1

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
     
Public Type TrackMouseEvent
      cbSize As Long
      dwFlags As Long
      hwndTrack As Long
      dwHoverTime As Long
End Type

Public Declare Function TrackMouseEvent Lib "user32.dll" ( _
       ByRef lpEventTrack As TrackMouseEvent) As Long

Public Const TME_CANCEL As Long = &H80000000
Public Const TME_LEAVE As Long = &H2
Public Const WM_MOUSELEAVE As Long = &H2A3

Public Const WM_CAPTURECHANGED As Long = &H215

Public Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetCapture Lib "user32.dll" () As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long

