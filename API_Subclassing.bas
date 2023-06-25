Attribute VB_Name = "APISubclass"
' *************************************************************
' Windows API: Subclassing Functions
' *************************************************************
' TODO: change to SetWindowLongPtr for 64-bit compatibility as well as 32-bit.
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




