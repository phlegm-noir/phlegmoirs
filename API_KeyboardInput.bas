Attribute VB_Name = "APIKeyb"
' *************************************************************
' Windows API: keyboard input
' *************************************************************


Public Const VK_RIGHT As Long = &H27
Public Const VK_LEFT As Long = &H25
Public Const VK_SHIFT As Long = &H10
Public Const VK_CONTROL As Long = &H11
Public Const VK_MENU As Long = &H12
Public Const VK_BACK As Long = &H8
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_TAB As Long = &H9


Public Const WM_KEYDOWN As Long = &H100
      'p1 = virtual key code
      'p2 = key data
      ' return should be zero

Public Declare Function GetKeyState Lib "user32.dll" ( _
       ByVal nVirtKey As Long) As Integer

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



