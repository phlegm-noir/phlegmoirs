Attribute VB_Name = "APIMenus"
' *************************************************************
' Windows API: menu functions
' *************************************************************


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




