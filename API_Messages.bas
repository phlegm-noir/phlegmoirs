Attribute VB_Name = "APIMsg"
' *************************************************************
' Windows API: message/notification constants
' *************************************************************

Public Const WM_USER As Long = &H400

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

Public Const WM_SETREDRAW As Long = &HB
      'p1 = true/false
      'p2 = 0
      ' zero if message is processed

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

Public Const LVM_FIRST As Long = &H1000

Public Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
      'p1=Index of the list view item. To cancel editing, set iItem to -1.
      'p2 = 0
      'returns handle to edit control if successful






