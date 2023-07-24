Attribute VB_Name = "APIMsg"
' *************************************************************
' Windows API: message/notification constants
' *************************************************************

Option Explicit
Option Compare Binary

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
    
Public Const EM_GETSEL As Long = &HB0


Public Const EM_EXGETSEL As Long = (WM_USER + 52)
      'p1 = 0
      'p2 = pointer to a CHARRANGE structure to receive selection range.
      ' if p2 = (0, -1) then range includes everything.

Public Const EM_GETSELTEXT As Long = (WM_USER + 62)

Public Const EM_GETLINECOUNT = 186

Public Const EM_CHARFROMPOS As Long = &HD7
      ' p1=0
      ' p2 = POINTL (but I'll use POINTAPI) structure of coordinates.
      ' ... I suspect these coordinates to mean line and column numbers.
      ' returns 0 based character position.

Public Const EM_GETMODIFY As Long = &HB8

Public Const EM_UNDO = &HC7
      ' p1 = 0
      ' p2 = 0
Public Const EM_REDO = WM_USER + 84
      'p1=0
      'p2=0
Public Const EM_SETUNDOLIMIT = WM_USER + 84
      ' p1 = maximum undo actions
      ' p2 = 0
Public Const EM_GETUNDONAME = WM_USER + 86
      ' p1 = 0
      ' p2 = 0
Public Const EM_GETREDONAME = WM_USER + 87
      ' p1 = 0
      ' p2 = 0
Public Const EM_SETTEXTMODE = WM_USER + 89
      ' p1 = new text mode
      ' p2 = 0
Public Const TM_PLAINTEXT = 1
Public Const TM_RICHTEXT = 2
Public Const TM_MULTILEVELUNDO = 8
Public Const EM_GETTEXTMODE = WM_USER + 90
      ' p1 = 0
      ' p2 = 0

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
Public Const MK_LBUTTON = &H1
Public Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_XBUTTON1 = &H20
Public Const MK_XBUTTON2 = &H40
      ' note that these MKs are modifier key constants;
      '     are not sendable messages
      
Public Const WM_MOUSEWHEEL As Long = &H20A
      ' lparam low = virtual keys
      ' lparam high = distance wheel is rotated, in multiples of WHEEL_DELTA
      ' wparam low = x-coordinate of pointer   (in pixels, in relation to the screen)
      ' wparam high = y-coordinate of pointer
Public Const WHEEL_DELTA As Long = 120

Public Const WM_HSCROLL As Long = &H114
      ' p1 low = scrolling request
      ' p1 high & p2 = nevermind.

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

Public Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
      ' parameters not unlike EM_GETCHARFORMAT
Public Const SCF_DEFAULT As Long = &H0
Public Const SCF_SELECTION As Long = &H1
Public Const SCF_ALL As Long = &H4

Public Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
      ' p1 = default character formatting, nonzero = current selection's formatting
      ' p2 = CHARFORMAT2 structure (or CHARFORMAT, for richedit < 2.0)

Public Const LVM_FIRST As Long = &H1000

Public Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
      'p1=Index of the list view item. To cancel editing, set iItem to -1.
      'p2 = 0
      'returns handle to edit control if successful
      
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)

Public Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
      ' p1 = sort parameter
      ' p2 = addressof callback function
      
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 115) 'Unicode

Public Type LVFINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type

Public Const LVFI_STRING As Integer = 2
Public Const LVM_FINDITEM As Integer = 4179
Public Const LVIF_TEXT As Long = &H1
Public Const LVIF_PARAM As Long = &H4

Public Type LVITEM
    mask As Integer
    iItem As Integer
    iSubItem As Integer
    State As Integer
    stateMask As Integer
    placeholder1 As Integer
    pszText As String
    placeholder2 As Integer
    cchTextMax As Long
    iImage As Integer
End Type

Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
      ' p1 = zero-based column index
      ' p2 = new column width in pixels
Public Const LVSCW_AUTOSIZE As Long = -1

Public Const WM_NOTIFY As Long = &H4E
      ' This is how events happen
      ' w = identifier of the sending control
      ' l = NMHDR

Public Type NMHDR
      hwndFrom As Long
      idfrom   As Long
      code     As Long
End Type

' WM_NOTIFY codes:
Public Const HDN_FIRST As Long = -300
Public Const HDN_BEGINTRACKA As Long = (HDN_FIRST - 6)
Public Const HDN_BEGINTRACKW As Long = (HDN_FIRST - 26)

' Most WM_NOTIFY message codes just point to a NMHDR.
' But some codes point to an entire NMHEADER.
' In *both cases*, a NMHDR exists at that address
' because the smaller data type is also the larger data type's first field.
Public Type NMHEADER
      hdr As NMHDR
      iItem As Long
      iButton As Long
      lPtrHDItem As Long
End Type

