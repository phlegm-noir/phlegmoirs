Attribute VB_Name = "modSubclass"
Option Explicit
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) _
                                                                            As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) _
                                                                              As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) _
                                                                            As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
Private Const GWL_WNDPROC = (-4)
 
Private Const WM_SIZING = &H214
 
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8
 
Private mlMinWidth As Long ' The minimum width in pixels
Private mlMinHeight As Long ' The minimum height in pixels
 
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const LAST_VISIBLE_COLUMN As Integer = 3
 
Private mPrevProc As Long
Private mPrevLvwProc As Long

Public Sub ListViewSpyHook(hwnd As Long)
      If mPrevLvwProc = 0& Then
            mPrevLvwProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewLvwProc)
            DebugLog "Setting a new window procedure for a ListView", 2
      End If
End Sub

Public Sub ListViewSpyUnhook(hwnd As Long)
      SetWindowLong hwnd, GWL_WNDPROC, mPrevLvwProc
      mPrevLvwProc = 0&
End Sub

' The Microsoft Common Control ListView can sort much faster than a 3rd party control.
' So that's why we are using this old control despite not having unicode and not having all the nice features.
'
' It can only handle a basic text sort. So it needs extra hidden columns to represent the real sorts, e.g. file size.
' (And the extra information in an extra column doesn't slow it down. It's still the fastest.)
'
' To complement the headache of invisible columns, it brings the extra headache that you cannot turn off resizing.
' So the user might happen to screw around and figure out they have a 0-width column they can expand.
'
' This window procedure stops the resize by interrupting the "begin track" event notification and throwing it away.
'
Public Function NewLvwProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      Select Case uMsg
            
            Case &H204E, &H31, &H84, &H200
                  ' Some msgs are too noisy to allow them a debug statement.
            
            Case WM_NOTIFY
                  Dim nmMsgInfo As NMHDR, nmHeaderInfo As NMHEADER
                  CopyMemory nmHeaderInfo, ByVal lParam, Len(nmHeaderInfo)
                  CopyMemory nmMsgInfo, ByVal lParam, Len(nmMsgInfo)
                  DebugLog "WM_NOTIFY (" & Hex(wParam) & ", " & Hex(lParam) & "); " & GetCodeName(nmMsgInfo.code), 0
                  
                  If nmMsgInfo.code = HDN_BEGINTRACKA And nmHeaderInfo.iItem > LAST_VISIBLE_COLUMN Then
                        NewLvwProc = 1                                 ' ^ iItem is a column index
                        Exit Function
                  End If
            
            Case Else
                  DebugLog GetMsgName(uMsg) & " (" & Hex(wParam) & ", " & Hex(lParam) & ")", 0
      End Select
      
      If mPrevLvwProc > 0& Then
            NewLvwProc = CallWindowProc(mPrevLvwProc, hwnd, uMsg, wParam, lParam)
      Else
            NewLvwProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
      End If
End Function

Public Sub SizeLimiterHook(hwnd As Long, Optional ByVal lMinWidthTwips As Long = -1, Optional ByVal lMinHeightTwips = -1)
      If (lMinWidthTwips > 0) Then mlMinWidth = lMinWidthTwips / Screen.TwipsPerPixelX
      If (lMinHeightTwips > 0) Then mlMinHeight = lMinHeightTwips / Screen.TwipsPerPixelY
      
      If mPrevProc = 0& Then
            mPrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWndProc)
      Else
            SetWindowLong hwnd, GWL_WNDPROC, AddressOf NewWndProc
      End If
      DebugLog "Form's min width has been set: " & lMinWidthTwips, 2
End Sub
 
Public Sub SizeLimiterUnhook(hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, mPrevProc
    mPrevProc = 0&
End Sub
 
' How to limit a form's size (min and max) but SMOOTHLY
'
' Thanks to crptcblade from vbforums.
'
' https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307
'
Public Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      On Error Resume Next
 
      Dim r As RECT
 
      If uMsg = WM_SIZING Then
            Call CopyMemory(r, ByVal lParam, Len(r))
            
            If (r.Right - r.Left < mlMinWidth) Then
                DebugLog "Width too small: " & r.Right - r.Left, 0
                Select Case wParam
                    Case WMSZ_LEFT, WMSZ_BOTTOMLEFT, WMSZ_TOPLEFT
                        r.Left = r.Right - mlMinWidth
                    Case WMSZ_RIGHT, WMSZ_BOTTOMRIGHT, WMSZ_TOPRIGHT
                        r.Right = r.Left + mlMinWidth
                End Select
            End If
            
            If (r.Bottom - r.Top < mlMinHeight) Then
                Select Case wParam
                    Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        r.Top = r.Bottom - mlMinHeight
                    Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        r.Bottom = r.Top + mlMinHeight
                End Select
            End If
            
            Call CopyMemory(ByVal lParam, r, Len(r))
            
            NewWndProc = 0&
            Exit Function
      End If
      
      
      If mPrevProc > 0& Then
            NewWndProc = CallWindowProc(mPrevProc, hwnd, uMsg, wParam, lParam)
      Else
            NewWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
      End If
 
End Function
