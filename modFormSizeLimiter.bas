Attribute VB_Name = "modFormSizeLimiter"
' How to limit a form's size (min and max)
'
' Many thanks to crptcblade from vbforums.
'
' https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307

Option Explicit
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) _
                                                                            As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hWnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) _
                                                                              As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, _
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
    RIGHT  As Long
    Bottom As Long
End Type
 
Private mPrevProc As Long
 
Public Sub SizeLimiterHook(hWnd As Long, Optional ByVal lMinWidthTwips As Long = -1, Optional ByVal lMinHeightTwips = -1)
      If (lMinWidthTwips > 0) Then mlMinWidth = lMinWidthTwips / Screen.TwipsPerPixelX
      If (lMinHeightTwips > 0) Then mlMinHeight = lMinHeightTwips / Screen.TwipsPerPixelY
      
      If mPrevProc = 0& Then
            mPrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWndProc)
      Else
            SetWindowLong hWnd, GWL_WNDPROC, AddressOf NewWndProc
      End If
      Debug.Print "Form's min width has been set: " & lMinWidthTwips
End Sub
 
Public Sub SizeLimiterUnhook(hWnd As Long)
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, mPrevProc)
    mPrevProc = 0&
    
End Sub
 
Public Function NewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      On Error Resume Next
 
      Dim r As RECT
 
      If uMsg = WM_SIZING Then
            Call CopyMemory(r, ByVal lParam, Len(r))
            
            'Keep the form only at least as wide as MIN_WIDTH
            If (r.RIGHT - r.Left < mlMinWidth) Then
                Debug.Print "Width too small: " & r.RIGHT - r.Left
                Select Case wParam
                    Case WMSZ_LEFT, WMSZ_BOTTOMLEFT, WMSZ_TOPLEFT
                        r.Left = r.RIGHT - mlMinWidth
                    Case WMSZ_RIGHT, WMSZ_BOTTOMRIGHT, WMSZ_TOPRIGHT
                        r.RIGHT = r.Left + mlMinWidth
                End Select
            End If
            
            'Keep the form only at least as tall as MIN_HEIGHT
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
            NewWndProc = CallWindowProc(mPrevProc, hWnd, uMsg, wParam, lParam)
      Else
            NewWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
      End If
 
End Function

