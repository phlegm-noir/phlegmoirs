VERSION 5.00
Begin VB.Form frmFullScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnClose 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Exit Fullscreen Mode (F11 or Esc)"
      Top             =   0
      Width           =   175
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4560
      Left            =   1320
      MousePointer    =   15  'Size All
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3600
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnClose_Click()
      Unload frmFullScreen
End Sub

Private Sub Form_Load()
      gfFullScreenMode = True
      Set gImageData.OutPic = Image1
      
      CopyDimensions
      Image1.Picture = frmMain.Image1.Picture
      
      gpOldfrmFullScreenProc = SetWindowLong(hWnd, GWL_WNDPROC, _
            AddressOf TrackMouseWheelFullScreen)
End Sub

Private Sub CopyDimensions()
      With frmMain.Image1
            Image1.Move .Left, .Top, .Width, .Height
      End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      
      'Debug.Print KeyCode
      With frmMain.sliZoom
            Select Case KeyCode
                  Case 107, 187 ' "+" and Keypad "+"
                        If Shift = 0 Then
                              frmMain.ImageZoomIn .SmallChange
                        ElseIf Shift = vbCtrlMask Then
                              frmMain.ImageZoomIn .LargeChange
                        End If
                  Case 109, 189 ' "-" and Keypad "-"
                        If Shift = 0 Then
                              frmMain.ImageZoomOut .SmallChange
                        ElseIf Shift = vbCtrlMask Then
                              frmMain.ImageZoomOut .LargeChange
                        End If
                  Case vbKey0, 106 ' 0 and Keypad "*" -- reset position and size.
                        .Value = 100
                        Image1.Move 0, 0, gImageData.DefaultWidth, gImageData.DefaultHeight
                  Case 107, 55   ' 7 and Keypad 7
                        .Value = .Value / 2
                  Case 104, 56   ' 8 and Keypad 8
                        .Value = .Value * 2
                  Case vbKeyDown
                        Image1.Top = Image1.Top + MoveIncrement
                  Case vbKeyUp
                        Image1.Top = Image1.Top - MoveIncrement
                  Case vbKeyLeft
                        Image1.Left = Image1.Left - MoveIncrement
                  Case vbKeyRight
                        Image1.Left = Image1.Left + MoveIncrement
                        
                  Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                        If Shift = 0 Then frmMain.BrowserExecuteNext
                  Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                        If Shift = 0 Then frmMain.BrowserExecutePrev
                        
                  Case vbKeyF11, vbKeyEscape
                        If Shift = 0 Then Unload frmFullScreen
            End Select
      End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not gImageData.Zoomed And Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            frmMain.BrowserExecuteNext
      ElseIf Not gImageData.Zoomed And Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            frmMain.BrowserExecutePrev
      ElseIf Not gImageData.Moved And Button = vbMiddleButton Then
            Unload frmFullScreen
      End If
      
      gImageData.Zoomed = False
      gImageData.Moved = False
End Sub

Private Sub Image1_DblClick()
      ' This needs to (effectively) call an Image1_mousedown... but with what parameters???
      Dim poiPrev As POINTAPI
      
      GetCursorPos poiPrev
      
      gImageData.PrevX = poiPrev.X * Screen.TwipsPerPixelX  ' TODO: FIX.  FORMULA DOESN'T WORK.
      gImageData.PrevY = poiPrev.Y * Screen.TwipsPerPixelY
      gImageData.Dragging = True
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      gImageData.PrevX = X
      gImageData.PrevY = Y
      If Button = vbLeftButton Then
            gImageData.Dragging = True
      End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If gImageData.Dragging Then
            Image1.Move Image1.Left + X - gImageData.PrevX, Image1.Top + Y - gImageData.PrevY, _
                  Image1.Width, Image1.Height
            If X <> gImageData.PrevX Or Y <> gImageData.PrevY Then gImageData.Moved = True
      End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' Mouse button lifted?  Stop the drag!
      gImageData.Dragging = False
      
      If Not gImageData.Moved And Not gImageData.Zoomed And Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            frmMain.BrowserExecuteNext
      ElseIf Not gImageData.Moved And Not gImageData.Zoomed And Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            frmMain.BrowserExecutePrev
      End If
      
      gImageData.Moved = False
      gImageData.Zoomed = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
      SetWindowLong hWnd, GWL_WNDPROC, gpOldfrmFullScreenProc
      gpOldfrmFullScreenProc = 0
      
      With Image1
            frmMain.Image1.Move .Left, .Top, .Width, .Height
            frmMain.Image1.Picture = .Picture
      End With
      Set gImageData.OutPic = frmMain.Image1
      gfFullScreenMode = False
      frmMain.Show
End Sub
