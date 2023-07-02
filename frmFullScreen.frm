VERSION 5.00
Begin VB.Form frmFullScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmFullScreen.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   8475
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFullScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   8190
      TabIndex        =   0
      Top             =   0
      Width           =   8190
      Begin VB.CommandButton btnClose 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFullScreen.frx":039C
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Exit Fullscreen Mode (F11 or Esc)"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   400
      End
      Begin VB.Label lblFileNameZoom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\pics\big ol zebra.jpg   79%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5595
         TabIndex        =   2
         Top             =   0
         Width           =   2550
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   0
         MousePointer    =   15  'Size All
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3600
      End
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit

Private Sub btnClose_Click()
      Unload frmFullScreen
End Sub

Private Sub CopyDimensions()
      With frmMain.Image1
            Image1.Move .Left, .Top, .Width, .Height
      End With
End Sub

Private Sub Form_Load()
      gbFullScreenMode = True
      Set gtImageData.OutPic = Image1
      Set gtImageData.SurroundingBox = picFullScreen
      
      picFullScreen.Move 0, 0, Screen.Width, Screen.Height
      Image1.Picture = frmMain.Image1.Picture
      If geImageSizingMode = eImageSizingMode.AlwaysFit Then
            frmMain.ImageZoomFit gtImageData.OutPic.Picture, frmMain.agEditor.tag
      Else
            CopyDimensions
      End If
      
      lblFileNameZoom.Left = picFullScreen.Width - lblFileNameZoom.Width
      lblFileNameZoom = frmMain.Caption & "  "
      
      glOldfrmFullScreenProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf TrackMouseWheel)
End Sub

Private Sub Form_Resize()
      picFullScreen.Move 0, 0, ScaleWidth, ScaleHeight
      lblFileNameZoom.Left = picFullScreen.Width - lblFileNameZoom.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
      SetWindowLong hwnd, GWL_WNDPROC, glOldfrmFullScreenProc
      glOldfrmFullScreenProc = 0
      
      With Image1
            frmMain.Image1.Move .Left, .Top, .Width, .Height
            frmMain.Image1.Picture = .Picture
      End With
      Set gtImageData.OutPic = frmMain.Image1
      Set gtImageData.SurroundingBox = frmMain.picEditor
      If geImageSizingMode = eImageSizingMode.AlwaysFit Then
            frmMain.ImageZoomFit gtImageData.OutPic.Picture, frmMain.agEditor.tag
      End If
      gbFullScreenMode = False
      frmMain.Show
End Sub

Private Sub Image1_DblClick()
      ' This needs to (effectively) call an Image1_mousedown... but with what parameters???
      Dim tPrev As POINTAPI
      
      GetCursorPos tPrev
      
      gtImageData.PrevX = tPrev.X * Screen.TwipsPerPixelX
      gtImageData.PrevY = tPrev.Y * Screen.TwipsPerPixelY
      gtImageData.Dragging = True
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      gtImageData.PrevX = X
      gtImageData.PrevY = Y
      If Button = vbLeftButton Then
            gtImageData.Dragging = True
      End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If gtImageData.Dragging Then
            Image1.Move Image1.Left + X - gtImageData.PrevX, Image1.Top + Y - gtImageData.PrevY, _
                  Image1.Width, Image1.Height
            If X <> gtImageData.PrevX Or Y <> gtImageData.PrevY Then gtImageData.Moved = True
      End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ' Mouse button lifted?  Stop the drag!
      gtImageData.Dragging = False
      
      picfullscreen_MouseUp Button, Shift, X, Y
End Sub

Private Sub picFullScreen_KeyDown(KeyCode As Integer, Shift As Integer)
      With frmMain.sliZoom
            Select Case KeyCode
                  Case 107, 187 ' "+" and Keypad "+"
                        If Shift = 0 Then
                              frmMain.ImageZoomIn .SmallChange
                              RedoCaption
                        ElseIf Shift = vbCtrlMask Then
                              frmMain.ImageZoomIn .LargeChange
                              RedoCaption
                        End If
                  Case 109, 189 ' "-" and Keypad "-"
                        If Shift = 0 Then
                              frmMain.ImageZoomOut .SmallChange
                              RedoCaption
                        ElseIf Shift = vbCtrlMask Then
                              frmMain.ImageZoomOut .LargeChange
                              RedoCaption
                        End If
                  Case vbKey0, 106 ' 0 and Keypad "*" -- reset position and size.
                        .Value = 100
                        RedoCaption
                        Image1.Move 0, 0, gtImageData.DefaultWidth, gtImageData.DefaultHeight
                  Case 103, 55   ' 7 and Keypad 7
                        .Value = .Value / 2
                        RedoCaption
                  Case 104, 56   ' 8 and Keypad 8
                        .Value = .Value * 2
                        RedoCaption
                  Case vbKeyDown
                        Image1.Top = Image1.Top + MOVE_INCREMENT
                  Case vbKeyUp
                        Image1.Top = Image1.Top - MOVE_INCREMENT
                  Case vbKeyLeft
                        Image1.Left = Image1.Left - MOVE_INCREMENT
                  Case vbKeyRight
                        Image1.Left = Image1.Left + MOVE_INCREMENT
                  
                  Case vbKeyHome
                        Image1.Top = 0
                  Case vbKeyEnd
                        Image1.Top = picFullScreen.Height - Image1.Height
                        
                  Case vbKeyPageUp
                        If Image1.Top < -picFullScreen.Height Then
                              Image1.Top = Image1.Top + picFullScreen.Height
                        ElseIf Image1.Top < 0 Then
                              Image1.Top = 0
                        End If
                        
                  Case vbKeyPageDown
                        If Image1.Top + Image1.Height > picFullScreen.Height * 2 Then
                              Image1.Top = Image1.Top - picFullScreen.Height
                        ElseIf Image1.Top + Image1.Height > picFullScreen.Height Then
                              Image1.Top = picFullScreen.Height - Image1.Height
                        End If
                        
                  Case vbKeySpace, vbKeyN, 221   ' Right Bracket "]"
                        If Shift = 0 Then frmMain.BrowserExecuteNext
                  Case vbKeyBack, vbKeyP, 219   ' Left Bracket "["
                        If Shift = 0 Then frmMain.BrowserExecuteNext True
                        
                  Case vbKeyF11, vbKeyEscape
                        If Shift = 0 Then Unload frmFullScreen
            End Select
      End With

End Sub

Private Sub picfullscreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not gtImageData.Zoomed And Not gtImageData.Moved And Button = vbLeftButton Then
            ' On a left click, we'll go to the next picture.  We spare no expense on ease of use.
            frmMain.BrowserExecuteNext
      ElseIf Not gtImageData.Zoomed And Not gtImageData.Moved And Button = vbRightButton Then
            ' On a right click, we go to the previous picture.
            ' Essentially, it'll means we don't need the toolbar open for picture manipulation.
            frmMain.BrowserExecuteNext True
      ElseIf Not gtImageData.Moved And Not gtImageData.Zoomed And Button = vbMiddleButton Then
            Unload frmFullScreen
      End If
      
      gtImageData.Zoomed = False
      gtImageData.Moved = False
End Sub

Private Sub RedoCaption()
      lblFileNameZoom = frmMain.agEditor.tag & " (" & frmMain.sliZoom.Value & "%)  "
End Sub

