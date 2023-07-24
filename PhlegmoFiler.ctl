VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PhlegmoFiler 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ScaleHeight     =   5670
   ScaleWidth      =   4140
   Begin VB.CommandButton btnScrollToTop 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2730
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Scroll To Top"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   390
   End
   Begin VB.CommandButton btnSyncContents 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2340
      MaskColor       =   &H80000001&
      Picture         =   "PhlegmoFiler.ctx":04B6
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Jump to the directory containing your open file... (Ctrl+F5)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnDeleteSelected 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1950
      MaskColor       =   &H00000000&
      Picture         =   "PhlegmoFiler.ctx":07F8
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Delete File (Del)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnRefresh 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1560
      MaskColor       =   &H80000005&
      Picture         =   "PhlegmoFiler.ctx":0B3A
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Refresh Files (F5)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnSort 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1170
      MaskColor       =   &H80000005&
      Picture         =   "PhlegmoFiler.ctx":0E7C
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Reverse the sort order (Ctrl+H)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnFolderUp 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   780
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":11BE
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Go up a directory (Left arrow key or Ctrl+F6)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnPathForward 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   390
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":1500
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Go forward a directory (Alt+Right)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.CommandButton btnPathBack 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFiler.ctx":197A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Go back a directory (Alt+Left)"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.ComboBox cboPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "PhlegmoFiler.ctx":1DF4
      Left            =   0
      List            =   "PhlegmoFiler.ctx":1DF6
      TabIndex        =   0
      Text            =   "c:\windows\system32"
      ToolTipText     =   "Type a directory into here, or select one below.  You can even specify a file extension.  Example:   c:\windows\*.dll"
      Top             =   0
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ilsFileIcons2 
      Left            =   2475
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8388863
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":1DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":214A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":249C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":27EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":2B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":2E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":31E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3536
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3888
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":3F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":427E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":45D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhlegmoFiler.ctx":4922
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwBrowser 
      Height          =   4335
      Left            =   0
      TabIndex        =   10
      Tag             =   "c:\test\"
      Top             =   780
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7646
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilsFileIcons2"
      SmallIcons      =   "ilsFileIcons2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "PhlegmoFiler.ctx":4C74
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Object.Tag             =   "0"
         Text            =   "[N]ame"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "Type"
         Text            =   "[T]ype"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Size"
         Text            =   "Si[z]e"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Modified"
         Text            =   "[M]odified"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "SortSize"
         Object.Tag             =   "Adding 0s (000054) to make #s text-sortable"
         Text            =   "SortSize"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblDivider 
      Height          =   25005
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "PhlegmoFiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INITIAL_WIDTH = 4140
Const RIGHT_MARGIN = 105
Const BOTTOM_MARGIN = 0
Const MIN_WIDTH = 995
Private mlInitialPointerX As Long
Private mlPrevPointerX As Long
Private mlMaxWidth As Long
Private miInitializings As Integer

Public Event ResizeHorizontal(ByVal lWidth As Long)

' Sent less frequently, use this to hard-limit the form's min-width via API calls
Public Event SeriousResize(ByVal lWidth As Long)

Public Sub SetMaxWidth(ByVal lMaxWidth As Long)
      mlMaxWidth = lMaxWidth
End Sub

Public Function ActualWidth() As Long
      ActualWidth = Width
End Function

Private Sub lblDivider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      If lblDivider.MousePointer = vbSizeWE And lblDivider.Tag = "" Then
            lblDivider.Tag = "Resizing"
            mlInitialPointerX = GetCursorPosX() * Screen.TwipsPerPixelX
            mlPrevPointerX = mlInitialPointerX
      End If
End Sub

Private Sub lblDivider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim lDeltaX As Long
      
      lDeltaX = GetCursorPosX() * Screen.TwipsPerPixelX - mlPrevPointerX
      If Abs(lDeltaX) > 15 And (lblDivider.Left + lDeltaX > MIN_WIDTH) And (lblDivider.Left + lDeltaX < mlMaxWidth) _
                  And lblDivider.MousePointer = vbSizeWE And lblDivider.Tag = "Resizing" Then
            lblDivider.Tag = "Busy"
            
            mlPrevPointerX = mlPrevPointerX + lDeltaX
            RearrangeControls lblDivider.Left + lDeltaX
            If lblDivider.Tag = "Busy" Then lblDivider.Tag = "Resizing"
      Else
            lblDivider.MousePointer = vbSizeWE
      End If
End Sub

Private Sub lblDivider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If lblDivider.MousePointer = vbSizeWE Then
            ' SaveSettingsToRegistry
            RaiseEvent SeriousResize(Width)
      End If
      lblDivider.MousePointer = 0
      lblDivider.Tag = ""
      Debug.Print "WE ARE NO LONGER RESIZING"
End Sub

Private Sub RearrangeControls(Optional ByVal lSupposedWidth As Long = -1)
      Dim lRightWall As Long, lToolbarRightEdge As Long
      
      If lSupposedWidth = -1 Then lSupposedWidth = Width - RIGHT_MARGIN
      Bound lSupposedWidth, MIN_WIDTH, mlMaxWidth
      
      lvwBrowser.Width = lSupposedWidth
      cboPath.Width = lSupposedWidth
      lblDivider.Left = lSupposedWidth
      If miInitializings > 0 Then
            Width = lSupposedWidth + RIGHT_MARGIN
            RaiseEvent ResizeHorizontal(Width)
      End If
      
      lToolbarRightEdge = btnSyncContents.Left + btnSyncContents.Width
      lRightWall = lvwBrowser.Left + lvwBrowser.Width - btnScrollToTop.Width - 30
      btnScrollToTop.Left = Max(lRightWall, lToolbarRightEdge)
End Sub

Private Sub UserControl_Initialize()
      mlMaxWidth = 10000
      RearrangeControls INITIAL_WIDTH - RIGHT_MARGIN
      miInitializings = miInitializings + 1
End Sub

Private Sub UserControl_Resize()
      lvwBrowser.Height = ScaleHeight - lvwBrowser.Top - BOTTOM_MARGIN
End Sub
