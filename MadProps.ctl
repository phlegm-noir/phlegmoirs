VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.3#0"; "VBCCR17.OCX"
Begin VB.UserControl MadProps 
   BackColor       =   &H80000004&
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   ScaleHeight     =   7725
   ScaleWidth      =   7665
   Begin VBCCR17.TabStrip TabZero 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12303
      BackColor       =   -2147483644
      TabMinWidth     =   100
      InitTabs        =   "MadProps.ctx":0000
   End
End
Attribute VB_Name = "MadProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INITIAL_WIDTH = 8070
Const TOP_MARGIN = 0
Const BOTTOM_MARGIN = 300
Const LEFT_MARGIN = 240
Const RIGHT_MARGIN = 300


Public Sub Resize()
      UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
      DebugLog "MadProps_Init (w, sw, h, sh): " & Width & ", " & ScaleWidth & ", " & Height & ", " & ScaleHeight, 0
End Sub

Private Sub UserControl_Resize()
      TabZero.Move LEFT_MARGIN, TOP_MARGIN, ScaleWidth - LEFT_MARGIN - RIGHT_MARGIN, ScaleHeight - TOP_MARGIN - BOTTOM_MARGIN
End Sub

