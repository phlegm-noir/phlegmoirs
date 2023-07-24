VERSION 5.00
Begin VB.UserControl PhlegmoFoto 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image ImageOne 
      Appearance      =   0  'Flat
      Height          =   20235
      Left            =   240
      Picture         =   "PhlegmoFoto.ctx":0000
      Top             =   240
      Width           =   16200
   End
End
Attribute VB_Name = "PhlegmoFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INITIAL_WIDTH = 8070
Const TOP_MARGIN = 0
Const BOTTOM_MARGIN = 15

Private mlDefaultHeight As Long
Private mlDefaultWidth As Long
Private mfPrevX As Single
Private mfPrevY As Single
Private mbDragging As Boolean
Private mbMoved As Boolean
Private mbZoomed As Boolean

Public Sub Resize()
      UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
      DebugLog "Foto_Init (w, sw, h, sh): " & Width & ", " & ScaleWidth & ", " & Height & ", " & ScaleHeight, 0
End Sub

Private Sub UserControl_Resize()
      ImageOne.Move 0, TOP_MARGIN, ScaleWidth, ScaleHeight - TOP_MARGIN - BOTTOM_MARGIN
End Sub

