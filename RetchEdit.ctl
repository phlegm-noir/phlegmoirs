VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.3#0"; "VBCCR17.OCX"
Begin VB.UserControl RechEdit 
   Alignable       =   -1  'True
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   8070
   Begin VBCCR17.RichTextBox TextBox 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   11880
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HideSelection   =   0   'False
      MultiLine       =   -1  'True
      ScrollBars      =   2
      SelectionBar    =   -1  'True
      TextMode        =   1
   End
End
Attribute VB_Name = "RechEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INITIAL_WIDTH = 8070
Const TOP_MARGIN = 100

Public Sub Resize()
      UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
      ' Debug.Print "Editor_Init (w, sw, h, sh): " & Width & ", " & ScaleWidth & ", " & Height & ", " & ScaleHeight
End Sub

Private Sub UserControl_Resize()
      ' Debug.Print "Editor_Resize (w, sw, h, sh): " & Width & ", " & ScaleWidth & ", " & Height & ", " & ScaleHeight
      
      TextBox.Move 0, TOP_MARGIN, ScaleWidth, ScaleHeight - TOP_MARGIN
End Sub
