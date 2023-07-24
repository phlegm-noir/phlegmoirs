VERSION 5.00
Begin VB.UserControl PhlegmoFinder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   600
   ScaleWidth      =   4095
   Begin VB.CommandButton btnReplace 
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
      Height          =   300
      Left            =   2610
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFinder.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Replace (Ctrl+R)"
      Top             =   270
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CheckBox chkFindOptions 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "More search options (Alt+period)"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnFindPrev 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1530
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFinder.ctx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find Previous (Shift+F3)"
      Top             =   270
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnFindNext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   450
      MaskColor       =   &H00FFFFFF&
      Picture         =   "PhlegmoFinder.ctx":0684
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Find Next (F3)"
      Top             =   270
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtReplace 
      Height          =   288
      Left            =   450
      MaxLength       =   50
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      ToolTipText     =   "Replace"
      Top             =   290
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton btnCloseFind 
      Appearance      =   0  'Flat
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
      Left            =   3810
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Close Find Dialog"
      Top             =   0
      Width           =   175
   End
   Begin VB.TextBox txtFind 
      Height          =   288
      Left            =   450
      MaxLength       =   50
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      ToolTipText     =   "Search within file (Ctrl+F)"
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblFind 
      Alignment       =   2  'Center
      Caption         =   "Find:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   465
   End
End
Attribute VB_Name = "PhlegmoFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Closing()

Private Sub btnCloseFind_Click()
      RaiseEvent Closing
End Sub
