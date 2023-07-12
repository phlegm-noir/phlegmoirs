VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About phlegmoirs"
   ClientHeight    =   1590
   ClientLeft      =   2340
   ClientTop       =   1755
   ClientWidth     =   4950
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1097.446
   ScaleMode       =   0  'User
   ScaleWidth      =   4648.306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0CCA
      ScaleHeight     =   331.934
      ScaleMode       =   0  'User
      ScaleWidth      =   331.934
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   585
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label lblWhatever 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000005&
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   0
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000005&
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   2
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit

Private Sub cmdOK_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      Me.Caption = "About " & App.title
      lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      lblTitle.Caption = App.title
End Sub
