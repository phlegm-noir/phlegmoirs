VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5976
   ClientLeft      =   2568
   ClientTop       =   1500
   ClientWidth     =   6648
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   6648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSample1 
      Caption         =   "Sample 1"
      Height          =   5268
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   6132
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Move Down"
         Height          =   612
         Left            =   2760
         TabIndex        =   12
         Top             =   3720
         Width           =   732
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Move Up"
         Height          =   612
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   732
      End
      Begin MSComctlLib.ListView lvwToolButtons 
         Height          =   4692
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   2292
         _ExtentX        =   4043
         _ExtentY        =   8276
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   5532
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5532
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3096
      TabIndex        =   0
      Top             =   5532
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMoveUp_Click()
      Dim iIndex As Integer
      Dim sTempTag As String, sTempText As String
      Dim fTempValue As Boolean
            
      With lvwToolButtons
            iIndex = .SelectedItem.Index
            If iIndex = 1 Then Exit Sub
            
            sTempTag = .ListItems(iIndex - 1).Tag
            .ListItems(iIndex - 1).Tag = .SelectedItem.Tag
            .SelectedItem.Tag = sTempTag
            
            fTempValue = .ListItems(iIndex - 1).Checked
            .ListItems(iIndex - 1).Checked = .SelectedItem.Checked
            .SelectedItem.Checked = fTempValue
            
            sTempText = .ListItems(iIndex - 1).Text
            .ListItems(iIndex - 1).Text = .SelectedItem.Text
            .SelectedItem.Text = sTempText
            
            .ListItems(iIndex - 1).Selected = True
      End With
End Sub

Private Sub btnMoveDown_Click()
      Dim iIndex As Integer
      Dim sTempTag As String, sTempText As String
      Dim fTempValue As Boolean
      
      With lvwToolButtons
            iIndex = .SelectedItem.Index
            If iIndex = .ListItems.Count Then Exit Sub
            
            sTempTag = .ListItems(iIndex + 1).Tag
            .ListItems(iIndex + 1).Tag = .SelectedItem.Tag
            .SelectedItem.Tag = sTempTag
            
            fTempValue = .ListItems(iIndex + 1).Checked
            .ListItems(iIndex + 1).Checked = .SelectedItem.Checked
            .SelectedItem.Checked = fTempValue
            
            sTempText = .ListItems(iIndex + 1).Text
            .ListItems(iIndex + 1).Text = .SelectedItem.Text
            .SelectedItem.Text = sTempText
            
            .ListItems(iIndex + 1).Selected = True
      End With
End Sub

Private Sub cmdApply_Click()
    MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
'    'handle ctrl+tab to move to the next tab
'    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
'        i = tbsOptions.SelectedItem.Index
'        If i = tbsOptions.Tabs.Count Then
'            'last tab so we need to wrap to tab 1
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
'        Else
'            'increment the tab
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
'        End If
'    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    lvwToolButtons.ListItems.Add , , "File Browser"
    lvwToolButtons.ListItems.Add , , "Status Bar"
    
    lvwToolButtons.ListItems.Add , , "New File"
    lvwToolButtons.ListItems.Add , , "Save"
    lvwToolButtons.ListItems.Add , , "Font"
    
    lvwToolButtons.ListItems.Add , , "Word Wrap"
    lvwToolButtons.ListItems.Add , , "Read Only"
    lvwToolButtons.ListItems.Add , , "Options"
    
    lvwToolButtons.ListItems.Add , , "Next File"
    lvwToolButtons.ListItems.Add , , "Previous File"
    lvwToolButtons.ListItems.Add , , "Back"
    lvwToolButtons.ListItems.Add , , "Forward"
    
    lvwToolButtons.ListItems.Add , , "Undo"
    lvwToolButtons.ListItems.Add , , "Redo"
    lvwToolButtons.ListItems.Add , , "Cut"
    lvwToolButtons.ListItems.Add , , "Copy"
    lvwToolButtons.ListItems.Add , , "Paste"
    lvwToolButtons.ListItems.Add , , "Select All"
    
    lvwToolButtons.ListItems.Add , , "Zoom"
    lvwToolButtons.ListItems.Add , , "Find"
    
      lvwToolButtons.ListItems(4).CreateDragImage
    
End Sub

