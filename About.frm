VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{35E55124-D2A7-4467-955F-19C1DCB7F1CB}#1.1#0"; "RichEdit.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "phlegmoirs - UNREGISTERED"
   ClientHeight    =   9825
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10410
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6781.39
   ScaleMode       =   0  'User
   ScaleWidth      =   9775.528
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   4200
      TabIndex        =   13
      Top             =   3360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
      _Version        =   393217
      TextRTF         =   $"About.frx":0000
   End
   Begin RECtl.RichEdit RichEdit1 
      Height          =   3015
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -9999997
      HideSelection   =   -1  'True
   End
   Begin VB.Timer timReg 
      Interval        =   2500
      Left            =   240
      Top             =   7080
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   3720
      Width           =   3855
      Begin VB.Label lblGreen 
         Alignment       =   2  'Center
         Caption         =   "Thank you for registering.  Your support helps us grow."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1815
         Left            =   600
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblRed 
         Alignment       =   2  'Center
         Caption         =   "You have thirty days to enjoy this software free of charge."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2775
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame fraRegistration 
      Caption         =   "Product Registration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
      Begin VB.OptionButton optUnregistered 
         Caption         =   "Unregistered"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optRegistered 
         Caption         =   "Registered"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "About.frx":008B
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   2640
      Width           =   1260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   1080
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "File Properties"
      TabPicture(0)   =   "About.frx":0955
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProperties1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraProperties2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraProperties2 
         Caption         =   "2"
         Height          =   1215
         Left            =   240
         TabIndex        =   29
         Top             =   4320
         Width           =   4575
      End
      Begin VB.Frame fraProperties1 
         Caption         =   "1"
         Height          =   3735
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   4575
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   480
            TabIndex        =   28
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "Accessed:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   480
            TabIndex        =   27
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "Modified:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   480
            TabIndex        =   26
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "Created:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   480
            TabIndex        =   25
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "Size:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   480
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "File Type:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblPropTitle 
            Alignment       =   1  'Right Justify
            Caption         =   "File Name:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   21
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   19
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   18
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2040
            TabIndex        =   17
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label lblPropValue 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2040
            TabIndex        =   16
            Top             =   2160
            Width           =   1935
         End
      End
   End
   Begin VB.Label lblYouJustPressed 
      Caption         =   "You've just pressed this keycode:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   2
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miRedMessage As Integer
Dim msRedReg(0 To 14) As String

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      lblYouJustPressed = "You've just pressed this KeyCode:  " & KeyCode
End Sub

Private Sub Form_Load()
      Me.Caption = "About " & App.Title
      lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      lblTitle.Caption = App.Title
      
      msRedReg(0) = "You have thirty days to enjoy this software free of charge."
      msRedReg(1) = "At the end of the thirty days, please kill yourself."
      msRedReg(2) = "Ha ha!  Just kidding!  We're not really like that, honest."
      msRedReg(3) = "Let's see YOU write something in large red letters and "
      msRedReg(4) = "not turn it into a suicide referral."
      msRedReg(5) = "But like I was trying to say before,"
      msRedReg(6) = "Amateur software developers have to eat, too."
      msRedReg(7) = "Babies.  We have to eat a lot of babies."
      msRedReg(8) = "You could very well be our next homicide victim."
      msRedReg(9) = "If you're a baby."
      msRedReg(10) = "Hey, why not register some software once in a while, asshole?"
      msRedReg(11) = "But personally, I like children."
      msRedReg(12) = "It's been a dream of mine to someday become one."
      msRedReg(13) = "These messages get nicer when you're registered."
      msRedReg(14) = "...or do they?"
     ' msRedReg() = ""
End Sub



Private Sub optRegistered_Click()
      lblRed.Visible = False
      lblGreen.Visible = True
End Sub

Private Sub optUnregistered_Click()
      lblRed.Visible = True
      lblGreen.Visible = False
End Sub

Private Sub RichEdit1_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long, LinkRange As RECtl.Range)
      Debug.Print "OCX: " & RichEdit1.RangeFromPoint(x, y).EndPos
      
      Dim poiTemp As POINTAPI, lRetVal As Long
      poiTemp.x = x / Screen.TwipsPerPixelX
      poiTemp.y = y / Screen.TwipsPerPixelY
      lRetVal = SendMessage(RichEdit1.hwnd, EM_CHARFROMPOS, ByVal 0, poiTemp)
      Debug.Print "API: " & lRetVal & " " & Timer
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      Debug.Print "OCX: " & RichEdit1.RangeFromPoint(x, y).EndPos
      'richtextbox1.
      Dim poiTemp As POINTAPI, lRetVal As Long
      poiTemp.x = x / Screen.TwipsPerPixelX
      poiTemp.y = y / Screen.TwipsPerPixelY
      lRetVal = SendMessage(RichEdit1.hwnd, EM_CHARFROMPOS, ByVal 0, poiTemp)
      Debug.Print "API: " & lRetVal & " " & Timer
End Sub

Private Sub timReg_Timer()
      miRedMessage = miRedMessage + 1
      If miRedMessage = UBound(msRedReg) Then miRedMessage = 0
      lblRed = msRedReg(miRedMessage)
End Sub
