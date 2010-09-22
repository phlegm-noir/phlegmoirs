VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "ImgEdit.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "phlegmoirs - UNREGISTERED"
   ClientHeight    =   9690
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10275
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6688.211
   ScaleMode       =   0  'User
   ScaleWidth      =   9648.756
   ShowInTaskbar   =   0   'False
   Begin ImgeditLibCtl.ImgEdit imgEdit 
      Height          =   6495
      Left            =   4320
      TabIndex        =   12
      Top             =   2880
      Width           =   5775
      _Version        =   131074
      _ExtentX        =   10186
      _ExtentY        =   11456
      _StockProps     =   96
      BorderStyle     =   1
      Image           =   "F:\Dox\My Pictures\oizo.gif"
      ImageControl    =   "ImgEdit1"
      UndoBufferSize  =   56558592
      OcrZoneVisibility=   -3516
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3516
      sReserved2      =   -3516
      lReserved1      =   1241728
      lReserved2      =   1241728
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
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
      Left            =   240
      TabIndex        =   5
      Top             =   7680
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
      Picture         =   "About.frx":0000
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
      Height          =   2385
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   3300
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
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
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
      Me.Caption = "About " & App.title
      lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      lblTitle.Caption = App.title
      
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
     
     imgEdit.Display
End Sub


Private Sub optRegistered_Click()
      lblRed.Visible = False
      lblGreen.Visible = True
End Sub

Private Sub optUnregistered_Click()
      lblRed.Visible = True
      lblGreen.Visible = False
End Sub


Private Sub timReg_Timer()
      miRedMessage = miRedMessage + 1
      If miRedMessage = UBound(msRedReg) Then miRedMessage = 0
      lblRed = msRedReg(miRedMessage)
End Sub
