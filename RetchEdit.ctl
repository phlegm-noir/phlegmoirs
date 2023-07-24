VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.3#0"; "VBCCR17.OCX"
Begin VB.UserControl RetchEdit 
   Alignable       =   -1  'True
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   8070
   Begin VBCCR17.RichTextBox rtEditor 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   11880
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
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
      TextRTF         =   "RetchEdit.ctx":0000
   End
End
Attribute VB_Name = "RetchEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INITIAL_WIDTH = 8070
Const TOP_MARGIN = 0
Const BOTTOM_MARGIN = 15

Private mbEditorLoading As Boolean
Private miTextEncoding As Integer

Public Sub Resize()
      UserControl_Resize
End Sub

Public Sub SetFont(ByVal oFont As StdFont)
      rtEditor.Font = oFont
End Sub

Public Sub ForceRefresh()
      rtEditor.Refresh
End Sub

Public Sub Save(Optional ByVal sDir As String, Optional ByVal sFileName As String)
      If sDir = "" Then sDir = "D:\workspace-vb6\phlegmoirs_v1\"
      If sFileName = "" Then sFileName = "save_" & Round(999999 * Rnd()) & ".txt"
      rtEditor.SaveFile sDir & sFileName, RtfLoadSaveFormatUnicodeText
      DebugLog "File saved: " & sDir & sFileName, 2
End Sub

Private Function RoundExcept1(ByVal fSize As Single) As Integer
      ' I don't like how 1.5 and 2.25 both show up as 2
      RoundExcept1 = Ternary(rtEditor.Font.Size > 1.5, Round(rtEditor.Font.Size), 1)
End Function

' iIndexStep = 1 for next, -1 for previous
Public Function NextFontSize(Optional ByVal iIndexStep As Integer = 1) As Integer
      Dim iIndex As Integer, iInitialSize As Integer, iFirstIndex As Integer, iLastIndex As Integer
      Dim sCache As String
      sCache = rtEditor.Text
      iInitialSize = RoundExcept1(rtEditor.Font.Size)
      
      If iIndexStep > 0 Then
            iFirstIndex = LBound(PreferredFontSizes())
            iLastIndex = UBound(PreferredFontSizes())
      Else
            iFirstIndex = UBound(PreferredFontSizes())
            iLastIndex = LBound(PreferredFontSizes())
      End If
      
      For iIndex = iFirstIndex To iLastIndex Step iIndexStep
            If (PreferredFontSizes()(iIndex) - iInitialSize) * iIndexStep > 0 Then
                  rtEditor.Font.Size = PreferredFontSizes()(iIndex)
                  DebugLog "New actual font size: " & rtEditor.Font.Size
                  
                  ' We have our list of font sizes that we want the editor to use...
                  ' but ultimately the editor decides if it liked the size or not.
                  ' If there was no change, we have to keep trying more sizes.
                  If RoundExcept1(rtEditor.Font.Size) <> iInitialSize Then
                        Exit For
                  End If
            End If
      Next iIndex
      rtEditor.Text = sCache ' workaround VBCCR bug; editor loses some unicode on font change
      NextFontSize = RoundExcept1(rtEditor.Font.Size)
End Function

Private Function PreferredFontSizes() As Variant()
      PreferredFontSizes = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 60, 72)
End Function

Public Function PrevFontSize() As Integer
      PrevFontSize = NextFontSize(-1)
End Function

Public Sub OpenFontDialog(ByRef sName As String, ByRef iSize As Integer)
      Dim NewFont As StdFont, sCache As String
      sCache = rtEditor.Text
      
      With New CommonDialog
            .HookEvents = True
            .flags = CdlCFScreenFonts Or CdlCFApply Or CdlCFLimitSize Or CdlCCFullOpen
            .FontName = rtEditor.Font.Name
            .FontBold = rtEditor.Font.Bold
            .FontItalic = rtEditor.Font.Italic
            .FontSize = rtEditor.Font.Size
            .FontStrikethru = rtEditor.Font.Strikethrough
            .FontUnderline = rtEditor.Font.Underline
            .Color = rtEditor.Font.Underline
            .FontCharset = rtEditor.Font.Charset
            .Min = 1
            .Max = 72
            If .ShowFont = True Then
                  Set NewFont = New StdFont
                  NewFont.Bold = NullOr(.FontBold, False)
                  NewFont.Charset = .FontCharset
                  NewFont.Italic = .FontItalic
                  NewFont.Name = .FontName
                  NewFont.Size = .FontSize
                  NewFont.Strikethrough = .FontStrikethru
                  NewFont.Underline = .FontUnderline
                  NewFont.Weight = .FontWeight
                  rtEditor.Font = NewFont
            End If
      End With
      rtEditor.Text = sCache ' workaround VBCCR bug; editor loses some unicode on font change
      
      sName = rtEditor.Font.Name
      iSize = RoundExcept1(rtEditor.Font.Size)
      DebugLog "New actual font size: " & rtEditor.Font.Size & "; NewFont.Size: " & iSize
      
End Sub

Private Sub UserControl_Initialize()
      DebugLog "Editor_Init (w, sw, h, sh): " & Width & ", " & ScaleWidth & ", " & Height & ", " & ScaleHeight, 0
      
'      rtEditor.LoadFile "D:\workspace-vb6\phlegmoirs_v1\utf8wbom.txt", RtfLoadSaveFormatText
      rtEditor.LoadFile "P:\workspace2000\phlegmoirs_v1\name1.txt", RtfLoadSaveFormatText
End Sub

Private Sub UserControl_Resize()
      rtEditor.Move 0, TOP_MARGIN, ScaleWidth, ScaleHeight - TOP_MARGIN - BOTTOM_MARGIN
End Sub
