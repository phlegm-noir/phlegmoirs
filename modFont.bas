Attribute VB_Name = "modFont"
Option Explicit
Option Compare Binary

Const CFE_BOLD As Long = &H1
Const CFE_ITALIC As Long = &H2
Const CFE_STRIKEOUT As Long = &H8
Const CFE_UNDERLINE As Long = &H4

Const CFM_FACE As Long = &H20000000
Const CFM_SIZE As Long = &H80000000
Const CFM_CHARSET As Long = &H8000000
Const CFM_BOLD As Long = &H1
Const CFM_COLOR As Long = &H40000000
Const CFM_ITALIC As Long = &H2
Const CFM_LINK As Long = &H20
Const CFM_OFFSET As Long = &H10000000
Const CFM_STRIKEOUT As Long = &H8
Const CFM_UNDERLINE As Long = &H4
Const CFM_WEIGHT As Long = &H400000

Const TWIPS_PER_POINT As Integer = 20 ' "twenty per point" is literally in the name twip

Function AllowedSizes() As Variant()
      AllowedSizes = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 60, 72)
End Function

Public Function GetNextFontSize(ByVal iFontSize As Integer) As Integer
      Dim iIndex As Integer
      iIndex = LBound(AllowedSizes())
      While iIndex < UBound(AllowedSizes()) And AllowedSizes()(iIndex) <= iFontSize
            iIndex = iIndex + 1
      Wend
      GetNextFontSize = AllowedSizes()(iIndex)
End Function

Public Function GetPrevFontSize(ByVal iFontSize As Integer) As Integer
      Dim iIndex As Integer
      iIndex = UBound(AllowedSizes())
      While iIndex > LBound(AllowedSizes()) And AllowedSizes()(iIndex) >= iFontSize
            iIndex = iIndex - 1
      Wend
      GetPrevFontSize = AllowedSizes()(iIndex)
End Function

Public Function GetRealFontName(ByVal lEditorHwnd As Long) As String
      Dim char2 As CHARFORMAT2
      Dim lRetVal As Long
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_FACE
      char2.dwEffects = 0
      lRetVal = SendMessage(lEditorHwnd, EM_GETCHARFORMAT, ByVal 0, char2)
      
      ' Make it 16-bit characters, and trim the fat.
      GetRealFontName = CstringToVBstring(StrConv(char2.szFaceName, vbUnicode))
End Function

Public Function GetRealFontSize(ByVal lEditorHwnd As Long) As Integer
      Dim char2 As CHARFORMAT2
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_SIZE '+ CFM_FACE
      char2.dwEffects = 0
      SendMessage lEditorHwnd, EM_GETCHARFORMAT, ByVal SCF_SELECTION, char2
      GetRealFontSize = char2.yHeight / TWIPS_PER_POINT
End Function

Public Function SetRealFontSize(ByVal lEditorHwnd As Long, ByVal sNewSize As Single) As Integer
      Dim char2 As CHARFORMAT2
      
      char2.cbSize = LenB(char2)
      char2.dwMask = CFM_SIZE
      char2.dwEffects = 0
      char2.yHeight = sNewSize * TWIPS_PER_POINT
      SendMessage lEditorHwnd, EM_SETCHARFORMAT, ByVal SCF_ALL, char2
      SetRealFontSize = char2.yHeight / TWIPS_PER_POINT  ' return this, see if it's moved or something weird like that.
End Function

Public Function SetRealStdFont(ByVal lEditorHwnd As Long, ByRef fnt As StdFont, _
      Optional lTextColor As Long = vbWindowText) As Long

      Dim char2 As CHARFORMAT2
      Dim sFontName As String
      Dim bDyn() As Byte, i As Integer
      
      With char2
            .cbSize = LenB(char2)
            ' Tell it which CHARFORMAT2 properties carry relevant data:
            .dwMask = CFM_SIZE + CFM_FACE + CFM_CHARSET + CFM_BOLD + _
                  CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE + CFM_WEIGHT
            
            .dwMask = .dwMask + CFM_COLOR
            .crTextColor = TranslateColor(lTextColor)
            
            If fnt.Bold Then .dwEffects = .dwEffects + CFE_BOLD
            If fnt.Italic Then .dwEffects = .dwEffects + CFE_ITALIC
            If fnt.Underline Then .dwEffects = .dwEffects + CFE_UNDERLINE
            If fnt.Strikethrough Then .dwEffects = .dwEffects + CFE_STRIKEOUT
            .yHeight = fnt.Size * TWIPS_PER_POINT
            .bCharSet = fnt.Charset
            .wWeight = fnt.Weight
            ' the font name takes some string manipulation...
            bDyn = StrConv(fnt.Name & Chr(0), vbFromUnicode)
            For i = LBound(bDyn) To UBound(bDyn)
                  .szFaceName(i) = bDyn(i)
            Next i
      End With
      SetRealStdFont = SendMessage(lEditorHwnd, EM_SETCHARFORMAT, ByVal SCF_ALL, char2)
End Function

Public Function GetRealStdFont(ByVal lEditorHwnd As Long, Optional ByRef lTextColor As Long) As StdFont
      ' OK, I put in a byref value to pass on the text color, which is not included in the StdFont type.
      ' The function returns a StdFont containing the rest of the font data.
      
      Dim char2 As CHARFORMAT2
      Dim lRetVal As Long
      Dim fntNew As New StdFont
      
      char2.cbSize = LenB(char2)
      ' Tell it which CHARFORMAT2 properties carry relevant data:
      char2.dwMask = CFM_SIZE + CFM_FACE + CFM_BOLD + CFM_COLOR + _
                  CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE
            ' I took out cfm_charset and cfm_weight, because they are set automatically
      lRetVal = SendMessage(lEditorHwnd, EM_GETCHARFORMAT, ByVal 0, char2)
      
      If lRetVal <> 0 Then
            With fntNew
                  .Size = char2.yHeight / TWIPS_PER_POINT
                  .Name = CstringToVBstring(StrConv(char2.szFaceName, vbUnicode))
                  .Bold = char2.dwEffects And CFE_BOLD
                  .Italic = char2.dwEffects And CFE_ITALIC
                  .Strikethrough = char2.dwEffects And CFE_STRIKEOUT
                  .Underline = char2.dwEffects And CFE_UNDERLINE
            End With
            lTextColor = char2.crTextColor
            Set GetRealStdFont = fntNew
      Else
            Set GetRealStdFont = Nothing
      End If
End Function

