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
      Dim tChar2 As CHARFORMAT2
      Dim lRetVal As Long
      
      tChar2.cbSize = LenB(tChar2)
      tChar2.dwMask = CFM_FACE
      tChar2.dwEffects = 0
      lRetVal = SendMessage(lEditorHwnd, EM_GETCHARFORMAT, ByVal 0, tChar2)
      
      ' Make it 16-bit characters, and trim the fat.
      GetRealFontName = CstringToVBstring(StrConv(tChar2.szFaceName, vbUnicode))
End Function

Public Function GetRealFontSize(ByVal lEditorHwnd As Long) As Integer
      Dim tChar2 As CHARFORMAT2
      
      tChar2.cbSize = LenB(tChar2)
      tChar2.dwMask = CFM_SIZE '+ CFM_FACE
      tChar2.dwEffects = 0
      SendMessage lEditorHwnd, EM_GETCHARFORMAT, ByVal SCF_SELECTION, tChar2
      GetRealFontSize = tChar2.yHeight / TWIPS_PER_POINT
End Function

Public Function GetRealStdFont(ByVal lEditorHwnd As Long, Optional ByRef lTextColor As Long) As StdFont
      ' OK, I put in a byref value to pass on the text color, which is not included in the StdFont type.
      ' The function returns a StdFont containing the rest of the font data.
      
      Dim tChar2 As CHARFORMAT2
      Dim lRetVal As Long
      Dim objNewFont As New StdFont
      
      tChar2.cbSize = LenB(tChar2)
      ' Tell it which CHARFORMAT2 properties carry relevant data:
      tChar2.dwMask = CFM_SIZE + CFM_FACE + CFM_BOLD + CFM_COLOR + _
            CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE
      
      ' I took out cfm_charset and cfm_weight, because they are set automatically
      lRetVal = SendMessage(lEditorHwnd, EM_GETCHARFORMAT, ByVal 0, tChar2)
      
      If lRetVal <> 0 Then
            With objNewFont
                  .Size = tChar2.yHeight / TWIPS_PER_POINT
                  .Name = CstringToVBstring(StrConv(tChar2.szFaceName, vbUnicode))
                  .Bold = tChar2.dwEffects And CFE_BOLD
                  .Italic = tChar2.dwEffects And CFE_ITALIC
                  .Strikethrough = tChar2.dwEffects And CFE_STRIKEOUT
                  .Underline = tChar2.dwEffects And CFE_UNDERLINE
            End With
            lTextColor = tChar2.crTextColor
            Set GetRealStdFont = objNewFont
      Else
            Set GetRealStdFont = Nothing
      End If
End Function

Public Function SetRealFontSize(ByVal lEditorHwnd As Long, ByVal fNewSize As Single) As Integer
      Dim tChar2 As CHARFORMAT2
      
      tChar2.cbSize = LenB(tChar2)
      tChar2.dwMask = CFM_SIZE
      tChar2.dwEffects = 0
      tChar2.yHeight = fNewSize * TWIPS_PER_POINT
      SendMessage lEditorHwnd, EM_SETCHARFORMAT, ByVal SCF_ALL, tChar2
      SetRealFontSize = tChar2.yHeight / TWIPS_PER_POINT  ' return this, see if it's moved or something weird like that.
End Function

Public Function SetRealStdFont(ByVal lEditorHwnd As Long, ByRef r_objNewFont As StdFont, _
      Optional lTextColor As Long = vbWindowText) As Long

      Dim tChar2 As CHARFORMAT2
      Dim cDyn() As Byte, iIndex As Integer
      
      With tChar2
            .cbSize = LenB(tChar2)
            ' Tell it which CHARFORMAT2 properties carry relevant data:
            .dwMask = CFM_SIZE + CFM_FACE + CFM_CHARSET + CFM_BOLD + _
                  CFM_ITALIC + CFM_STRIKEOUT + CFM_UNDERLINE + CFM_WEIGHT
            
            .dwMask = .dwMask + CFM_COLOR
            .crTextColor = TranslateColor(lTextColor)
            
            If r_objNewFont.Bold Then .dwEffects = .dwEffects + CFE_BOLD
            If r_objNewFont.Italic Then .dwEffects = .dwEffects + CFE_ITALIC
            If r_objNewFont.Underline Then .dwEffects = .dwEffects + CFE_UNDERLINE
            If r_objNewFont.Strikethrough Then .dwEffects = .dwEffects + CFE_STRIKEOUT
            .yHeight = r_objNewFont.Size * TWIPS_PER_POINT
            .bCharSet = r_objNewFont.Charset
            .wWeight = r_objNewFont.Weight
            ' the font name takes some string manipulation...
            cDyn = StrConv(r_objNewFont.Name & Chr(0), vbFromUnicode)
            For iIndex = LBound(cDyn) To UBound(cDyn)
                  .szFaceName(iIndex) = cDyn(iIndex)
            Next iIndex
      End With
      SetRealStdFont = SendMessage(lEditorHwnd, EM_SETCHARFORMAT, ByVal SCF_ALL, tChar2)
End Function

