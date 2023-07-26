Attribute VB_Name = "modUtils"
Option Explicit
Option Compare Binary

Public Const MAX_DWORD As Currency = 4294967295@ ' double word = 8 bytes = 16 ^ 8 - 1

Public Function Max(ByVal v1 As Variant, ByVal v2 As Variant) As Variant
      If (v1 > v2) Then
            Max = v1
      Else
            Max = v2
      End If
End Function

Public Function Min(ByVal v1 As Variant, ByVal v2 As Variant) As Variant
      If (v1 > v2) Then
            Min = v2
      Else
            Min = v1
      End If
End Function

' Cuts off value (in place) if it is beneath the floor or above the ceiling
Public Sub Bound(ByRef rvValue As Variant, ByVal vFloor As Variant, ByVal vCeiling As Variant)
      If rvValue > vCeiling Then rvValue = vCeiling
      If rvValue < vFloor Then rvValue = vFloor
End Sub

Public Function Ternary(ByVal bCondition As Boolean, ByVal vIfTrue As Variant, ByVal vIfFalse As Variant) As Variant
      If bCondition Then
            Ternary = vIfTrue
      Else
            Ternary = vIfFalse
      End If
End Function

Public Function GetCursorPosX() As Long
      Dim tRect As POINTAPI
      GetCursorPos tRect
      GetCursorPosX = tRect.X
End Function

Public Function NullOr(ByVal FirstThing As Variant, ByVal OtherThing As Variant)
      If Not IsNull(FirstThing) Then
            NullOr = FirstThing
      Else
            NullOr = OtherThing
      End If
End Function

Public Function TrimTrailingSlash(ByVal sPath As String) As String
      If Right(sPath, 1) = "\" Then
            TrimTrailingSlash = Left(sPath, Len(sPath) - 1)
      Else
            TrimTrailingSlash = sPath
      End If
End Function

Public Function ParentDirectoryOf(ByVal sPath As String)
      Dim iSlash As Integer
      
      If sPath = "\" Then
            ParentDirectoryOf = ""
      Else
            iSlash = InStrRev(sPath, "\", Len(sPath) - 1)
            ParentDirectoryOf = Left(sPath, iSlash)
      End If
End Function

Public Function SnipPath(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipPath = Right(sPath, Len(sPath) - iSlash)
End Function

Public Function SnipFileName(ByVal sPath As String) As String
      Dim iSlash As Integer
      iSlash = InStrRev(sPath, "\")
      SnipFileName = Left(sPath, iSlash)
End Function

Public Function CstringToVBstring(ByVal sCstring As String) As String
      ' Removes first null character and anything following it.
      On Error GoTo CONVERSION_ERROR
      Dim lNullPos As Long
      
      lNullPos = InStr(1, sCstring, Chr(0))
      If lNullPos = 0 Then
            CstringToVBstring = sCstring
      Else
            CstringToVBstring = Left(sCstring, lNullPos - 1)
      End If
      Exit Function
CONVERSION_ERROR:
      DebugLog "CONVERSION ERROR: " & sCstring, 2
End Function

Public Function FormatNonLocalFileTime(rtNlft As FILETIME) As String
      ' example date string:   2005-03-15 6:14:21
      
      Dim tLocalTime As FILETIME
      Dim tSysTime As SYSTEMTIME
      
      FileTimeToLocalFileTime rtNlft, tLocalTime
      FileTimeToSystemTime tLocalTime, tSysTime
      With tSysTime
            FormatNonLocalFileTime = .wYear & "-" & Format(.wMonth, "00") & "-" & Format(.wDay, "00") _
                  & ", " & Format(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00")
      End With
End Function

' Long (unsigned, positive values) only range from 1 to HALF of 4294967296 (16 ^ 8)
' So for the upper half of them, a DWORD from the API cannot be represented by a Long
'
' Whereas Currency goes up to 922,337,203,685,477.5807
' Even without leveraging the decimal places, that's enough bytes to describe any file
' in the high terabytes, close to a petabyte.
'
Public Function SignedLongToCurrency(ByVal lLng As Long) As Currency
      SignedLongToCurrency = lLng
      If lLng < 0 Then
            SignedLongToCurrency = CCur(lLng) + MAX_DWORD + 1
      End If
End Function

Public Function GetBigFileSize(ByRef rtWfd As WIN32_FIND_DATA, ByVal sFileName As String, _
      ByVal sDir As String, ByRef roFso As Object) As Currency
      
      GetBigFileSize = SignedLongToCurrency(rtWfd.nFileSizeLow)
      If rtWfd.nFileSizeHigh > 0 Then
            GetBigFileSize = rtWfd.nFileSizeHigh * (MAX_DWORD + 1) + GetBigFileSize
      End If
      LogBigFileSize rtWfd, sFileName, sDir, roFso, GetBigFileSize
End Function

Function FormatBytes(ByVal oBytes, iPrecision As Integer) As String
      ' Takes a quantity of bytes as a currency value (because it's 64-bit),
      ' format it to read like:
      
      ' 45.2 MB
      ' 300.2 KB
      ' 666 Bytes
      ' 99.4444 GB
      ' 20000 TB  (not supporting terabytes at the moment)
      
      ' ...with iPrecision digits after the demical.
      
      If oBytes = 1 Then
            FormatBytes = CStr(oBytes) & " byte"
      ElseIf oBytes < 1024@ Then
            FormatBytes = CStr(oBytes) & " bytes"
      ElseIf oBytes < 1048576@ Then
            FormatBytes = CStr(Round(oBytes / 1024@, iPrecision)) & " KB"
      ElseIf oBytes < 1073741824@ Then
            FormatBytes = CStr(Round(oBytes / 1048576@, iPrecision)) & " MB"
      ElseIf oBytes < 1099511627776@ Then
            FormatBytes = CStr(Round(oBytes / 1073741824@, iPrecision)) & " GB"
      Else
            FormatBytes = "Size Unknown"
      End If
End Function
