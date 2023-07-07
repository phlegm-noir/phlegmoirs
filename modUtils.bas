Attribute VB_Name = "modUtils"
Option Explicit
Option Compare Binary

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
