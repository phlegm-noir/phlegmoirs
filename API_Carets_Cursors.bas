Attribute VB_Name = "APICaret"
' *************************************************************
' Windows API: carets and cursors
' *************************************************************

Option Explicit
Option Compare Binary

Public Type POINTAPI
      X As Long
      Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32.dll" ( _
      ByRef lpPoint As POINTAPI) As Long


