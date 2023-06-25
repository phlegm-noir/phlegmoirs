Attribute VB_Name = "APIRegistry"
' *************************************************************
' Windows API: Registry Functions
' *************************************************************

Option Explicit
Option Compare Binary

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
      ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal Reserved As Long, _
      ByVal lpClass As String, _
      ByVal dwOptions As Long, _
      ByVal samDesired As Long, _
      ByRef lpSecurityAttributes As Any, _
      ByRef phkResult As Long, _
      ByRef lpdwDisposition As Long) As Long
       
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_ALL_ACCESS = &HF003F
Public Const KEY_QUERY_VALUE = &H1

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
      ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      ByRef phkResult As Long) As Long

Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      lpData As String, _
      ByVal cbData As Long) As Long
      
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      lpData As Long, _
      ByVal cbData As Long) As Long
      
Public Declare Function RegSetValueExAny Lib "advapi32.dll" Alias "RegSetValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      lpData As Any, _
      ByVal cbData As Long) As Long
      
Public Const REG_BINARY As Long = 3 ' most any type, seems to be trouble for fixed length strings
Public Const REG_SZ As Long = 1 ' string value
Public Const REG_NONE As Long = 0
Public Const REG_DWORD As Long = 4 ' 32 bit number


Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      ByRef lpType As Long, _
      ByRef lpData As String, _
      ByRef lpcbData As Long) As Long
      
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      ByRef lpType As Long, _
      ByRef lpData As Long, _
      ByRef lpcbData As Long) As Long
      
Public Declare Function RegQueryValueExAny Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      ByRef lpType As Long, _
      ByRef lpData As Any, _
      ByRef lpcbData As Long) As Long
      
      
Public Declare Function RegCloseKey Lib "advapi32.dll" ( _
      ByVal hKey As Long) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
       ByVal hKey As Long, _
       ByVal lpValueName As String) As Long


