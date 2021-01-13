Attribute VB_Name = "iniFile"
Option Explicit
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const Default = Empty

Function WriteINI(opt1 As String, opt2 As String, opt3 As String)
Dim ret
Dim ret3 As String
    ret3 = Trim(opt3)
    ret = WritePrivateProfileString(opt1, opt2, ret3, "D:\USER.ini")
End Function

Function ReadINI(opt1 As String, opt2 As String, vlen As Integer) As String
Dim ret
Dim tmpstring As String * 512
    ret = GetPrivateProfileString(opt1, opt2, Default, tmpstring, vlen, "D:\USER.ini")
    ReadINI = Replace(tmpstring, Chr(0), "")
End Function

Function ReadINI_Delta(opt1 As String, opt2 As String, vlen As Integer) As String
Dim ret
Dim tmpstring As String * 512
    ret = GetPrivateProfileString(opt1, opt2, Default, tmpstring, vlen, ReadINI("CONFIGURATION", "FILE INI DELTA", 512))
    ReadINI_Delta = Replace(tmpstring, Chr(0), "")
End Function
