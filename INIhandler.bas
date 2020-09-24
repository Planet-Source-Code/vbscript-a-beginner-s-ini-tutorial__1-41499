Attribute VB_Name = "INIhandling"
'--------------------------------------------------------------------------
' This module makes writing to an INI file very easy.  See the examples
' below for what needs to go into the VB code for this to work.

' Write Example:
' --------------
' WriteINI "Section", "Setting", Value, App.Path & "\settings.ini"
'
' Read Example:
' -------------
' Variable = ReadINI("Section", "Setting", App.Path & "\settings.ini")
'--------------------------------------------------------------------------

Option Explicit

Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, getprivateprofilestring(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function

