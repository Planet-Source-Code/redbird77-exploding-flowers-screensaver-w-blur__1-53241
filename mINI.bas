Attribute VB_Name = "mINI"
Option Explicit

' Short and simple interface to an INI file.
' ------------------------------------------
' redbird77@earthlink.net (2004.04.14)

' Structure of an INI file
' ------------------------
' [Section]
' Key=Value

' Functions contained in mINI.bas
' -------------------------------
' PutValue
' DelValue
' GetValue

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub PutValue(ByVal sSection As String, ByVal sKey As String, _
                    ByVal sValue As String, ByVal sFile As String)
Dim r As Long

    r = WritePrivateProfileString(sSection, sKey, sValue, sFile)
    Debug.Assert r
    
End Sub

Public Sub DelValue(ByVal sSection As String, ByVal sKey As String, _
                    ByVal sFile As String)

Dim r As Long

    ' Pass NULL as the value thus deleting the key-value pair.
    r = WritePrivateProfileString(sSection, sKey, vbNullString, sFile)
    Debug.Assert r
    
End Sub

Public Function GetValue(ByVal sSection As String, ByVal sKey As String, _
                         ByVal sFile As String, _
                         Optional ByVal sDefault As String = "") As String

Dim sBuf As String, r As Long, lPos As Long
    
    sBuf = String$(1024, vbNullChar)
    
    r = GetPrivateProfileString(sSection, sKey, sDefault, sBuf, Len(sBuf), sFile)
    
    lPos = InStr(sBuf, vbNullChar)
    
    If lPos Then GetValue = Left$(sBuf, lPos - 1)
    
End Function
