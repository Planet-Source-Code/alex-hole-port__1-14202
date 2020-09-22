Attribute VB_Name = "WorkWithINI"

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long


Public Function GETini(ByVal App As String, ByVal Key As String, ByVal FileName As String)
Dim buf As String * 256
Dim length As Long

    length = GetPrivateProfileString( _
        App, Key, "<no value>", _
        buf, Len(buf), FileName)
    GETini = Left$(buf, length)
End Function

' Set the value.
Public Sub WRITEini(ByVal App As String, ByVal Key As String, ByVal FileName As String, ByVal value As String)
    WritePrivateProfileString _
        App, Key, _
        value, FileName
    End Sub

