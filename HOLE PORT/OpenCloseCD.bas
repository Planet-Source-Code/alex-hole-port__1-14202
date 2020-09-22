Attribute VB_Name = "OpenCloseCD"
Public Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As String, ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Public Function OPENCD()
mciSendString "set CDAudio door open", returnstring, 127, 0
End Function

Public Function CLOSECD()
mciSendString "set cdaudio door closed", returnstring, 127, 0
End Function
