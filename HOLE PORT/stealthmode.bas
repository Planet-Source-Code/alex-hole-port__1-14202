Attribute VB_Name = "stealthmode"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Sub HideExWindow(Class)
Dim HideW%
HideW% = FindWindow(Class, vbNullString)
Call ShowWindow(HideW%, 0)
End Sub
