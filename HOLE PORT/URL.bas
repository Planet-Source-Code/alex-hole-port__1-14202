Attribute VB_Name = "URL"
Public Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Const SW_SHOWNORMAL = 1

Public Function LOADURL(ByVal form As form, ByVal Site As String)
Dim iret As Long
    iret = ShellExecute(form.hwnd, _
        vbNullString, _
        Site, _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL)
End Function
