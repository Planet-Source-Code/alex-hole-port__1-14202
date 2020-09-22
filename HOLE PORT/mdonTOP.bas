Attribute VB_Name = "mdonTOP"


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Sub SetOnTop(frm As Form)
SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Public Sub SetOffTop(frm As Form)
SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Sub
