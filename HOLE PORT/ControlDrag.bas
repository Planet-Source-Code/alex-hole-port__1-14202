Attribute VB_Name = "ControlDrag"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub Window_Drag(Ctrl As Control)
    ReleaseCapture
    SendMessage Ctrl.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

