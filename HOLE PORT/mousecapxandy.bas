Attribute VB_Name = "MouseCap"
Dim blnStartCapture As Boolean
Dim blnAmCapturing As Boolean

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Type POINTAPI
    x As Long
    y As Long
End Type


Public Function StartMouseCapX()
Dim ptPoint As POINTAPI
Dim RetVal As Variant
ptPoint.x = x
ptPoint.y = y
RetVal = ClientToScreen(hWnd, ptPoint)
StartMouseCapX = ptPoint.x
End Function
Public Function StartMouseCapY()
Dim ptPoint As POINTAPI
Dim RetVal As Variant
ptPoint.x = x
ptPoint.y = y
RetVal = ClientToScreen(hWnd, ptPoint)
StartMouseCapY = ptPoint.y
End Function

