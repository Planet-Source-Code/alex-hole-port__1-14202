Attribute VB_Name = "RandomMouseMove"

Dim time1 As Integer
Dim x1 As Integer
Dim y1 As Integer
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Dim more As Boolean
Dim speed As Integer
Dim previous As POINTAPI
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Sub RandomMouse()
Dim X As Integer

On Error GoTo err


    For X = 0 To 25
    PauseData.Pause 0.2
    X = X + 2
Dim current1 As POINTAPI
r = GetCursorPos(current1)
r = SetCursorPos(current1.X + Int(Rnd * 10) + 1, current1.Y + Int(Rnd * 10) + 1)
DoEvents
    Next X
err:
Exit Sub

End Sub
