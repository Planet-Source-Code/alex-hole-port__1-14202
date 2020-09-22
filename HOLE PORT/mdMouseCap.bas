Attribute VB_Name = "mdMouseCap"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type


Public Sub MouseCap()
Dim pt As POINTAPI

    GetCursorPos pt
    client.txtX.Text = pt.x
    client.txtY.Text = pt.y
    client.wskClient.SendData "<<X>>" & pt.x & "," & pt.y
    PauseData.Pause 1

End Sub
