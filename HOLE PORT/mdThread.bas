Attribute VB_Name = "mdThread"
Public MyThread As Long


Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long


Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long





Sub SpawnThread()
    MyThread = SetTimer(0&, 0&, 500&, AddressOf MySub)
End Sub

Sub MySub()
Dim AppText As String
Dim App


App = GetSetting("MainServer", "Blocked", "BlockedApp")
AppText = GetSetting("MainServer", "BlockedTEXT", "BlockedTEXT")

    SETING = FindWindow(App, AppText)
If SETING <> 0 Then ShowWindow SETING, SW_HIDE

End Sub



