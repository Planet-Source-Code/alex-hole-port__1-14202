Attribute VB_Name = "ChangeBackground"
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20

Public Sub changeBackground(ByVal FileName As String)
 On Error GoTo err
SystemParametersInfo SPI_SETDESKWALLPAPER, 0, FileName, 0
err:
 Exit Sub
End Sub
