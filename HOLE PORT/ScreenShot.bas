Attribute VB_Name = "shot"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'keybd_event vbKeySnapshot, 1 for full or 0 for active windows, 0&, 0&
'DoEvents
'Form1.Picture = Clipboard.GetData(vbCFBitmap)
