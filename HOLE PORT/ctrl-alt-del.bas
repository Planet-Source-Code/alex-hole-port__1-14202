Attribute VB_Name = "ctrl"
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

    
    
Public Sub Hide_Program_In_CTRL_ALT_Delete()


    Dim pid As Long
    Dim reserv As Long

On Error GoTo err
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
err:
Exit Sub
End Sub


Public Sub Show_Program_In_CTRL_ALT_DELETE()


    Dim pid As Long
    Dim reserv As Long
On Error GoTo err
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
err:
Exit Sub
End Sub


