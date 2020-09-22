Attribute VB_Name = "WinThreads"
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public txtTitle As String

Type wndClass
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

'C language TypeDef to hold the size information for a given window.
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
       
       Public Const SW_SHOW = 5
       
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
                                                        lpdwProcessId As Long) As Long

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, _
                                                                  ByVal nIndex As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                   ByVal lpWindowName As String) As Long
                                                                                         
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                  ByVal lpClassName As String, _
                                                                  ByVal nMaxCount As Long) _
                                                                  As Long

Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                                                        ByVal lpString As String, ByVal cch As Long) _
                                                                        As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
                                                                    ByVal dwNewLong As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
                                                  hWndInsertAfter As Long, ByVal X As Long, _
                                                  ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
                                                  ByVal wFlags As Long) As Long

Declare Function GetActiveWindow Lib "user32" () As Long

Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetForegroundWindow Lib "user32" () As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                ByVal wMsg As Long, ByVal wParam As Long, _
                                                                lParam As Any) As Long

Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, _
                                                                  ByVal lpClassName As String, _
                                                                  lpWndClass As wndClass) As Long

'----------------------------------------------------------------------------------------------------------
Public Const WM_ACTIVATE = &H6
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
'----------------------------------------------------------------------------------------------------------



Public Function GetOpenWindowNames(Index As Integer) As Long

On Error GoTo ErrorHandl
'----------------------------------------------------------------------------------------------------------
'Name:        Function GetOpenWindowNames()
'
'Purpose:     To retrieve all open windows in the system.
'
'Parameters:  N/A
'
'Return:      NONE
'----------------------------------------------------------------------------------------------------------

'Declare local variables
Dim lngDeskTopHandle As Long    'Used to hold the value of the Desktop handle.
Dim lngHand As Long             'Used to hold each windows handle as it loops.
Dim strName As String * 255     'Fixed length string passed to GetWindowText API call.
Dim lngWindowCount As Long      'Counter used to return the numberof open windows in the system.

'Get the handle for the desktop.
lngDeskTopHandle = GetDesktopWindow()

'Get the first child of the desktop window.
'(Note: The desktop is the parent of all windows in the system.
lngHand = GetWindow(lngDeskTopHandle, GW_CHILD)

'set the window counter to 1.
lngWindowCount = 1

'Loop while there are still open windows.
Do While lngHand <> 0
     
     'Get the title of the next window in the window list.
     GetWindowText lngHand, strName, Len(strName)
     
     'Get the sibling of the current window.
     lngHand = GetWindow(lngHand, GW_HWNDNEXT)
     
     'Make sure the window has a title; and if it does add it to the list.
     If Left$(strName, 1) <> vbNullChar Then
     PauseData.Pause 0.5
'          frmClassFinder.lstOpenWindows.AddItem Left$(strName, InStr(1, strName, vbNullChar))
          frmMain.wskWINT(Index).SendData "\Wnt" & Left$(strName, InStr(1, strName, vbNullChar))
          lngWindowCount = lngWindowCount + 1
          PauseData.Pause 0.5
          frmMain.wskWINT(Index).SendData "\lng" & lngWindowCount
          End If
Loop

'Return the number of windows opened.
'GetOpenWindowNames = lngWindowCount
'PauseData.Pause 0.3
'frmMain.wskWINT(Index).SendData "\lnFinal" & lngWindowCount
PauseData.Pause 1
frmMain.wskWINT(Index).SendData "<<FINISHED>>"
ErrorHandl:

Exit Function

End Function

Public Sub GetClass(Index As Integer)



Dim lngHand As Long
Dim strName As String * 255
Dim wndClass As wndClass
Dim rctTemp As RECT
Dim lblClassName As String

On Error GoTo err

lngHand = FindWindow(vbNullString, txtTitle)


GetClassName lngHand, strName, Len(strName)


If Left$(strName, 1) = vbNullChar Then
     
Else
     lblClassName = "Class Name: " & strName
     GetWindowThreadProcessId lngHand, lngProcID
     GetWindowRect lngHand, rctTemp
End If


frmMain.wskWINT(Index).SendData "\class" & lblClassName
PauseData.Pause 0.5
frmMain.wskWINT(Index).SendData "\prID" & lngProcID
PauseData.Pause 0.5
frmMain.wskWINT(Index).SendData "\top" & rctTemp.Top
PauseData.Pause 0.5
frmMain.wskWINT(Index).SendData "\bot" & rctTemp.Bottom
PauseData.Pause 0.5
frmMain.wskWINT(Index).SendData "\left" & rctTemp.Left
PauseData.Pause 0.5
frmMain.wskWINT(Index).SendData "\right" & rctTemp.Right


err:
Exit Sub

End Sub


Public Sub Activate()
Dim lngHand As Long

On Error GoTo err

lngHand = FindWindow(vbNullChar, Trim$(txtTitle))
ShowWindow lngHand, SW_RESTORE

err:
Exit Sub

End Sub

Public Sub Hide()
Dim lngHide As Long

On Error GoTo err

lngHand = FindWindow(vbNullChar, Trim$(txtTitle))
ShowWindow lngHide, SW_HIDE

err:
Exit Sub

End Sub


Public Sub Destroy()
Dim lngHide As Long
lngHand = FindWindow(vbNullChar, Trim$(txtTitle))
ShowWindow lngHide, WM_DESTROY
End Sub

