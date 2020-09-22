Attribute VB_Name = "mdlWinsock"
'Filename: mdlWinsock.bas
'Desc    : Sample IRC client (raw data) using winsock API's.
'Author  : Jay Freeman (saurik)
'Comapany: SaurikSoft (www.saurik.com)
'Modified: 2:00 AM CST, October 4, 1997

Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)

Public intSocket As Integer
Public OldWndProc As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lct As String
    Dim lcv As Integer
    Select Case uMsg
        Case 1025
            Debug.Print "NOTIFICATION - " & wParam & " - " & lParam
            If wParam = intSocket Then
                Dim WSAEvent As Long
                Dim WSAError As Long
                WSAEvent = WSAGetSelectEvent(lParam)
                WSAError = WSAGetAsyncError(lParam)
                Select Case WSAEvent
                    'FD_READ    = &H1    = 1
                    'FD_WRITE   = &H2    = 2
                    'FD_OOB     = &H4    = 4
                    'FD_ACCEPT  = &H8    = 8
                    'FD_CONNECT = &H10   = 16
                    'FD_CLOSE   = &H20   = 32
                    Case FD_CLOSE
                        intSocket = 0
                        frmWinsock.cmdConnect.Caption = "Connect"
                        Output "*** Socket Closed (Server)"
                    Case FD_CONNECT
                        Select Case WSAError
                            Case 0
                                frmWinsock.cmdConnect.Caption = "Disconnect"
                                SendIt "USER " & frmWinsock.txtNick & " " & GetLocalHostName & " " & GetAscIP(GetHostByNameAlias(GetLocalHostName)) & " :" & frmWinsock.txtComment & vbCrLf
                                SendIt "NICK " & frmWinsock.txtNick & vbCrLf
                            Case Else
                                Output "Error Occured: " & GetWSAErrorString(WSAError) & vbCrLf
                        End Select
                        frmWinsock.cmdConnect.Enabled = True
                    Case FD_READ
                        Dim strBuf As String
                        Dim buflen As Long
                        strBuf = String(16384, 0)
                        buflen = recv(intSocket, strBuf, Len(strBuf), 0)
                        If buflen > -1 Then Output Left(strBuf, buflen)
                End Select
            End If
        Case Else: WindowProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, ByVal lParam)
    End Select
End Function
Public Sub SendIt(ByVal What As String)
    SendData intSocket, What
    Output "SENT: " & What
End Sub
Public Sub Output(ByVal What As String)
    frmWinsock.txtWinsock.SelStart = Len(frmWinsock.txtWinsock)
    frmWinsock.txtWinsock.SelText = What
    frmWinsock.txtWinsock.SelStart = Len(frmWinsock.txtWinsock)
End Sub
