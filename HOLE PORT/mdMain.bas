Attribute VB_Name = "mdMain"
Public txtTitle As String
Dim SMTPServer As String
Dim SystemName As String
Dim Sender As String
Dim Recepient As String
Dim Message As String
Dim NL As String

Dim TextL As String
Dim TextL2 As String
Dim IPSend As Boolean
Dim cnt As Integer

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


Dim KEY_ALL_ACCESS As Variant
Dim a As Variant
Dim Texter$
