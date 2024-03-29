Attribute VB_Name = "SendData"
Option Explicit

' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ '
' \\                                                                \\ '
' \\  (modFileTransfer)  Server                                     \\ '
' \\                                                                \\ '
' \\ You will notice a definite simalarity in the Client module     \\ '
' \\ and server module for file transfer.
' \\                                                                \\ '
' \\                                                                \\ '
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ '

' Author    : gh0ul
' E-Mail    : gh0ul@hotmail.com
' ICQ#      : 31047555
' OS        : Win NT 4.0 Service Pack 3.0
' Platform  : VB 5.0
'
'

Declare Function GetTickCount Lib "kernel32" () As Long


Public Const Port = 1256                ' Port to listen on
Public Const MAX_CHUNK = 4169           ' Max size of sendable data

Public bInconnection     As Boolean     ' True if connected




' --- a function for pausing

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

' --- SendFile() Function
'
' Sends a file from one computer to another via WinSock

Sub SendFile(Fname As String)
    Dim DataChunk As String
    Dim passes As Long
    
    '
    ' send over the filename so the app knows where
    ' to store the file.
    SendData "OpenFile," & Fname$
    ' pause to give app time to get ready
    Pause 200
    
    ' open the file to be sent
    Open Fname$ For Binary As #1
       
        Do While Not EOF(1)
          ' update the Pass Variable
          passes& = passes& + 1
          ' get some of the file data
          DataChunk$ = Input(MAX_CHUNK, #1)
          ' send it to the server
          SendData DataChunk$
          ' report status
' ** // ** '
' IMPORTANT: comment out the code below when sending files
' larger than 500Kb. It makes the function CRAWL otherwise
             
' comment the above line to increase speed

          ' pause to give the Client time to procees this
          ' information
          Pause 200
          DoEvents
        Loop ' loop until all data is sent
        
        ' transfer done, notify the server to close the file
        SendData "CloseFile,"
        ' re-init byte counter and update status
        passes& = 0
    Close #1
End Sub

' --- send data function this is merely a better way to access
' the winsock "SendData" function. does it's own error
' checking

Sub SendData(sData As String)
    On Error GoTo ErrH

    Dim TimeOut As Long
    
    client.wskClient.SendData sData
    
    Do Until (client.wskClient.State = 0) Or (TimeOut < 10000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 10000 Then Exit Do
    Loop
    
ErrH:
    Exit Sub
End Sub


' GetFileName()
'
' Extract the file name and extension only from
' the full path.

Function GetFileName(Fname As String) As String
    ' return the filename given the path
    Dim i As Integer
    Dim tempStr As String
    
    For i% = 1 To Len(Fname$)
       ' look for the "\"
       tempStr$ = Right$(Fname$, i%)
       
       If Left$(tempStr$, 1) = "\" Then
         GetFileName$ = Mid$(tempStr$, 2, Len(tempStr$))
         Exit Function
       End If
    Next i
End Function



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'======================================================================
' (EvalData Function)
'
'  Purpose - Extract data from a given string, to the right or left
'            of a specified character.
'
'  Parameters:
'     sIncoming - The String you want to extract data from.
'     iRtLt     - Extract from the Left, 1.
'                 Extract from the right, 2.
'     sDivider  - The character that seperates the data in
'                 the string. <default = ",">
'
'  Returns:
'     (type)String
'     Returns the data to the right or left of strDivider.
'======================================================================
             
' THis function can be used to seperate endless bits of
' data sent withth SendData Function. Although it can get a little
' cumbersome with really long lists.

' If you would like more info on exactly how to accomplish this
' E-mail me or Message on ICQ and I will show you.

Public Function EvalData(sIncoming As String, iRtLt As Integer, _
                  Optional sDivider As String) As String
   Dim i As Integer
   Dim tempStr As String
   ' Storage for the current Divider
   Dim sSplit As String
   
   ' the current character used to divide the data
   If sDivider = "" Then
      sSplit = ","
   Else
      sSplit = sDivider
   End If
   
   ' getting the right or left?
   Select Case iRtLt
        
      Case 1
          ' remove the data to the Left of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Left(sIncoming, i)
            
            If Right(tempStr, 1) = sSplit Then
              EvalData = Left(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
          
      Case 2
          ' remove the data to the Right of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Right(sIncoming, i)
            
            If Left(tempStr, 1) = sSplit Then
              EvalData = Right(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
   End Select
   
End Function




