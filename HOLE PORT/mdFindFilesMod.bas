Attribute VB_Name = "mdFindFilesMod"
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public NodeKeySent As String
Public NodePathSent As String
Public Const MAX_PATH = 260

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80




Public Sub findfilesapi(DirPath As String, strpath As String, Index As Integer)
Dim FindData As WIN32_FIND_DATA
Dim FindHandle As Long
Dim FindNextHandle As Long
Dim filestring As String
DirPath = Trim$(DirPath)


On Error GoTo err


If Right(DirPath, 1) <> "\" Then
  DirPath = DirPath & "\"
End If

FindHandle = FindFirstFile(DirPath & "*.*", FindData)

If FindHandle <> 0 Then
  If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
    
    If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
      filestring = DirPath & Trim$(FindData.cFileName) & "\"
      frmMain.wskFL(Index).SendData "\TreeView" & FindData.cFileName
      PauseData.Pause 0.5
    End If
  Else
    filestring = DirPath & Trim$(FindData.cFileName)
    frmMain.wskFL(Index).SendData "\ListView1" & FindData.cFileName
    PauseData.Pause 0.5
  End If

If FindHandle <> 0 Then
  Do
  
    DoEvents
    
    FindNextHandle = FindNextFile(FindHandle, FindData)
    If FindNextHandle <> 0 Then
      If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        
        If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
          filestring = DirPath & Trim$(FindData.cFileName) & "\"
          frmMain.wskFL(Index).SendData "\TreeView" & FindData.cFileName
          PauseData.Pause 0.5
           End If
      Else
        filestring = DirPath & Trim$(FindData.cFileName)
        frmMain.wskFL(Index).SendData "\ListView1" & FindData.cFileName
        PauseData.Pause 0.5
    End If
    Else
      Exit Do
      End If
  Loop
End If


Call FindClose(FindHandle)

End If

err:
Exit Sub

End Sub



Public Function SendData(sData As String, Index As Integer) As Boolean
    
    Dim TimeOut As Long

On Error GoTo err


    SendData = False
    ' send data
    frmMain.wskFileSend(Index).SendData sData

    ' check for timeout or closed socket
    Do Until (frmMain.wskFileSend(Index).State = 0) Or (TimeOut < 100000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 100000 Then Exit Do
    Loop
    ' ok.... sent
    SendData = True
    Exit Function

err:
Exit Function

End Function
