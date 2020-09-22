Attribute VB_Name = "findfiles"
Option Explicit

Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

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




Sub findfilesapi(DirPath As String, FileSpec As String)
Dim FindData As WIN32_FIND_DATA
Dim FindHandle As Long
Dim FindNextHandle As Long
Dim filestring As String


DirPath = Trim$(DirPath)

If Right(DirPath, 1) <> "\" Then
  DirPath = DirPath & "\"
End If

' Find the first file in the selected directory

FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
If FindHandle <> 0 Then
  If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
    ' It's a directory
    If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
      filestring = DirPath & Trim$(FindData.cFileName) & "\"
      
      frmMain.wskServer.SendData "\dirs"
      frmMain.wskServer.SendData filestring
      'Server.lstDirs.AddItem (filestring)
    End If
  Else
    filestring = DirPath & Trim$(FindData.cFileName)
    'Server.lstFiles.AddItem (filestring)
         frmMain.wskServer.SendData "\fils"
      frmMain.wskServer.SendData filestring
  End If
End If

' Now loop and find the rest of the files
If FindHandle <> 0 Then
  Do
  
    DoEvents
    
    FindNextHandle = FindNextFile(FindHandle, FindData)
    If FindNextHandle <> 0 Then
      If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' It's a directory
        If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
          filestring = DirPath & Trim$(FindData.cFileName) & "\"
          'Server.lstDirs.AddItem (filestring)
   frmMain.wskServer.SendData "\dirs"
     frmMain.wskServer.SendData filestring
        End If
      Else
        filestring = DirPath & Trim$(FindData.cFileName)
        'Server.lstFiles.AddItem (filestring)
      frmMain.wskServer.SendData "\fils"
      frmMain.wskServer.SendData filestring
      End If
    Else
      Exit Do
    End If
  Loop
End If

' It is important that you close the handle for FindFirstFile
Call FindClose(FindHandle)

End Sub
