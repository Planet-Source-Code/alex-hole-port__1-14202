Attribute VB_Name = "MAIN"

'**************************************
' Name: A Must Have .bas File For VB Pro
'     gramming
' Description:A Must Have StartupModule.
'     bas File. Lots Of Options.
' By: T. L. Phillips
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'
'Warranty:
'Code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************


Option Explicit
Public DBa(1 To 100) As String
Public AppPath
Public DallorGet
Public FirstLoad
Public KeyBoardType
Public KeyBoardRepeatDelay
Public KeyBoardRepeatSpeed
Public KeyBoardCaretFlashSpeed
Public CurDate
Public Ret As String
Public ReturnINIdat
Public INIFileFound
Public ShortFName
Public title
Public FileInfoName As String
Public FileInfoPathName As String
Public FileInfoSize As String
Public FileInfoLastModified As String
Public FileInfoLastAccessed As String
Public FileInfoAttributeHidden As String
Public FileInfoAttributeSystem As String
Public FileInfoAttributeReadOnly As String
Public FileInfoAttributeArchive As String
Public FileInfoAttributeTemporary As String
Public FileInfoAttributeNormal As String
Public FileInfoAttributeCompressed As String
Public VBSysDir
Public DirChkSize
Public Cd_Rom
Public Msg
Public DatGet
Public Word
Public StartTime
Public WordD
Public WordK
Public Dat
Public DOt
Public IsFileThere
Public Playinfo
Public DelConFirm
Public FlPath
Public sDType
Public GetWinDir
Public FlName
Public ShortPN
Public GWinDir
Public SupSound
Public DriveFreeSpace
Public DOSWinActive As String
Public Const GW_HWNDNEXT = 2
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_UNKNOWN = 0
Public Const AUDIO_NONE = 0
Public Const AUDIO_WAVE = 1
Public Const AUDIO_MIDI = 2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_CLOSE = &H10
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETKEYBOARDDELAY = 22


Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long


Declare Function GetWindowDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long


Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long


Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long


Declare Function GetCaretBlinkTime Lib "user32" () As Long


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Declare Function GetDesktopWindow Lib "user32" () As Long


Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Integer) As Integer


Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Integer) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long


Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Private Declare Function fCreateShellGroup Lib "STKIT432.DLL" _
    (ByVal lpstrDirName As String) As Long


Private Declare Function fCreateShellLink Lib "STKIT432.DLL" _
    (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
    ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long


Private Declare Function fRemoveShellLink Lib "STKIT432.DLL" _
    (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long


Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String ' only used if FOF_SIMPLEPROGRESS
    End Type


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type


Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
    End Type


Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTimeAs As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260 'MUST be set to 260
    cAlternate As String * 14
    End Type


Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
    End Type


Type POINTAPI
    X As Long
    Y As Long
    End Type
    Const SWP_NOZORDER = &H4


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Const HKEY_LOCAL_MACHINE = &H80000002


Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Public Const SND_ALIAS = &H10000
    Public Const SND_ALIAS_ID = &H110000
    Public Const SND_ALIAS_START = 0
    Public Const SND_APPLICATION = &H80
    Public Const SND_ASYNC = &H1
    Public Const SND_FILENAME = &H20000
    Public Const SND_LOOP = &H8
    Public Const SND_MEMORY = &H4
    Public Const SND_NODEFAULT = &H2
    Public Const SND_NOSTOP = &H10
    Public Const GWL_STYLE = (-16)
    Public Const ES_NUMBER = &H2000
    Public Const SND_NOWAIT = &H2000
    Public Const SND_PURGE = &H40
    Public Const SND_RESERVED = &HFF000000
    Public Const SND_RESOURCE = &H40004
    Public Const SND_SYNC = &H0
    Public Const SND_TYPE_MASK = &H170007
    Public Const SND_VALID = &H1F
    Public Const SND_VALIDFLAGS = &H17201F
    Private Const ERROR_SUCCESS = 0&
    Private Const APINULL = 0&
    Private ReturnCode As Long
    Private Target As String


Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
    End Type


Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
    End Type
    Global Const WM_USER = &H400
    Global UserhWnd As Long


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Const NORMAL_PRIORITY_CLASS = &H20&
    Private Const INFINITE = -1&


Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long


Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
    Private lShowCursor As Long


Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long


Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long


Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
    
    Private Const FO_DELETE = &H3
    Private Const FOF_ALLOWUNDO = &H40
    Private Const FOF_SILENT = &H4
    Private Const FOF_NOCONFIRMATION = &H10


Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
    "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long


Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long


Declare Function GetActiveWindow Lib "user32" () As Long


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long


Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, _
    ByVal lpstrBffer As String, ByVal uLength As Long) As Long


Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Declare Function sndPlaySoundByte Lib "winmm.dll" Alias "sndPlaySoundA" _
    (lpszSoundName As Byte, ByVal uFlags As Long) As Long


Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data) ' Get information of file using API call


    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error" ' If file was not found, return error as name
    Else
        Findfile = Win32Data ' Else return results
    End If
    plngRtn = FindClose(plngFirstFileHwnd) ' It is important that you close the handle for FindFirstFile
End Function


Function REGGETSTRING$(hInKey As Long, ByVal subkey$, ByVal valname$)
    Dim v$, RetVal$, hSubKey As Long, dwType As Long, SZ As Long
    Dim r As Long
    RetVal$ = ""
    Const KEY_ALL_ACCESS As Long = &HF0063
    Const ERROR_SUCCESS As Long = 0
    Const REG_SZ As Long = 1
    r = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
    If r <> ERROR_SUCCESS Then GoTo Quit_Now
    SZ = 256: v$ = String$(SZ, 0)
    r = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)


    If r = ERROR_SUCCESS And dwType = REG_SZ Then
        RetVal$ = Left$(v$, SZ)
    Else
        RetVal$ = "--Not String--"
    End If
    If hInKey = 0 Then r = RegCloseKey(hSubKey)
Quit_Now:
    REGGETSTRING$ = RetVal$
End Function


Public Function ActiveConnection() As Boolean
    '
    'Usage:
    ' ActiveConnection
    ' Msgbox ActiveConnection 'True = Connec
    '     ted to Internet \ False = Not Connected
    '     to Internet
    '
    Dim hKey As Long
    Dim lpSubKey As String
    Dim phkResult As Long
    Dim lpValueName As String
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As Long
    Dim lpcbData As Long
    ActiveConnection = False
    ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\RemoteAccess", phkResult)


    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = "Remote Connection"
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        ReturnCode = RegQueryValueEx(hKey, lpValueName, _
        lpReserved, lpType, ByVal lpData, lpcbData)
        lpcbData = Len(lpData)
        ReturnCode = RegQueryValueEx(hKey, lpValueName, _
        lpReserved, lpType, lpData, lpcbData)


        If ReturnCode = ERROR_SUCCESS Then


            If lpData = 0 Then
                ActiveConnection = False
            Else
                ActiveConnection = True
            End If
        End If
        RegCloseKey (hKey)
    End If
End Function


Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
    Dim buf As String * 256
    Dim title As String
    Dim length As Long
    ' Get the window's title.
    length = GetWindowText(app_hWnd, buf, Len(buf))
    title = Left$(buf, length)
    ' See if this is the target window.


    If InStr(title, Target) <> 0 Then
        ' Kill the window.
        SendMessage app_hWnd, WM_CLOSE, 0, 0
    End If
    ' Continue searching.


EnumCallback = 1
End Function


Public Function FindWindowPartial(ByVal TitlePart As String) As Long
    '
    'Used By FindDosWin
    '
    Dim hWndTmp As Long
    Dim nRet As Integer
    Dim TitleTmp As String
    TitlePart = UCase$(TitlePart)
    hWndTmp = FindWindow(0&, 0&)


    Do Until hWndTmp = 0
        TitleTmp = Space$(256)
        nRet = GetWindowText(hWndTmp, TitleTmp, Len(TitleTmp))


        If nRet Then
            TitleTmp = UCase$(VBA.Left$(TitleTmp, nRet))


            If InStr(TitleTmp, TitlePart) Then
                FindWindowPartial = hWndTmp
                Exit Do
            End If
        End If
        hWndTmp = GetWindow(hWndTmp, GW_HWNDNEXT)
    Loop
End Function


Function GETCURRUSER() As String
    '
    'Usage:
    'USERNAME = GETCURRUSER()
    'Msgbox USERNAME
    '
    GETCURRUSER = REGGETSTRING$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
End Function


Function GETCURRORG() As String
    '
    'Usage:
    ' GETCURRORG
    ' Msgbox USERORG
    '
    GETCURRORG = REGGETSTRING$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
End Function


Function STRIPNULLS(startStrg$) As String
    Dim c%, item$
    c% = 1


    Do


        If Mid$(startStrg$, c%, 1) = Chr$(0) Then
            item$ = Mid$(startStrg$, 1, c% - 1)
            startStrg$ = Mid$(startStrg$, c% + 1, Len(startStrg$))
            STRIPNULLS$ = item$
            Exit Function
        End If
        c% = c% + 1
    Loop
End Function


Function App_Path() As String
    '
    'Usage:
    ' App_Path
    ' msgbox App_Path
    '
    Dim X
    X = App.Path
    If Right$(X, 1) <> "\" Then X = X + "\"
    App_Path = UCase$(X)
End Function


Sub CenterForm(WhatForm As Form)
    '
    'Usage:
    ' CenterForm Form1
    '
    If WhatForm.WindowState <> 0 Then Exit Sub
    WhatForm.Move (Screen.Width - WhatForm.Width) \ 2, (Screen.Height - WhatForm.Height) \ 2
End Sub


Public Sub CenterFormTop(frm As Form)
    '
    'Usage:
    'CenterFormTop Form1
    '


    With frm
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / (Screen.Height)
    End With
End Sub


Public Sub CenterFormBottom(frm As Form)
    '
    'Usage:
    'CenterFormBottom Form1
    '


    With frm
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height)
    End With
End Sub


Public Sub CenterFormBottomRight(frm As Form)
    '
    'Usage:
    'CenterFormBottomRight Form1
    '


    With frm
        .Left = (Screen.Width - .Width) / 1
        .Top = (Screen.Height - .Height)
    End With
End Sub


Public Sub CenterFormBottomLeft(frm As Form)
    '
    'Usage:
    'CenterFormBottomLeft Form1
    '


    With frm
        .Left = 0
        .Top = (Screen.Height - .Height)
    End With
End Sub


Public Sub CenterFormTopRight(frmForm As Form)
    '
    'Usage:
    'CenterFormTopRight Form1
    '


    With frmForm
        .Left = (Screen.Width - .Width) / 1
        .Top = (Screen.Height - .Height) / 2000
    End With
End Sub


Public Sub CenterFormTopLeft(frmForm As Form)
    '
    'Usage:
    'CenterFormTopLeft Form1
    '


    With frmForm
        .Left = 0
        .Top = 0
    End With
End Sub


Sub DeKrypt()
    '
    'Usage:
    'Dat = "TEST"
    'DeKrypt
    'Msgbox WordD
    '
    Dim i, Strg$, h$, J$
    WordD = ""


    For i = 1 To Len(Dat)
        WordD = WordD & Chr(Asc(Mid(Dat, i, 1)) - 1)
    Next i
End Sub


Sub Krypt()
    '
    'Usage:
    'Dat = "TEST"
    'Krypt
    'Msgbox WordK
    '
    Dim i, Strg$, h$, J$
    WordK = ""


    For i = 1 To Len(Dat)
        WordK = WordK & Chr(Asc(Mid(Dat, i, 1)) + 1)
    Next i
End Sub


Sub Detect_CD_Rom()
    '
    'Usage:
    'Detect_CD_ROM
    'Msgbox CD_ROM
    '
    Dim r&, allDrives$, JustOneDrive$, pos%, DriveType&
    Dim CDfound As Integer
    allDrives$ = Space$(64)
    r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
    allDrives$ = Left$(allDrives$, r&)


    Do
        pos% = InStr(allDrives$, Chr$(0))


        If pos% Then
            JustOneDrive$ = Left$(allDrives$, pos%)
            allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
            DriveType& = GetDriveType(JustOneDrive$)


            If DriveType& = DRIVE_CDROM Then
                CDfound% = True
                Exit Do
            End If
        End If
    Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM


    If CDfound% Then
        Cd_Rom = Trim(UCase$(JustOneDrive$))
    Else: Cd_Rom = "?"
    End If
End Sub


Sub HandW(FORMID As Form)
    '
    'Form Hieght And Width
    '
    'Usage:
    ' HandW Form1
    '
    Dim a, b
    Dat = ""
    a = FORMID.Height
    b = FORMID.Width
    Dat = "Hieght = " & a & " Width = " & b
    Msg = Dat
    MsgBx
End Sub


Sub LandT(FORMID As Form)
    '
    'Form Left And Top
    '
    'Usage:
    ' LandT Form1
    '
    Dim a, b
    Dat = ""
    a = FORMID.Left
    b = FORMID.Top
    Dat = "Left = " & a & " Top = " & b
    Msg = Dat
    MsgBx
End Sub


Sub MidiPlay(NamePath As String)
    '
    'Usage:
    'MidiPlay "Test.mid"
    '
    OpenMidi NamePath
    PlayMidi
End Sub


Sub OpenMidi(sfile As String)
    '
    'Used by MidiPlay SUB
    '
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim sError As String * 255
    lResult = GetShortPathName(sfile, sShortFile, Len(sShortFile))
    sfile = Left(sShortFile, lResult)
    lResult = mciSendString("open " & sfile & " type sequencer alias mcitest", ByVal 0&, 0, 0)


    If lResult Then
        lResult = mciGetErrorString(lResult, sError, 255)
        Debug.Print "open: " & sError
    End If
End Sub


Sub PlayMidi()
    '
    'Used by MidiPlay SUB
    '
    Dim lResult As Integer
    Dim sError As String * 255
    lResult = mciSendString("play mcitest", ByVal 0&, 0, 0)


    If lResult Then
        lResult = mciGetErrorString(lResult, sError, 255)
        Debug.Print "play: " & sError
    End If
End Sub


Sub StopMidi()
    '
    'Usage:
    ' StopMidi 'Stop Any Midi File Playing
    '
    Dim lResult As Integer
    Dim sError As String * 255
    lResult = mciSendString("close mcitest", "", 0&, 0&)


    If lResult Then
        lResult = mciGetErrorString(lResult, sError, 255)
        Debug.Print "stop: " & sError
    End If
End Sub


Sub Timeout(duration)
    '
    'Usage:
    ' Timeout (1)
    '
    StartTime = Timer


    Do While Timer - StartTime < duration


        DoEvents
        Loop
    End Sub


Sub MsgBx()
    '
    'Usage:
    'Msg = "Test Message"
    'MsgBx
    '


    If Msg = "" Then
        Msg = "NO MESSAGE TO DISPLAY"
    End If
    MsgBox Msg, vbOKOnly, title
End Sub


Sub YN_Msgbox()
    '
    'Usage:
    'Title = "Test Title"
    'Msg = "Quit?"
    'YN_Msgbox
    'If Word = "Y" then
    'Msgbox "Yes!"
    'End if
    'If Word = "N" then
    'Msgbox "No!"
    'End if
    '
    Dim style, CTXT, HELP, Response
    Word = ""
    style = vbYesNo + vbDefaultButton2
    CTXT = 1000
    Response = MsgBox(Msg, style, title, HELP, CTXT)


    If Response = vbYes Then
        Word = "Y"
    Else
        Word = "N"
    End If
End Sub


Public Sub PlayWav(SFileName As String, Optional Mode)
    '
    'Usage:
    'PlayWav "test.wav",1 'Plays Wav With Ou
    '     t Delay.
    'PlayWav "test.wav",2 'Plays Wav With De
    '     lay.
    '
    Dim lReturn As Long
    On Error GoTo ErrorHandleFile
    If IsMissing(Mode) Then Mode = SND_ASYNC Or SND_NODEFAULT


    If (Mode And SND_ALIAS) <> SND_ALIAS Then


        If Len(Dir(Trim$(SFileName))) = 0 Then
            Exit Sub
        End If
    End If
    lReturn = sndPlaySound(SFileName, Mode)
ErrorHandleFile:
End Sub


Sub StayOnTop(the As Form)
    '
    'Usage:
    'StayOnTop Form1
    '
    Dim SetWinOnTop%
    SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


Sub NumRND(NMBR As Long)
    '
    'Usage:
    'NumRND 999999999 'Nine Number Max.
    'Msgbox Dat
    '
    Randomize
    Dat = Int(NMBR * Rnd)
End Sub


Sub NumTextOnly(KeyR)
    '
    'Usage:
    'NumTextOnly KeyAscii 'Place This Code I
    '     n The TextBox_KeyPressed Sub
    '
    Const numbers$ = "0123456789"


    If KeyR <> 8 Then


        If InStr(numbers, Chr(KeyR)) = 0 Then
            KeyR = 0
            Exit Sub
        End If
    End If
End Sub


Sub NumTextOnlyWithDash(KeyR)
    '
    'Usage:
    'NumTextOnlyWithDash KeyAscii 'Place Thi
    '     s Code In The TextBox_KeyPressed Sub
    '
    Const numbers$ = "0123456789-"


    If KeyR <> 8 Then


        If InStr(numbers, Chr(KeyR)) = 0 Then
            KeyR = 0
            Exit Sub
        End If
    End If
End Sub


Sub NumTextOnlyWithDOT(KeyR, DataText As TextBox)
    '
    'Usage:
    'NumTextOnlyWithDOT KeyAscii, text1 'Pla
    '     ce This Code In The TextBox_KeyPressed S
    '     ub
    '
    Dim a, b, c, USEdot
    USEdot = True
    If FirstLoad = True Then Exit Sub
    a = Len(DataText)
    b = 1


    Do Until b = a
        If b > a Then Exit Sub
        c = Mid$(DataText, b, 1)


        If c = "." Then
            USEdot = False
        End If
        b = b + 1
    Loop
    Const numbers$ = "0123456789."
    'If USEdot = False Then
    'numbers$ = "0123456789"
    'Else
    'numbers$ = "0123456789."
    'End If


    If KeyR <> 8 Then


        If InStr(numbers, Chr(KeyR)) = 0 Then
            KeyR = 0
            Exit Sub
        End If
    End If
End Sub


Sub FormRunLeft(the As Form)
    '
    'Usage:
    'FormRunLeft Form1
    '
    Dim counter
    counter = the.Left


    Do: DoEvents
        counter = counter + 100
        the.Left = counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunRight(the As Form)
    '
    'Usage:
    'FormRunRight Form1
    '
    Dim counter
    counter = the.Left


    Do: DoEvents
        counter = counter + 100
        the.Left = the.Left - counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunDown(the As Form)
    '
    'Usage:
    'FormRunDown Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Top = counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunUp(the As Form)
    '
    'Usage:
    'FormRunUp Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Top = the.Top - counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunLeftUp(the As Form)
    '
    'Usage:
    'FormRunLeftUp Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Left = the.Left - counter
        the.Top = the.Top - counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunRightUp(the As Form)
    '
    'Usage:
    'FormRunRightUp Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Left = the.Left + counter
        the.Top = the.Top - counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunRightDown(the As Form)
    '
    'Usage:
    'FormRunRightDown Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Left = the.Left + counter
        the.Top = the.Top + counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub FormRunLeftDown(the As Form)
    '
    'Usage:
    'FormRunLeftDown Form1
    '
    Dim counter
    counter = the.Top


    Do: DoEvents
        counter = counter + 100
        the.Left = the.Left - counter
        the.Top = the.Top + counter
    Loop Until counter >= Screen.Width + the.Width
End Sub


Sub LimitText(KeyR, LimitDat)
    '
    'Usage:
    'LimitText KeyAscii, "ABC.1" 'Place This
    '     Code In The TextBox_KeyPressed Sub
    '
    ' Const
    Dim numbers$
    numbers$ = LimitDat


    If KeyR <> 8 Then


        If InStr(numbers, Chr(KeyR)) = 0 Then
            KeyR = 0
            Exit Sub
        End If
    End If
End Sub


Sub WebLink(WeBLnk)
    '
    'Usage:
    '
    Dim WL, nResult
    WL = "start.exe " & WeBLnk
    nResult = Shell(WL, vbHide)
End Sub


Public Sub ExecCmd(cmdline$)
    '
    ' Shell the Application then
    ' Wait for the shelled application
    ' to finish.
    '
    'Usage:
    'ExecCmd "calc.exe"
    '
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&
    start.cb = Len(start)
    Ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to fi
    '     nish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
End Sub


Sub DirSize(DirChk)
    '
    'Usage:
    'DirSize "c:\windows"
    'Msg = "Total bytes used = " + DirChkSiz
    '     e
    'MsgBx
    '
    Dim FileName As String
    Dim FileSize As Currency
    Dim Directory As String


    If Len(DirChk) = 3 Then
        Directory = DirChk
    Else
        Directory = DirChk & "\"
    End If
    FileName = Dir$(Directory & "*.*")
    FileSize = 0


    Do While FileName <> ""
        FileSize = FileSize + FileLen(Directory & FileName)
        FileName = Dir$
    Loop
    DirChkSize = Str$(FileSize)
End Sub


Sub SupportSound()
    '
    'Usage:
    ' SupportSound
    '
    'Return Value Supsound>> True = Yes - Fa
    '     lse = No
    '
    Dim i As Integer
    i = waveOutGetNumDevs()


    If i > 0 Then
        SupSound = True
    Else
        SupSound = False
    End If
End Sub


Function WindowsSysDir() As String
    '
    'Usage:
    ' WindowsSysDir
    ' Msg = VBSysDir
    ' msgbx
    '
    Dim Gwdvar As String, Gwdvar_Length As Integer
    Gwdvar = Space(255)
    Gwdvar_Length = GetSystemDirectory(Gwdvar, 255)
    VBSysDir = Left(Gwdvar, Gwdvar_Length)
End Function


Public Function AddBackslash(s As String) As String
    '
    'Used By Other Sub's
    '


    If Len(s) > 0 Then


        If Right$(s, 1) <> "\" Then
            AddBackslash = s + "\"
        Else
            AddBackslash = s
        End If
    Else
        AddBackslash = "\"
    End If
End Function


Public Function RemoveBackslash(s As String) As String
    '
    'Used By Other Sub's
    '
    Dim i As Integer
    i = Len(s)


    If i <> 0 Then


        If Right$(s, 1) = "\" Then
            RemoveBackslash = Left$(s, i - 1)
        Else
            RemoveBackslash = s
        End If
    Else
        RemoveBackslash = ""
    End If
End Function


Public Function GetWindowsDirectory() As String
    '
    'Usage:
    ' GetWindowsDirectory
    ' Msgbox GetWinDir
    '
    Dim s As String
    Dim i As Integer
    i = GetWindowsDirectoryA("", 0)
    s = Space(i)
    Call GetWindowsDirectoryA(s, i)
    GetWinDir = AddBackslash(Left$(s, i - 1))
End Function


Public Function FileExists(ByVal strPathName As String) As Integer
    '
    'Usage:
    ' FileExists "c:\test.exe"
    ' MsgBox IsFileThere
    '
    Dim intFileNum As Integer
    On Error Resume Next


    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    IsFileThere = IIf(Err, False, True)
    Close intFileNum
    Err = 0
End Function


Public Function GetPath(s As String) As String
    '
    'Usage:
    ' GetPath "c:\t.bat"
    ' MsgBox FlPath
    '
    Dim i As Integer
    Dim J As Integer
    i = 0
    J = 0
    i = InStr(s, "\")


    Do While i <> 0
        J = i
        i = InStr(J + 1, s, "\")
    Loop


    If J = 0 Then
        FlPath = ""
    Else
        FlPath = Left$(s, J)
    End If
End Function


Public Function GetFile(s As String) As String
    '
    'Usage:
    ' GetFile "c:\t.bat"
    ' MsgBox FlName
    '
    Dim i As Integer
    Dim J As Integer
    i = 0
    J = 0
    i = InStr(s, "\")


    Do While i <> 0
        J = i
        i = InStr(J + 1, s, "\")
    Loop


    If J = 0 Then
        FlName = ""
    Else
        FlName = Right$(s, Len(s) - J)
    End If
End Function


Public Function sDriveType(sDrive As String) As String
    '
    'Usage:
    ' sDriveType "c"
    ' MsgBox sDType
    '
    Dim lRet As Long
    lRet = GetDriveTypeA(sDrive & ":\")


    Select Case lRet
        Case 0
        sDType = "Unknown"
        Case 1
        sDType = "Drive Not Found"
        Case DRIVE_CDROM:
        sDType = "CD-ROM Drive"
        Case DRIVE_REMOVABLE:
        sDType = "Removable Drive"
        Case DRIVE_FIXED:
        sDType = "Fixed Drive"
        Case DRIVE_REMOTE:
        sDType = "Remote Drive"
    End Select
End Function


Public Function ShellDelete(ParamArray vntFileName() As Variant) As Boolean
    '
    'Usage:
    ' ShellDelete "c:\test.exe"
    '
    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT


    For i = LBound(vntFileName) To UBound(vntFileName)
        sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar


    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
    End With
    i = SHFileOperation(SHFileOp)


    If i = 0 Then
        DelConFirm = True
    Else
        DelConFirm = False
    End If
End Function


Public Sub ShadeForm(f As Form, Optional StartColor As Variant, Optional Fstep As Variant, Optional Cstep As Variant)
    '
    'Colors:
    'vbBlack
    'vbRed
    'vbGreen
    'vbYellow
    'vbBlue
    'vbMagenta
    'vbCyan
    'vbWhite
    '
    ' StartColor is what color to start with
    '     .
    '(Default = vbBlue)
    '
    ' Fstep is the number of steps to use to
    '     fill the form.
    '(Default = 64)
    '
    ' Cstep is the color step (change in col
    '     or per step).
    '(Default = 4)
    '
    'Usage:
    ' ShadeForm StartUp, vbRed, 64, 4
    '
    Dim FillStep As Single
    Dim c As Long
    Dim FillArea As RECT
    Dim i As Integer
    Dim oldm As Integer
    Dim hBrush As Long
    Dim C2(1 To 3) As Long
    Dim cs2(1 To 3) As Long
    Dim fs As Long
    Dim cs As Integer
    
    fs = IIf(IsMissing(Fstep), 64, CLng(Fstep))
    cs = IIf(IsMissing(Cstep), 4, CInt(Cstep))
    c = IIf(IsMissing(StartColor), vbBlue, CLng(StartColor))
    oldm = f.ScaleMode
    f.ScaleMode = vbPixels
    FillStep = f.ScaleHeight / fs
    FillArea.Left = 0
    FillArea.Right = f.ScaleWidth
    FillArea.Top = 0
    C2(1) = c And 255#
    cs2(1) = IIf(C2(1) > 0, cs, 0)
    C2(2) = (c \ 256#) And 255#
    cs2(2) = IIf(C2(2) > 0, cs, 0)
    C2(3) = (c \ 65536#) And 255#
    cs2(3) = IIf(C2(3) > 0, cs, 0)


    For i = 1 To fs
        FillArea.Bottom = FillStep * i
        hBrush = CreateSolidBrush(RGB(C2(1), C2(2), C2(3)))
        FillRect f.hdc, FillArea, hBrush
        DeleteObject hBrush
        
        C2(1) = (C2(1) - cs2(1)) And 255#
        C2(2) = (C2(2) - cs2(2)) And 255#
        C2(3) = (C2(3) - cs2(3)) And 255#
        
        FillArea.Top = FillArea.Bottom
    Next i
    f.ScaleMode = oldm
End Sub


Public Sub HideMouse()
    '
    'Usage:
    ' HideMouse
    '
    Dim result As Integer


    Do
        lShowCursor = lShowCursor - 1
        result = ShowCursor(False)
    Loop Until result < 0
End Sub


Public Sub ShowMouse()
    '
    'Usage:
    'ShowMouse
    '


    If lShowCursor > 0 Then


        Do While lShowCursor <> 0
            ShowCursor (False)
            lShowCursor = lShowCursor - 1
        Loop
    ElseIf lShowCursor < 0 Then


        Do While lShowCursor <> 0
            ShowCursor (True)
            lShowCursor = lShowCursor + 1
        Loop
    End If
End Sub


Public Function CanPlaySound() As Integer
    '
    'Usage:
    'CanPlaySound
    'Msgbox Playinfo
    '
    Dim i As Integer
    i = AUDIO_NONE


    If waveOutGetNumDevs > 0 Then
        i = AUDIO_WAVE
    End If


    If midiOutGetNumDevs > 0 Then
        i = i + AUDIO_MIDI
    End If
    If i = 1 Then Playinfo = "WAV ONLY"
    If i = 2 Then Playinfo = "MID ONLY"
    If i = 3 Then Playinfo = "WAV AND MID"
End Function


Public Sub GetBytes(ChkDrive)
    '
    'Usage:
    ' GetBytes
    ' Msgbox DriveFreeSpace
    '
    Dim ApiRes As Long
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TotalNumberOfClusters As Long
    Dim FreeBytes As Long
    Dim drvStr As String
    Dim spaceInt As Integer
    drvStr = ChkDrive
    spaceInt = InStr(drvStr, " ")
    If spaceInt > 0 Then drvStr = Left$(drvStr, spaceInt - 1)
    If Right$(drvStr, 1) <> "\" Then drvStr = drvStr & "\"
    Dim NumberOFreeClusters
    ApiRes = GetDiskFreeSpace(drvStr, SectorsPerCluster, BytesPerSector, NumberOFreeClusters, TotalNumberOfClusters)
    FreeBytes = NumberOFreeClusters * SectorsPerCluster * BytesPerSector
    DriveFreeSpace = FreeBytes
End Sub


Public Sub FormatFloppy()
    '
    'Usage:
    ' FormatFloppy
    '
    Dim sBuffer As String, Windir As String, Procs As String, X
    Dim lResult As Long
    Dim K
    sBuffer = String$(255, 0)
    lResult = GetWindowDirectory(sBuffer, Len(sBuffer))
    Windir = Trim(sBuffer)
    Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,SHFormatDrive"
    Call CenterDialog("Format - 3½ Floppy (A:)")
    X = Shell(Procs, 1)
    Call CenterDialog("Format - 3½ Floppy (A:)")
    K = LockWindowUpdate(0)
End Sub


Public Sub CenterDialog(WinText As String)
    '
    'This Sub Is Used By FormatFloppy
    '


    DoEvents
        On Error Resume Next
        Dim D3 As Long
        D3 = LockWindowUpdate(GetDesktopWindow())
        Dim wdth%
        Dim hght%
        Dim Scrwdth%
        Dim Scrhght%
        Dim lpDlgRect As RECT
        Dim lpdskrect As RECT
        Dim X%, Y%
        Dim hTaskBar As Long
        hTaskBar = FindWindow(0&, WinText)
        Call GetWindowRect(hTaskBar, lpDlgRect)
        wdth% = lpDlgRect.Right - lpDlgRect.Left
        hght% = lpDlgRect.Bottom - lpDlgRect.Top
        Call GetWindowRect(GetDesktopWindow(), lpdskrect)
        Scrwdth% = lpdskrect.Right - lpdskrect.Left
        Scrhght% = lpdskrect.Bottom - lpdskrect.Top
        
        X% = (Scrwdth% - wdth%) / 2
        Y% = (Scrhght% - hght%) / 2
        Call SetWindowPos(hTaskBar, 0, X%, Y%, 0, 0, SWP_NOZORDER Or SWP_NOSIZE)


        DoEvents
        End Sub


Public Sub ChkFileStats(File_Name_To_Chk)
    '
    'Usage:
    ' ChkFileStats "C:\TEST.EXE"
    ' MsgBox FileInfoName 'File Name Without
    '     Path
    ' MsgBox FileInfoPathName ' File Name Wi
    '     th Path
    ' MsgBox FileInfoSize 'File Size
    ' MsgBox FileInfoLastModified 'File Last
    '     Modified
    ' MsgBox FileInfoLastAccessed 'File Last
    '     Accessed
    ' MsgBox FileInfoAttributeHidden 'File A
    '     ttribute Hidden? True/False
    ' MsgBox FileInfoAttributeSystem 'File A
    '     ttribute System? True/False
    ' MsgBox FileInfoAttributeReadOnly 'File
    '     Attribute Read Only? True/False
    ' MsgBox FileInfoAttributeArchive 'File
    '     Attribute Archive? True/False
    ' MsgBox FileInfoAttributeTemporary 'Fil
    '     e Attribute Temporary? True/False
    ' MsgBox FileInfoAttributeNormal 'File A
    '     ttribute Normal? True/False
    ' MsgBox FileInfoAttributeCompressed 'Fi
    '     le Attribute Compressed? True/False
    '
    Dim ftime As SYSTEMTIME
    Dim tfilename As String
    tfilename = File_Name_To_Chk
    Dim filedata As WIN32_FIND_DATA
    filedata = Findfile("c:\command.com")
    FileInfoName = UCase$(File_Name_To_Chk)
    FileInfoPathName = UCase$(tfilename)
    GetFile FileInfoName
    FileInfoName = FlName


    If filedata.nFileSizeHigh = 0 Then
        
        FileInfoSize = filedata.nFileSizeLow & " Bytes"
    Else
        FileInfoSize = filedata.nFileSizeHigh & "Bytes"
    End If
    Call FileTimeToSystemTime(filedata.ftCreationTimeAs, ftime)
    Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime)
    FileInfoLastModified = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear
    Call FileTimeToSystemTime(filedata.ftLastAccessTime, ftime)
    FileInfoLastAccessed = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
        
        FileInfoAttributeHidden = True
    Else
        FileInfoAttributeHidden = False
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM Then
        FileInfoAttributeSystem = True
    Else
        FileInfoAttributeSystem = False
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
        FileInfoAttributeReadOnly = True
    Else
        FileInfoAttributeReadOnly = False
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE Then
        FileInfoAttributeArchive = True
    Else
        FileInfoAttributeArchive = False
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY Then
        FileInfoAttributeTemporary = True
    Else
        FileInfoAttributeTemporary = True
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL Then
        FileInfoAttributeNormal = True
    Else
        FileInfoAttributeNormal = False
    End If


    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED Then
        FileInfoAttributeCompressed = True
    Else
        FileInfoAttributeCompressed = False
    End If
End Sub


Public Sub FindDosWin(ByVal WndCap As String)
    '
    'Usage:
    'FindDosWin UCase$(Text11.Text)
    'Msgbox DOSWinActive 'True = DOS Window
    '     Is Active \ False = DOS Window Is Not Ac
    '     tive
    '
    Dim hWndFrame As Long
    hWndFrame = FindWindowPartial(WndCap)


    If hWndFrame = 0 Then


        DOSWinActive = False
            Exit Sub
        End If


        DOSWinActive = True
        End Sub
    


Sub makeShortCut(sExecutable As String, sShortcut, sArguments, PlaceInWhere)
    '
    'Usage:
    'makeShortCut "c:\test.exe", Testexe, ""
    '     , (DESKTOP or STARTMENU or PATH TO PLACE
    '     SHORTCUT)
    '
    On Error GoTo py
    Dim lRet As Integer
    Dim DestPth, CreatedPth
    PlaceInWhere = UCase$(PlaceInWhere)
    Short_Name sExecutable
    sExecutable = ShortFName
    FileExists sExecutable


    If IsFileThere = False Then
        Msg = "ERROR! Short Cut File You Want To Link To Does Not Exists"
        MsgBx
        Exit Sub
    End If


    If PlaceInWhere = "STARTMENU" Then
        lRet = fCreateShellLink("", sShortcut, sExecutable, sArguments)
        Exit Sub
    End If
    GetWindowsDirectory


    If PlaceInWhere = "DESKTOP" Then
        CreatedPth = GetWinDir & "startm~1\programs\" & sShortcut & ".pif"
        DestPth = GetWinDir & "desktop\" & sShortcut & ".pif"
    Else
        CreatedPth = GetWinDir & "startm~1\programs\" & sShortcut & ".pif"
        DestPth = PlaceInWhere & sShortcut & ".pif"
        lRet = fCreateShellLink("", sShortcut, sExecutable, sArguments)
    End If


    If PlaceInWhere = "DESKTOP" Then
        FileExists DestPth


        If IsFileThere = True Then
            ShellDelete DestPth
        End If
        lRet = fCreateShellLink("", sShortcut, sExecutable, sArguments)
    End If
    Name CreatedPth As DestPth
    Exit Sub
py:
End Sub


Public Function Short_Name(Long_Path As String) As String
    '
    'Usage:
    'Short_Name "C:\PathNameToProgram\test.e
    '     xe"
    'ShortFname
    Dim Short_Path As String
    Dim Answer As Long
    Short_Path = Space(250)
    Answer = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))
    ShortFName = Left$(Short_Path, Answer)
End Function


Public Sub TerminateTask(app_name As String)
    '
    'Usage:
    ' TerminateTask "Active WIndow Name You
    '     Want To Kill"
    '
    Target = app_name


EnumWindows AddressOf EnumCallback, 0
End Sub


Public Sub WriteINI(FileName As String, Section As String, Key As String, Text As String)
    '
    'Usage:
    'WriteINI "c:\test.ini", "section name",
    '     "key name", "text data"
    '
    WritePrivateProfileString Section, Key, Text, FileName
End Sub


Public Function ReadINI(FileName As String, Section As String, Key As String)
    '
    'Usage:
    'ReturnINIdat = ReadINI("c:\test.ini", "
    '     section name", "key name")
    'Msgbox INIFileFound 'True = File Found
    '     \ False = File Found
    Dim RetLen
    INIFileFound = True
    FileExists FileName


    If IsFileThere = False Then
        INIFileFound = False
        Exit Function
    End If
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), FileName)
    Ret = Left$(Ret, RetLen)
    ReadINI = Ret
End Function


Sub GetKeyboardInfo()
    Dim r As Long
    Dim t As String
    Dim K As Long
    Dim Q As Long
    K = GetKeyboardType(0)
    If K = 1 Then t = "PC or compatible 83-key keyboard"
    If K = 2 Then t = "Olivetti 102-key keyboard"
    If K = 3 Then t = "AT or compatible 84-key keyboard"
    If K = 4 Then t = "Enhanced(IBM) 101-102-key keyboard"
    If K = 5 Then t = "Nokia 1050 keyboard"
    If K = 6 Then t = "Nokia 9140 keyboard"
    If K = 7 Then t = "Japanese keyboard"
    KeyBoardType = t
    Q = SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, r, 0)
    KeyBoardRepeatDelay = r
    Q = SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, r, 0)
    KeyBoardRepeatSpeed = r
    KeyBoardCaretFlashSpeed = GetCaretBlinkTime
End Sub
'here


Sub OpenCD_ROMDoor()
    '
    'Usage:
    ' OpenCD_ROMDoor
    '
    'retvalue = mciSendString("set CDAudio d
    '     oor open", returnstring, 127, 0)
End Sub


Sub CloseCD_ROMDoor()
    '
    'Usage:
    ' CloseCD_ROMDoor
    '
    'retvalue = mciSendString("set CDAudio d
    '     oor closed", returnstring, 127, 0)
End Sub


Sub Search32(dPath$, dpattern$, SFileName)
    '
    'Usage:
    'Search32 "C:\", "*.WAV", "c:\DIR.TXT"
    '||| ||Name Of File To Save Files Found.
    '
    '|Files To Search For Wildcards Can Be U
    '     sed.
    'Directory To Start Search In. If Path =
    '     "C:\Windows" The Search Will Search
    'The Windows Directory Then All It's Sub
    '     Directories.
    '
    Close #10
    Open SFileName For Output As 10
    Call dirloop(dPath$, dpattern$)
    Close #10
End Sub


Sub dirloop(thispath As String, thispattern As String)
    '
    'Used By Search32
    '
    Dim thisfile, thesefiles, thesedirs, X, checkfile
    If Right$(thispath, 1) <> "\" Then thispath = thispath + "\"
    thisfile = Dir$(thispath + thispattern, 0)


    Do While thisfile <> ""
        Print #10, LCase$(thispath + thisfile)
        thisfile = Dir$
    Loop
    
    thisfile = Dir$(thispath + "*.", 0)
    thesefiles = 0
    ReDim filelist(10)


    Do While thisfile <> ""
        thesefiles = thesefiles + 1


        If (thesefiles Mod 10) = 0 Then
            ReDim Preserve filelist(thesefiles + 10)
        End If
        filelist(thesefiles) = thisfile
        thisfile = Dir$
    Loop
    thisfile = Dir$(thispath + "*.", 16)
    checkfile = 1
    thesedirs = 0
    ReDim dirlist(10)


    Do While thisfile <> ""


        If thisfile = "." Or thisfile = ".." Then
        ElseIf thisfile = filelist(checkfile) Then
            checkfile = checkfile + 1
        Else
            thesedirs = thesedirs + 1
            If (thesedirs Mod 10) = 0 Then ReDim Preserve dirlist(thesedirs + 10)
            dirlist(thesedirs) = thisfile
        End If
        thisfile = Dir$
    Loop


    For X = 1 To thesedirs
        Call dirloop(thispath + dirlist(X), thispattern): DoEvents
    Next X
End Sub


Sub GetDate()
    'Usage:
    ' GetDate
    '
    ' CurDate = Current Computer Date
    '
    CurDate = Date
End Sub


Sub ClearAllTextBoxes(frmTarget As Form)
    'Usage:
    'ClearAllTextBoxes Form1
    '
    Dim i, ctrltarget


    For i = 0 To (frmTarget.Controls.Count - 1)
        Set ctrltarget = frmTarget.Controls(i)


        If TypeOf ctrltarget Is TextBox Then
            ctrltarget.Text = ""
        End If
    Next i
End Sub


Sub GetAPPpath()
    Dim X
    X = App.Path
    If Right$(X, 1) <> "\" Then X = X + "\"
    AppPath = UCase$(X)
End Sub


Sub DallorPeriodSet(Tdat As TextBox)
    'Usage:
    '
    ' DallorPeriodSet Text1
    ' msgbox DallorGet
    '
    Dim a, b, Mrk1, c, d, C1, C2, C3, C4, C5
    DallorGet = "0"
    If Tdat = "" Or Val(Tdat) = 0 Then Exit Sub
    Mrk1 = False
    a = Len(Tdat.Text) + 1
    b = 1
    d = 0


    Do Until b = a
        c = Mid$(Tdat, b, 1)
        If c = "." Then Mrk1 = True
        If Mrk1 = True Then d = d + 1
        DBa(b) = c
        b = b + 1
    Loop
    d = d - 1
    If d = 0 Then d = 2
    c = Tdat
    'no period


    If d = -1 And Mrk1 = False Then
        c = c & ".00"
        DallorGet = c
        Exit Sub
    End If
    'over flow 5.00573


    If d > 2 Then
        Dim v
        d = False


        For b = Len(c) To 1 Step -1


            If DBa(b) = "." Then
            Else


                If Val(DBa(b)) >= 5 Then


                    If b - 2 <= 0 Then
                        '
                    Else


                        If DBa(b - 2) = "." Then
                            d = True
                        Else


                            If b - 1 <= 0 Then
                                '
                            Else
                                If d = False Then DBa(b - 1) = Val(DBa(b - 1)) + 1
                            End If
                        End If
                    End If
                End If
                Dim t, Y
                Y = c
                c = ""


                For t = 1 To Len(Y)
                    c = c & DBa(t)
                Next t
            End If
        Next b
        Dim e, f
        a = 1
        b = ""
        e = 0
        Mrk1 = False


        Do Until a = Len(c) + 1
            d = Mid$(c, a, 1)
            If d = "." Then Mrk1 = True
            If Mrk1 = False Then f = f & d


            If Mrk1 = True And e <= 2 Then
                f = f & d
                e = e + 1
            End If
            a = a + 1
        Loop
        DallorClean f
        f = DallorGet
        DallorGet = f
        Exit Sub
    End If


    For b = 1 To d
        c = c & "0"
    Next b
    DallorClean c
    c = DallorGet
    DallorGet = c
End Sub


Sub DallorClean(DDat)
    On Error GoTo yu
    Dim a, b, c, f, Mrk1
    DallorGet = ""
    a = 1
    c = 0
    Mrk1 = False


    Do Until a = Len(DDat) + 1
        b = Mid$(DDat, a, 1)
        If b = "." Then Mrk1 = True
        If Mrk1 = False Then f = f & b


        If Mrk1 = True Then
            c = c + 1


            If c <= 3 Then
                f = f & b
            End If
        End If
        a = a + 1
    Loop
    a = 1
    Mrk1 = False


    Do Until a = Len(f) + 1


        If Mid$(f, a, 1) = "." Then
            b = a
            Mrk1 = True
        End If
        a = a + 1
    Loop
    'If Mrk1 = False Then f = f & "."
    If Val(Mid$(f, b, Len(f))) = 3 Then f = f & "00"
    If Val(Mid$(f, b, Len(f))) = 4 Then f = f & "0"
    If Mrk1 = False Then f = f & ".00"
    DallorGet = f
    Exit Sub
yu:
    Exit Sub
End Sub


Sub addletter(frm As Form, newletter As String, oldcaption As String)
    'Used By AnimateCaption
    Dim total As Integer, spaces As Integer, temp, X
    total = Len(temp)
    spaces = (frm.Width / 50) - (total)


    For X = spaces To Len(temp) Step -1
        frm.Caption = oldcaption & Space(X) & newletter


        DoEvents
        Next X
    End Sub


Sub AnimateCaption(CapData, MEfrm As Form)
    'Usage:
    '
    ' AnimateCaption Form1
    '
    MEfrm.Show
    MEfrm.Caption = ""
    Dim a, t
    a = CapData


    For t = 1 To Len(a)
        addletter MEfrm, Mid$(a, t, 1), MEfrm.Caption
    Next t
End Sub


Sub DisableX(FormNameHere As Form)
    'Usage:
    '
    ' DisableX Form1
    '
    Dim hMenu As Long
    Dim menuItemCount As Long
    hMenu = GetSystemMenu(FormNameHere.hwnd, 0)


    If hMenu Then
        menuItemCount = GetMenuItemCount(hMenu)
        Call RemoveMenu(hMenu, menuItemCount - 1, MF_REMOVE Or MF_BYPOSITION)
        Call RemoveMenu(hMenu, menuItemCount - 2, MF_REMOVE Or MF_BYPOSITION)
        Call DrawMenuBar(FormNameHere.hwnd)
    End If
    
End Sub



