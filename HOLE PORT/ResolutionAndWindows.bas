Attribute VB_Name = "ResolutionAndWindows"
Option Explicit

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type DEVMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


Public Sub ChangeResolution(ByVal optRes As Integer)
Dim DevM    As DEVMODE
Dim lResult As Long
Dim iAns    As Integer
'
' Retrieve info about the current graphics mode
' on the current display device.
'
lResult = EnumDisplaySettings(0, 0, DevM)
'
' Set the new resolution. Don't change the color
' depth so a restart is not necessary.
'
With DevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
    If optRes = 0 Then
        .dmPelsWidth = 640  'ScreenWidth
        .dmPelsHeight = 480 'ScreenHeight
    ElseIf optRes = 1 Then
        .dmPelsWidth = 800
        .dmPelsHeight = 600
    Else
        .dmPelsWidth = 1024
        .dmPelsHeight = 768
    End If
    '.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
End With
'
' Change the display settings to the specified graphics mode.
'
lResult = ChangeDisplaySettings(DevM, CDS_TEST)
Select Case lResult
    Case DISP_CHANGE_RESTART
        iAns = MsgBox("You must restart your computer to apply these changes." & _
            vbCrLf & vbCrLf & "Do you want to restart now?", _
            vbYesNo + vbSystemModal, "Screen Resolution")
        If iAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
        Call ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        MsgBox "Screen resolution changed", vbInformation, "Resolution Changed"
    Case Else
        MsgBox "Mode not supported", vbSystemModal, "Error"
End Select
End Sub

Public Sub RebootComp(ByVal optShut As Integer)
Dim lMode As Long

If optShut = 0 Then
    lMode = EWX_LOGOFF
ElseIf optShut = 1 Then
    lMode = EWX_REBOOT
ElseIf optShut = 2 Then
    lMode = EWX_SHUTDOWN
Else: lMode = EWX_FORCE
End If

Call ExitWindowsEx(lMode, 0)
End Sub
