Attribute VB_Name = "mdlReboot"
'Call prcedure ExitWin with one of the contants
'listed below.
'
'eg. ExitWin(EW_REBOOT)
'
Option Explicit

Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer

Global Const EW_REBOOT = &H43
Global Const EW_RESTART = &H42
Global Const EW_EXIT = 0
Function ExitWin(ActionIn As Long) As Integer


    Dim intRetVal As Integer
    intRetVal = ExitWindows(ActionIn, 0)
    ExitWin = intRetVal

End Function


