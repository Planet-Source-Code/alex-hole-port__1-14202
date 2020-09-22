Attribute VB_Name = "mdlFileManager"

Public Function FillFileList(Index As Integer)
Dim newdir As String


On Error GoTo err


frmMain.wskF(Index).SendData "filelist.Clear"

PauseData.Pause 1

If frmMain.Dir1.Path = "C:\" Then
 For X = 0 To frmMain.Dir1.ListCount - 1
 newdir = "<" + Right(frmMain.Dir1.List(X), Len(frmMain.Dir1.List(X)) - Len(frmMain.Dir1.Path)) + ">"
 frmMain.wskF(Index).SendData "\newdir" & newdir
 PauseData.Pause 0.5
 Next X
Else
 frmMain.wskF(Index).SendData "filelist.AddItem" & "<..>"
 PauseData.Pause 0.5
 For X = 0 To frmMain.Dir1.ListCount - 1
 newdir = "<" + Right(frmMain.Dir1.List(X), Len(frmMain.Dir1.List(X)) - Len(frmMain.Dir1.Path) - 1) + ">"
 frmMain.wskF(Index).SendData "\newdir" & newdir
 PauseData.Pause 0.5
 Next X
End If

For X = 1 To frmMain.File1.ListCount - 1
frmMain.wskF(Index).SendData "<additem>" & frmMain.File1.List(X)
PauseData.Pause 0.5
Next X
frmMain.wskF(Index).SendData "<txtDir>" & frmMain.Dir1.Path
PauseData.Pause 0.5
frmMain.wskF(Index).SendData "<<FINISHED>>"

err:
Exit Function

End Function


Public Sub Status(ByVal Msg$)
Dim w As Winsock

On Error GoTo err


For Each w In frmMain.wskServer
If w.State = 7 Then frmMain.wskServer(w.Index).SendData "StatusInfo" & Msg$
Next w

err:
Exit Sub

End Sub
