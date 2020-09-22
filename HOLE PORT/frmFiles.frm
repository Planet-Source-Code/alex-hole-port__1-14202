VERSION 5.00
Begin VB.Form frmFiles 
   BackColor       =   &H80000011&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "files"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000011&
      Caption         =   "Execute file"
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000011&
      Caption         =   "Download file"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000011&
      Caption         =   "Delete file"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim LoopIndex
Dim resultQ

If Option1(1).value = True Then

For LoopIndex = 0 To client.lstFiles.ListCount - 1

If client.lstFiles.Selected(LoopIndex) Then
    
    PauseData.Pause 0.3
    
    client.wskClient.SendData "\fileSize" & client.lstFiles.List(LoopIndex)
    
    PauseData.Pause 0.5
    
resultQ = MsgBox("Do you wish to download file  " & client.lstFiles.List(LoopIndex) & " it is " & flSize & " bytes" & " ?", vbOKCancel, "??")
 
 If resultQ = vbOK Then
 client.cmSave.ShowSave
 Else
 Exit Sub
 End If
 
 If Not vbCancel Then
 
SavePath = client.cmSave.FileName
 
PauseData.Pause 1

client.lblSize.Caption = flSize
client.ProgressBar1.Min = 0
client.ProgressBar1.Max = flSize
 
client.wskClient.SendData "\filePath" & client.lstFiles.List(LoopIndex)

End If
End If

Next LoopIndex

ElseIf Option1(0).value = True Then

    For LoopIndex = 0 To client.lstFiles.ListCount - 1

    If client.lstFiles.Selected(LoopIndex) Then
    
    PauseData.Pause 0.3
    
    client.wskClient.SendData "\deleteFile" & client.lstFiles.List(LoopIndex)
    
    End If
    
    Next LoopIndex

Else

For LoopIndex = 0 To client.lstFiles.ListCount - 1

If client.lstFiles.Selected(LoopIndex) Then
    
    PauseData.Pause 0.3
    
    client.wskClient.SendData "\execFile" & "," & client.lstFiles.List(LoopIndex)
    
End If

Next LoopIndex

End If

frmFiles.Hide
End Sub

Private Sub Option1_Click(Index As Integer)


If Option1(0).value = True Then
    Option1(1).value = False And Option1(2).value = False
ElseIf Option1(1).value = True Then
    Option1(0).value = False And Option1(2).value = False
ElseIf Option1(2).value = True Then
    Option1(0).value = False And Option1(1).value = False
End If


End Sub
