VERSION 5.00
Begin VB.Form frmManager 
   BorderStyle     =   0  'None
   Caption         =   "File Manager"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8280
      Width           =   6495
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   6495
   End
   Begin VB.ListBox filelist 
      Height          =   7470
      ItemData        =   "frmManager.frx":0000
      Left            =   120
      List            =   "frmManager.frx":0007
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label lblFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   6495
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
If filelist.Enabled = True Then
Unload frmManager
Else
MsgBox "Cannot unload because download is still in progress...", vbCritical, "BitcH"
End If
End Sub

Private Sub filelist_DblClick()
 If Dir1 = "C:\" Then
  frmManager.lblFile.Caption = Dir1 + frmManager.filelist.Text
 Else
  frmManager.lblFile.Caption = Dir1 + "\" + frmManager.filelist.Text
 End If
End Sub

Private Sub filelist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
x = InStr(1, frmManager.filelist.Text, "<")
If x = 1 Then
 y = InStr(1, frmManager.filelist.Text, ">")
 z = Mid(frmManager.filelist.Text, x + 1, y - 2)
 If Dir1 = "C:\" Then
  Dir1 = Dir1 + z
 Else
  Dir1 = Dir1 + "\" + z
 End If
 client.wskF.SendData "SETDIR" & Dir1
 PauseData.Pause 2
 client.wskF.SendData "FillFileList"
End If
Else
client.wskF.SendData "Dir1.Path"
PauseData.Pause 1
If frmManager.filelist.Text = "<..>" Then
 frmManager.lblFile.Caption = ""
 Exit Sub
End If
x = InStr(1, frmManager.filelist.Text, "<")
If x = 1 Then
   y = InStr(1, frmManager.filelist.Text, ">")
   z = Mid(frmManager.filelist.Text, x + 1, y - 2)
  If Dir1 = "C:\" Then
  frmManager.lblFile.Caption = Dir1 + z + "\"
  Else
  frmManager.lblFile.Caption = Dir1 + "\" + z + "\"
  End If
End If
End If
End Sub

