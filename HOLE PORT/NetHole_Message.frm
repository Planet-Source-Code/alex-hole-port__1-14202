VERSION 5.00
Begin VB.Form NetHole_Message 
   BackColor       =   &H80000011&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NEM"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5325
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtAnswer 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   1335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "NetHole_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
frmMain.wskServer(Indx).SendData "X\D" & txtAnswer.Text
txtAnswer.Text = ""
End Sub

