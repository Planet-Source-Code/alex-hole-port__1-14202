VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "THIS IS JUST A DEMO....FINAL PROJECT WILL BE DONE IN A WEEK..."
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "Created By Alex Kokoshenko"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Beep
Form1.Hide
client.Show
End Sub
