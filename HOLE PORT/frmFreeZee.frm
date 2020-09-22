VERSION 5.00
Begin VB.Form frmFreeZee 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   4320
   End
End
Attribute VB_Name = "frmFreeZee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyQ And Shift = 1 Then
        mdonTOP.SetOffTop frmFreeZee
        End
    Else
        Beep 1, 1
    End If
End Sub

Private Sub Form_Load()
Hide_Program_In_CTRL_ALT_Delete
App.TaskVisible = False
frmFreeZee.Height = Screen.Height
frmFreeZee.Width = Screen.Width
frmFreeZee.Top = 0
frmFreeZee.Left = 0
keybd_event vbKeySnapshot, 1, 0&, 0&
DoEvents
frmFreeZee.Picture = Clipboard.GetData(vbCFBitmap)
mdonTOP.SetOnTop frmFreeZee
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Beep 1, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mdonTOP.SetOffTop frmFreeZee
End Sub

Private Sub Timer1_Timer()
Hide_Program_In_CTRL_ALT_Delete
End Sub
