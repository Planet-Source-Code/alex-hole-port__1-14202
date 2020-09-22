VERSION 5.00
Begin VB.Form frmPic 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9360
      Top             =   120
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyQ And Shift = 1 Then
           mdonTOP.SetOffTop frmPic
        End
    Else
        Beep 1, 1
    End If
End Sub

Private Sub Form_Load()
Hide_Program_In_CTRL_ALT_Delete
mdonTOP.SetOnTop frmPic
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Beep 1, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mdonTOP.SetOffTop frmPic
End Sub

Private Sub Timer1_Timer()
ctrl.Hide_Program_In_CTRL_ALT_Delete
End Sub
