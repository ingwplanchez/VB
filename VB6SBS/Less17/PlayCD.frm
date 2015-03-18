VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Audio CD Player"
   ClientHeight    =   2520
   ClientLeft      =   2340
   ClientTop       =   2595
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play CD"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1920
      Picture         =   "PlayCD.frx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Place an audio CD in CD-ROM drive and click Play CD!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    'specify CD Audio type (from CD-ROM drive)
    MMControl1.DeviceType = "CDAudio"
    MMControl1.Command = "Open"
    Image1.Visible = True
    End Sub

Private Sub cmdQuit_Click()
    'stop CD audio if quit button clicked
    MMControl1.Command = "Stop"
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'always close device when finished
    MMControl1.Command = "Close"
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    'hide instruction bitmap before playing CD
    Image1.Visible = False
End Sub
