VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Burn Barrel"
   ClientHeight    =   3570
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4755
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65
      Left            =   240
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      Picture         =   "Smoke.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   2400
      Picture         =   "Smoke.frx":030A
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image5 
      DragIcon        =   "Smoke.frx":0614
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1680
      Picture         =   "Smoke.frx":091E
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image4 
      DragIcon        =   "Smoke.frx":0C28
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   960
      Picture         =   "Smoke.frx":0F32
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image3 
      DragIcon        =   "Smoke.frx":123C
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   240
      Picture         =   "Smoke.frx":1546
      Tag             =   "Fire"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image2 
      DragIcon        =   "Smoke.frx":1850
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   240
      Picture         =   "Smoke.frx":1B5A
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3480
      Picture         =   "Smoke.frx":1E64
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Throw everything away, and then drop in the match."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = False
    If Source.Tag = "Fire" Then
        Image1.Picture = Image6.Picture
        Picture1.Visible = True
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If Picture1.Top > 0 Then
        Picture1.Move Picture1.Left - 50, Picture1.Top - 75
    Else
        Picture1.Visible = False
        Timer1.Enabled = False
    End If
End Sub





