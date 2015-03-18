VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Burn Barrel"
   ClientHeight    =   3570
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4650
   Begin VB.Image Image6 
      Height          =   855
      Left            =   2520
      Picture         =   "Dragdrop.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image5 
      DragIcon        =   "Dragdrop.frx":030A
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1680
      Picture         =   "Dragdrop.frx":0614
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image4 
      DragIcon        =   "Dragdrop.frx":091E
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   960
      Picture         =   "Dragdrop.frx":0C28
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image3 
      DragIcon        =   "Dragdrop.frx":0F32
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   240
      Picture         =   "Dragdrop.frx":123C
      Tag             =   "Fire"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image2 
      DragIcon        =   "Dragdrop.frx":1546
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   240
      Picture         =   "Dragdrop.frx":1850
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3600
      Picture         =   "Dragdrop.frx":1B5A
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
    End If
End Sub




