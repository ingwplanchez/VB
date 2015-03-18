VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Password"
   ClientHeight    =   1770
   ClientLeft      =   2355
   ClientTop       =   2025
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4440
   Begin VB.CommandButton Command1 
      Caption         =   "Try Password"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   240
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your password within 15 seconds."
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Text1.Text = "secret" Then
        Timer1.Enabled = False
        MsgBox ("Welcome to the system!")
        End
    Else
        MsgBox ("Sorry, friend, I don't know you.")
    End If
End Sub

Private Sub Timer1_Timer()
    MsgBox ("Sorry, your time is up.")
    End
End Sub

