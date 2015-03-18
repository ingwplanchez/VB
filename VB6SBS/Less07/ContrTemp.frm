VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Contraseña"
   ClientHeight    =   1770
   ClientLeft      =   2355
   ClientTop       =   2025
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4440
   Begin VB.CommandButton Command1 
      Caption         =   "Probar contraseña"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
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
      Caption         =   "Introduzca su contraseña en 15 segundos"
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
    If Text1.Text = "secreto" Then
        Timer1.Enabled = False
        MsgBox ("¡Bienvenido al sistema!")
        End
    Else
        MsgBox ("Lo siento, amigo, no le conozco.")
    End If
End Sub

Private Sub Timer1_Timer()
    MsgBox ("Lo siento, su tiempo ha expirado.")
    End
End Sub

