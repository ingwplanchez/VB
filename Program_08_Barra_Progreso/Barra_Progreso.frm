VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "PROGRESO: "
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdReiniciar 
      Caption         =   "&Reiniciar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Barra de Porgreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label BarraProgreso 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Porcentaje As Integer
Private Sub cmdReiniciar_Click()
    Label3.Caption = 0 & "%"
    Form1.Caption = "PROGRESO: " & 0 & "%"
    BarraProgreso.Width = 15
    Timer1.Enabled = True 'Habilita el tiempo
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub Timer1_Timer()
    If BarraProgreso.Width < 4215 Then
        Porcentaje = (BarraProgreso.Width / 4215) * 100
        BarraProgreso.Width = BarraProgreso.Width + 100 'Suma 100 twip a la anchura del control
        Label3.Caption = Porcentaje & "%"
        Form1.Caption = "PROGRESO: " & Porcentaje & "%"
    Else
        BarraProgreso.Width = 4215
        Label3.Caption = 100 & "%"
        Form1.Caption = "PROGRESO: " & 100 & "%"
        Timer1.Enabled = False 'Deshabilita el tiempo
        MsgBox ("El proceso ha finalizado con exito.")
    End If
End Sub
 
