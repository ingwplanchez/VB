VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Personal Alarm"
   ClientHeight    =   3165
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5190
   Icon            =   "Problema2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   5190
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   360
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Configurar alarma y minimizar"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3840
      Picture         =   "Problema2.frx":030A
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Hora de aviso (00:00 - 23:59)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Avisador Personal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   480
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

Private Sub Command1_Click()
    Form1.WindowState = 1  'Minimizar formulario
    Timer1.Enabled = True  'Poner en marcha temporizador
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Timer1_Timer()
    HoraActual = Format(Time, "hh:mm")
    If HoraActual = Text1.Text Then
        Beep
        MsgBox (Text2.Text), , "Alarma Personal"
        Timer1.Enabled = False
        Form1.WindowState = 0 'Restaura formulario
    End If
End Sub

