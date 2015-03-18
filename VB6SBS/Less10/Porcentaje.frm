VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Siete Afortunado"
   ClientHeight    =   3975
   ClientLeft      =   3030
   ClientTop       =   2130
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6015
   Begin VB.CommandButton Command2 
      Caption         =   "Fin"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jugar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblPorcentaje 
      Alignment       =   2  'Center
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblGanadas 
      Alignment       =   2  'Center
      Caption         =   "Ganadas: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   1680
      Picture         =   "Porcentaje.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Siete Afortunado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Image1.Visible = False         ' ocultar monedas
    Label1.Caption = Int(Rnd * 10) ' generar números
    Label2.Caption = Int(Rnd * 10)
    Label3.Caption = Int(Rnd * 10)
        Jugadas = Jugadas + 1
    'si algún número es 7 mostrar una pila de monedas y pitar
    If (Label1.Caption = 7) Or (Label2.Caption = 7) _
      Or (Label3.Caption = 7) Then
        Image1.Visible = True
        Beep
        Ganadas = Ganadas + 1
        lblGanadas.Caption = "Ganadas: " & Ganadas
    End If
    lblPorcentaje.Caption = Porcentaje(Ganadas, Jugadas)
End Sub

Private Sub Command2_Click()
    End
End Sub





