VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Errores"
   ClientHeight    =   2985
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6885
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prueba"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Salida"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "¿Qué edad tienes?"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Práctica de depuración: ¿Puede encontrar el error?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Edad = Text1.Text
    
    If Edad > 13 And Edad < 20 Then
        Text2.Text = "Eres un adolescente."
    Else
        Text2.Text = "No eres un adolescente."
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

