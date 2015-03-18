VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Probador de operadores avanzados"
   ClientHeight    =   2385
   ClientLeft      =   1650
   ClientTop       =   1845
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4980
   Begin VB.Frame Frame1 
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   1935
      Begin VB.OptionButton Option4 
         Caption         =   "Concatenación"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Exponenciación"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Resto entero (Mod)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "División entera (\)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Resultado"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Variable 2"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Variable 1"
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
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim Primero, Segundo   'declaración de variables
    
    Primero = Text1.Text  'capturar números introducidos en los cuadros de rótulo
    Segundo = Text2.Text
    'si se pulsa el primer botón, ejecutar división entera
    If Option1.Value = True Then
        Label1.Caption = Primero \ Segundo
    End If
    'si se pulsa el segundo botón, calcular resto división entera
    If Option2.Value = True Then
        Label1.Caption = Primero Mod Segundo
    End If
    'si se pulsa el tercer botón, calcular exponenciación
    If Option3.Value = True Then
        Label1.Caption = Primero ^ Segundo
    End If
    'si se pulsa el cuarto botón, realizar concatenación de cadenas
    If Option4.Value = True Then
        Label1.Caption = Primero & Segundo
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub


