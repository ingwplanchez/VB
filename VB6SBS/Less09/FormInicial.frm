VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bienvenida"
   ClientHeight    =   3075
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6555
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   """Productos de Calidad para la oficina y el hogar"""
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   6  'Cross
      Height          =   975
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008080&
      FillColor       =   &H00008080&
      FillStyle       =   6  'Cross
      Height          =   1935
      Left            =   240
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   120
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Ventanas Noroeste"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    End
End Sub


