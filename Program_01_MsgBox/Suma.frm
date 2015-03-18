VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Suma"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "C&errar"
      Height          =   435
      Left            =   2040
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdSumar 
      Caption         =   "&Sumar"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Segundo Valor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Primer Valor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
    End
End Sub

Private Sub cmdSumar_Click()
    Text3.Text = Val(Text1.Text) + Val(Text2.Text)
    ' Val inica que el contenido de la caja de tecto
    ' sera tratado como numeros y no como cadena de texto
End Sub
