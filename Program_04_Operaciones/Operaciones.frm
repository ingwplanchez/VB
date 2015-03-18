VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Operaciones"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResultado 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operaciones"
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   5055
      Begin VB.OptionButton OptDividir 
         Caption         =   "Dividir"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton OptMultiplicar 
         Caption         =   "Multiplicar"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton OptRestar 
         Caption         =   "Restar"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OptSumar 
         Caption         =   "Sumar"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox TxtSegundoValor 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox TxtPrimerValor 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resultado:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Segundo Valor:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Primer Valor: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptSumar_Click()
    txtResultado.Text = Val(TxtPrimerValor.Text) + Val(TxtSegundoValor.Text)
End Sub
Private Sub OptRestar_Click()
    txtResultado.Text = Val(TxtPrimerValor.Text) - Val(TxtSegundoValor.Text)
End Sub
Private Sub OptMultiplicar_Click()
    txtResultado.Text = Val(TxtPrimerValor.Text) * Val(TxtSegundoValor.Text)
End Sub
Private Sub OptDividir_Click()
    If Val(TxtSegundoValor.Text) = 0 Then
        MsgBox ("No se puede dividir por cero.")
    Else
        txtResultado.Text = Val(TxtPrimerValor.Text) / Val(TxtSegundoValor.Text)
    End If
End Sub
