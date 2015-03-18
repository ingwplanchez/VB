VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CONFIGURACION DE PROPIEDADDES PARA LOS OBJETOS
' Control       Propiedad       Valor
' Text1         Text            (Vacio)
' Command1      Caption         Aceptar
' Objeto: Text1
' Propiedad: Text

Private Sub Command1_Click()
   Text1.Text = "Hola Mundo"
End Sub

Private Sub Command2_Click()
    End
End Sub
