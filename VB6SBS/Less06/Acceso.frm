VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6090
   Begin VB.CommandButton Command1 
      Caption         =   "Acceso"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
     NombreUsuario = InputBox("Introduzca su nombre.")
    If NombreUsuario = "Laura" Then
        MsgBox ("¡Bienvenida, Laura!  ¿Preparada para comenzar?")
        Form1.Picture = _
          LoadPicture("c:\vb6sbs\less06\pcomputr.wmf")
    ElseIf NombreUsuario = "Marcos" Then
        MsgBox ("¡Bienvenido, Marcos!  ¿Listo para ver su Rolodex?")
        Form1.Picture = _
          LoadPicture("c:\vb6sbs\less06\rolodex.wmf")
    Else
        MsgBox ("Lo siento, no le conozco.")
        End   'salir del programa
    End If
End Sub
