VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Temperaturas"
   ClientHeight    =   3300
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdMostrarTemps 
      Caption         =   "Mostrar Temperaturas"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdIntroducirTemps 
      Caption         =   "Introducir Temperaturas"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMostrarTemps_Click()
    Print "Temperaturas máximas:"
    Print
    For i% = 1 To Días
        Print "Día "; i%, Temperaturas(i%)
        Total! = Total! + Temperaturas(i%)
    Next i%
    Print
    Print "La media de las temperaturas máximas es: "; Total! / Días
End Sub

Private Sub cmdIntroducirTemps_Click()
    Cls
    Días = InputBox("¿Cuántos días?", "Crear Array")
    If Días > 0 Then ReDim Temperaturas(Días)
    Indicador$ = "Introduzca la temperatura más elevada"
    For i% = 1 To Días
        Título$ = "Día " & i%
        Temperaturas(i%) = InputBox(Indicador$, Título$)
    Next i%
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
