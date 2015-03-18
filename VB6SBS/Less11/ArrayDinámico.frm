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
    Print "Temperaturas m�ximas:"
    Print
    For i% = 1 To D�as
        Print "D�a "; i%, Temperaturas(i%)
        Total! = Total! + Temperaturas(i%)
    Next i%
    Print
    Print "La media de las temperaturas m�ximas es: "; Total! / D�as
End Sub

Private Sub cmdIntroducirTemps_Click()
    Cls
    D�as = InputBox("�Cu�ntos d�as?", "Crear Array")
    If D�as > 0 Then ReDim Temperaturas(D�as)
    Indicador$ = "Introduzca la temperatura m�s elevada"
    For i% = 1 To D�as
        T�tulo$ = "D�a " & i%
        Temperaturas(i%) = InputBox(Indicador$, T�tulo$)
    Next i%
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
