Attribute VB_Name = "Module1"
Option Base 1         'Define como 1 la base del array
Public Marcador(2, 9) As Variant
Public Entrada As Integer

Sub SumarPuntuaciones()
'SumarPuntuaciones es un procedimiento público que calcula el total y
'muestra las carreras en el array Marcador.

    For i% = 1 To 9   'empleo de un bucle para sumar las puntuaciones
        PuntosVisitante% = PuntosVisitante% + Marcador(1, i%)
        PuntosCasa% = PuntosCasa% + Marcador(2, i%)
    Next i%           'muestra las puntuaciones en un cuadro
    Form1.CurrentX = 5000
    Form1.CurrentY = 1050
    Form1.Print PuntosVisitante%
    Form1.CurrentX = 5000
    Form1.CurrentY = 1400
    Form1.Print PuntosCasa%
End Sub


