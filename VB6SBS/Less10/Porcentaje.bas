Attribute VB_Name = "Module1"
Public Ganadas
Public Jugadas
Function Porcentaje(Exitos, Intentos) As String
    Tasa = Exitos / Intentos
    Porcentaje = Format(Tasa, "0.0%")
End Function

