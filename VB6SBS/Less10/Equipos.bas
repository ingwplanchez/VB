Attribute VB_Name = "Module1"
    Sub A�adirNombre(Equipo$, CadenaDevuelta$)
        Indicador$ = "Introduzca un nuevo empleado de " & Equipo$
        Nm$ = InputBox(Indicador$, "Cuadro de entrada")
        NuevaL�nea$ = Chr(13) + Chr(10)
        CadenaDevuelta$ = Nm$ & NuevaL�nea$
    End Sub

