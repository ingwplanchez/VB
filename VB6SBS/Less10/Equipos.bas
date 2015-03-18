Attribute VB_Name = "Module1"
    Sub AñadirNombre(Equipo$, CadenaDevuelta$)
        Indicador$ = "Introduzca un nuevo empleado de " & Equipo$
        Nm$ = InputBox(Indicador$, "Cuadro de entrada")
        NuevaLínea$ = Chr(13) + Chr(10)
        CadenaDevuelta$ = Nm$ & NuevaLínea$
    End Sub

