Attribute VB_Name = "Module1"
Option Base 1      'Comenzar array en 1
Public strArray$() 'declarar array dinámico para ordenación

Sub ShellSort(ordenar$(), numDeElementos%)
'El subprograma ShellSort ordena los elementos contenidos
'en el array ordenar$() en orden descendente y devuelve el
'resultado al procedimiento que ha llamado.

longitud% = numDeElementos% \ 2
Do While longitud% > 0
    For i% = longitud% To numDeElementos% - 1
        j% = i% - longitud% + 1
        For j% = (i% - longitud% + 1) To 1 Step -longitud%
            If ordenar$(j%) <= ordenar$(j% + longitud%) Then Exit For
            'intercambia los elementos del array que no
            'estén ordenados
            temp$ = ordenar$(j%)
            ordenar$(j%) = ordenar$(j% + longitud%)
            ordenar$(j% + longitud%) = temp$
        Next j%
    Next i%
    longitud% = longitud% \ 2
Loop

End Sub
