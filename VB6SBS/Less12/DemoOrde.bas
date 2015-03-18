Attribute VB_Name = "Module1"
Option Base 1      'Comenzar array en 1
Public strArray$() 'declarar un array dinámico para ordenar

Sub ShellSort(sort$(), numOfElements%)
'El subprograma ShellSort ordena los elementos contenidos en el
'array sort$()en orden descendente y lo devuelve al
'procedimiento que ha llamado.

span% = numOfElements% \ 2
Do While span% > 0
    For i% = span% To numOfElements% - 1
        j% = i% - span% + 1
        For j% = (i% - span% + 1) To 1 Step -span%
            If sort$(j%) <= sort$(j% + span%) Then Exit For
            'intercambia los elementos del array que se encuentren desordenadors
            temp$ = sort$(j%)
            sort$(j%) = sort$(j% + span%)
            sort$(j% + span%) = temp$
        Next j%
    Next i%
    span% = span% \ 2
Loop

End Sub
