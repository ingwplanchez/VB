VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nota r�pida"
   ClientHeight    =   4230
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6720
   Begin VB.TextBox txtNote 
      Height          =   3015
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2,53906e-29
   End
   Begin VB.Label lblFile 
      Caption         =   "Escriba su nota y almac�nela en el disco"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuItemAbrir 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu mnuItemCerrar 
         Caption         =   "&Cerrar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuItemGuardar 
         Caption         =   "&Guardar como..."
      End
      Begin VB.Menu mnuItemOrdenarTexto 
         Caption         =   "&Ordenar Texto"
      End
      Begin VB.Menu mnuItemFecha 
         Caption         =   "&Insertar Fecha"
      End
      Begin VB.Menu mnuItemSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuItemCerrar_Click()
    txtNote.Text = ""            'limpia el cuadro de texto
    lblFile.Caption = "Escriba el texto o abra un archivo para ordenar su contenido."
    mnuItemCerrar.Enabled = False 'desactiva el mandato Cerrar
    mnuItemAbrir.Enabled = True   'activa el mandato Abrir
End Sub

Private Sub mnuItemFecha_Click()
    Wrap$ = Chr$(13) & Chr$(10) 'a�ade la fecha al texto
    txtNote.Text = Date$ & Wrap$ & txtNote.Text
End Sub

Private Sub mnuItemSalir_Click()
    End                         'salir del programa
End Sub

Private Sub mnuItemAbrir_Click()
    Wrap$ = Chr$(13) + Chr$(10)  'crea el car�cter wrap (salto)
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen       'muestra el cuadro de di�logo Abrir
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11  'muestra un reloj de arena
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo TooBig:    'define un manejador de error
        Do Until EOF(1)          'lee l�neas del archivo
            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        lblFile.Caption = CommonDialog1.FileName
        txtNote.Text = AllText$  'muestra el archivo
        txtNote.Enabled = True
        mnuItemCerrar.Enabled = True
        mnuItemAbrir.Enabled = False 'permite el desplazamiento
CleanUp:
        Form1.MousePointer = 0   'vuelve a configurar el rat�n
        Close #1                 'cierra el archivo
    End If
    Exit Sub
TooBig:             'el manejador de error muestra un mensaje
    MsgBox ("El archivo especificado es demasiado largo.")
    Resume CleanUp: 'a continuaci�n, salta a la rutina CleanUp

End Sub

Private Sub mnuItemGuardar_Click()
'nota: todo el archivo se almacenar� como una �nica cadena
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowSave      'muestra el cuadro de di�logo Guardar
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1
        Print #1, txtNote.Text  'guarda el texto en un archivo
        Close #1                'cerrar archivo
    End If
End Sub

Private Sub mnuItemOrdenarTexto_Click()
'calcula el n�mero de l�neas existente en el objeto cuadro de texto(txtNote)
lineCount% = 0  'esta variable almacena el n�mero total de l�neas
charsInFile% = Len(txtNote.Text)  'obtiene el n�mero total de caracteres contenidos en el cuadro
For i% = 1 To charsInFile%  'desplaza un car�cter en el cuadro
    letter$ = Mid(txtNote.Text, i%, 1) 'introduce el siguiente car�cter en letter$
    If letter$ = Chr$(13) Then 'si se encuentra un retorno de carro (�final de la l�nea!)
        lineCount% = lineCount% + 1  'va a la l�nea siguiente (a�ade al contador)
        i% = i% + 1   'pasa el car�cter de alimentaci�n de l�nea (que siempre sigue a un r. de c.)
    End If
Next i%

'crea un array para almacenar el texto contenido en el cuadro
ReDim strArray$(lineCount%) 'crea un array del tama�o adecuado
curline% = 1
ln$ = ""  'utiliza ln$ para construir l�neas de un �nico car�cter
For i% = 1 To charsInFile%     'hace un bucle por todo el texto
    letter$ = Mid(txtNote.Text, i%, 1)
    If letter$ = Chr$(13) Then 'si encuentra un retorno de carro
        curline% = curline% + 1    'incrementa el contador de l�nea
        i% = i% + 1            'salta el car�cter de alimentaci�n de l�nea
        ln$ = ""               'borra la l�nea y salta a la siguiente
    Else
        ln$ = ln$ & letter$    'a�ade una letra a la l�nea
        strArray$(curline%) = ln$  'y la introduce en el array
   End If
Next i%

'ordenar array
ShellSort strArray$(), lineCount%

'finalmente, muestra el array ordenado en el cuadro
txtNote.Text = ""
Wrap$ = Chr$(13) & Chr$(10) 'a�ade la fecha a la cadena
curline% = 1
For i% = 1 To lineCount%
    txtNote.Text = txtNote.Text & strArray$(curline%) & Wrap$
    curline% = curline% + 1
Next i%

End Sub
