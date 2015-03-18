VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Nota rápida"
   ClientHeight    =   4380
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6720
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtNote 
      Height          =   3015
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2,53906e-29
   End
   Begin VB.Label lblFile 
      Caption         =   "Introduzca texto o abra un archivo para ordenarlo"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3855
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
    txtNote.Text = ""            'borrar el cuadro de texto
    lblFile.Caption = "Escriba texto o abra un archivo para ordenarlo."
    mnuItemCerrar.Enabled = False 'desactiva el mandato Cerrar
    mnuItemAbrir.Enabled = True   'activa el mandato Abrir
End Sub

Private Sub mnuItemFecha_Click()
    Salto$ = Chr$(13) & Chr$(10) 'añade la fecha a la cadena
    txtNote.Text = Date$ & Salto$ & txtNote.Text
End Sub

Private Sub mnuItemSalir_Click()
    End                         'salir del programa
End Sub

Private Sub mnuItemAbrir_Click()
    Salto$ = Chr$(13) + Chr$(10)  'crear carácter Salto
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen       'muestra el cuadro de diálogo Abrir
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11  'muestra un reloj de arena
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo DemGrande:    'define el manejador de error
        Do Until EOF(1)          'luego, lee líneas del archivo
            Line Input #1, LineaDeTexto$
            TodoTexto$ = TodoTexto$ & LineaDeTexto$ & Salto$
        Loop
        lblFile.Caption = CommonDialog1.FileName
        txtNote.Text = TodoTexto$  'muestra archivo
        txtNote.Enabled = True
        mnuItemCerrar.Enabled = True
        mnuItemAbrir.Enabled = False 'permite el desplazamiento
Limpiar:
        Form1.MousePointer = 0   'reconfigura el ratón
        Close #1                 'cerrar archivo
    End If
    Exit Sub
DemGrande:             'el manejador de error muestra un mensaje
    MsgBox ("El archivo especificado es demasiado grande.")
    Resume Limpiar: 'a continuación, salta a la rutina Limpiar
End Sub

Private Sub mnuItemGuardar_Click()
'nota: se almacena el archivo completo es una cadena
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowSave      'muestra el cuadro de diálogo Guardar
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1
        Print #1, txtNote.Text  'almacena la cadena en un archivo
        Close #1                'cerrar archivo
    End If
End Sub

Private Sub mnuItemOrdenarTexto_Click()
'calcula el número de caracteres contenido en el archivo
ContLinea% = 0
CarEnArchivo% = Len(txtNote.Text)
If CarEnArchivo% < 2 Then Exit Sub 'renuncia si no hay nada que ordenar

'en caso contrario, comienza la ordenación mostrando una barra de progreso
ProgressBar1.Visible = True
ProgressBar1.Min = 1
ProgressBar1.Max = CarEnArchivo%   'define max para la barra de progreso
ProgressBar1.Value = 1            'define el valor inicial
'calcula el número de líneas contenidas en el cuadro de texto
For i% = 1 To CarEnArchivo%
    letra$ = Mid(txtNote.Text, i%, 1)
    ProgressBar1.Value = i%    'mueve la barra de progreso
    If letra$ = Chr$(13) Then 'si se encuentra un retorno de carro
        ContLinea% = ContLinea% + 1  'incrementa el contador de línea
        i% = i% + 1      'salta el carácter de alimentación de línea
    End If
Next i%

'reconfigura la barra de progreso para la siguiente fase de la
'ordenación
ProgressBar1.Value = 1
ProgressBar1.Max = ContLinea%

'crear un array para almacenar el texto en el cuadro de texto
ReDim strArray$(ContLinea%) 'crear array del tamaño adecuado
lineaactual% = 1
ln$ = ""  'utilizar ln$ para ir añadiendo un carácter a la línea
For i% = 1 To CarEnArchivo%     'hacer bucle por todo el texto
    letra$ = Mid(txtNote.Text, i%, 1)
    If letra$ = Chr$(13) Then 'si se encuentra un retorno de carro
        ProgressBar1.Value = lineaactual%  'mostrar progreso
        lineaactual% = lineaactual% + 1    'incrementar contador de línea
        i% = i% + 1            'saltar alimentación de línea
        ln$ = ""               'borrar línea y pasar a la siguiente
    Else
        ln$ = ln$ & letra$    'agregar letra a la línea
        strArray$(lineaactual%) = ln$  'e introducirla en el array
   End If
Next i%

'ordenar array
ShellSort strArray$(), ContLinea%

'mostrar el array ordenado en el cuadro de texto
txtNote.Text = ""
Salto$ = Chr$(13) & Chr$(10) 'agregar Fecha a la cadena
lineaactual% = 1
For i% = 1 To ContLinea%
    txtNote.Text = txtNote.Text & strArray$(lineaactual%) & Salto$
    lineaactual% = lineaactual% + 1
Next i%

'ocultar barra de progreso
ProgressBar1.Visible = False
End Sub

