VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nota rápida"
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
      FontSize        =   1,17491e-38
   End
   Begin VB.Label lblFile 
      Caption         =   "Escriba su nota y almacénela en el disco"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuAbrirItem 
         Caption         =   "&Abrir Archivo Cifrado..."
      End
      Begin VB.Menu mnuItemGuardar 
         Caption         =   "&Guardar Archivo Cifrado..."
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
Private Sub mnuItemFecha_Click()
    Wrap$ = Chr$(13) & Chr$(10) 'añadir la fecha a la cadena
    txtNote.Text = Date$ & Wrap$ & txtNote.Text
End Sub

Private Sub mnuItemSalir_Click()
    End                         'salir del programa
End Sub

Private Sub mnuItemGuardar_Click()
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowSave           'muestra el cuadro de diálogo Guardar
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11      'muestra el reloj de arena
        lblFile.Caption = CommonDialog1.FileName
        'almacena el texto con el algoritmo de cifrado (código ASCII + 1)
        encrypt$ = ""  'inicializa la cadena de cifrado
        charsInFile% = Len(txtNote.Text) 'calcula la longitud de la cadena
        For i% = 1 To charsInFile%   'para cada carácter perteneciente al archivo
            letter$ = Mid(txtNote.Text, i%, 1) 'lee el siguiente carácter
            'obtiene el código ASCII del carácter y le añade una unidad
            encrypt$ = encrypt$ & Chr$(Asc(letter$) + 1)
        Next i%
        Open CommonDialog1.FileName For Output As #1 'abre el archivo
        Print #1, encrypt$           'almacena el texto cifrado en el archivo
        txtNote.Text = encrypt$
        Close #1                     'cierra el archivo
        CommonDialog1.FileName = ""  'borra el nombre del archivo
        Form1.MousePointer = 0       'vuelve a configurar el ratón
    End If
End Sub

Private Sub mnuAbrirItem_Click()
    Wrap$ = Chr$(13) + Chr$(10) 'crea un carácter de salto
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen      'muestra el cuadro de diálogo Abrir
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11 'muestra el reloj de arena
        Open CommonDialog1.FileName For Input As #1 'abre el archivo
        On Error GoTo Problem:  'define el manejador de error
        Do Until EOF(1)         'copia cada una de las líneas de texto en
            Line Input #1, LineOfText$  'la cadena AllText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        'ahora, descifra la cadena restando una unidad al código ASCII
        decrypt$ = ""   'inicia la cadena para su descifrado
        charsInFile = Len(AllText$)  'obtiene la longitud de la cadena
        For i% = 1 To charsInFile    'hace un bucle para cada carácter
            letter$ = Mid(AllText$, i%, 1)  'obtiene el carácter mediante Mid
            decrypt$ = decrypt$ & Chr$(Asc(letter) - 1) 'resta 1
        Next i%                       'y construye una nueva cadena
        txtNote.Text = decrypt$ 'muestra la cadena una vez convertida
        txtNote.Enabled = True  'activa las barras de desplazamiento
        lblFile.Caption = CommonDialog1.FileName 'define el titular
CleanUp:                        'cuando termine...
        Form1.MousePointer = 0  'vuelve a modificar el icono del ratón
        Close #1                'cierra el archivo
        CommonDialog1.FileName = ""   'borra el nombre del archivo
    End If
    Exit Sub
Problem:  'si existe un problema, muestra el mensaje apropiado
    MsgBox "Error en la apertura del archivo", , Err.Description
    lblFile.Caption = ""        'elimina el titular
    txtNote.Text = ""           'borra el cuadro de texto
    Resume CleanUp:   'finalmente, termina con la rutina CleanUp
End Sub
