VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nota Rápida"
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
      Caption         =   "Escriba su nota y almacénela en el disco."
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
    CommonDialog1.ShowSave      'muestra el cuadro de diálogo Guardar
    If CommonDialog1.FileName <> "" Then
        'obtiene el código de cifrado y lo utiliza para cifrar el archivo
        code = InputBox("Introduzca el código de cifrado", , 1)
        If code = "" Then Exit Sub  'si se seleciona Cancelar, salir de la sub
        Form1.MousePointer = 11     'mostrar el reloj de arena
        charsInFile% = Len(txtNote.Text) 'calcula la longitud de la cadena
        Open CommonDialog1.FileName For Output As #1 'abrir archivo
        For i% = 1 To charsInFile%  'para cada carácter perteneciente al archivo
            letter$ = Mid(txtNote.Text, i%, 1) 'lee el siguiente carácter
            'convierte a un número ASCII y utiliza Xor para cifrarlo
            Print #1, Asc(letter$) Xor code; 'y lo almacena en un archivo
        Next i%
        Close #1                'cierra el archivo
        CommonDialog1.FileName = ""  'borra el nombre del archivo
        Form1.MousePointer = 0  'vuelve a configurar el ratón
    End If
End Sub

Private Sub mnuAbrirItem_Click()
    Wrap$ = Chr$(13) + Chr$(10) 'crea un carácter de salto
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen      'muestra el cuadro de diálogo Abrir
    If CommonDialog1.FileName <> "" Then
        'obtiene el código de cifrado para convertir en texto el archivo cifrado
        code = InputBox("Introduzca el código de cifrado", , 1)
        If code = "" Then Exit Sub 'si se seleciona Cancelar, salir de la sub
        Form1.MousePointer = 11 'mostrar reloj de arena
        Open CommonDialog1.FileName For Input As #1 'abrir archivo
        On Error GoTo Problem:  'define el manejador de error
        decrypt$ = ""   'inicia la cadena para su descifrado
        Do Until EOF(1)         'hasta que se alcanza el final del archivo
            Input #1, Number&   'lee los números cifrados
            e$ = Chr$(Number& Xor code) 'los convierte con Xor
            decrypt$ = decrypt$ & e$    'construye una cadena
        Loop
        lblFile.Caption = CommonDialog1.FileName 'define un titular
        txtNote.Text = decrypt$ 'muestra la cadena una vez convertida
        txtNote.Enabled = True  'activa las barras de desplazamiento
CleanUp:                        'cuando termine...
        Form1.MousePointer = 0  'vuelve a modificar el icono del ratón
        Close #1                'cierra el archivo
        CommonDialog1.FileName = ""  'borra el nombre del archivo
    End If
    Exit Sub
Problem:  'si existe un problema, muestra el mensaje apropiado
    If Err.Number = 5 Then  'problema Chr$, implica una clave errónea
        MsgBox ("Clave de cifrado incorrecta")
    Else  'para otros problemas (como archivo demasiado grande) mostrar error
        MsgBox "Error en la apertura del archivo", , Err.Description
    End If
    Resume CleanUp:   'finalmente, termina con la rutina CleanUp
End Sub
