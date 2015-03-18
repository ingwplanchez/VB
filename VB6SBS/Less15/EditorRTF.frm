VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Editor RTF"
   ClientHeight    =   4275
   ClientLeft      =   2775
   ClientTop       =   2580
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6075
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Rich Text Format (*.RTF)|*.RTF|All Files (*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"EditorRTF.frx":0000
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuAbrirItem 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu mnuCerrarItem 
         Caption         =   "&Cerrar"
      End
      Begin VB.Menu mnuGuardarComoItem 
         Caption         =   "&Guardar como..."
      End
      Begin VB.Menu mnuImprimirItem 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuSalirItem 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuCortarItem 
         Caption         =   "Cor&tar"
      End
      Begin VB.Menu mnuCopiarItem 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnuPegarItem 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu mnuBuscarItem 
         Caption         =   "&Buscar..."
      End
   End
   Begin VB.Menu mnuFormato 
      Caption         =   "&Formato"
      Begin VB.Menu mnuConvMayItem 
         Caption         =   "Convertir &Mayúsculas"
      End
      Begin VB.Menu mnuFuenteItem 
         Caption         =   "&Fuente..."
      End
      Begin VB.Menu mnuNegritaItem 
         Caption         =   "&Negrita"
      End
      Begin VB.Menu mnuCursivaItem 
         Caption         =   "&Cursiva"
      End
      Begin VB.Menu mnuSubrayadoItem 
         Caption         =   "&Subrayado"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarar CambiosNoGuardados como una variable pública
'Buleana (Verdadero/Falso)con el fin de supervisar el
'estado actual del texto (almacenado o no almacenado)
'Si el texto se encuentra actualizado, el procedimiento
'de suceso denominado RichTextBox1_Change define esta
'variable como True.
Dim CambiosNoGuardados As Boolean

Private Sub mnuNegritaItem_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub

Private Sub mnuCerrarItem_Click()
    Dim Indicador As String
    Dim Respuesta As Integer
    'Saltar hasta el manejador de error si se pulsa el botón
    'Cancelar
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    If CambiosNoGuardados = True Then
        Indicador = "¿Quiere almacenar sus cambios?"
        Respuesta = MsgBox(Indicador, vbYesNo)
        If Respuesta = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
        End If
    End If
    RichTextBox1.Text = ""  'borrar cuadro de texto
    CambiosNoGuardados = False
ManejadorErr:
    'Cancelar botón pulsado.
    Exit Sub
End Sub

Private Sub mnuCopiarItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
End Sub

Private Sub mnuCortarItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
    RichTextBox1.SelRTF = ""
End Sub

Private Sub mnuSalirItem_Click()
    Dim Indicador As String
    Dim Respuesta As Integer
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    If CambiosNoGuardados = True Then
        Indicador = "¿Desea almacenar sus cambios?"
        Respuesta = MsgBox(Indicador, vbYesNo)
        If Respuesta = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
            CambiosNoGuardados = False
        End If
    End If
    End 'Una vez almacenado el archivo, salir del programa
ManejadorErr:
    'Cancelar botón pulsado(volver al programa)
End Sub

Private Sub mnuBuscarItem_Click()
    Dim CadBuscada As String  'texto utilizado en la búsqueda
    Dim PosLocaliz As Integer  'ubicación del texto buscado
    CadBuscada = InputBox("Introduzca la palabra a localizar", "Buscar")
    If CadBuscada <> "" Then  'Si la cadena de búsqueda no está vacía
        'localizar la primera aparición de la palabra
        PosLocaliz = RichTextBox1.Find(CadBuscada, , , _
            rtfWholeWord)
        'si se localiza la palabra(si no es -1)
        If PosLocaliz <> -1 Then
        'utilice el método Span para seleccionar la palabra
        ' (dirección hacia delante)
            RichTextBox1.Span " ", True, True
        Else
            MsgBox "Cadena buscada no encontrada", , "Buscar"
        End If
    End If
End Sub

Private Sub mnuFuenteItem_Click()
    'Forzar un error si el usuario pulsa el botón Cancelar
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    'Definir banderas para efectos especiales y
    'para todas las fuentes disponibles
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    'Mostrar el cuadro de diálogo de fuentes
    CommonDialog1.ShowFont
    'Definir las propiedades de formato teniendo en cuenta
    'las opciones elegidas por el usuario:
    RichTextBox1.SelFontName = CommonDialog1.FontName
    RichTextBox1.SelFontSize = CommonDialog1.FontSize
    RichTextBox1.SelColor = CommonDialog1.Color
    RichTextBox1.SelBold = CommonDialog1.FontBold
    RichTextBox1.SelItalic = CommonDialog1.FontItalic
    RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
    RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
ManejadorErr:
    'salir del procedimiento si el usuario pulsa Cancelar
End Sub

Private Sub mnuCursivaItem_Click()
    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub

Private Sub mnuConvMayItem_Click()
    RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub mnuImprimirItem_Click()
    'Imprimir el documento actual en la impresora definida
    'por defecto
    RichTextBox1.SelPrint (Printer.hDC)
End Sub

Private Sub mnuSubrayadoItem_Click()
    RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub

Private Sub mnuAbrirItem_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName, rtfRTF
ManejadorErr:
    'si se pulsa Cancelar, salir del procedimiento
End Sub

Private Sub mnuPegarItem_Click()
    RichTextBox1.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuGuardarComoItem_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    CommonDialog1.ShowSave
    'Guardar el archivo especificado en formato RTF
    RichTextBox1.SaveFile CommonDialog1.FileName, rtfRTF
    CambiosNoGuardados = False
ManejadorErr:
    'Cancelar botón pulsado
End Sub

Private Sub RichTextBox1_Change()
    'Definir la variable pública CambiosNoGuardados como True
    'cada vez que se modifique el texto contenido en el
    'cuadro de texto Rich.
    CambiosNoGuardados = True
End Sub
