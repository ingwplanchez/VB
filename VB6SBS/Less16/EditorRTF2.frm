VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor RTF 2"
   ClientHeight    =   4200
   ClientLeft      =   2760
   ClientTop       =   2565
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7800
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3825
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"EditorRTF2.frx":0000
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Rich Text Format (*.RTF)|*.RTF|All Files (*.*)|*.*"
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
         Caption         =   "Convertir &May�sculas"
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
'Declarar CambiosNoGuardados como una variable p�blica
'Buleana (Verdadero/Falso)con el fin de supervisar el
'estado actual del texto (almacenado o no almacenado)
'Si el texto se encuentra actualizado, el procedimiento
'de suceso denominado RichTextBox1_Change define esta
'variable como True.
Dim CambiosNoGuardados As Boolean

Private Sub Form_Load()
    'Definir valores iniciales para el control Slider
    Slider1.Left = RichTextBox1.Left  'alinear al cuadro de texto
    Slider1.Width = RichTextBox1.Width
    'nota: todas las medidas del deslizador est�n en twips
    Slider1.Max = RichTextBox1.Width
    Slider1.TickFrequency = Slider1.Max * 0.1
    Slider1.LargeChange = Slider1.Max * 0.1
    Slider1.SmallChange = Slider1.Max * 0.01
End Sub

Private Sub mnuNegritaItem_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub

Private Sub mnuCerrarItem_Click()
    Dim Indicador As String
    Dim Respuesta As Integer
    'Saltar hasta el manejador de error si se pulsa el bot�n
    'Cancelar
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    If CambiosNoGuardados = True Then
        Indicador = "�Quiere almacenar sus cambios?"
        Respuesta = MsgBox(Indicador, vbYesNo)
        If Respuesta = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
            'mostrar nombre de archivo (sin la ruta) en la barra de estado
            StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
        End If
    End If
    RichTextBox1.Text = ""  'borrar cuadro de texto
    StatusBar1.Panels(1).Text = ""
    CambiosNoGuardados = False
ManejadorErr:
    'Cancelar bot�n pulsado.
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
        Indicador = "�Desea almacenar sus cambios?"
        Respuesta = MsgBox(Indicador, vbYesNo)
        If Respuesta = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
            'mostrar nombre de archivo (sin la ruta) en la barra de estado
            StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
            CambiosNoGuardados = False
        End If
    End If
    End 'Una vez almacenado el archivo, salir del programa
ManejadorErr:
    'Cancelar bot�n pulsado(volver al programa)
End Sub

Private Sub mnuBuscarItem_Click()
    Dim CadBuscada As String  'texto utilizado en la b�squeda
    Dim PosLocaliz As Integer  'ubicaci�n del texto buscado
    CadBuscada = InputBox("Introduzca la palabra a localizar", "Buscar")
    If CadBuscada <> "" Then  'Si la cadena de b�squeda no est� vac�a
        'localizar la primera aparici�n de la palabra
        PosLocaliz = RichTextBox1.Find(CadBuscada, , , _
            rtfWholeWord)
        'si se localiza la palabra(si no es -1)
        If PosLocaliz <> -1 Then
        'utilice el m�todo Span para seleccionar la palabra
        ' (direcci�n hacia delante)
            RichTextBox1.Span " ", True, True
        Else
            MsgBox "Cadena buscada no encontrada", , "Buscar"
        End If
    End If
End Sub

Private Sub mnuFuenteItem_Click()
    'Forzar un error si el usuario pulsa el bot�n Cancelar
    CommonDialog1.CancelError = True
    On Error GoTo ManejadorErr:
    'Definir banderas para efectos especiales y
    'para todas las fuentes disponibles
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    'Mostrar el cuadro de di�logo de fuentes
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
    'Llamar al procedimiento de suceso RichTextBox1_SelChange
    'para actualizar la barra de estado con el nombre de la fuente
    'utilizada en el texto seleccionado.
    RichTextBox1_SelChange
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
    'display filename (without path) on status bar
    StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
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
    'mostrar nombre de archivo(sin la ruta) en la barra de estado
    StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
    CambiosNoGuardados = False
ManejadorErr:
    'bot�n Cancelar pulsado
End Sub

Private Sub RichTextBox1_Change()
    'Definir la variable p�blica CambiosNoGuardados como True
    'cada vez que se modifique el texto contenido en el
    'cuadro de texto Rich.
    CambiosNoGuardados = True
End Sub

Private Sub RichTextBox1_SelChange()
    'si s�lo se ha seleccionado una fuente, mostrar su nombre
    'en la barra de estado (si hay varias fuentes seleccionadas
    'se devolver� el valor Null)
    If IsNull(RichTextBox1.SelFontName) Then
        StatusBar1.Panels(2).Text = ""
    Else
        StatusBar1.Panels(2).Text = RichTextBox1.SelFontName
    End If
    'si s�lo se ha seleccionado un tipo de sangrado, mostrar su nombre
    'en la barra de estado (si hay varios estilos seleccionados
    'se devolver� el valor Null)
    If Not IsNull(RichTextBox1.SelIndent) Then
        Slider1.Value = RichTextBox1.SelIndent
    End If
End Sub

Private Sub Slider1_Scroll()
    RichTextBox1.SelIndent = Slider1.Value
End Sub

