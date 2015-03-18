VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Visualizador de texto"
   ClientHeight    =   4155
   ClientLeft      =   1125
   ClientTop       =   1770
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5910
   Begin VB.TextBox txtArchivo 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label lblArchivo 
      Caption         =   "Carga un archivo de texto con el mandato Abrir"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4815
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
    txtArchivo.Text = ""            'limpia el cuadro de texto
    lblArchivo.Caption = "Carga un archivo de texto con el mandato Abrir"
    mnuItemCerrar.Enabled = False 'desactiva el mandato Cerrar
    mnuItemAbrir.Enabled = True   'Activa el mandato Abrir
    txtArchivo.Enabled = False      'desactiva el cuadro de texto
End Sub

Private Sub mnuItemSalir_Click()
    End                          'Sale del programa
End Sub

Private Sub mnuItemAbrir_Click()
    Salto$ = Chr$(13) + Chr$(10)  'crea un carácter salto de línea
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen       'muestra el cuadro de diálogo Abrir
    If CommonDialog1.filename <> "" Then
        Form1.MousePointer = 11  'muestra un reloj de arena
        Open CommonDialog1.filename For Input As #1
        On Error GoTo MuyGrande:    'define el manejador de error
        Do Until EOF(1)          'lee líneas del archivo
            Line Input #1, LíneaDeTexto$
            TodoElTexto$ = TodoElTexto$ & LíneaDeTexto$ & Salto$
        Loop
        lblArchivo.Caption = CommonDialog1.filename
        txtArchivo.Text = TodoElTexto$  'muestra archivo
        txtArchivo.Enabled = True
        mnuItemCerrar.Enabled = True
        mnuItemAbrir.Enabled = False 'activa desplazamiento
Reiniciar:
        Form1.MousePointer = 0   'redefine el ratón
        Close #1                 'cierra el archivo
    End If
    Exit Sub
MuyGrande:             'el manejador de error muestra un mensaje
    MsgBox ("El archivo especificado es demasiado largo.")
    Resume Reiniciar: 'salta a la rutina Reiniciar
End Sub

