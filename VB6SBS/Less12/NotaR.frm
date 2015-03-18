VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
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
      _Version        =   327680
      FontSize        =   2.53906e-29
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba su nota y almacénela en el disco"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuItemGuardar 
         Caption         =   "&Guardar como..."
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
    Salto$ = Chr$(13) & Chr$(10) 'añade la fecha al texto
    txtNote.Text = Date$ & Salto$ & txtNote.Text
End Sub

Private Sub mnuItemSalir_Click()
    End                         'Salir del programa
End Sub

Private Sub mnuItemGuardar_Click()
'nota: todo el archivo se almacenará como una única cadena
    CommonDialog1.Filter = "Archivos de texto (*.TXT)|*.TXT"
    CommonDialog1.ShowSave      'muestra el cuadro de diálogo Guardar
    If CommonDialog1.filename <> "" Then
        Open CommonDialog1.filename For Output As #1
        Print #1, txtNote.Text  'guarda el texto en un archivo
        Close #1                'cerrar archivo
    End If
End Sub

