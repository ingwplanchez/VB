VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Importar documentos HTML"
   ClientHeight    =   4695
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6990
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   80
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar HTML"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtURLbox 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "http://www.microsoft.com"
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txtNote 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1440
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2,53906e-29
   End
   Begin VB.Label Label1 
      Caption         =   "Introducir el  URL de un documento HTML y pulse el botón Importar HTML."
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuItemHTML 
         Caption         =   "Guardar como &HTML..."
      End
      Begin VB.Menu mnuItemGuardar 
         Caption         =   "&Guardar como texto..."
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
'declarar una variable para el URL introducido
Dim strUrl As String

Private Sub Command1_Click()
    On Error GoTo manejadorerror
    strUrl = txtURLbox.Text
    'verificar si hay, al menos, 11 caracteres ("http://www.")
    If Len(strUrl) > 11 Then
        'copiar el documento html en el cuadro de texto
        txtNote.Text = Inet1.OpenURL(strUrl)
    Else
        MsgBox "Introduzca un nombre de documento válido en el cuadro URL"
    End If
    Exit Sub
manejadorerror:
    MsgBox "Error en la conexión con el URL", , Err.Description
End Sub

Private Sub mnuItemSalir_Click()
    Unload Form1                       'salir del programa
End Sub

Private Sub mnuItemHTML_Click()
'nota: todo el archivo se almacena en una única cadena
CommonDialog1.DefaultExt = "HTM"
CommonDialog1.Filter = "archivos HTML (*.HTML;*.HTM)|*.HTML;HTM"
CommonDialog1.ShowSave      'mostrar cuadro de diálogo Guardar
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtNote.Text  'almacenar cadena en el archivo
    Close #1                'cerrar archivo
End If
End Sub

Private Sub mnuItemGuardar_Click()
'nota: todo el archivo se almacena en una única cadena
CommonDialog1.DefaultExt = "TXT"
CommonDialog1.Filter = "archivos de texto (*.TXT)|*.TXT"
CommonDialog1.ShowSave      'mostrar cuadro de diálogo Guardar
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtNote.Text  'almacenar cadena en el archivo
    Close #1                'cerrar archivo
End If
End Sub

