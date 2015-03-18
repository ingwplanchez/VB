VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Ampliador de mapas de bits"
   ClientHeight    =   3990
   ClientLeft      =   1590
   ClientTop       =   1800
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   7695
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   2,54016e-29
   End
   Begin VB.Image Image6 
      Height          =   330
      Left            =   6600
      Picture         =   "Ampliador.frx":0000
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   330
      Left            =   6120
      Picture         =   "Ampliador.frx":018A
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   840
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Área de suelta de la impresora"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ventana visualización mapas de bits"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   2895
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Un programa que amplia e imprime los mapas de bits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuAbrirItem 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu mnuCerrarItem 
         Caption         =   "&Cerrar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSalirItem 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    Image1.Picture = Source.Picture
    mnuCerrarItem.Enabled = True
End Sub

Private Sub Image2_DragDrop(Source As Control, X As Single, Y As Single)
    Image2.Picture = Source.Picture
End Sub

Private Sub Image3_DragDrop(Source As Control, X As Single, Y As Single)
    Image3.Picture = Source.Picture
End Sub

Private Sub Image4_DragDrop(Source As Control, X As Single, Y As Single)
    Image4.Picture = Source.Picture
End Sub

Private Sub Image5_DragDrop(Source As Control, X As Single, Y As Single)
    Load Form2
    Form2.Image1.Picture = Source.Picture
    On Error GoTo ErrorImpresión
    Form2.PrintForm
    Unload Form2
    Exit Sub
    
ErrorImpresión:
    MsgBox (Err.Description), , "Error de Impresión"
    MsgBox ("Función de impresión desactivada"), , "Ampliador de mapas de bits"
    Image5.Enabled = False
    Image5.Picture = Image6.Picture
    Resume Next
End Sub

Private Sub mnuCerrarItem_Click()
    Image1.Picture = LoadPicture("")
    mnuCerrarItem.Enabled = False
End Sub

Private Sub mnuSalirItem_Click()
    End
End Sub

Private Sub mnuAbrirItem_Click()
    CommonDialog1.Filter = "Bitmaps (*.BMP)|*.BMP"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.filename)
    mnuCerrarItem.Enabled = True
End Sub

