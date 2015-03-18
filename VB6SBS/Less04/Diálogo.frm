VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2055
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
   Begin VB.Menu mnureloj 
      Caption         =   "&Reloj"
      Begin VB.Menu mnuHoraItem 
         Caption         =   "&Hora"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFechaItem 
         Caption         =   "&Fecha"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuColortextoItem 
         Caption         =   "&Color del texto"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbrirItem_Click()
    CommonDialog1.Filter = "Metafiles (*.WMF)|*.WMF"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.filename)
    mnuCerrarItem.Enabled = True
End Sub

Private Sub mnuCerrarItem_Click()
    Image1.Picture = LoadPicture("")
    mnuCerrarItem.Enabled = False
End Sub

Private Sub mnuColortextoItem_Click()
    CommonDialog1.Flags = &H1&
    CommonDialog1.ShowColor
    Label1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnuFechaItem_Click()
    Label1.Caption = Date
End Sub

Private Sub mnuHoraItem_Click()
    Label1.Caption = Time
End Sub

Private Sub mnuSalirItem_Click()
    End
End Sub
