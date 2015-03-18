VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   1935
   ClientTop       =   1815
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   5415
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2,54016e-29
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Un programa que visualiza mapas de bits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4335
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
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
    mnuCerrarItem.Enabled = True
End Sub

