VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Reproductor archivos AVI"
   ClientHeight    =   3105
   ClientLeft      =   1770
   ClientTop       =   2040
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5910
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdReproducir 
      Caption         =   "Reproducir .avi"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   " Abrir.avi"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "¡Abra su vídeo favorito en formato .avi y reprodúzcalo!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbrir_Click()
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Filter = "Video (*.AVI)|*.AVI"
    CommonDialog1.ShowOpen
    MMControl1.FileName = CommonDialog1.FileName
    MMControl1.Command = "Open"
Errhandler:
    'Si pulsa Cancelar, salir del procedimiento
End Sub

Private Sub cmdReproducir_Click()
    MMControl1.Command = "Play"
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub Form_Load()
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "AVIVideo"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

