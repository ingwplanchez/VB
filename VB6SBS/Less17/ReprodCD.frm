VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Reproductor de CD de Audio"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   2595
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdReproducir 
      Caption         =   "Reproducir CD"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   1800
      Picture         =   "ReprodCD.frx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Introduzca un CD de audio en la unidad de CD-ROM y pulse reproducir para comenzar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReproducir_Click()
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    'Especifique el tipo de CD de Audio (desde la unidad de CD-ROM)
    MMControl1.DeviceType = "CDAudio"
    MMControl1.Command = "Open"
    Image1.Visible = True
    End Sub

Private Sub cmdSalir_Click()
    'detener la reproducción del CD de audio si se pulsa
    'el botón Salir
    MMControl1.Command = "Stop"
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cerrar siempre el dispositivo cuando se termine
    MMControl1.Command = "Close"
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    'ocultar el dibujo con las instrucciones antes de
    'comenzar la reproducción del CD
    Image1.Visible = False
End Sub
