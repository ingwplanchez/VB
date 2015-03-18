VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bienvenida"
   ClientHeight    =   3195
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6555
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Continuar"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   """Productos de Calidad para la oficina y el hogar"""
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   6  'Cross
      Height          =   975
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008080&
      FillColor       =   &H00008080&
      FillStyle       =   6  'Cross
      Height          =   1935
      Left            =   240
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   120
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Ventanas Noroeste"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = "c:\vb6sbs\less17\applause.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub
