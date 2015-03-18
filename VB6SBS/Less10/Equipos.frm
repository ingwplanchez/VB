VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Asignar los equipos departamentales"
   ClientHeight    =   3375
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMkt 
      Caption         =   "Añadir Nombre"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdVentas 
      Caption         =   "Añadir Nombre"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtMkt 
      Height          =   1575
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtVentas 
      Height          =   1575
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblMkt 
      Caption         =   "Marketing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblVentas 
      Caption         =   "Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub cmdVentas_Click()
    AñadirNombre "Ventas", PosiciónVentas$
    txtVentas.Text = txtVentas.Text & PosiciónVentas$
End Sub

Private Sub cmdMkt_Click()
    AñadirNombre "Marketing", PosiciónMkt$
    txtMkt.Text = txtMkt.Text & PosiciónMkt$
End Sub

Private Sub lblSalir_Click()
    End
End Sub

