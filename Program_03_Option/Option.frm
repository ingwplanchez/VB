VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Option"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtContenido 
      Height          =   4455
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.Frame frmColorLetra 
      Caption         =   "Colores de Letra"
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
      Begin VB.OptionButton OptAzulLetra 
         Caption         =   "Azul"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton OptNegroLetra 
         Caption         =   "Negro"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton optBlancoLetra 
         Caption         =   "Blanco"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton OptMagentaLetra 
         Caption         =   "Magenta"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptCyanLetra 
         Caption         =   "Cyan"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmColorFondo 
      Caption         =   "Colores de fondo"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton OptBlancoFondo 
         Caption         =   "Blanco"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton OptAmarilloFondo 
         Caption         =   "Amarillo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton OptAzulFondo 
         Caption         =   "Azul"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton OptVerdeFondo 
         Caption         =   "Verde"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptRojoFondo 
         Caption         =   "Rojo"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OptRojoFondo_Click()
    txtContenido.BackColor = vbRed      ' Fondo Rojo
End Sub
Private Sub OptVerdeFondo_Click()
    txtContenido.BackColor = vbGreen    ' Fondo Verde
End Sub
Private Sub OptAzulFondo_Click()
    txtContenido.BackColor = vbBlue     ' Fondo Azul
End Sub
Private Sub OptAmarilloFondo_Click()
    txtContenido.BackColor = vbYellow   ' Fondo Amarillo
End Sub
Private Sub OptBlancoFondo_Click()
    txtContenido.BackColor = vbWhite    ' Fondo Blanco
End Sub
Private Sub OptCyanLetra_Click()
   txtContenido.ForeColor = vbCyan      ' Color de Letra Cyan
End Sub
Private Sub OptMagentaLetra_Click()
    txtContenido.ForeColor = vbMagenta  ' Color de Letra Magenta
End Sub
Private Sub optBlancoLetra_Click()
    txtContenido.ForeColor = vbWhite    ' Color de Letra Blanco
End Sub
Private Sub OptNegroLetra_Click()
    txtContenido.ForeColor = vbBlack    ' Color de Letra Negro
End Sub
Private Sub OptAzulLetra_Click()
    txtContenido.ForeColor = vbBlue      ' Color de Letra Azul
End Sub
Private Sub cmdSalir_Click()
    End
End Sub
