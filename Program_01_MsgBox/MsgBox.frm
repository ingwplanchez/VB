VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MsgBox"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHabilitar 
      Caption         =   "Habilitar"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeshabilitar 
      Caption         =   "Deshabilitar"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaludo 
      Caption         =   "Saludo"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CONFIGURACION DE PROPIEDADDES PARA LOS OBJETOS
' Control       Propiedad       Valor
' Text1         Text            (Vacio)
' Command1      Caption         Saludo
' Command2      Caption         Limpiar
' Command3      Caption         Deshabilitar
' Command4      Caption         Habilitar
' Command5      Caption         Salir
' RENOMBRAMIENTO DE BOTONES
' Control       Propiedad       Valor
' Command1      Name            cmdSaludo
' Command2      Name            cmdLimpiar
' Command3      Name            cmdDeshabilitar
' Command4      Name            cmdHabilitar
' Command5      Name            cmdSalir

'Boton Desahabilitar
Private Sub cmdDeshabilitar_Click()
    Text1.Text = " "                            ' Limpia El cuadro de texto
    Text1.Text = "Desabilitado."   ' Se muestra en el cuadro de texto
    cmdSaludo.Enabled = False                   'Deshabilita el boton <<Saludo>>
    MsgBox ("Botòn saludo desahabilitado.")     ' Msj en cuadro de texto
End Sub

'Boton Habilitar
Private Sub cmdHabilitar_Click()
    Text1.Text = " "                            ' Limpia El cuadro de texto
    Text1.Text = "Habilitado. "                  ' Se muestra en el cuadro de texto
    cmdSaludo.Enabled = True                    ' Habilita el boton <<Saludo>>
    MsgBox ("Botòn saludo habilitado.")         ' Msj en cuadro de texto
End Sub

' Boton limpiar
Private Sub cmdLimpiar_Click()
    
    Text1.Text = " "                            ' Limpia El cuadro de texto
End Sub

' Boton Salir
Private Sub cmdSalir_Click()
    End                                         ' Salir del programa
End Sub

' Boton Saludo
Private Sub cmdSaludo_Click()
    Text1.Text = "Presione Limpiar"           ' Se muestra en el cuadro de texto
    MsgBox ("Bienvenido a visual basic 6.0.")   ' Msj en cuadro de texto
End Sub


