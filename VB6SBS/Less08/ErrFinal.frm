VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Verificador de Unidad"
   ClientHeight    =   3915
   ClientLeft      =   2205
   ClientTop       =   1815
   ClientWidth     =   4905
   Icon            =   "ErrFinal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4905
   Begin VB.CommandButton Command1 
      Caption         =   "Comprobar Unidad"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"ErrFinal.frx":030A
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error GoTo ErrorDisco
    Image1.Picture = LoadPicture("a:\prntout2.wmf")
    Exit Sub  'salir procedimiento
    
ErrorDisco:
    If Err.Number = 71 Then  'si EL DISCO NO ESTÁ PREPARADO
        MsgBox ("Por favor, cierre la puerta de la unidad."), , _
          "Disco no preparado"
        Resume
    Else
        MsgBox ("No puedo encontrar prntout2.wmf en A:\."), , _
          "Archivo no encontrado"
        Resume PararPrueba
    End If
PararPrueba:
End Sub


