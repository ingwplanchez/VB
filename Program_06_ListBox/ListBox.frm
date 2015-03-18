VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Paises"
   ClientHeight    =   4350
   ClientLeft      =   2385
   ClientTop       =   1695
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8175
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtCantidadPaises 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox lstPaises 
      Height          =   2595
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtPais 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Paises agregados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lista de paises"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Introduzca el pais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    'Verifica que la caja no se deje vacía
    If Len(TxtPais.Text) = 0 Then
    MsgBox ("No puede dejar la caja vacía.")
    Else
        lstPaises.AddItem TxtPais.Text ' Agrega el país en el control ListBox
        TxtPais.Text = "" ' Limpia la caja de texto
        TxtPais.SetFocus ' Hace que el cursor se mantenga sobre la caja
        txtCantidadPaises.Text = lstPaises.ListCount 'Pone el número de países agregados
    End If
End Sub
Private Sub cmdEliminar_Click()
    On Error GoTo Error 'Verificar si ocurre un error tratar de borrar un elemento.
    lstPaises.RemoveItem (lstPaises.ListIndex) 'Borra el elemento
    txtCantidadPaises.Text = lstPaises.ListCount
    Exit Sub 'Indica que lo que esta debajo solo ocurrirá cuando pase algún error.
    
Error:
    MsgBox ("No existen elementos seleccionados.")
End Sub

Private Sub cmdSalir_Click()
    End 'Finaliza la aplicación
End Sub

