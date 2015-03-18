VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Prueba de datos"
   ClientHeight    =   3990
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7365
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Dato ejemplo"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione un tipo de datos"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Tipos fundamentales de datos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3735
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
    'Estas líneas añaden elementos al cuadro de lista List1
    List1.AddItem "Entero"
    List1.AddItem "Entero Largo"
    List1.AddItem "Simple precisión"
    List1.AddItem "Doble precisión"
    List1.AddItem "Monetario"
    List1.AddItem "Cadena"
    List1.AddItem "Booleano"
    List1.AddItem "Fecha"
    List1.AddItem "Variante"
End Sub

Private Sub List1_Click()
    'Sección de declaración de variables
    Dim Pajaros%, Ingresos&, Precio!, Pi#, Deuda@, Perro$, Total
    Dim Bandera As Boolean
    Dim Aniversario As Date
    
    'Select Case procesa la elección del usuario
    Select Case List1.ListIndex
    Case 0
        Pajaros% = 37
        Label4.Caption = Pajaros%
    Case 1
        Ingresos& = 350000
        Label4.Caption = Ingresos&
    Case 2
        Precio! = -1234.123
        Label4.Caption = Precio!
    Case 3
        Pi# = 3.1415926535
        Label4.Caption = Pi#
    Case 4
        Deuda@ = 299950.95
        Label4.Caption = Deuda@
    Case 5
        Perro$ = "Pastor alemán de pura raza"
        Label4.Caption = Perro$
    Case 6  'True se almacena como -1 en el código, False como 0
        Bandera = True
        Label4.Caption = Bandera
    Case 7  'Observe el símbolo # y la función Format
        Aniversario = #11/19/1963#
        Label4.Caption = Format$(Aniversario, "dddd, mmmm dd, yyyy")
    Case 8
        Precio = 99.95
        Label4.Caption = Precio
    End Select
End Sub

