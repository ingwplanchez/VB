VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Visor de datos"
   ClientHeight    =   4620
   ClientLeft      =   2115
   ClientTop       =   2070
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtTime 
      DataField       =   "DaysAndTimes"
      DataSource      =   "datStudent"
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtDept 
      DataField       =   "Department"
      DataSource      =   "datStudent"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtProf 
      DataField       =   "Prof"
      DataSource      =   "datStudent"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "ClassName"
      DataSource      =   "datStudent"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Data datStudent 
      Caption         =   "Students.mdb"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb6Sbs\Less03\Students.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "Classes"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblProf 
      Caption         =   "Profesor"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblDept 
      Caption         =   "Departamento"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Caption         =   "D�as/Horas"
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
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   240
      X2              =   5400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblTitle 
      Caption         =   "T�tulo de la clase"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image imgBook 
      Height          =   735
      Left            =   5040
      Picture         =   "BuscarDatos.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblDescription 
      Caption         =   "Un visor de bases de datos que muestra informaci�n de clase de la actual Tabla de Horarios de la Universidad"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblHead 
      Caption         =   "Lista de cursos Universitarios"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    mensaje$ = "Introduzca el t�tulo completo del curso."
    'obtiene la cadena que se utilizar� en la b�squeda del t�tulo
    SearchStr$ = InputBox(mensaje$, "B�squeda del curso")
    datStudent.Recordset.Index = "ClassName"   'usa Nombre Clase
    datStudent.Recordset.Seek "=", SearchStr$  'y busca
    If datStudent.Recordset.NoMatch Then       'si no encuentra ninguno
        datStudent.Recordset.MoveFirst         'va al primer registro.
    End If
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
