VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Examinador de datos"
   ClientHeight    =   4620
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "Añadir"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtYear 
      DataField       =   "Year Published"
      DataSource      =   "datBiblio"
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtISBN 
      DataField       =   "ISBN"
      DataSource      =   "datBiblio"
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      DataSource      =   "datBiblio"
      Height          =   645
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Data datBiblio 
      Caption         =   "Biblio.mdb"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb6Sbs\Extras\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Titles"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblISBN 
      Caption         =   "ISBN:"
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
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblYear 
      Caption         =   "Año"
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
      TabIndex        =   0
      Top             =   3360
      Width           =   495
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
      Caption         =   "Título libro:"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image imgBook 
      Height          =   735
      Left            =   4920
      Picture         =   "Infolibro.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblDescription 
      Caption         =   "Una lista de libros sobre base de datos y programación"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblHead 
      Caption         =   "Base de datos Bibliográfica"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAñadir_Click()
    mensaje$ = "Introduzca un nuevo registro y pulse el botón flecha izquierda."
    reply = MsgBox(mensaje$, vbOKCancel, "Añadir Registro")
    If reply = vbOK Then            'si el usuario pulsa Aceptar
        txtTitle.SetFocus           'mueve el cursor al cuadro de título
        datBiblio.Recordset.AddNew  'y obtén un nuevo registro
        'define el campo PubID como 14 (este campo es necesario
        datBiblio.Recordset.PubID = 14  'para Biblio.mdb)
    End If
End Sub

Private Sub cmdBuscar_Click()
    mensaje$ = "Introduzca el título completo del libro."
    'obtiene la cadena que se utilizará en la búsqueda del título
    SearchStr$ = InputBox(mensaje$, "Búsqueda del libro")
    datBiblio.Recordset.Index = "Title"      'usa título
    datBiblio.Recordset.Seek "=", SearchStr$ 'y busca
    If datBiblio.Recordset.NoMatch Then      'si no encuentra ninguno
        datBiblio.Recordset.MoveFirst        'va al primer registro.
    End If
End Sub
Private Sub cmdBorrar_Click()
    mensaje$ = "¿Seguro que quiere borrar este registro?"
    respuesta = MsgBox(mensaje$, vbOKCancel, "Borrar registro")
    If respuesta = vbOK Then         'si el usuario pulsa Aceptar
        datBiblio.Recordset.Delete   'borra el registro actual
        datBiblio.Recordset.MoveNext 'mueve al siguiente registro
    End If
End Sub
Private Sub cmdSalir_Click()
    End
End Sub
