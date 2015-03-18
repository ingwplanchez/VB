VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Revista Musical"
   ClientHeight    =   4005
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6165
   Begin VB.Data Data1 
      Caption         =   "Talent"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb5Sbs\Less12\Talent.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Artists"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Revisar Ortografía"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Comentarios"
      DataSource      =   "Data1"
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   5895
   End
   Begin VB.TextBox Text4 
      DataField       =   "Estado"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "Ciudad"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "Dirección"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Bloc de Notas de la Revista Musical"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim X As Object       'crea un objeto de Word
Set X = CreateObject("Word.Application")
    
X.Visible = False     'oculta Word
X.Documents.Add       'abre un nuevo documento
X.Selection.Text = Text5.Text  'copia cuadro de texto al documento
X.ActiveDocument.CheckGrammar  'ejecuta el corrector ortográfico/gramatical
Text5.Text = X.Selection.Text  'copia el texto corregido en VB
X.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
X.Application.Quit    'sale de Word

Set X = Nothing       'libera la variable objeto
End Sub

Private Sub Command2_Click()
    End
End Sub


