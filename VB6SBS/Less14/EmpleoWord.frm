VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Revisor Ortográfico Personal"
   ClientHeight    =   2760
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6405
   Begin VB.CommandButton Command2 
      Caption         =   "Fin"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Revisar ortografía"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba una o más palabras en el cuadro de texto y pulse Revisar Ortografía"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim X As Object      'creación de una variable objeto Word
Set X = CreateObject("Word.Application")
X.Visible = False    'ocultar Word
X.Documents.Add      'abrir un nuevo documento
X.Selection.Text = Text1.Text  'copiar cuadro de texto al documento
X.ActiveDocument.CheckSpelling 'ejecutar corrector ortográfico/gramática
Text1.Text = X.Selection.Text  'copiar los resultados
X.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
X.Quit               'salir Word

Set X = Nothing      'liberar variable objeto
End Sub



Private Sub Command2_Click()
    End
End Sub



