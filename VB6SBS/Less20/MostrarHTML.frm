VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mostrar documento HTML"
   ClientHeight    =   2730
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6990
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "http://www.microsoft.com/"
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar HTML"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Introduzca el URL de un documento HTML y pulse el botón mostrar HTML."
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarar una variable para el URL actual
Public Explorer As SHDocVw.InternetExplorer

Private Sub Command1_Click()
    On Error GoTo manejadorerror
    Set Explorer = New SHDocVw.InternetExplorer
    Explorer.Visible = True
    Explorer.Navigate Combo1.Text
    Exit Sub
manejadorerror:
    MsgBox "Error visualizando el archivo", , Err.Description
End Sub

Private Sub Form_Load()
'Añadir unos pocos servidores web al cuadro combo durante
'la puesta en marcha
    'página inicial de Microsoft Corp.
    Combo1.AddItem "http://www.microsoft.com/"
    'página inicial de Microsoft Press
    Combo1.AddItem "http://mspress.microsoft.com/"
    'página inicial de Microsoft Visual Basic Programming
    Combo1.AddItem "http://www.microsoft.com/vbasic/"
    'recursos de Fawcette Publication para programación en VB
    Combo1.AddItem "http://www.windx.com"
    'página inicial de VB de Carl y Gary (no-Microsoft)
    Combo1.AddItem "http://www.apexsc.com/vb/"
End Sub

