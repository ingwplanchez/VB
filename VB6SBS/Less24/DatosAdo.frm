VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Control ADO"
   ClientHeight    =   3885
   ClientLeft      =   2805
   ClientTop       =   2415
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4965
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "PhoneNumber"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Text            =   "Texto2"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "LastName"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Text            =   "Texto1"
      Top             =   960
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   1
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Student Records"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Student Records"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Students"
      Caption         =   "Students.mdb"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Realice una doble pulsación sobre el campo para almacenar todos los datos en un archivo"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Registros totales:"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Si no se encuentra ya en el último registro, pasar al siguiente
If Not Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveNext
End If

End Sub

Private Sub Command2_Click()

'Si no se encuentra ya en el primer registro, pasar al anterior
If Not Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MovePrevious
End If

End Sub

Private Sub Form_Load()

'Poblar el cuadro de lista con nombres de campo
For i = 1 To Adodc1.Recordset.Fields.Count - 1
    List1.AddItem Adodc1.Recordset.Fields(i).Name
Next i

'Mostrar el número total de registros
Label1.Caption = "Registros totales: " & _
    Adodc1.Recordset.RecordCount

End Sub

Private Sub List1_DblClick()

'Crear una constante para almacenar el nombre del archivo de texto
Const myFile = "c:\vb6sbs\less24\names.txt"

'Abrir archivo utilizando Append (para poder trabajar con varios campos)
Open myFile For Append As #1
Print #1, String$(30, "-")  'imprimir una línea punteada

Adodc1.Recordset.MoveFirst  'mover al primer registro
x = List1.ListIndex + 1     'escoger el elemento pulsado

'Para cada registro de la base de datos, escribir campo en el disco
For i = 1 To Adodc1.Recordset.RecordCount
    Print #1, Adodc1.Recordset.Fields(x).Value
    Adodc1.Recordset.MoveNext
Next i

'Imprimir un mensaje y cerrar archivo
MsgBox Adodc1.Recordset.Fields(x).Name & _
    " el campo ha sido escrito en " & myFile
Close #1
Adodc1.Recordset.MoveFirst

End Sub
