VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Visor hoja"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "Seattle"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuscarTxt 
      Caption         =   "Buscar texto"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOrdenar 
      Caption         =   "Ordenar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb6Sbs\Less03\Students.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Students"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "VerRej.frx":0000
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
    Unload Form2
End Sub

Private Sub cmdBuscarTxt_Click()
    'Seleccionar toda la hoja y borrar negrita
    '(para eliminar el resultado de las operaciones de
    'búsqueda anteriores)
    MSFlexGrid1.FillStyle = flexFillRepeat
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.CellFontBold = False
    
    'Iniciar ProgressBar para supervisar la búsqueda
    ProgressBar1.Min = 0
    ProgressBar1.Max = MSFlexGrid1.Rows - 1
    ProgressBar1.Visible = True
    
    'Buscar la hoja celda a celda para encontrar el texto
    MSFlexGrid1.FillStyle = flexFillSingle
    For i = 0 To MSFlexGrid1.Cols - 1
        For j = 1 To MSFlexGrid1.Rows - 1
        'Mostrar la ubicación actual de ProgressBar
        ProgressBar1.Value = j
            'Si coincide la celda actual encontrar el cuadro de texto
            If InStr(MSFlexGrid1.TextMatrix(j, i), _
            Text1.Text) Then
                '...seleccionar la celda y poner en negrita
                MSFlexGrid1.Col = i
                MSFlexGrid1.Row = j
                MSFlexGrid1.CellFontBold = True
            End If
        Next j
    Next i
    ProgressBar1.Visible = False 'ocultar ProgressBar
End Sub

Private Sub cmdOrdenar_Click()
    'Definir columna 2 (LastName, Apellido) como clave de ordenación
    MSFlexGrid1.Col = 2
    'Ordenar la hoja en orden ascendente
    MSFlexGrid1.Sort = 1
End Sub

