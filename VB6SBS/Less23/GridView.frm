VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Grid View"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4440
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "Seattle"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdFindTxt 
      Caption         =   "Find Text"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
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
      Bindings        =   "GridView.frx":0000
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
Private Sub cmdClose_Click()
    Unload Form2
End Sub

Private Sub cmdFindTxt_Click()
    'Select entire grid and remove bold formatting
    '(to remove the results of previous finds)
    MSFlexGrid1.FillStyle = flexFillRepeat
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.CellFontBold = False
    
    'Initialize ProgressBar to track search
    ProgressBar1.Min = 0
    ProgressBar1.Max = MSFlexGrid1.Rows - 1
    ProgressBar1.Visible = True
    
    'Search the grid cell by cell for find text
    MSFlexGrid1.FillStyle = flexFillSingle
    For i = 0 To MSFlexGrid1.Cols - 1
        For j = 1 To MSFlexGrid1.Rows - 1
        'Display current row location on ProgressBar
        ProgressBar1.Value = j
            'If current cell matches find text box
            If InStr(MSFlexGrid1.TextMatrix(j, i), _
            Text1.Text) Then
                '...select cell and format bold
                MSFlexGrid1.Col = i
                MSFlexGrid1.Row = j
                MSFlexGrid1.CellFontBold = True
            End If
        Next j
    Next i
    ProgressBar1.Visible = False 'hide ProgressBar
End Sub

Private Sub cmdSort_Click()
    'Set column 2 (LastName) as the sort key
    MSFlexGrid1.Col = 2
    'Sort grid in ascending order
    MSFlexGrid1.Sort = 1
End Sub

