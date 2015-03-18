VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Talent EIS"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "State"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Talent"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb6Sbs\Less14\Talent.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Artists"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "WorkPhone"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "Musiceis.frx":0000
      SourceDoc       =   "C:\VB6SBS\LESS14\SALES_98.XLS"
      TabIndex        =   8
      Top             =   3120
      Width           =   3495
   End
   Begin VB.OLE OLE2 
      Class           =   "Excel.Sheet.8"
      Height          =   2655
      Left            =   3720
      OleObjectBlob   =   "Musiceis.frx":3218
      SourceDoc       =   "C:\Vb6Sbs\Less14\SALES_98.xls!Sheet1![SALES_98.xls]Sheet1 Chart 2"
      TabIndex        =   7
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Seattle Beat Talent Tracker"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   240
      Picture         =   "Musiceis.frx":5630
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2295
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
OLE1.Update
OLE2.Update
End Sub
