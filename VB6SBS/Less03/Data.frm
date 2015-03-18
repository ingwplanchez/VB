VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   2010
   ClientTop       =   2340
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4215
   Begin VB.TextBox Text1 
      DataField       =   "Instructor"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Students.mdb"
      Connect         =   "Access"
      DatabaseName    =   "C:\Vb6Sbs\Less03\Students.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Instructors"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Instructor"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
