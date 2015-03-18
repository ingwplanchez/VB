VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Instructor"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Students.mdb"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6SBS\Less03\Students.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Instructors"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Profesor"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
