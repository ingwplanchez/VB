VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Data Browser"
   ClientHeight    =   4410
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtYear 
      DataField       =   "Year Published"
      DataSource      =   "datBiblio"
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtISBN 
      DataField       =   "ISBN"
      DataSource      =   "datBiblio"
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      DataSource      =   "datBiblio"
      Height          =   525
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
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblYear 
      Caption         =   "Year:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   3120
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
      Caption         =   "Book title:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image imgBook 
      Height          =   735
      Left            =   4920
      Picture         =   "BookInfo.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblDescription 
      Caption         =   "A list of books about databases and database programming."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblHead 
      Caption         =   "The Database Bibliography"
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

Private Sub cmdAdd_Click()
    prompt$ = "Enter the new record, and then click the left arrow button."
    reply = MsgBox(prompt$, vbOKCancel, "Add Record")
    If reply = vbOK Then             'if the user clicks OK
        txtTitle.SetFocus            'move cursor to Title box
        datBiblio.Recordset.AddNew   'and get new record
        'set PubID field to 14 (this field is required
        datBiblio.Recordset.PubID = 14  'by biblio.mdb)
    End If
End Sub

Private Sub cmdDelete_Click()
    prompt$ = "Do you really want to delete this record?"
    reply = MsgBox(prompt$, vbOKCancel, "Delete Record")
    If reply = vbOK Then             'if the user clicks OK
        datBiblio.Recordset.Delete   'delete current record
        datBiblio.Recordset.MoveNext 'move to next record
    End If
End Sub

Private Sub cmdFind_Click()
    prompt$ = "Enter the full (complete) book title."
    'get the string to be used in the Title field search
    SearchStr$ = InputBox(prompt$, "Book Search")
    datBiblio.Recordset.Index = "Title"      'use Title
    datBiblio.Recordset.Seek "=", SearchStr$ 'and search
    datBiblio.Recordset.Index = "PrimaryKey" 'reset primary key
    If datBiblio.Recordset.NoMatch Then      'if no match
        datBiblio.Recordset.MoveFirst        'go to first record
    End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
