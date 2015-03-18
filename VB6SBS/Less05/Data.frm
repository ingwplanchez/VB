VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Data Test"
   ClientHeight    =   3990
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7365
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Sample data"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Choose a data type"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fundamental Data Types"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   3015
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
    'these lines add items to the List1 list box
    List1.AddItem "Integer"
    List1.AddItem "Long integer"
    List1.AddItem "Single precision"
    List1.AddItem "Double precision"
    List1.AddItem "Currency"
    List1.AddItem "String"
    List1.AddItem "Boolean"
    List1.AddItem "Date"
    List1.AddItem "Variant"
End Sub

Private Sub List1_Click()
    'Variable declaration section
    Dim Birds%, Loan&, Price!, Pie#, Debt@, Dog$, Total
    Dim Flag As Boolean
    Dim Birthday As Date
    
    'Select Case processes the user's choice
    Select Case List1.ListIndex
    Case 0
        Birds% = 37
        Label4.Caption = Birds%
    Case 1
        Loan& = 350000
        Label4.Caption = Loan&
    Case 2
        Price! = -1234.123
        Label4.Caption = Price!
    Case 3
        Pie# = 3.1415926535
        Label4.Caption = Pie#
    Case 4
        Debt@ = 299950.95
        Label4.Caption = Debt@
    Case 5
        Dog$ = "German Wire-haired Pointer"
        Label4.Caption = Dog$
    Case 6  'True is stored as -1 in code, False as 0
        Flag = True
        Label4.Caption = Flag
    Case 7  'Note # symbol and Format function here
        Birthday = #11/19/63#
        Label4.Caption = Format$(Birthday, "dddd, mmmm dd, yyyy")
    Case 8
        Price = 99.95
        Label4.Caption = Price
    End Select
End Sub

