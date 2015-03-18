VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5475
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione un país"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Programa de bienvenida internacional"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4695
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
    List1.AddItem "Inglaterra"
    List1.AddItem "Alemania"
    List1.AddItem "España"
    List1.AddItem "Italia"
End Sub

Private Sub List1_Click()
    Label3.Caption = List1.Text
    Select Case List1.ListIndex
    Case 0
        Label4.Caption = "Hello, programmer"
    Case 1
        Label4.Caption = "Hallo, Programmierer"
    Case 2
        Label4.Caption = "Hola, programador"
    Case 3
        Label4.Caption = "Ciao, programmatori"
    End Select
End Sub

