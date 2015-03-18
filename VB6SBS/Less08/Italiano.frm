VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Italiano paso a paso"
   ClientHeight    =   3270
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4545
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Pulse dos veces sobre la palabra para ver su definición"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Verbos de esta semana"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Vocabulario italiano"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2655
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
    List1.AddItem "aprire"
    List1.AddItem "ascoltare"
    List1.AddItem "bere"
    List1.AddItem "cucinare"
    List1.AddItem "dormire"
    List1.AddItem "pagare"
    List1.AddItem "scrivere"
End Sub

Private Sub List1_DblClick()
    Select Case List1.ListIndex
    Case 0
        Def = "abrir"
    Case 1
        Def = "escuchar"
    Case 2
        Def = "beber"
    Case 3
        Def = "cocinar"
    Case 4
        Def = "dormir"
    Case 5
        Def = "pagar"
    Case 6
        Def = "escribir"
    End Select
    
    MsgBox (Def), , List1.Text
End Sub



