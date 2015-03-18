VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Italian Step by Step"
   ClientHeight    =   3270
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4545
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Double-click word for definition."
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "This week's verbs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Italian Vocabulary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   1
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
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Base = "0{05A2C84E-0A5B-11D0-93C3-444553540000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
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
        Def = "to open"
    Case 1
        Def = "to listen"
    Case 2
        Def = "to drink"
    Case 3
        Def = "to cook"
    Case 4
        Def = "to sleep"
    Case 5
        Def = "to pay, pay for"
    Case 6
        Def = "to write"
    End Select
    
    Load Form2
    Form2.Label1 = List1.Text
    Form2.Text1 = Def
    Form2.Show
End Sub



