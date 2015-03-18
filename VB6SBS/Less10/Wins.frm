VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Lucky Seven"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spin"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblWins 
      Alignment       =   2  'Center
      Caption         =   "Wins: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   1680
      Picture         =   "Wins.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Lucky Seven"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Image1.Visible = False         ' hide coins
    Label1.Caption = Int(Rnd * 10) ' pick numbers
    Label2.Caption = Int(Rnd * 10)
    Label3.Caption = Int(Rnd * 10)
    'if any caption is 7 display coin stack and beep
    If (Label1.Caption = 7) Or (Label2.Caption = 7) _
      Or (Label3.Caption = 7) Then
        Image1.Visible = True
        Beep
        Wins = Wins + 1
        lblWins.Caption = "Wins: " & Wins
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub
