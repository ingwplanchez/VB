VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Scoreboard"
   ClientHeight    =   3120
   ClientLeft      =   1170
   ClientTop       =   2160
   ClientWidth     =   5865
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5865
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdNextInning 
      Caption         =   "Next Inning"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtHome 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtAway 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblHome2 
      Caption         =   "Mariners"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblAway2 
      Caption         =   "Yankees"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   2760
      X2              =   5400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   735
      Left            =   2760
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblHome1 
      Caption         =   "Mariners:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblAway1 
      Caption         =   "Yankees:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblInning 
      Caption         =   "Inning 1 Scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Yankees vs. Mariners"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Baseball Scoreboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
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

Private Sub cmdNextInning_Click()
    'put scores in the array each inning
    Scoreboard(1, Inning) = txtAway.Text
    Scoreboard(2, Inning) = txtHome.Text
    
    'then display the scores in the score box
    '(CurrentX and CurrentY control the cursor)
    CurrentX = 2626 + (Inning * 224)
    CurrentY = 1050
    Print txtAway.Text
    CurrentX = 2626 + (Inning * 224)
    CurrentY = 1400
    Print txtHome.Text
    
    'change to the next inning
    Inning = Inning + 1
    'and if the game is over, display the results
    If Inning > 9 Then
        cmdNextInning.Enabled = False
        AddUpScores  'this procedure (in BASEBALL.BAS)
    Else             'calculates the score
        lblInning.Caption = "Inning " & Inning & " Scores"
    End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Inning = 1         'initialize Inning to 1
    CurrentX = 2850    'place cursor at top of box
    CurrentY = 750
    Show 'enable output during load and print header
    Print "1   2   3   4   5   6   7   8   9";
    Print "     Final"
End Sub

