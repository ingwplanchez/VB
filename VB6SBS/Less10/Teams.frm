VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Assign Department Teams"
   ClientHeight    =   3375
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMkt 
      Caption         =   "Add Name"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdSales 
      Caption         =   "Add Name"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMkt 
      Height          =   1575
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSales 
      Height          =   1575
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblMkt 
      Caption         =   "Marketing"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblSales 
      Caption         =   "Sales"
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
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSales_Click()
    AddName "Sales", SalesPosition$
    txtSales.Text = txtSales.Text & SalesPosition$
End Sub

Private Sub cmdMkt_Click()
    AddName "Marketing", MktPosition$
    txtMkt.Text = txtMkt.Text & MktPosition$
End Sub

Private Sub lblQuit_Click()
    End
End Sub



