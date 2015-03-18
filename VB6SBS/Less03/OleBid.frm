VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.OLE OLE3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Paint.Picture"
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   4080
      OleObjectBlob   =   "OleBid.frx":0000
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OLE OLE2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   2160
      OleObjectBlob   =   "OleBid.frx":1B618
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Word.Document.8"
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   240
      OleObjectBlob   =   "OleBid.frx":1CC30
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Site drawings"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Bid calculator"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Estimate scratchpad"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "A construction estimate front end featuring Word, Excel, and Paint"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Bid Estimator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
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
