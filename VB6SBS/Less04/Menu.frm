VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   1950
   ClientTop       =   2565
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Menu mnuClock 
      Caption         =   "&Clock"
      Begin VB.Menu mnuTimeItem 
         Caption         =   "&Time"
      End
      Begin VB.Menu mnuDateItem 
         Caption         =   "&Date"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDateItem_Click()
    Label1.Caption = Date
End Sub

Private Sub mnuTimeItem_Click()
    Label1.Caption = Time
End Sub
