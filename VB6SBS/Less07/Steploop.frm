VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Display Icons"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   28
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   26
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   24
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   22
      Left            =   480
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 22 To 28 Step 2
    Image1(i).Picture = _
      LoadPicture("c:\vb6sbs\less07\misc" & i & ".ico")
Next i
End Sub
