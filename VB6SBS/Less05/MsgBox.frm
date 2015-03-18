VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Box Tester"
   ClientHeight    =   2385
   ClientLeft      =   1635
   ClientTop       =   1860
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4950
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input Box"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Dim Prompt, FullName
    Prompt = "Please enter your name."
    
    FullName = InputBox$(Prompt)
    MsgBox (FullName), , "Input Results"
End Sub


Private Sub Command2_Click()
    End
End Sub

