VERSION 5.00
Begin VB.Form frmMainForm 
   Caption         =   "Naming Conventions"
   ClientHeight    =   3555
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   3930
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblInstructions 
      Caption         =   "To exit the program, click Quit."
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Welcome to the program!"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdQuit_Click()
    End
End Sub


