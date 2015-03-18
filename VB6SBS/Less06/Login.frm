VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5625
   Begin VB.CommandButton Command1 
      Caption         =   "Log in"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    UserName = InputBox("Enter your first name.")
    If UserName = "Laura" Then
        MsgBox ("Welcome, Laura!  Ready to start your PC?")
        Form1.Picture = _
          LoadPicture("c:\vb6sbs\less06\pcomputr.wmf")
    ElseIf UserName = "Marc" Then
        MsgBox ("Welcome, Marc!  Ready to display your Rolodex?")
        Form1.Picture = _
          LoadPicture("c:\vb6sbs\less06\rolodex.wmf")
    Else
        MsgBox ("Sorry, I don't recognize you.")
        End   'quit the program
    End If
End Sub

