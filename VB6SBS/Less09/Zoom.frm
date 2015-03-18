VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Aproximación a la Tierra"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1200
      Picture         =   "Zoom.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    Image1.Height = Image1.Height + 200
    Image1.Width = Image1.Width + 200
End Sub
