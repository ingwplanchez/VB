VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Temperatures"
   ClientHeight    =   3300
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdDisplayTemps 
      Caption         =   "Display Temperatures"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnterTemps 
      Caption         =   "Enter Temperatures"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDisplayTemps_Click()
    Print "High temperatures:"
    Print
    For i% = 1 To Days
        Print "Day "; i%, Temperatures(i%)
        Total! = Total! + Temperatures(i%)
    Next i%
    Print
    Print "Average high temperature:  "; Total! / Days
End Sub

Private Sub cmdEnterTemps_Click()
    Cls
    Days = InputBox("How many days?", "Create Array")
    If Days > 0 Then ReDim Temperatures(Days)
    Prompt$ = "Enter the high temperature."
    For i% = 1 To Days
        Title$ = "Day " & i%
        Temperatures(i%) = InputBox(Prompt$, Title$)
    Next i%
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
