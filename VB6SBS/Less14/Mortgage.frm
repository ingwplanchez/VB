VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mortgage"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "'$'#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "360"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "0.09"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Pmt"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Principal"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Months"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Interest"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Monthly Payment Calculator"
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
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim xl As Object   'create object for Excel
Dim loanpmt        'declare return value
                   'if all fields contain values
If Text1.Text <> "" And Text2.Text <> "" _
And Text3.Text <> "" Then  'create object and call Pmt
    Set xl = CreateObject("Excel.Sheet")
    loanpmt = xl.application.WorksheetFunction.Pmt _
        (Text1.Text / 12, Text2.Text, Text3.Text)
    MsgBox "The monthly payment is " & _
        Format(Abs(loanpmt), "$#.##"), , "Mortgage"
    xl.application.quit
    Set xl = Nothing
Else
    MsgBox "All fields required", , "Mortgage"
End If

End Sub
