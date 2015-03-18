VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hipoteca"
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
      Text            =   "0,09"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular Pago"
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
      Caption         =   "Meses"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Interés"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cálculo de Pagos mensuales"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim xl As Object   'crear objeto para Excel
Dim loanpmt        'declarar el valor a transferir
                   'si todos los campos contienen valores
If Text1.Text <> "" And Text2.Text <> "" _
And Text3.Text <> "" Then  'crear un objeto y llamar a Pmt
    Set xl = CreateObject("Excel.Sheet")
    loanpmt = xl.Application.WorksheetFunction.Pmt _
        (Text1.Text / 12, Text2.Text, Text3.Text)
    MsgBox "El pago mensual es " & _
        Format(Abs(loanpmt), "$#.##"), , "Hipoteca"
    xl.Application.Quit
    Set xl = Nothing
Else
    MsgBox "Tiene que llenar todos los campos", , "Hipoteca"
End If

End Sub
