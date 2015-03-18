VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   1095
   ClientTop       =   2070
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5910
   Begin VB.CommandButton Command1 
      Caption         =   "Add Rows"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

With MSFlexGrid1
'Add four entries to table with each click
.AddItem "North" & vbTab & "45,000" & vbTab & "53,000"
.AddItem "South" & vbTab & "20,000" & vbTab & "25,000"
.AddItem "East" & vbTab & "38,000" & vbTab & "77,300"
.AddItem "West" & vbTab & "102,000" & vbTab & "87,500"
End With

End Sub

Private Sub Form_Load()

With MSFlexGrid1    'use shorthand "With" notation

'Create headings for Columns 1 and 2
.TextMatrix(0, 1) = "Q1 1999"
.TextMatrix(0, 2) = "Q2 1999"

'Select headings
.Row = 0
.Col = 1
.RowSel = 0
.ColSel = 2

'Format headings with bold and align on center
.FillStyle = flexFillRepeat 'fill entire selection
.CellFontBold = True
.CellAlignment = flexAlignCenterCenter

'Add three entries for first row
.TextMatrix(1, 0) = "International"  'title column (0)
.TextMatrix(1, 1) = "55000"          'col 1
.TextMatrix(1, 2) = "83000"          'col 2
End With

End Sub
