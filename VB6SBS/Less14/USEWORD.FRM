VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Personal Spelling Checker"
   ClientHeight    =   2760
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5970
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Spelling"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter one or more words in the text box, and then click Check Spelling."
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim X As Object      'create Word object variable
Set X = CreateObject("Word.Application")
X.Visible = False    'hide Word
X.Documents.Add      'open a new document
X.Selection.Text = Text1.Text  'copy text box to document
X.ActiveDocument.CheckSpelling 'run spell/grammar check
Text1.Text = X.Selection.Text  'copy results back
X.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
X.Quit               'quit Word

Set X = Nothing      'release object variable
End Sub



Private Sub Command2_Click()
    End
End Sub



