VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Quick Note"
   ClientHeight    =   4230
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6720
   Begin VB.TextBox txtNote 
      Height          =   3015
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.53906e-29
   End
   Begin VB.Label Label1 
      Caption         =   "Type your note and then save it to disk."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuItemSave 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuItemDate 
         Caption         =   "Insert &Date"
      End
      Begin VB.Menu mnuItemExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuItemDate_Click()
    Wrap$ = Chr$(13) & Chr$(10) 'add date to string
    txtNote.Text = Date$ & Wrap$ & txtNote.Text
End Sub

Private Sub mnuItemExit_Click()
    End                         'quit program
End Sub

Private Sub mnuItemSave_Click()
'note: the entire file is stored in a string
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
    CommonDialog1.ShowSave      'display Save dialog
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1
        Print #1, txtNote.Text  'save string to file
        Close #1                'close file
    End If
End Sub

