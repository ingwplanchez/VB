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
   Begin VB.Label lblFile 
      Caption         =   "Type text or open file for sorting."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuItemOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuItemClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuItemSave 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuItemSortText 
         Caption         =   "Sor&t Text"
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
Private Sub mnuItemClose_Click()
    txtNote.Text = ""            'clear text box
    lblFile.Caption = "Type text or open file for sorting."
    mnuItemClose.Enabled = False 'dim Close command
    mnuItemOpen.Enabled = True   'enable Open command
End Sub

Private Sub mnuItemDate_Click()
    Wrap$ = Chr$(13) & Chr$(10) 'add date to string
    txtNote.Text = Date$ & Wrap$ & txtNote.Text
End Sub

Private Sub mnuItemExit_Click()
    End                         'quit program
End Sub

Private Sub mnuItemOpen_Click()
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen       'display Open dialog box
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11  'display hourglass
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo TooBig:    'set error handler
        Do Until EOF(1)          'then read lines from file
            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        lblFile.Caption = CommonDialog1.FileName
        txtNote.Text = AllText$  'display file
        txtNote.Enabled = True
        mnuItemClose.Enabled = True
        mnuItemOpen.Enabled = False 'enable scroll
CleanUp:
        Form1.MousePointer = 0   'reset mouse
        Close #1                 'close file
    End If
    Exit Sub
TooBig:             'error handler displays message
    MsgBox ("The specified file is too large.")
    Resume CleanUp: 'then jumps to CleanUp routine

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

Private Sub mnuItemSortText_Click()
'determine number of lines in text box object (txtNote)
lineCount% = 0  'this variable holds the total number of lines
charsInFile% = Len(txtNote.Text)  'get total characters in box
For i% = 1 To charsInFile%  'move one char at a time through box
    letter$ = Mid(txtNote.Text, i%, 1) 'put next char in letter$
    If letter$ = Chr$(13) Then 'if carriage ret found (end of line!)
        lineCount% = lineCount% + 1  'go to next line (add to count)
        i% = i% + 1   'skip linefeed char (which always follows cr)
    End If
Next i%

'build an array to hold the text in the text box
ReDim strArray$(lineCount%) 'create array of proper size
curline% = 1
ln$ = ""  'use ln$ to build lines one character at a time
For i% = 1 To charsInFile%     'loop through text again
    letter$ = Mid(txtNote.Text, i%, 1)
    If letter$ = Chr$(13) Then 'if carriage return found
        curline% = curline% + 1    'increment line count
        i% = i% + 1            'skip linefeed char
        ln$ = ""               'clear line and go to next
    Else
        ln$ = ln$ & letter$    'add letter to line
        strArray$(curline%) = ln$  'and put in array
   End If
Next i%

'sort array
ShellSort strArray$(), lineCount%

'then display sorted array in text box
txtNote.Text = ""
Wrap$ = Chr$(13) & Chr$(10) 'add date to string
curline% = 1
For i% = 1 To lineCount%
    txtNote.Text = txtNote.Text & strArray$(curline%) & Wrap$
    curline% = curline% + 1
Next i%

End Sub
