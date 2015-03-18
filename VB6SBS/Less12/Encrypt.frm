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
      FontSize        =   1.17491e-38
   End
   Begin VB.Label lblFile 
      Caption         =   "Type your note and then save it to disk."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open Encrypted File..."
      End
      Begin VB.Menu mnuItemSave 
         Caption         =   "&Save Encrypted File..."
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
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
    CommonDialog1.ShowSave           'display Save dialog
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11      'display hourglass
        lblFile.Caption = CommonDialog1.FileName
        'save text with encryption scheme (ASCII code + 1)
        encrypt$ = ""  'initialize encryption string
        charsInFile% = Len(txtNote.Text) 'find string length
        For i% = 1 To charsInFile%   'for each character in file
            letter$ = Mid(txtNote.Text, i%, 1) 'read next char
            'determine ASCII code of char and add one to it
            encrypt$ = encrypt$ & Chr$(Asc(letter$) + 1)
        Next i%
        Open CommonDialog1.FileName For Output As #1 'open file
        Print #1, encrypt$           'save encrypted text to file
        txtNote.Text = encrypt$
        Close #1                     'close file
        CommonDialog1.FileName = ""  'clear filename
        Form1.MousePointer = 0       'reset mouse
    End If
End Sub

Private Sub mnuOpenItem_Click()
    Wrap$ = Chr$(13) + Chr$(10) 'create wrap character
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
    CommonDialog1.ShowOpen      'display Open dialog box
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11 'display hourglass
        Open CommonDialog1.FileName For Input As #1 'open file
        On Error GoTo Problem:  'set error handler
        Do Until EOF(1)         'copy each line of text to
            Line Input #1, LineOfText$  'AllText$ string
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        'now, decrypt string by subtracting one from ASCII code
        decrypt$ = ""   'initialize string for decryption
        charsInFile = Len(AllText$)  'get length of string
        For i% = 1 To charsInFile    'loop once for each char
            letter$ = Mid(AllText$, i%, 1)  'get char with Mid
            decrypt$ = decrypt$ & Chr$(Asc(letter) - 1) 'subtract 1
        Next i%                       'and build new string
        txtNote.Text = decrypt$ 'then display converted string
        txtNote.Enabled = True  'and enable scroll bars
        lblFile.Caption = CommonDialog1.FileName 'set caption
CleanUp:                        'when finished...
        Form1.MousePointer = 0  'reset mouse
        Close #1                'close file
        CommonDialog1.FileName = ""   'clear filename
    End If
    Exit Sub
Problem:  'if there is a problem, display appropriate message
    MsgBox "Error Opening File", , Err.Description
    lblFile.Caption = ""        'remove caption
    txtNote.Text = ""           'clear text box
    Resume CleanUp:   'finally, finish with CleanUp routine
End Sub
