VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Text Browser"
   ClientHeight    =   4155
   ClientLeft      =   1125
   ClientTop       =   1770
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5910
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.53906e-29
   End
   Begin VB.Label lblFile 
      Caption         =   "Load a text file with the Open command."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4815
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
    txtFile.Text = ""            'clear text box
    lblFile.Caption = "Load a text file with the Open command."
    mnuItemClose.Enabled = False 'dim Close command
    mnuItemOpen.Enabled = True   'enable Open command
    txtFile.Enabled = False      'disable text box
End Sub

Private Sub mnuItemExit_Click()
    End                          'quit program
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
        txtFile.Text = AllText$  'display file
        txtFile.Enabled = True
        mnuItemClose.Enabled = True
        mnuItemOpen.Enabled = False
CleanUp:
        Form1.MousePointer = 0   'reset mouse
        Close #1                 'close file
    End If
    Exit Sub
TooBig:             'error handler displays message
    MsgBox ("The specified file is too large.")
    Resume CleanUp: 'then jumps to CleanUp routine
End Sub

