VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Download HTML Documents"
   ClientHeight    =   4695
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6990
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   80
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download HTML"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtURLbox 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "http://www.microsoft.com"
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txtNote 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1440
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.53906e-29
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the URL for a valid HTML document and click Download HTML."
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuItemHTML 
         Caption         =   "Save As &HTML..."
      End
      Begin VB.Menu mnuItemSave 
         Caption         =   "&Save As Text..."
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
'declare a variable for the current URL
Dim strUrl As String

Private Sub Command1_Click()
    On Error GoTo errorhandler
    strUrl = txtURLbox.Text
    'check for at least 11 characters ("http://www.")
    If Len(strUrl) > 11 Then
        'copy html document into text box
        txtNote.Text = Inet1.OpenURL(strUrl)
    Else
        MsgBox "Enter valid document name in the URL box"
    End If
    Exit Sub
errorhandler:
    MsgBox "Error opening URL", , Err.Description
End Sub

Private Sub mnuItemExit_Click()
    Unload Form1                       'quit program
End Sub

Private Sub mnuItemHTML_Click()
'note: the entire file is stored in a string
CommonDialog1.DefaultExt = "HTM"
CommonDialog1.Filter = "HTML files (*.HTML;*.HTM)|*.HTML;HTM"
CommonDialog1.ShowSave      'display Save dialog
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtNote.Text  'save string to file
    Close #1                'close file
End If
End Sub

Private Sub mnuItemSave_Click()
'note: the entire file is stored in a string
CommonDialog1.DefaultExt = "TXT"
CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
CommonDialog1.ShowSave      'display Save dialog
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtNote.Text  'save string to file
    Close #1                'close file
End If
End Sub

