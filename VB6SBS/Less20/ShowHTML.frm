VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Display HTML Document"
   ClientHeight    =   2730
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6990
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "http://www.microsoft.com/"
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display HTML"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the URL for a valid HTML document and click Display HTML"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare a variable for the current URL
Public Explorer As SHDocVw.InternetExplorer

Private Sub Command1_Click()
    On Error GoTo errorhandler
    Set Explorer = New SHDocVw.InternetExplorer
    Explorer.Visible = True
    Explorer.Navigate Combo1.Text
    Exit Sub
errorhandler:
    MsgBox "Error displaying file", , Err.Description
End Sub

Private Sub Form_Load()
'Add a few useful web sites to combo box at startup
    'Microsoft Corp. home page
    Combo1.AddItem "http://www.microsoft.com/"
    'Microsoft Press home page
    Combo1.AddItem "http://mspress.microsoft.com/"
    'Microsoft Visual Basic Programming home page
    Combo1.AddItem "http://www.microsoft.com/vbasic/"
    'Fawcette Publication resources for VB programming
    Combo1.AddItem "http://www.windx.com"
    'Carl and Gary's VB home page (non-Microsoft)
    Combo1.AddItem "http://www.apexsc.com/vb/"
End Sub

