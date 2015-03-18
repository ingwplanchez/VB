VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5445
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   120
      Normal          =   0   'False
      Pattern         =   "*.bmp;*.wmf;*.ico"
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1605
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    SelectedFile = File1.Path & "\" & File1.FileName
    Image1.Picture = LoadPicture(SelectedFile)
End Sub

