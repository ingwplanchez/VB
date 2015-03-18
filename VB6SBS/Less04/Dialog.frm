VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   4096
      FontSize        =   2.52734e-29
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuCloseItem 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "&Clock"
      Begin VB.Menu mnuTimeItem 
         Caption         =   "&Time"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDateItem 
         Caption         =   "&Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTextColorItem 
         Caption         =   "TextCo&lor..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCloseItem_Click()
Image1.Picture = LoadPicture("")
mnuCloseItem.Enabled = False

End Sub

Private Sub mnuDateItem_Click()
    Label1.Caption = Date
End Sub

Private Sub mnuExitItem_Click()
    End
End Sub

Private Sub mnuOpenItem_Click()
    CommonDialog1.Filter = "Metafiles (*.WMF)|*.WMF"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
    mnuCloseItem.Enabled = True
End Sub

Private Sub mnuTextColorItem_Click()
    CommonDialog1.Flags = &H1&
    CommonDialog1.ShowColor
    Label1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnuTimeItem_Click()
    Label1.Caption = Time
End Sub
