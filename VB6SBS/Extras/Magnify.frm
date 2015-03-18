VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Magnify"
   ClientHeight    =   3360
   ClientLeft      =   1935
   ClientTop       =   1815
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   5415
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.54016e-29
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A program that magnifies bit maps."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3855
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuCloseItem_Click()
    Image1.Picture = LoadPicture("")
    mnuCloseItem.Enabled = False
    Form1.Caption = "Magnify"
End Sub

Private Sub mnuExitItem_Click()
    End
End Sub

Private Sub mnuOpenItem_Click()
    CommonDialog1.Filter = "Bitmaps (*.BMP)|*.BMP"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
    Form1.Caption = "Magnify " & CommonDialog1.FileName
    mnuCloseItem.Enabled = True
End Sub

