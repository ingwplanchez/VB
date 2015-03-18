VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Botones Gráficos"
   ClientHeight    =   2505
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5745
   Begin VB.Image Image9 
      Height          =   330
      Left            =   600
      Picture         =   "Botones.frx":0000
      Top             =   2040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image8 
      Height          =   330
      Left            =   600
      Picture         =   "Botones.frx":018A
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   600
      Picture         =   "Botones.frx":0314
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image6 
      Height          =   330
      Left            =   120
      Picture         =   "Botones.frx":049E
      Top             =   2040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   330
      Left            =   120
      Picture         =   "Botones.frx":0628
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   120
      Picture         =   "Botones.frx":07B2
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   2760
      Picture         =   "Botones.frx":093C
      Tag             =   "Up"
      Top             =   720
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   2160
      Picture         =   "Botones.frx":0AC6
      Tag             =   "Up"
      Top             =   720
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1560
      Picture         =   "Botones.frx":0C50
      Tag             =   "Up"
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Pulse los botones para dar formato al texto de muestra"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Texto de muestra"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image1.Tag = "Up" Then
        Image1.Picture = Image4.Picture
        Label2.FontBold = True
        Image1.Tag = "Down"
    Else
        Image1.Picture = Image7.Picture
        Label2.FontBold = False
        Image1.Tag = "Up"
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image2.Tag = "Up" Then
        Image2.Picture = Image5.Picture
        Label2.FontItalic = True
        Image2.Tag = "Down"
    Else
        Image2.Picture = Image8.Picture
        Label2.FontItalic = False
        Image2.Tag = "Up"
    End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image3.Tag = "Up" Then
        Image3.Picture = Image6.Picture
        Label2.FontUnderline = True
        Image3.Tag = "Down"
    Else
        Image3.Picture = Image9.Picture
        Label2.FontUnderline = False
        Image3.Tag = "Up"
    End If
End Sub






