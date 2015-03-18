VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.OLE OLE2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Paint.Picture"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   4080
      OleObjectBlob   =   "OleProy.frx":0000
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   855
      Left            =   2160
      OleObjectBlob   =   "OleProy.frx":AA18
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OLE OLE4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Word.Document.8"
      DisplayType     =   1  'Icon
      Height          =   855
      Left            =   360
      OleObjectBlob   =   "OleProy.frx":CE30
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Planos de la obra"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Cálculo de costes"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Memoria de calidades"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Estimación de proyectos utilizando Word, Excel y Paint"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Proyectos Urbanísticos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

