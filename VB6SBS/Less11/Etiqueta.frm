VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Trabajando con colecciones"
   ClientHeight    =   3825
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdButton 
      Caption         =   "Mover Objetos"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Tag             =   "Bot�n"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image imgBanana 
      Height          =   480
      Left            =   360
      Picture         =   "Etiqueta.frx":0000
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image imgStopWatch 
      Height          =   480
      Left            =   360
      Picture         =   "Etiqueta.frx":030A
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgEye 
      Height          =   480
      Left            =   360
      Picture         =   "Etiqueta.frx":0614
      Top             =   840
      Width           =   480
   End
   Begin VB.Image imgEar 
      Height          =   480
      Left            =   360
      Picture         =   "Etiqueta.frx":091E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub cmdButton_Click()
        For Each Ctrl In Controls
            If Ctrl.Tag <> "Bot�n" Then
                Ctrl.Left = Ctrl.Left + 200
            End If
        Next Ctrl
    End Sub

