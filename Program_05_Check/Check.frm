VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdMostrarLibro 
      Caption         =   "&Mostrar Libro"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Libros de Programaciòn"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox ChkLibro4 
         Caption         =   "Cobol 2.0, Autor: Borland."
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CheckBox ChkLibro3 
         Caption         =   "Delphi 6.0, Tomo I. Autor: Borland."
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox ChkLibro2 
         Caption         =   "Pascal Estructurado 7.0, Tomo II.Autor: Borland."
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox ChkLibro1 
         Caption         =   "Visual Basic, Tomo I Autor: Microsoft."
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMostrarLibro_Click()

    If ChkLibro1.Value = 1 Then
        MsgBox (ChkLibro1.Caption)  ' Muestra el Nombre del libro
    End If

    If ChkLibro2.Value = 1 Then
        MsgBox (ChkLibro2.Caption)  ' Muestra el Nombre del libro
    End If

    If ChkLibro3.Value = 1 Then     ' Muestra el Nombre del libro
        MsgBox (ChkLibro3.Caption)
    End If

    If ChkLibro4.Value = 1 Then     ' Muestra el Nombre del libro
        MsgBox (ChkLibro4.Caption)
    End If
    
    If ChkLibro1.Value = 0 And ChkLibro2.Value = 0 And ChkLibro3.Value = 0 And _
    ChkLibro4.Value = 0 Then
        MsgBox ("Seleccione un libro de la lista.")
End If
    
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
