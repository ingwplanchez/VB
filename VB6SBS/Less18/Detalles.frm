VERSION 5.00
Begin VB.Form frmDetalles 
   Caption         =   "Detalles de la memoria del sistema"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblAvailVirtual 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label lblTotalVirtual 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label lblAvailPage 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label lblTotalPage 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblAvailPhys 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblTotalPhys 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Ocultar el formulario cuando el usuario pulse el
    'botón Cerrar (pero continua en memoria)
    frmDetalles.Hide
End Sub

Private Sub Form_Load()
    'Utiliza el tipo memInfo para mostrar los detalles
    'del empleo de la memoria
    lblTotalPhys.Caption = "Memoria física total:(RAM): " & _
        memInfo.dwTotalPhys / 1024 & " KB"
    lblAvailPhys.Caption = "Memoria física libre:(RAM): " & _
        memInfo.dwAvailPhys / 1024 & " KB"
    lblTotalPage.Caption = "KB totales en el archivo actual de paginación: " & _
        memInfo.dwTotalPageFile / 1024
    lblAvailPage.Caption = "KB libres en el archivo actual de paginación: " & _
        memInfo.dwAvailPageFile / 1024
    lblTotalVirtual.Caption = "Memoria virtual total: " & _
        memInfo.dwTotalVirtual / 1024 & " KB"
    lblAvailVirtual.Caption = "Memoria virtual libre: " & _
        memInfo.dwAvailVirtual / 1024 & " KB"
End Sub
