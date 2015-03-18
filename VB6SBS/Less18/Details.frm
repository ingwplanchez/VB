VERSION 5.00
Begin VB.Form frmDetails 
   Caption         =   "System Memory Details"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblAvailVirtual 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblTotalVirtual 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label lblAvailPage 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblTotalPage 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblAvailPhys 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label lblTotalPhys 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Hide form when Close button clicked (but keep in mem)
    frmDetails.Hide
End Sub

Private Sub Form_Load()
    'Use memInfo type to display memory usage details
    lblTotalPhys.Caption = "Total physical memory (RAM): " & _
        memInfo.dwTotalPhys / 1024 & " KB"
    lblAvailPhys.Caption = "Free physical memory (RAM): " & _
        memInfo.dwAvailPhys / 1024 & " KB"
    lblTotalPage.Caption = "Total KB in current paging file: " & _
        memInfo.dwTotalPageFile / 1024
    lblAvailPage.Caption = "Free KB in current paging file: " & _
        memInfo.dwAvailPageFile / 1024
    lblTotalVirtual.Caption = "Total virtual memory: " & _
        memInfo.dwTotalVirtual / 1024 & " KB"
    lblAvailVirtual.Caption = "Free virtual memory: " & _
        memInfo.dwAvailVirtual / 1024 & " KB"
End Sub
