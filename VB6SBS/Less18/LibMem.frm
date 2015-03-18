VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Memoria libre del sistema"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin ComctlLib.ProgressBar pgbVirtMem 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pgbPhysMem 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblVirtUsed 
      Caption         =   "Memoria virtual libre:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblPhysUsed 
      Caption         =   "Memoria física libre:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Muestra el formulario frmDetalles si el usuario pulsa
    'el botón Detalles
    Load frmDetalles
    frmDetalles.Show
End Sub

Private Sub Command2_Click()
    Unload frmDetalles    'descarga ambos formularios para salir
    Unload Form1
End Sub

Private Sub Form_Load()
    'Determina la longitud del tipo memInfo
    memInfo.dwLength = Len(memInfo)
    'Llama al API GlobalMemoryStatus para configurar la
    'barra de progreso
    Call GlobalMemoryStatus(memInfo)
    pgbPhysMem.Min = 0
    pgbPhysMem.Max = memInfo.dwTotalPhys
    pgbVirtMem.Min = 0
    pgbVirtMem.Max = memInfo.dwTotalVirtual
End Sub

Private Sub Timer1_Timer()
    Dim PhysUsed
    Dim VirtUsed
    'Llama al API GlobalMemoryStatus API para obtener
    'información sobre el empleo de la memoria
    Call GlobalMemoryStatus(memInfo)
    PhysUsed = memInfo.dwTotalPhys - memInfo.dwAvailPhys
    pgbPhysMem.Value = PhysUsed
    'Muestra el uso de la memoria mediante etiquetas y
    'una barra de progreso
    lblPhysUsed.Caption = "Empleo de memoria física: " & _
        Format(PhysUsed / memInfo.dwTotalPhys, "0.00%")
    VirtUsed = memInfo.dwTotalVirtual - memInfo.dwAvailVirtual
    pgbVirtMem.Value = VirtUsed
    lblVirtUsed.Caption = "Empleo de memoria virtual: " & _
        Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
End Sub
