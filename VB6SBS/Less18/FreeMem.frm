VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Free System Memory"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Details"
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
      Caption         =   "Free Virtual Memory: "
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblPhysUsed 
      Caption         =   "Free Physical Memory: "
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
    'Display frmDetails form if user clicks Details button
    Load frmDetails
    frmDetails.Show
End Sub

Private Sub Command2_Click()
    Unload frmDetails       'unload both forms to quit
    Unload Form1
End Sub

Private Sub Form_Load()
    'Determine length of memInfo type
    memInfo.dwLength = Len(memInfo)
    'Call GlobalMemoryStatus API to setup progress bars
    Call GlobalMemoryStatus(memInfo)
    pgbPhysMem.Min = 0
    pgbPhysMem.Max = memInfo.dwTotalPhys
    pgbVirtMem.Min = 0
    pgbVirtMem.Max = memInfo.dwTotalVirtual
End Sub

Private Sub Timer1_Timer()
    Dim PhysUsed
    Dim VirtUsed
    'Call GlobalMemoryStatus API to get memory usage info
    Call GlobalMemoryStatus(memInfo)
    PhysUsed = memInfo.dwTotalPhys - memInfo.dwAvailPhys
    pgbPhysMem.Value = PhysUsed
    'Display memory usage with labels and progress bars
    lblPhysUsed.Caption = "Physical Memory Usage: " & _
        Format(PhysUsed / memInfo.dwTotalPhys, "0.00%")
    VirtUsed = memInfo.dwTotalVirtual - memInfo.dwAvailVirtual
    pgbVirtMem.Value = VirtUsed
    lblVirtUsed.Caption = "Virtual Memory Usage: " & _
        Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
End Sub
