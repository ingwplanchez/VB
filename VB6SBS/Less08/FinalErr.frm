VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Drive Tester"
   ClientHeight    =   3915
   ClientLeft      =   2205
   ClientTop       =   1815
   ClientWidth     =   4905
   Icon            =   "FinalErr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4905
   Begin VB.CommandButton Command1 
      Caption         =   "Check Drive"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"FinalErr.frx":030A
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error GoTo DiskError
    Image1.Picture = LoadPicture("a:\prntout2.wmf")
    Exit Sub  'exit procedure
    
DiskError:
    If Err.Number = 71 Then  'if DISK NOT READY
        MsgBox ("Please close the drive latch."), , _
          "Disk Not Ready"
        Resume
    Else
        MsgBox ("I can't find prntout2.wmf in A:\."), , _
          "File Not Found"
        Resume StopTrying
    End If
StopTrying:
End Sub


