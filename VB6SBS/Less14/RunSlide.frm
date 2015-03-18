VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Microsoft PowerPoint Automation"
   ClientHeight    =   3225
   ClientLeft      =   3495
   ClientTop       =   2415
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Click to view Presentation"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   $"RunSlide.frx":0000
      Height          =   1455
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "PowerPoint and Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim ppt As Object       'dim object variable
Dim reply, prompt       'dim variables for msgbox

prompt = "Press spacebar to move from slide to slide" & _
    " in the presentation." & vbCrLf & "Ready to start?"
reply = MsgBox(prompt, vbYesNo, "Amazing PowerPoint Facts")

If reply = vbYes Then
    Set ppt = CreateObject("PowerPoint.Application.8")
    ppt.Visible = True      'open and run presentation
    ppt.Presentations.Open "c:\vb6sbs\less14\pptfacts.ppt"
    ppt.ActivePresentation.SlideShowSettings.Run
    Set ppt = Nothing       'release object variable
End If

End Sub
