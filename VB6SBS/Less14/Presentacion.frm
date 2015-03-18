VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Automatización de Microsoft PowerPoint"
   ClientHeight    =   3225
   ClientLeft      =   3495
   ClientTop       =   2415
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Pulsar para ver Presentación"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   $"Presentacion.frx":0000
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "PowerPoint y Visual Basic"
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

Dim ppt As Object        'declaración de variable objeto
Dim respuesta, mensaje   'declaración de variables para msgbox

mensaje = "Pulse la barra espaciadora para moverse al " & _
    "siguiente mensaje de la presentación." & vbCrLf & "¿Listo para comenzar?"
respuesta = MsgBox(mensaje, vbYesNo, "Características sorprendentes de PowerPoint")

If respuesta = vbYes Then
    Set ppt = CreateObject("PowerPoint.Application.8")
    ppt.Visible = True      'abrir y ejecutar la presentación
    ppt.Presentations.Open "c:\vb6sbs\less14\pptfacts.ppt"
    ppt.ActivePresentation.SlideShowSettings.Run
    Set ppt = Nothing       'liberar la variable objeto
End If

End Sub
