VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Automatización de Microsoft Outlook"
   ClientHeight    =   3405
   ClientLeft      =   2115
   ClientTop       =   2070
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   4935
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "EnviarCorreo.frx":0000
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar mensaje de prueba"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   $"EnviarCorreo.frx":0024
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Este procedimiento utiliza Automatización para introducir
'un mensaje de prueba en la bandeja de salida de
'Microsoft Outlook (si está conectado y Outlook está abierto,
'Outlook enviará también el mensaje a su servicio de correo
'electrónico). El programa Outlook es necesario y podrá comprobar
'que la operación de envío es más rápida y eficiente si Outlook
'ha sido ejecutado previamente.

Dim out As Object           'crea una variable objeto
'asigna Outlook.Application a la variable objeto
Set out = CreateObject("Outlook.Application")

With out.CreateItem(olMailItem) 'empleo del objeto Outlook
    'inserte nuevos destinatarios, uno por vez, utilizando el método Add
    '(estos nombres son ficticios--introduzca los suyos propios)
    .Recipients.Add "maria@xxx.com"  'Para: campo
    .Recipients.Add "casey@xxx.com"  'Para: campo
    'para introducir usuarios en el campo CC:, especifique el tipo olCC
    .Recipients.Add("mike_halvorson@classic.msn.com").Type = olCC
    .Subject = "Mensaje de prueba"  'incluye el contenido del campo Asunto:
    .Body = Text1.Text  'copia el mensaje del cuadro de texto
    'inserta anexos, uno por vez, utilizando el método Add
    .Attachments.Add "c:\vb6sbs\less14\smile.bmp"
    'finalmente, copia el mensaje a la bandeja de salida de Outlook con Send
    .Send
End With

End Sub
