VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Transferencia de archivos mediante FTP"
   ClientHeight    =   5610
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6990
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Visualización"
      Height          =   975
      Left            =   4560
      TabIndex        =   6
      Top             =   840
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "No mostrar"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mostrar en cuadro"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtLocalPath 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Text            =   "C:\Vb6Sbs\Less19\disclaimer.txt"
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtServerPath 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "disclaimer.txt"
      Top             =   1560
      Width           =   3975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   80
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transferir ahora"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtURLbox 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "ftp://ftp.microsoft.com"
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtNote 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   $"FTP.frx":0000
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarar variables para URL, ubicación del archivo en el
'servidor y ruta de destino para el archivo en el disco fijo
Dim strUrl As String           'URL es un servidor ftp
Dim strSource As String
Dim strDest As String

Private Sub Command1_Click()
'Conectar con el servidor ftp y copiar archivos en el disco fijo
strUrl = txtURLbox.Text        'obtener URL del usuario
strSource = txtServerPath.Text 'obtener la ruta del archivo fuente
strDest = txtLocalPath.Text    'obtener la ruta destino
'Usar el método Execute y la operación GET para copiar el archivo
Inet1.Execute strUrl, "GET " & strSource & " " & strDest
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
'Este suceso se activa cuando el control finaliza varias
'tareas, como conectarse y registrar los errores
Dim strAllText As String 'declara dos variables para
Dim strLine As String    'mostrar el archivo de texto
'Cuando la transferencia haya terminado o se produzca un
'error procesar apropiadamente el estado
Select Case State
Case icError   'si existe un error, describirlo
    If Inet1.ResponseCode = 80 Then 'existe el archivo
        MsgBox "¡El archivo existe! Por favor, especifique un nuevo destino"
    Else       'si el código no es 80, muestra un error desconocido
        MsgBox Inet1.ResponseInfo, , "Ha fallado la transferencia del archivo"
    End If
Case icResponseCompleted         'si ftp ha tenido éxito
    If Option1.Value = True Then 'y se ha pedido su visualización
        Open strDest For Input As #1 'abrirlo en un cuadro de texto
        Do Until EOF(1)
            Line Input #1, strLine   'leer cada línea
            strAllText = strAllText & strLine & vbCrLf
        Loop
        Close #1
        txtNote.Text = strAllText    'copiar a cuadro de texto
    Else  'si el usuario selecciona no visualizar el texto (opción por defecto)
        txtNote.Text = ""  'tan sólo resta por mostrar un mensaje de finalización
        MsgBox "Transferencia completa", , strDest
    End If
End Select
End Sub

