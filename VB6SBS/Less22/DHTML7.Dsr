VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin {90290CCD-F27D-11D0-8031-00C04FB6C701} DHTMLPage1 
   ClientHeight    =   12345
   ClientLeft      =   1815
   ClientTop       =   1545
   ClientWidth     =   13290
   _ExtentX        =   23442
   _ExtentY        =   21775
   SourceFile      =   ""
   BuildFile       =   "C:\Vb6Sbs\Less22\DHTML7.htm"
   BuildMode       =   0
   TypeLibCookie   =   495
   AsyncLoad       =   0   'False
   id              =   "DHTMLPage1"
   ShowBorder      =   -1  'True
   ShowDetail      =   -1  'True
   AbsPos          =   -1  'True
   HTMLDocument    =   "DHTML7.dsx":0000
End
Attribute VB_Name = "DHTMLPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Function Button1_onclick() As Boolean
'Declarar la variable local x para contar las jugadas
'ganadas (copiadas a Property entre jugadas)
Dim x

'Generar tres n�meros aleatorios
Num1.innerText = Int(Rnd * 10)
Num2.innerText = Int(Rnd * 10)
Num3.innerText = Int(Rnd * 10)

'Si alguno de los tres n�meros es 7 mostrar una pila de
'monedas y reproducir aplausos
If Num1.innerText = 7 Or Num2.innerText = 7 Or _
    Num3.innerText = 7 Then
    'Si la jugada es ganadora, reproducir el archivo.wav
    '(applause.wav)
    MMControl1.Command = "Prev"  'rebobinar si es necesario
    MMControl1.Command = "Play"  'reproducir el archivo .wav
    'e incrementar el contador de Property
    x = GetProperty("Wins")
    Result.innerText = "Wins: " & x + 1
    PutProperty "Wins", x + 1
End If
End Function

Private Sub DHTMLPage_Load()
'Aplicar subrayado para la cabecera
LuckyHead.Style.textDecorationUnderline = True
'Asignar el color azul a los n�meros
Num.Style.Color = "blue"

'inicializar el generador de n�meros aleatorios
'para obtener aut�nticos n�meros aleatorios
Randomize
'Mostrar pila de monedas
Image1.src = "c:\vb6sbs\less22\coins.wmf"

'Configurar y abrir el control Multimedia MCI
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"
MMControl1.FileName = "c:\vb6sbs\less22\applause.wav"
MMControl1.Command = "Open"

'Utilizar la funci�n GetProperty para determinar si se ha
'ganado anteriormente y se ha almacenado el hecho en
'Property (un lugar de almacenamiento que persiste
'durante las operaciones de carga y de descarga de la
'p�gina HTML). Con este c�digo podr� almacenar el n�mero
'de jugadas ganadas aunque se acceda a la p�gina "Sobre 7
'Afortunado" o a cualquier otra p�gina Web.
Result.innerText = "Wins: " & GetProperty("Wins")

End Sub

