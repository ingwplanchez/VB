VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Microsoft Outlook Automation"
   ClientHeight    =   3270
   ClientLeft      =   2115
   ClientTop       =   2070
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "SendMail.frx":0000
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send test message"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"SendMail.frx":0019
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'This procedure uses Automation to place a test message in
'the Microsoft Outlook outbox. (If you are online and
'Outlook is open, Outlook will also send the message to
'your email service.) The Outlook program is required, and
'you'll find that the send operation is quicker and more
'memory efficient if Outlook is already running.

Dim out As Object           'create object variable
'assign Outlook.Application to object variable
Set out = CreateObject("Outlook.Application")

With out.CreateItem(olMailItem) 'using the Outlook object
    'insert recipients one at a time with the Add method
    '(these names are fictitious--replace with your own)
    .Recipients.Add "maria@xxx.com"  'To: field
    .Recipients.Add "casey@xxx.com"  'To: field
    'to place users in the CC: field, specify olCC type
    .Recipients.Add("mike_halvorson@classic.msn.com").Type = olCC
    .Subject = "Test Message"  'include a subject field
    .Body = Text1.Text  'copy message text from text box
    'insert attachments one at a time with the Add method
    .Attachments.Add "c:\vb6sbs\less14\smile.bmp"
    'finally, copy message to Outlook outbox with Send
    .Send
End With

End Sub
