VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RTF Edit 2"
   ClientHeight    =   4485
   ClientLeft      =   2760
   ClientTop       =   2565
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6075
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4230
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Filename"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Font"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "10:18 AM"
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "7/16/98"
            Object.ToolTipText     =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Rich Text Format (*.RTF)|*.RTF|All Files (*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"RTFEdit2.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuCloseItem 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSaveAsItem 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuPrintItem 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCutItem 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPasteItem 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuFindItem 
         Caption         =   "&Find..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuAllcapsItem 
         Caption         =   "&All Caps"
      End
      Begin VB.Menu mnuFontItem 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuBoldItem 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuItalicItem 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuUnderlineItem 
         Caption         =   "&Underline"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare UnsavedChanges as a public Boolean (True/False)
'variable to track the current save state of the text.
'When the text is updated, the RichTextBox1_Change event
'procedure sets this variable to True.
Dim UnsavedChanges As Boolean

Private Sub Form_Load()
    'Set initial values for Slider control
    Slider1.Left = RichTextBox1.Left  'align to text box
    Slider1.Width = RichTextBox1.Width
    'note: all slider measurements in twips (form default)
    Slider1.Max = RichTextBox1.Width
    Slider1.TickFrequency = Slider1.Max * 0.1
    Slider1.LargeChange = Slider1.Max * 0.1
    Slider1.SmallChange = Slider1.Max * 0.01
End Sub

Private Sub mnuBoldItem_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub

Private Sub mnuCloseItem_Click()
    Dim Prompt As String
    Dim Reply As Integer
    'jump to error handler if the Cancel button is clicked
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    If UnsavedChanges = True Then
        Prompt = "Would you like to save your changes?"
        Reply = MsgBox(Prompt, vbYesNo)
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
            'display filename (without path) on status bar
            StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
    StatusBar1.Panels(1).Text = ""
    UnsavedChanges = False
Errhandler:
    'Cancel button clicked.
    Exit Sub
End Sub

Private Sub mnuCopyItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
End Sub

Private Sub mnuCutItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
    RichTextBox1.SelRTF = ""
End Sub

Private Sub mnuExitItem_Click()
    Dim Prompt As String
    Dim Reply As Integer
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    If UnsavedChanges = True Then
        Prompt = "Would you like to save your changes?"
        Reply = MsgBox(Prompt, vbYesNo)
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
            'display filename (without path) on status bar
            StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
            UnsavedChanges = False
        End If
    End If
    End 'after file has been saved, quit program
Errhandler:
    'Cancel button clicked (return to program)
End Sub

Private Sub mnuFindItem_Click()
    Dim SearchStr As String  'text used for search
    Dim FoundPos As Integer  'location of found text
    SearchStr = InputBox("Enter search word", "Find")
    If SearchStr <> "" Then  'if search string not empty
        'find the first occurrence of the whole word
        FoundPos = RichTextBox1.Find(SearchStr, , , _
            rtfWholeWord)
        'if the word is found (if not -1)
        If FoundPos <> -1 Then
        'use Span method to select word (forward direction)
            RichTextBox1.Span " ", True, True
        Else
            MsgBox "Search string not found", , "Find"
        End If
    End If
End Sub

Private Sub mnuFontItem_Click()
    'Force an error if the user clicks Cancel
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    'Set flags for special effects and all available fonts
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    'Display font dialog box
    CommonDialog1.ShowFont
    'Set formatting properties with user selections:
    RichTextBox1.SelFontName = CommonDialog1.FontName
    RichTextBox1.SelFontSize = CommonDialog1.FontSize
    RichTextBox1.SelColor = CommonDialog1.Color
    RichTextBox1.SelBold = CommonDialog1.FontBold
    RichTextBox1.SelItalic = CommonDialog1.FontItalic
    RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
    RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
    'Call RichTextBox1_SelChange event procedure to
    'update status bar with font name of selected text
    RichTextBox1_SelChange
Errhandler:
    'exit procedure if the user clicks Cancel
End Sub

Private Sub mnuItalicItem_Click()
    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub

Private Sub mnuAllcapsItem_Click()
    RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub mnuPrintItem_Click()
    'Prints the current document using the device
    'handle of the current printer
    RichTextBox1.SelPrint (Printer.hDC)
End Sub

Private Sub mnuUnderlineItem_Click()
    RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub

Private Sub mnuOpenItem_Click()
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName, rtfRTF
    'display filename (without path) on status bar
    StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
Errhandler:
    'if Cancel clicked, then exit procedure
End Sub

Private Sub mnuPasteItem_Click()
    RichTextBox1.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuSaveAsItem_Click()
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    CommonDialog1.ShowSave
    'save specified file in RTF format
    RichTextBox1.SaveFile CommonDialog1.FileName, rtfRTF
    'display filename (without path) on status bar
    StatusBar1.Panels(1).Text = CommonDialog1.FileTitle
    UnsavedChanges = False
Errhandler:
    'Cancel button clicked
End Sub

Private Sub RichTextBox1_Change()
    'Set public variable UnsavedChanges to True each time
    'the text in the Rich textbox is modified.
    UnsavedChanges = True
End Sub

Private Sub RichTextBox1_SelChange()
    'if there is one font in the selection, then display
    'it on the status bar (multiple fonts return Null)
    If IsNull(RichTextBox1.SelFontName) Then
        StatusBar1.Panels(2).Text = ""
    Else
        StatusBar1.Panels(2).Text = RichTextBox1.SelFontName
    End If
    'if there is one indent style in the selection, display
    'it on the slider bar (multiple styles return Null)
    If Not IsNull(RichTextBox1.SelIndent) Then
        Slider1.Value = RichTextBox1.SelIndent
    End If
End Sub

Private Sub Slider1_Scroll()
    RichTextBox1.SelIndent = Slider1.Value
End Sub
