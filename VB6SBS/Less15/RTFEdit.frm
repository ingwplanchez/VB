VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "RTF Edit"
   ClientHeight    =   4275
   ClientLeft      =   2775
   ClientTop       =   2580
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6075
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3600
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
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"RTFEdit.frx":0000
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
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
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
    UnsavedChanges = False
Errhandler:
    'Cancel button clicked
End Sub

Private Sub RichTextBox1_Change()
    'Set public variable UnsavedChanges to True each time
    'the text in the Rich textbox is modified.
    UnsavedChanges = True
End Sub
