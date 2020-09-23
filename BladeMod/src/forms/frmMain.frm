VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "BladeMod"
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9270
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTabs 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   9240
      TabIndex        =   3
      Top             =   360
      Width           =   9270
      Begin MSComctlLib.TabStrip tabMain 
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   8700
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   4290
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11668
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   794
            MinWidth        =   794
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   714
            MinWidth        =   707
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   13
            Style           =   2
         EndProperty
      EndProperty
      Begin VB.CommandButton cmdFont 
         Caption         =   "FONT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4980
         TabIndex        =   5
         Top             =   30
         Width           =   615
      End
      Begin VB.ComboBox cmbSize 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   5640
         List            =   "frmMain.frx":0049
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   8640
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00A5
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01B7
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02C9
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03DB
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04ED
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05FF
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0711
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0823
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0935
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A47
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B59
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C6B
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D7D
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
         Begin VB.Menu mnuFile_NewDoc 
            Caption         =   "&Document"
         End
      End
      Begin VB.Menu mnuFile_Open 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFile_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFile_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_PageSetup 
         Caption         =   "Page &Setup ..."
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFile_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit_Redo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "C&opy"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuFormat_Font 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView_Tip 
         Caption         =   "&Tip Of The Day"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Status 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindow_THorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindow_TVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindow_Arrange 
         Caption         =   "Arrange &Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_TipOfTheDay 
         Caption         =   "Show &Tip Of The Day"
      End
      Begin VB.Menu mnuHelp_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_VoteBladeMod 
         Caption         =   "&Vote for BladeMod (PlanetSourceCode)"
      End
      Begin VB.Menu mnuHelp_DLBM 
         Caption         =   "&Download BladeMod"
      End
      Begin VB.Menu mnuHelp_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_Vote 
         Caption         =   "Vote for Pass&Gen 2.5 (PlanetSourceCode)"
      End
      Begin VB.Menu mnuHelp_DLPG25 
         Caption         =   "Download &PassGen 2.5"
      End
      Begin VB.Menu mnuAbout_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_Site 
         Caption         =   "Blade Software Web &Site"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSize_Click()
On Error Resume Next
        '
        ' If there are no document windows open
        ' Show a message in the status bar to remind the user
        If tabMain.Tabs.Count = 0 Then
               'Status.Panels(1).Text = "WARNING: No active documents."
               Exit Sub
        End If
        '
        ' Set text size and set the focus
        ActiveForm.rtfText.SelFontSize = cmbSize.Text
        ActiveForm.rtfText.SetFocus
        ' THIS IS ONLY A GUI BEAUTIFIER
        Call LooseFocus
End Sub

Private Sub cmdFont_Click()
        On Error Resume Next
        Call mnuFormat_Font_Click
        ' THIS IS ONLY A GUI BEAUTIFIER
        Call LooseFocus
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
tmp = FreeFile
Open App.Path & "\" & App.EXEName & ".INI" For Binary As FreeFile
        If LOF(tmp) > 0 Then
                Call LoadSettings
        Else
                Me.Show
                frmTip.Show 0, frmMain
        End If
Close

' ERROR 91: Startup document 'Edit1' didn't load properly
' Occurs when you try to access the RTF control when it's not present
' Set Program Title & build number
Me.Caption = BuildTitle
'
' Remove the first tab present
tabMain.Tabs.Remove 1
'
' Add a blank document
'Call NewDocument
'
' 8 point font as default
cmbSize.ListIndex = 2
'
' set text to 8 points
'ActiveForm.rtfText.Font.Size = cmbSize.Text

' Display tips if required
If INIFile.TipOfTheDay = True Then
        Me.Show
        frmTip.Show 0, frmMain
End If

If Err <> 0 Then
        MsgBox "Error loading " & BuildTitle & vbCrLf & _
                "Error: " & Err.Number & " - " & Err.Description & ")"
        Err.Clear
        End
End If
'
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
' If files are dropped onto the main form...proc them and open
' documents for each one.
For i = 1 To Data.Files.Count
        Call NewDocument(Data.Files(i))
        ActiveForm.rtfText.LoadFile Data.Files(i)
Next i
End Sub

Private Sub MDIForm_Resize()
        On Error Resume Next
        ' Set the width of the tab sheet to the size of the form
        tabMain.Width = frmMain.Width - 280
End Sub

Private Sub MDIForm_Terminate()
'
' End program 'gracefully'
Call EndProgram
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'
' End program 'gracefully'
Call EndProgram
End Sub

Private Sub mnuEdit_Copy_Click()
        On Error Resume Next
        
        ' Copy selected text to the cliboard
        Clipboard.SetText ActiveForm.rtfText.SelRTF
End Sub

Private Sub mnuEdit_Cut_Click()
        On Error Resume Next
        ' Cut the selected text to the clipboard
        Clipboard.SetText ActiveForm.rtfText.SelRTF
        ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub mnuEdit_Paste_Click()
        On Error Resume Next
        ActiveForm.rtfText.SelRTF = Clipboard.GetText
End Sub


Private Sub mnuFile_Exit_Click()
' End program using common end call
Call EndProgram
End Sub

Private Sub mnuFile_NewDoc_Click()
' open new document
Call NewDocument
End Sub

Private Sub mnuFile_Open_Click()

        ' Open a file
        Dim sFile As String
        With dlgMain
                .DialogTitle = "Open"
                .CancelError = False
                'ToDo: set the flags and attributes of the common dialog control
                .Filter = "All Files (*.*)|*.*"
                .ShowOpen
                If Len(.fileName) = 0 Then
                        Exit Sub
                End If
                sFile = .fileName
                Call NewDocument(sFile)
        End With
End Sub

Private Sub mnuFile_Print_Click()
        On Error Resume Next
        ' Don't do anything if a document isn't open
        If ActiveForm Is Nothing Then Exit Sub
        '
        ' Open a file
        With dlgMain
                .DialogTitle = "Print"
                .CancelError = True
                .Flags = cdlPDReturnDC + cdlPDNoPageNums
                If ActiveForm.rtfText.SelLength = 0 Then
                        .Flags = .Flags + cdlPDAllPages
                Else
                        .Flags = .Flags + cdlPDSelection
                End If
                .ShowPrinter
                If Err <> MSComDlg.cdlCancel Then
                        ActiveForm.rtfText.SelPrint .hDC
                End If
        End With
End Sub

Private Sub mnuFile_PageSetup_Click()
        On Error Resume Next
        '
        ' Show printer dialog
        With dlgMain
                .DialogTitle = "Page Setup"
                .CancelError = True
                .ShowPrinter
        End With
End Sub

Private Sub mnuFile_Save_Click()
        On Error Resume Next
        ' Don't do anything if a document isn't open
        If ActiveForm Is Nothing Then Exit Sub
        Dim sFile As String
                '
                ' Save a file
                ' COMPLETE SO FILES THAT WERE LOADED ARE AUTOMATICALLY SAVED
                ' TO THE SOURCE FILE
                With dlgMain
                        .DialogTitle = "Save"
                        .CancelError = False
                        .fileName = ActiveForm.Caption
                        'ToDo: set the flags and attributes of the common dialog control
                        .Filter = "All Files (*.*)|*.*"
                        .ShowSave
                        If Len(.fileName) = 0 Then
                                Exit Sub
                        End If
                        sFile = .fileName
                End With
                ActiveForm.rtfText.SaveFile sFile
                sFile = ActiveForm.Caption
                ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFile_SaveAs_Click()
        Dim sFile As String
        ' Don't do anything if a document isn't open
        If ActiveForm Is Nothing Then Exit Sub
        '
        ' Save file as
        With dlgMain
                .DialogTitle = "Save As"
                .CancelError = False
                'ToDo: set the flags and attributes of the common dialog control
                .Filter = "All Files (*.*)|*.*"
                .ShowSave
                If Len(.fileName) = 0 Then
                        Exit Sub
                End If
                sFile = .fileName
        End With
        ActiveForm.Caption = sFile
        ActiveForm.rtfText.SaveFile sFile
End Sub

Private Sub mnuFormat_Font_Click()
On Error Resume Next
        '
        ' If there are no document windows open
        ' Show a message in the status bar to remind the user
        If tabMain.Tabs.Count = 0 Then
               'Status.Panels(1).Text = "WARNING: No active documents."
                Exit Sub
        End If
        '
        ' Loose the focus on the control that was clicked
        ' THIS IS ONLY A GUI BEAUTIFIER
        Call LooseFocus
        '
        ' Set flags to show all properties...striketru, underline, colour, etc...
        dlgMain.Flags = &H317   ' Flags property must be set
        '
        ' Grab the curently selected text properties
        dlgMain.FontName = ActiveForm.rtfText.SelFontName
        dlgMain.FontSize = ActiveForm.rtfText.SelFontSize
        dlgMain.FontStrikethru = ActiveForm.rtfText.SelStrikeThru
        dlgMain.FontUnderline = ActiveForm.rtfText.SelUnderline
        dlgMain.Color = ActiveForm.rtfText.SelColor
        dlgMain.FontBold = ActiveForm.rtfText.SelBold
        dlgMain.FontItalic = ActiveForm.rtfText.SelItalic
        '
        ' Display Font common dialog box.
        dlgMain.ShowFont
        '
        ' Set the selected text properties to match the selected properties
        ActiveForm.rtfText.SelFontName = dlgMain.FontName
        ActiveForm.rtfText.SelFontSize = dlgMain.FontSize
        ActiveForm.rtfText.SelStrikeThru = dlgMain.FontStrikethru
        ActiveForm.rtfText.SelUnderline = dlgMain.FontUnderline
        ActiveForm.rtfText.SelBold = dlgMain.FontBold
        ActiveForm.rtfText.SelItalic = dlgMain.FontItalic
        ActiveForm.rtfText.SelColor = dlgMain.Color
        '
        ' Loose the focus on the control that was clicked
        ' THIS IS ONLY A GUI BEAUTIFIER
        Call LooseFocus
End Sub

Private Sub mnuHelp_About_Click()
'
' Display program name and version (with build number)
MsgBox "Thanks to System33 for testing" & vbCrLf & _
        App.ProductName & " " & App.FileDescription & " (Build:" & BuildNumber & ")" & vbCrLf & _
        "by John Pettit", vbOKOnly, "About BladeMod"
End Sub

Private Sub mnuHelp_DLBM_Click()
Shell "explorer ""http://lantis.anu.edu.au/blade/projects/blademod/blademod.zip""", vbNormalFocus
End Sub

Private Sub mnuHelp_DLPG25_Click()
Shell "explorer ""http://lantis.anu.edu.au/blade/projects/passgen/pg25.zip""", vbNormalFocus
End Sub

Private Sub mnuHelp_Site_Click()
Shell "explorer ""http://webone.com.au/~jpettit""", vbNormalFocus
End Sub

Private Sub mnuHelp_TipOfTheDay_Click()
                frmTip.Show 0, frmMain
End Sub

Private Sub mnuHelp_Vote_Click()
Shell "explorer ""http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=7980""", vbNormalFocus
End Sub

Private Sub mnuHelp_VoteBladeMod_Click()
Shell "explorer ""http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=9568""", vbNormalFocus
End Sub

Private Sub mnuView_Status_Click()
'
' Toggle status bar's visibility state & save changes to menu
mnuView_Status.Checked = Not (mnuView_Status.Checked)
        If mnuView_Status.Checked = True Then
                Status.Visible = True
        Else
                Status.Visible = False
        End If
End Sub

Private Sub mnuView_Tip_Click()
        If mnuView_Tip.Checked = True Then
                mnuView_Tip.Checked = False
                INIFile.TipOfTheDay = False
        Else
                mnuView_Tip.Checked = True
                INIFile.TipOfTheDay = True
        End If
End Sub

Private Sub mnuWindow_Arrange_Click()
        ' Arrange child windows
        Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindow_Cascade_Click()
        ' Cascade windows
        Me.Arrange vbCascade
End Sub

Private Sub mnuWindow_THorizontal_Click()
        ' tile windows horizontally
        Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindow_TVertical_Click()
        ' tile windows vertically
        Me.Arrange vbTileVertical
End Sub


Private Sub tabMain_Click()
        On Error Resume Next
        '
        ' set focus to the selected tab
        Document(tabMain.SelectedItem.Index).SetFocus
        ActiveDocument = frmMain.tabMain.SelectedItem.Index
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        On Error Resume Next
        
        ' Handle toolbar clicks
        Select Case Button.Key
                Case "New"
                        Call NewDocument
                        
                Case "Open"
                        mnuFile_Open_Click
                Case "Save"
                        mnuFile_Save_Click
                Case "Print"
                        mnuFile_Print_Click
                Case "Cut"
                        mnuEdit_Cut_Click
                Case "Copy"
                        mnuEdit_Copy_Click
                Case "Paste"
                        mnuEdit_Paste_Click
                Case "Bold"
                        ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
                        Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
                Case "Italic"
                        ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
                        Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
                Case "Underline"
                        ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
                        Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
                Case "Align Left"
                        ActiveForm.rtfText.SelAlignment = rtfLeft
                Case "Center"
                        ActiveForm.rtfText.SelAlignment = rtfCenter
                Case "Align Right"
                        ActiveForm.rtfText.SelAlignment = rtfRight
        End Select
End Sub

Private Sub tmrStats_Timer()
        Status.Panels(1).Text = Len(ActiveForm.rtfText.Text)
End Sub
