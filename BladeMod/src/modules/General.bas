Attribute VB_Name = "General"
Public Document(256) As frmDocument     ' Array of document forms
Public ActiveDocument As Integer        ' Holds various info to differentiate documents
Public lDocumentCount As Long           ' Count of documents

' Variables to remove problem of frmDocument appearing as a child
' since when closing a form it then gets activated twice before
' closing.
Public Terminating As Boolean
Public Recovering As Boolean
Public INIFile As SavedSettings

Type SavedSettings
        ' standard window variables
        Top As Integer
        Left As Integer
        Width As Integer
        Height As Integer
        WindowState As Integer
        TipOfTheDay As Boolean
        
        ' default font settings
        Font As String
        Size As Integer
        Bold As Boolean
        Italic As Boolean
        StrikeTrough As Boolean
        Underline As Boolean
        
        ' document related
        OpenDocs(32) As String
        LastClipBoardText As String
        EnableUndo As Boolean
        
        
End Type

Public Function BuildNumber()
        ' Calculates the build number from App.Minor and App.Revision
        tempminor = App.Minor
        tempRevision = App.Revision
        BuildNumber = App.Minor & String(3 - Len(tempRevision), "0") & App.Revision
End Function


Public Sub NewDocument(Optional fileName As String)
        If fileName = "" Then
                lDocumentCount = lDocumentCount + 1
                Set Document(lDocumentCount) = New frmDocument
                Document(lDocumentCount).Caption = "Edit " & lDocumentCount
                frmMain.tabMain.Tabs.Add frmMain.tabMain.Tabs.Count + 1, Document(lDocumentCount).Caption, Document(lDocumentCount).Caption
                'frmMain.tabMain.Tabs(frmMain.tabMain.Tabs.Count).Selected = True
                frmMain.tabMain.Tabs(frmMain.tabMain.Tabs.Count).Tag = lDocumentCount
                ActiveDocument = frmMain.tabMain.Tabs.Count
                Document(lDocumentCount).Tag = lDocumentCount
                'Document(lDocumentCount).Width = frmMain.Width * 0.75
                If Document(lDocumentCount).Caption <> "doc" Then _
                        Document(lDocumentCount).Show
        Else
                On Error Resume Next
                For i = 1 To lDocumentCount
                        If Document(i).Caption = fileName Then Exit Sub
                Next i
                lDocumentCount = lDocumentCount + 1
                Set Document(lDocumentCount) = New frmDocument
                Document(lDocumentCount).Caption = fileName
                frmMain.tabMain.Tabs.Add frmMain.tabMain.Tabs.Count + 1, Document(lDocumentCount).Caption, Document(lDocumentCount).Caption
                frmMain.tabMain.Tabs(frmMain.tabMain.Tabs.Count).Selected = True
                frmMain.tabMain.Tabs(frmMain.tabMain.Tabs.Count).Tag = lDocumentCount
                ActiveDocument = Document(lDocumentCount).Caption
                frmMain.ActiveForm.rtfText.LoadFile fileName
                Document(lDocumentCount).Tag = lDocumentCount
                Document(lDocumentCount).Show
        End If
        
End Sub

Public Sub LooseFocus()
        On Error Resume Next
        ActiveForm.Status.SetFocus
End Sub

Public Function BuildTitle()
        BuildTitle = App.ProductName & " - " & App.FileDescription & " (Build: " & BuildNumber & ")"
End Function

Public Sub EndProgram()
        '
        Call SaveSettings
        Unload frmDocument      ' Remove frmDocument from memory
        Unload frmMain          ' Remove frmMain from memory
        Close                   ' Close all open files and I/O devices
        End                     ' Terminate program execution
End Sub

Private Sub SaveSettings()
        On Error Resume Next
        '
        With frmMain
                INIFile.Top = .Top
                INIFile.Left = .Left
                INIFile.Width = .Width
                INIFile.Height = .Height
                INIFile.WindowState = .WindowState
                INIFile.TipOfTheDay = .mnuView_Tip.Checked
        End With
        ' Open INI file
        tmp = FreeFile
        Kill App.Path & "\" & App.EXEName & ".INI"
        Open App.Path & "\" & App.EXEName & ".INI" For Binary As FreeFile
                Put tmp, 1, INIFile
        Close
End Sub

Public Sub LoadSettings()
        On Error Resume Next
        '
        ' Open INI file
        tmp = FreeFile
        Open App.Path & "\" & App.EXEName & ".INI" For Binary As FreeFile
                Get tmp, 1, INIFile
        Close

With frmMain
        .Top = INIFile.Top
        .Left = INIFile.Left
        .Width = INIFile.Width
        .Height = INIFile.Height
        .WindowState = INIFile.WindowState
        .mnuView_Tip.Checked = INIFile.TipOfTheDay
        If INIFile.TipOfTheDay = True Then
                frmMain.mnuView_Tip.Checked = True
        Else
                frmMain.mnuView_Tip.Checked = False
        End If
End With
End Sub
