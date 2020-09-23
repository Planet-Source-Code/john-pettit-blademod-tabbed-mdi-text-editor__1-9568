VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "doc"
   ClientHeight    =   2250
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   4725
   Visible         =   0   'False
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1935
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDocument.frx":0000
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
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Activate()
        On Error Resume Next
        If Terminating = False Then
                If Recovering = False Then
                        ActiveDocument = frmMain.tabMain.SelectedItem.Index
                        frmMain.tabMain.Tabs(Me.Caption).Selected = True
                Else
                        Recovering = False
                End If
        Else
                Recovering = True
                Terminating = False
        End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        On Error Resume Next
        Me.SetFocus
        frmMain.tabMain.Tabs.Remove Me.Caption
        If frmMain.tabMain.Tabs.Count = 0 Then lDocumentCount = 0
        Terminating = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
'If Me.Caption = "doc" Then Exit Sub
        If Me.Width > 1000 And Me.Height > 1000 Then
                rtfText.Width = Me.Width - 85
                rtfText.Height = Me.Height - 385
        End If
End Sub

Private Sub rtfText_Change()
        On Error Resume Next
        frmMain.Status.Panels(2).Text = Len(Me.rtfText.Text) & "b"
End Sub
