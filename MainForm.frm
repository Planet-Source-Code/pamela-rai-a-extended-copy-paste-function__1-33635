VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extended Copy/Paste"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Rich1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
      _Version        =   393217
      TextRTF         =   $"MainForm.frx":0000
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu CopyPrivate 
         Caption         =   "Copy Private"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu PastePrivate 
         Caption         =   "Paste Private"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Copy_Click()
Clipboard.SetText Rich1.SelText
End Sub

Private Sub CopyPrivate_Click()
If Rich1.SelText > "" Then
frmCopylist.List1.AddItem Rich1.SelText
End If
End Sub

Private Sub PastePrivate_Click()
frmCopylist.Left = GetX * 15
frmCopylist.Top = GetY * 15
frmCopylist.Show 1
End Sub

Private Sub Rich1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Me.PopupMenu Edit
End If
End Sub

Private Sub Paste_Click()
Rich1.SelText = Clipboard.GetText
End Sub
