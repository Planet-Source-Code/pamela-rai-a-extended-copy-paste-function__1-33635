VERSION 5.00
Begin VB.Form frmCopylist 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bnClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton bnClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmCopylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bnClear_Click()
List1.Clear
End Sub

Private Sub bnClose_Click()
frmCopylist.Hide
End Sub

Private Sub List1_Click()
MainForm.Rich1.SelText = List1.List(List1.ListIndex)
frmCopylist.Hide
End Sub
