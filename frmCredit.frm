VERSION 5.00
Begin VB.Form frmCredit 
   Caption         =   "Credits"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   5685
   Begin VB.TextBox txtCredit 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
mdiMain.mnuCredit.Checked = True
OPEN_frmCredit = True
End Sub

Public Sub updateCredit()
txtCredit.Text = strCredits
End Sub

Private Sub Form_Resize()
' this keeps the text at size with the window

' forget it if the goddamned thing is minimzied
If Me.WindowState = vbMinimized Then Exit Sub
If mdiMain.WindowState = vbMinimized Then Exit Sub

txtCredit.Width = Me.Width - txtCredit.left - 300
txtCredit.Height = Me.Height - 705
End Sub

Private Sub Form_Unload(Cancel As Integer)
OPEN_frmCredit = False
mdiMain.mnuCredit.Checked = False
End Sub

Private Sub txtCredit_Change()
strCredits = txtCredit.Text
End Sub
