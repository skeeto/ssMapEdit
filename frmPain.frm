VERSION 5.00
Begin VB.Form frmPain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pain Table"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2160
   Icon            =   "frmPain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   2160
   Begin VB.TextBox txtPain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblPain 
      Caption         =   "Entry 0:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmPain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
OPEN_frmPain = True
mdiMain.mnuPain.Checked = True

Dim i As Integer ' crappy loop variable

' the initial hieght of the window is 925
' and each additional value is 360 high
Me.Height = 925 + (360 * PAIN_MAX)

' load each thingy up and place it
For i = 1 To PAIN_MAX
Load txtPain(i)
txtPain(i).Visible = True
txtPain(i).top = txtPain(0).top + (360 * i)
txtPain(i).left = txtPain(0).left

Load lblPain(i)
lblPain(i).top = lblPain(i - 1).top + 360
lblPain(i).left = lblPain(0).left
lblPain(i).Caption = "Entry " & i & ":"
lblPain(i).Visible = True
Next i

txtPain(0).Height = txtPain(1).Height

' set values
updatePain

End Sub

Public Sub updatePain()
For i = 0 To txtPain.Count - 1
    txtPain(i).Text = arrPain(i)
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
OPEN_frmPain = False
mdiMain.mnuPain.Checked = False
End Sub

Private Sub txtPain_Change(Index As Integer)
txtPain(Index).Text = Val(txtPain(Index).Text)
arrPain(Index) = txtPain(Index).Text
If varUpdating = False Then updateDisplay
End Sub
