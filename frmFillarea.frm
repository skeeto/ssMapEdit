VERSION 5.00
Begin VB.Form frmFillarea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill Area"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmFillarea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6105
   Begin VB.TextBox txtFillFreq 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "32"
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox picBarColor 
      Height          =   315
      Left            =   4680
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbPain 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cmbColor 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbStyle 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkWhat 
      Caption         =   "Outside Border"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkWhat 
      Caption         =   "Vertical Barriers"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkWhat 
      Caption         =   "Horizontal Barriers"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "units per barrier (256 units = 1 sector)"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Caption         =   "Fill frequency:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblBarPain 
      Alignment       =   2  'Center
      Caption         =   "-100"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Caption         =   "Pain Index:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Caption         =   "Color Index:"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Caption         =   "Style/Behavior:"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Just click and drag the area in which you want to fill with barriers."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmFillarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkWhat_Click(Index As Integer)
Select Case Index

    Case 0
        If chkWhat(0).Value = 1 Then
            toolFillHorizontal = True
        Else
            toolFillHorizontal = False
        End If
    Case 1
        If chkWhat(1).Value = 1 Then
            toolFillVertical = True
        Else
            toolFillVertical = False
        End If
    Case 2
        If chkWhat(2).Value = 1 Then
            toolFillBorder = True
        Else
            toolFillBorder = False
        End If
        
End Select
End Sub

Private Sub cmbColor_Click()
toolColorIndex = Val(cmbColor.Text)
picBarColor.BackColor = RGB(arrColor(Val(cmbColor.Text), 0), arrColor(Val(cmbColor.Text), 1), arrColor(Val(cmbColor.Text), 2))
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbPain_Click()
toolPainIndex = Val(cmbPain.Text)
lblBarPain.Caption = arrPain(Val(cmbPain.Text))
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbStyle_Click()
toolXpar = cmbStyle.ItemData(cmbStyle.ListIndex)
End Sub

Private Sub Form_Load()

' initialize
varTool = "Fill Area"
mdiMain.mnuFillarea.Checked = True
OPEN_frmFillarea = True

toolFillHorizontal = True
toolFillVertical = True
toolFillBorder = True

txtFillFreq_Change

' set up everything
For i = 0 To COLOR_MAX
    cmbColor.AddItem i, i
Next i
cmbColor.ListIndex = 0

For i = 0 To COLOR_MAX
    cmbPain.AddItem i, i
Next i
cmbPain.ListIndex = 0

cmbStyle.Clear
For i = 0 To xparBarrierCount
    cmbStyle.AddItem xparBarrier(i, 1), i
    cmbStyle.ItemData(i) = Val(xparBarrier(i, 0))
Next i
cmbStyle.ListIndex = 0

End Sub

Public Sub updateFillarea()
cmbPain_Click
cmbColor_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.mnuFillarea.Checked = False
OPEN_frmFillarea = False

varTool = ""
End Sub

Private Sub txtFillFreq_Change()
txtFillFreq.Text = Int(Val(txtFillFreq.Text))
If txtFillFreq.Text < 1 Then txtFillFreq.Text = 1
toolFillFrequency = txtFillFreq.Text
End Sub
