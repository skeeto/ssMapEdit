VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Table"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3240
   Begin VB.PictureBox picColor 
      Height          =   360
      Index           =   0
      Left            =   2760
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtColorR 
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
      Left            =   840
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtColorG 
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
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtColorB 
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
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblColor 
      Caption         =   "Entry 0:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()

mdiMain.mnuColor.Checked = True
OPEN_frmColor = True

Dim i As Integer ' crappy loop variable

' the initial hieght of the window is 925
' and each additional value is 360 high
Me.Height = 925 + (360 * COLOR_MAX)

' load each thingy up, place it, and set its value
For i = 1 To COLOR_MAX

Load picColor(i)
picColor(i).top = picColor(i - 1).top + 360
picColor(i).left = picColor(0).left
picColor(i).Visible = True

Load txtColorR(i)
txtColorR(i).top = txtColorR(i - 1).top + 360
txtColorR(i).left = txtColorR(0).left
txtColorR(i).Visible = True

Load txtColorG(i)
txtColorG(i).top = txtColorG(i - 1).top + 360
txtColorG(i).left = txtColorG(0).left
txtColorG(i).Visible = True

Load txtColorB(i)
txtColorB(i).top = txtColorB(i - 1).top + 360
txtColorB(i).left = txtColorB(0).left
txtColorB(i).Visible = True

Load lblColor(i)
lblColor(i).top = lblColor(i - 1).top + 360
lblColor(i).left = lblColor(0).left
lblColor(i).Caption = "Entry " & i & ":"
lblColor(i).Visible = True
Next i

txtColorR(0).Height = txtColorR(1).Height
txtColorG(0).Height = txtColorG(1).Height
txtColorB(0).Height = txtColorB(1).Height

updateColors

End Sub

Public Sub updateColors()
Dim i As Integer ' loop variable

' load the values in which will also bring up coloration
For i = 0 To picColor.Count - 1
    txtColorR(i).Text = arrColor(i, 0)
    txtColorG(i).Text = arrColor(i, 1)
    txtColorB(i).Text = arrColor(i, 2)
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.mnuColor.Checked = False
OPEN_frmColor = False
End Sub

Private Sub txtColorG_Change(Index As Integer)
On Error Resume Next

' set the array thingy when the textbox changes
txtColorG(Index).Text = Val(txtColorG(Index).Text)
If txtColorG(Index).Text < 0 Then txtColorG(Index).Text = 0
If txtColorG(Index).Text > 255 Then txtColorG(Index).Text = 255
arrColor(Index, 1) = txtColorG(Index).Text

' recolor the sample box
picColor(Index).BackColor = RGB(txtColorR(Index).Text, txtColorG(Index).Text, txtColorB(Index).Text)

' update everything
If varUpdating = False Then updateDisplay

End Sub

Private Sub txtColorB_Change(Index As Integer)
On Error Resume Next

' set the array thingy when the textbox changes
txtColorB(Index).Text = Val(txtColorB(Index).Text)
If txtColorB(Index).Text < 0 Then txtColorB(Index).Text = 0
If txtColorB(Index).Text > 255 Then txtColorB(Index).Text = 255
arrColor(Index, 2) = txtColorB(Index).Text

' recolor the sample box
picColor(Index).BackColor = RGB(txtColorR(Index).Text, txtColorG(Index).Text, txtColorB(Index).Text)

' update everything
If varUpdating = False Then updateDisplay

End Sub

Private Sub txtColorR_Change(Index As Integer)
On Error Resume Next

' set the array thingy when the textbox changes
txtColorR(Index).Text = Val(txtColorR(Index).Text)
If txtColorR(Index).Text < 0 Then txtColorR(Index).Text = 0
If txtColorR(Index).Text > 255 Then txtColorR(Index).Text = 255
arrColor(Index, 0) = txtColorR(Index).Text

' recolor the sample box
picColor(Index).BackColor = RGB(txtColorR(Index).Text, txtColorG(Index).Text, txtColorB(Index).Text)

' update everything
If varUpdating = False Then updateDisplay

End Sub
