VERSION 5.00
Begin VB.Form frmSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Selector"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5445
   Begin VB.Frame fraSelect 
      Caption         =   "Powerup"
      Height          =   1935
      Index           =   5
      Left            =   0
      TabIndex        =   67
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label lblPUP 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   74
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblPUP 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   73
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblPUP 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   72
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X position:"
         Height          =   255
         Index           =   25
         Left            =   -360
         TabIndex        =   71
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Respawn time (s):"
         Height          =   255
         Index           =   24
         Left            =   2400
         TabIndex        =   70
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y position:"
         Height          =   255
         Index           =   17
         Left            =   -360
         TabIndex        =   69
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblPUP 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "This is a super duper rocket thingy..."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   68
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Zone"
      Height          =   1935
      Index           =   4
      Left            =   0
      TabIndex        =   47
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   66
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   65
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   64
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   63
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblZone 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "No bullets"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   61
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblZone 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Team 7"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   59
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   58
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   57
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   56
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   55
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblZone 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Goal Zone"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Right/Bottom:"
         Height          =   255
         Index           =   23
         Left            =   -360
         TabIndex        =   53
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current:"
         Height          =   255
         Index           =   22
         Left            =   2160
         TabIndex        =   52
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left/Top:"
         Height          =   255
         Index           =   21
         Left            =   -360
         TabIndex        =   51
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   255
         Index           =   20
         Left            =   2760
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waypoint ID:"
         Height          =   255
         Index           =   19
         Left            =   2760
         TabIndex        =   49
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pain:"
         Height          =   255
         Index           =   18
         Left            =   2760
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Star"
      Height          =   1935
      Index           =   3
      Left            =   0
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox picStar 
         Height          =   255
         Left            =   3600
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   46
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   45
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   44
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Radius:"
         Height          =   255
         Index           =   16
         Left            =   -720
         TabIndex        =   42
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature:"
         Height          =   255
         Index           =   15
         Left            =   1920
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   40
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   39
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblStar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reach:"
         Height          =   255
         Index           =   14
         Left            =   1920
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   34
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X position:"
         Height          =   255
         Index           =   12
         Left            =   -720
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y Position:"
         Height          =   255
         Index           =   11
         Left            =   -720
         TabIndex        =   32
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mass:"
         Height          =   255
         Index           =   10
         Left            =   -720
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Vertical Barrier"
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox picvBar 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   255
         Index           =   9
         Left            =   -720
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   255
         Index           =   8
         Left            =   -720
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         Height          =   255
         Index           =   7
         Left            =   -720
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pain:"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "bouncy"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblvBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Horizontal Barrier"
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox picHBar 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "bouncy"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblHBar 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pain:"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   255
         Index           =   0
         Left            =   -720
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   255
         Index           =   1
         Left            =   -720
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         Height          =   255
         Index           =   2
         Left            =   -720
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame fraSelect 
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click somewhere on the map to select an object"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   840
         Width           =   5415
      End
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Caption         =   "Right-clicking will instantly delete an object"
      Height          =   255
      Left            =   840
      TabIndex        =   75
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' This is one of the key things because it removes stuff the user does not want
Private Sub cmdDelete_Click()
deleteSelected
showProperties
updateDisplay
varSave = False
End Sub

Private Sub Form_Load()
' set the tool
varTool = "Select"
OPEN_frmSelector = True
mdiMain.mnuSelector.Checked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
' unset the tool before unloading
varTool = ""
unselectAll
OPEN_frmSelector = False
mdiMain.mnuSelector.Checked = False
End Sub

Public Sub showProperties()

Dim i As Integer ' loop variable

' clear everything
For i = 1 To fraSelect.Count - 1
    fraSelect(i).Visible = False
Next i

' horizontal barrier
If selHorizontal > -1 Then
    fraSelect(1).Visible = True
    lblHBar(0).Caption = arrHorizontal(selHorizontal, 0)
    lblHBar(1).Caption = arrHorizontal(selHorizontal, 1)
    lblHBar(2).Caption = arrHorizontal(selHorizontal, 2)
    lblHBar(3).Caption = arrHorizontal(selHorizontal, 3)
    picHBar.BackColor = RGB(arrColor(arrHorizontal(selHorizontal, 3), 0), arrColor(arrHorizontal(selHorizontal, 3), 1), arrColor(arrHorizontal(selHorizontal, 3), 2))
    For i = 0 To xparBarrierCount
        If Val(xparBarrier(i, 0)) = arrHorizontal(selHorizontal, 4) Then lblHBar(4).Caption = xparBarrier(i, 1)
    Next i
    lblHBar(5).Caption = arrPain(arrHorizontal(selHorizontal, 5))
End If

' vertical barrier
If selVertical > -1 Then
    fraSelect(2).Visible = True
    lblvBar(0).Caption = arrVertical(selVertical, 0)
    lblvBar(1).Caption = arrVertical(selVertical, 1)
    lblvBar(2).Caption = arrVertical(selVertical, 2)
    lblvBar(3).Caption = arrVertical(selVertical, 3)
    picvBar.BackColor = RGB(arrColor(arrVertical(selVertical, 3), 0), arrColor(arrVertical(selVertical, 3), 1), arrColor(arrVertical(selVertical, 3), 2))
    For i = 0 To xparBarrierCount
        If Val(xparBarrier(i, 0)) = arrVertical(selVertical, 4) Then lblvBar(4).Caption = xparBarrier(i, 1)
    Next i
    lblvBar(5).Caption = arrPain(arrVertical(selVertical, 5))
End If

' star
If selStar > -1 Then
    fraSelect(3).Visible = True
    lblStar(0).Caption = arrStar(selStar, 1)
    lblStar(1).Caption = arrStar(selStar, 2)
    lblStar(2).Caption = arrStar(selStar, 3)
    lblStar(3).Caption = arrStar(selStar, 4)
    lblStar(4).Caption = arrStar(selStar, 5)
    lblStar(5).Caption = arrStar(selStar, 6)
    lblStar(6).Caption = arrStar(selStar, 7)
    lblStar(7).Caption = arrStar(selStar, 8)
    lblStar(8).Caption = arrStar(selStar, 9)
    picStar.BackColor = RGB(arrStar(selStar, 7), arrStar(selStar, 8), arrStar(selStar, 9))
End If

'zone
If selZone > -1 Then
    fraSelect(4).Visible = True
    If arrZone(selZone, 0) = 1 Then lblZone(0).Caption = "Normal Zone"
    If arrZone(selZone, 0) = 2 Then lblZone(0).Caption = "Goal Zone"
    If arrZone(selZone, 0) = 3 Then lblZone(0).Caption = "Waypoint Zone"
    lblZone(1).Caption = arrZone(selZone, 1)
    lblZone(2).Caption = arrZone(selZone, 2)
    lblZone(3).Caption = arrZone(selZone, 3)
    lblZone(4).Caption = arrZone(selZone, 4)
    If arrZone(selZone, 6) = 1 Then lblZone(6).Caption = "belongs to ship 1"
    If arrZone(selZone, 6) = 2 Then lblZone(6).Caption = "belongs to ship 2"
    If arrZone(selZone, 6) = 3 Then lblZone(6).Caption = "belongs to ship 3"
    If arrZone(selZone, 6) = 4 Then lblZone(6).Caption = "belongs to ship 4"
    If arrZone(selZone, 6) = 5 Then lblZone(6).Caption = "belongs to ship 5"
    If arrZone(selZone, 6) = 6 Then lblZone(6).Caption = "belongs to ship 6"
    If arrZone(selZone, 6) = 7 Then lblZone(6).Caption = "belongs to ship 7"
    If arrZone(selZone, 6) = 8 Then lblZone(6).Caption = "belongs to ship 8"
    If arrZone(selZone, 6) = 9 Then lblZone(6).Caption = "belongs to ODD ships"
    If arrZone(selZone, 6) = 10 Then lblZone(6).Caption = "belongs to EVEN ships"
    If arrZone(selZone, 6) = 11 Then lblZone(6).Caption = "belongs to all ships"
    lblZone(7).Caption = arrZone(selZone, 7)
    If arrZone(selZone, 8) = 0 Then lblZone(8).Caption = "No effect on weapons"
    If arrZone(selZone, 8) = 1 Then lblZone(8).Caption = "Ship trigger disabled, bullets live"
    If arrZone(selZone, 8) = 2 Then lblZone(8).Caption = "Ship trigger OK, bullets expire"
    If arrZone(selZone, 8) = 3 Then lblZone(8).Caption = "Trigger and bullets disabled"
    lblZone(9).Caption = arrZone(selZone, 9)
    lblZone(10).Caption = arrZone(selZone, 10)
    lblZone(11).Caption = arrZone(selZone, 11)
    lblZone(12).Caption = arrZone(selZone, 12)
End If

'PUP
If selPUP > -1 Then
    fraSelect(5).Visible = True
    For i = 0 To stylePUPCount
        If arrPUP(selPUP, 0) = stylePUP(i, 0) Then lblPUP(0).Caption = stylePUP(i, 1)
    Next i
    lblPUP(1).Caption = arrPUP(selPUP, 1)
    lblPUP(2).Caption = arrPUP(selPUP, 2)
    lblPUP(3).Caption = arrPUP(selPUP, 3)
End If
End Sub
