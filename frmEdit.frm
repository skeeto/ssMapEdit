VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing Tools"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Powerups"
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Spawns"
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Zones"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Stars"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Barriers"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Barriers"
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4815
      Begin VB.ComboBox cmbPain 
         Height          =   315
         ItemData        =   "frmEdit.frx":08CA
         Left            =   1680
         List            =   "frmEdit.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.PictureBox picBarColor 
         Height          =   315
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   1275
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmEdit.frx":08CE
         Left            =   1680
         List            =   "frmEdit.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmEdit.frx":08D2
         Left            =   1680
         List            =   "frmEdit.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   3015
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
         Left            =   3120
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblBar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pain Index:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblBar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color Index:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblBar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Style / Behavior:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblBar 
         BackStyle       =   0  'Transparent
         Caption         =   "The next barrier will be drawn with the following properties:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Stars"
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrCheck 
         Interval        =   1000
         Left            =   4560
         Top             =   -120
      End
      Begin VB.PictureBox picStarColor 
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   35
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtRGB 
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
         Left            =   1200
         TabIndex        =   32
         Text            =   "200"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtRadius 
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
         Left            =   1200
         TabIndex        =   29
         Text            =   "128"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtTemp 
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
         Left            =   1200
         TabIndex        =   24
         Text            =   "30"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtReach 
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
         Left            =   1200
         TabIndex        =   22
         Text            =   "2048"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtMass 
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
         Left            =   1200
         TabIndex        =   20
         Text            =   "1000"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtRGB 
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
         Index           =   1
         Left            =   1800
         TabIndex        =   33
         Text            =   "200"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtRGB 
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
         Index           =   2
         Left            =   2400
         TabIndex        =   34
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblStarCount 
         Alignment       =   2  'Center
         Caption         =   "16"
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stars remaining:"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         Caption         =   "the standard is 128"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   30
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Radius:"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         Caption         =   "the standard size is 1000"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblTip 
         Caption         =   "a negative value heals"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   26
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         Caption         =   "(hint: one sector is 256 units wide)"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   25
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reach:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblStar 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mass:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Zones"
      Height          =   2655
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtCurY 
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
         Left            =   1920
         TabIndex        =   61
         Text            =   "0"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtCurX 
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
         Left            =   1440
         TabIndex        =   60
         Text            =   "0"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtZonePain 
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
         Left            =   3600
         TabIndex        =   59
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtWPID 
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
         Left            =   3600
         TabIndex        =   58
         Text            =   "0"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtFriction 
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
         Left            =   1440
         TabIndex        =   57
         Text            =   "0"
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox cmbWeapon 
         Height          =   315
         ItemData        =   "frmEdit.frx":08D6
         Left            =   1800
         List            =   "frmEdit.frx":08E6
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbTeam 
         Height          =   315
         ItemData        =   "frmEdit.frx":0928
         Left            =   1800
         List            =   "frmEdit.frx":094D
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbZoneStyle 
         Height          =   315
         ItemData        =   "frmEdit.frx":09B1
         Left            =   1800
         List            =   "frmEdit.frx":09BE
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current (X, Y):"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   56
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waypoint ID:"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   55
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pain:"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   54
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   53
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Weapon Effect:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zone ownership:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zone type:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Spawns"
      Height          =   2655
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdRand 
         Caption         =   "Randomize"
         Height          =   495
         Left            =   3120
         TabIndex        =   46
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtHeading 
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
         Height          =   345
         Left            =   1440
         TabIndex        =   42
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   40
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   39
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSpawnPos 
         Alignment       =   2  'Center
         Caption         =   "0, 0"
         Height          =   255
         Left            =   1440
         TabIndex        =   45
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblSpawnLabel 
         Caption         =   "Spawn position:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "degrees from galactic North"
         Height          =   255
         Left            =   2520
         TabIndex        =   43
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblSpawnLabel 
         Caption         =   "Spawn heading:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSpawn 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   960
         TabIndex        =   38
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Powerups"
      Height          =   2655
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox chkRandomPUP 
         Caption         =   "Disable the 20 random powerups for this map"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtRespawn 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   65
         Text            =   "1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmbPUPStyle 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblRespawn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Respawn time:"
         Height          =   255
         Left            =   840
         TabIndex        =   64
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblPUP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type/Style:"
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkRandomPUP_Click()
If chkRandomPUP.Value = 1 Then
    disableAllRandomPups = True
Else
    disableAllRandomPups = False
End If

varSave = False
updateDisplay
End Sub

Private Sub cmbColor_Click()
toolColorIndex = Val(cmbColor.Text)
picBarColor.BackColor = RGB(arrColor(Val(cmbColor.Text), 0), arrColor(Val(cmbColor.Text), 1), arrColor(Val(cmbColor.Text), 2))
toolColorIndex = Val(cmbColor.Text)
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbPain_Click()
toolPainIndex = Val(cmbPain.Text)
lblBarPain.Caption = arrPain(Val(cmbPain.Text))
toolPainIndex = Val(cmbPain.Text)
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbPUPStyle_Click()
toolPUPStyle = cmbPUPStyle.ItemData(cmbPUPStyle.ListIndex)
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbStyle_Click()
toolXpar = cmbStyle.ItemData(cmbStyle.ListIndex)
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbTeam_Click()
toolTeam = cmbTeam.ListIndex + 1
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbWeapon_Click()
toolBullets = cmbWeapon.ListIndex
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmbZoneStyle_Click()
toolZoneStyle = cmbZoneStyle.ListIndex + 1
If varUpdating = False Then updateDisplay
End Sub

Private Sub cmdMenu_Click(Index As Integer)
Dim i As Integer ' loop variable

' clear everything
For i = 0 To fraMenu.Count - 1
fraMenu(i).Visible = False
Next i

'bring up the desired menu
fraMenu(Index).Visible = True

' set the tool variable
Select Case Index

    Case 0
        varTool = "Barrier"
    
    Case 1
        varTool = "Star"
    
    Case 2
        varTool = "Zone"
    
    Case 3
        varTool = "Spawn"

    Case 4
        varTool = "PUP"
        
End Select

End Sub

Private Sub cmdMove_Click(Index As Integer)
' set the caption
If Index = 1 Then lblSpawn.Caption = lblSpawn.Caption + 1
If Index = 0 Then lblSpawn.Caption = lblSpawn.Caption - 1

toolCurrentSpawn = Val(lblSpawn.Caption) - 1

' check for validity
If lblSpawn.Caption < 1 Then lblSpawn.Caption = 1
If lblSpawn.Caption > SPAWN_MAX + 1 Then lblSpawn.Caption = SPAWN_MAX + 1
End Sub

Private Sub cmdRand_Click()
arrSpawn(lblSpawn.Caption - 1, 0) = -1
arrSpawn(lblSpawn.Caption - 1, 1) = -1
updateDisplay
End Sub

Public Sub updateEdit()

Dim i As Integer ' loop variable

' barrier
cmbColor_Click
cmbPain_Click

' random powerups
If disableAllRandomPups = True Then
    chkRandomPUP.Value = 1
Else
    chkRandomPUP.Value = 0
End If

End Sub

Public Sub Form_Load()
varTool = "Barrier"

mdiMain.mnuEdit.Checked = True
OPEN_frmEdit = True

' fix everything up
updateEdit

' set all the property boxes:
' barrier
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

' star
txtRGB_Change 0
txtMass_Change
txtReach_Change
txtTemp_Change
txtRadius_Change

' spawn
toolCurrentSpawn = 0
toolHeading = 0

' zone
cmbZoneStyle.ListIndex = 0
cmbTeam.ListIndex = 0
cmbWeapon.ListIndex = 0
cmbZoneStyle_Click
cmbTeam_Click
cmbWeapon_Click
txtFriction_Change
txtWPID_Change
txtZonePain_Change
txtCurX_Change
txtCurY_Change

' PUP
txtRespawn_Change
cmbPUPStyle.Clear
For i = 0 To stylePUPCount
    cmbPUPStyle.AddItem stylePUP(i, 1), i
    cmbPUPStyle.ItemData(i) = Val(stylePUP(i, 0))
Next i
cmbPUPStyle.ListIndex = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
varTool = ""
OPEN_frmEdit = False
mdiMain.mnuEdit.Checked = False
End Sub

Private Sub tmrCheck_Timer()

Dim i As Integer ' loop variable

lblStarCount.Caption = 0

For i = 0 To STAR_MAX
    If arrStar(i, 0) = 0 Then lblStarCount.Caption = lblStarCount.Caption + 1
Next i

End Sub

Private Sub txtCurX_Change()
txtCurX.Text = Val(txtCurX.Text)
toolCurrentX = txtCurX.Text
updateDisplay
End Sub

Private Sub txtCurY_Change()
txtCurY.Text = Val(txtCurY.Text)
toolCurrentY = txtCurY.Text
updateDisplay
End Sub

Private Sub txtFriction_Change()
txtFriction.Text = Val(txtFriction.Text)
toolFriction = txtFriction.Text
updateDisplay
End Sub

Private Sub txtHeading_Change()
txtHeading.Text = Val(txtHeading.Text)
If txtHeading.Text < 0 Then txtHeading.Text = 0
If txtHeading.Text > 359 Then txtHeading.Text = 359

toolHeading = txtHeading.Text
End Sub

Private Sub txtMass_Change()
toolMass = Val(txtMass.Text)
updateDisplay
End Sub

Private Sub txtRadius_Change()
toolRadius = Val(txtRadius.Text)
updateDisplay
End Sub

Private Sub txtReach_Change()
toolReach = Val(txtReach.Text)
updateDisplay
End Sub

Private Sub txtRespawn_Change()
txtRespawn.Text = Val(txtRespawn.Text)
toolRespawn = txtRespawn.Text
updateDisplay
End Sub

Private Sub txtRGB_Change(Index As Integer)
Dim i As Integer ' loop variable

' make sure values are alright
For i = 0 To 2
txtRGB(i).Text = Val(txtRGB(i))
If txtRGB(i).Text < 0 Then txtRGB(i) = 0
If txtRGB(i).Text > 255 Then txtRGB(i) = 255
Next i

' set the tool variables
toolColorR = txtRGB(0).Text
toolColorG = txtRGB(1).Text
toolColorB = txtRGB(2).Text

'  recolor the sample
picStarColor.BackColor = RGB(txtRGB(0).Text, txtRGB(1).Text, txtRGB(2).Text)

updateDisplay
End Sub

Private Sub txtTemp_Change()
toolTemperature = txtTemp.Text
updateDisplay
End Sub

Private Sub txtWPID_Change()
txtWPID.Text = Int(Val(txtWPID.Text))
toolWaypointID = txtWPID.Text
updateDisplay
End Sub

Private Sub txtZonePain_Change()
txtZonePain.Text = Val(txtZonePain.Text)
toolPain = txtZonePain.Text
updateDisplay
End Sub
