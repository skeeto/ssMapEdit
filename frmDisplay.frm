VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Options"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4695
   Begin VB.CheckBox chkZone 
      Caption         =   "Show Zones"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtStarMult 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Text            =   "10"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtGridSnap 
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
      Left            =   3720
      TabIndex        =   9
      Text            =   "1"
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picGrid 
      Height          =   255
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox chkSpawn 
      Caption         =   "Show Spawn Positions"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CheckBox chkStar 
      Caption         =   "Show Stars"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkPUP 
      Caption         =   "Show Powerups"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox chkBarriers 
      Caption         =   "Show Barriers"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Show Sector Grid"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblStarFactor 
      Alignment       =   1  'Right Justify
      Caption         =   "Star Size Multiplier:"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblOne 
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblGridSnap 
      Alignment       =   1  'Right Justify
      Caption         =   "Grid Snap Ratio:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblGridColor 
      Alignment       =   1  'Right Justify
      Caption         =   "Sector Grid Color:"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBarriers_Click()
DISP_showBarriers = chkBarriers.Value
frmMap.drawMap
End Sub

Private Sub chkGrid_Click()
DISP_showGrid = chkGrid.Value
frmMap.drawMap
End Sub

Private Sub chkPUP_Click()
DISP_showPUP = chkPUP.Value
frmMap.drawMap
End Sub

Private Sub chkSpawn_Click()
DISP_showSpawn = chkSpawn.Value
frmMap.drawMap
End Sub

Private Sub chkStar_Click()
DISP_showStar = chkStar.Value
frmMap.drawMap
End Sub

Private Sub chkZone_Click()
DISP_showZone = chkZone.Value
frmMap.drawMap
End Sub

Private Sub Form_Load()
OPEN_frmDisplay = True
mdiMain.mnuDisplay.Checked = True

' this intiates the window with the current settings
chkGrid.Value = DISP_showGrid
chkStar.Value = DISP_showStar
chkSpawn.Value = DISP_showSpawn
chkPUP.Value = DISP_showPUP
chkBarriers.Value = DISP_showBarriers
chkZone.Value = DISP_showZone

picGrid.BackColor = DISP_gridColor
txtGridSnap.Text = 1 / DISP_gridSnap
txtStarMult.Text = DISP_starMult
End Sub

Private Sub Form_Unload(Cancel As Integer)
OPEN_frmDisplay = False
mdiMain.mnuDisplay.Checked = False
End Sub

Private Sub picGrid_Click()
Dim R As Integer '
Dim G As Integer ' color values
Dim B As Integer '

' ask the user for values:
' This is a sloppy way to do it, so you can make up your
' own way to do it if you want. Whatever :-P
R = Val(InputBox("Red value? (0 - 255)", "Red", 0))
G = Val(InputBox("Green value? (0 - 255)", "Green", 127))
B = Val(InputBox("Blue value? (0 - 255)", "Blue", 0))

' check to make sure things are alright
If R > 255 Then R = 255
If R < 0 Then R = 0
If G > 255 Then G = 255
If G < 0 Then G = 0
If B > 255 Then B = 255
If B < 0 Then B = 0


' set everything

DISP_gridColor = RGB(R, G, B)

picGrid.BackColor = DISP_gridColor

frmMap.drawMap

End Sub

Private Sub txtGridSnap_Change()
txtGridSnap.Text = Val(txtGridSnap.Text)
If txtGridSnap.Text = 0 Then txtGridSnap.Text = 1
DISP_gridSnap = 1 / txtGridSnap.Text
End Sub

Private Sub txtStarMult_Change()
DISP_starMult = Val(txtStarMult)
frmMap.drawMap
End Sub
