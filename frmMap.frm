VERSION 5.00
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frmMap.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7770
   ScaleWidth      =   8055
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   120
      MouseIcon       =   "frmMap.frx":0BD4
      MousePointer    =   99  'Custom
      ScaleHeight     =   7575
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Image imgMouse 
         Height          =   240
         Left            =   2760
         Picture         =   "frmMap.frx":0EDE
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label lblSpawn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape shpStar 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim secWidth As Single         ' holds the sector width in pixels for
Dim unitWidth As Single        ' the current display size
Dim mapWidth As Integer        ' holds picMap's width value
Dim mouseX As Integer          ' last mousedown X
Dim mouseY As Integer          ' last mousedown Y

Private Sub Form_Load()

OPEN_frmMap = True
mdiMain.mnuMap.Checked = True

Dim i As Integer ' dumb loop variable

' load all the shape objects to act as stars
For i = 1 To STAR_MAX
    Load shpStar(i)
Next i

' load all the label objects to act as stars
For i = 1 To SPAWN_MAX
    Load lblSpawn(i)
Next i

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseDown Button, Shift, X - picMap.left, Y - picMap.top
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseMove Button, Shift, X - picMap.left, Y - picMap.top
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseUp Button, Shift, X - picMap.left, Y - picMap.top
End Sub

Private Sub Form_Resize()
' this keeps the map at size with the window

If Me.Width > (Me.Height - 285) Then
Me.Height = Me.Width + 285
Else
Me.Width = Me.Height - 285
End If

picMap.Width = Me.Width - picMap.left - 250
picMap.Height = Me.Height - 655

' find out the size of the window and get the sector and unit size
mapWidth = picMap.Width
secWidth = mapWidth / 32
unitWidth = mapWidth / 8192

drawMap
End Sub

' This is a very important function. It is the key to the entire program.
' It goes step by step to draw each element of the map data ignoring the
' items checked off to not be displayed.
'
' This is pretty much the entire display engine.
Public Sub drawMap()

Dim i As Integer, j As Integer ' throwaway loop variables
Dim zColor As Single           ' keeps track of zone coloring

' begin by clearing the map off
picMap.Cls
For i = 0 To SPAWN_MAX
    lblSpawn(i).Visible = False
Next i
For i = 0 To STAR_MAX
    shpStar(i).Visible = False
Next i

' draw the zones first so that they lie under everything
If DISP_showZone = 1 Then
If zoneCount > -1 Then
    For i = 0 To zoneCount
        If arrZone(i, 0) > 0 Then
            zColor = DISP_zoneColor(arrZone(i, 0) - 1)
            If arrZone(i, 6) = 9 Then zColor = RGB(100, 100, 100)   ' ODD zone
            If arrZone(i, 6) = 10 Then zColor = RGB(200, 200, 200)  ' EVEN zone
            picMap.Line (arrZone(i, 1) * unitWidth, mapWidth - (arrZone(i, 2) * unitWidth))-(arrZone(i, 3) * unitWidth, mapWidth - (arrZone(i, 4) * unitWidth)), zColor, BF
        End If
    Next i
End If
End If

' Draw the grid
If DISP_showGrid = 1 Then
For i = 0 To 32
picMap.Line (i * secWidth, 0)-(i * secWidth, picMap.Height), DISP_gridColor
picMap.Line (0, i * secWidth)-(picMap.Width, i * secWidth), DISP_gridColor
Next i
End If

' Draw the barriers
If DISP_showBarriers = 1 Then
' horizontal
If horizontalCount > -1 Then
    For i = 0 To horizontalCount
        picMap.Line (arrHorizontal(i, 0) * unitWidth, mapWidth - (arrHorizontal(i, 2) * unitWidth))-(arrHorizontal(i, 1) * unitWidth, mapWidth - (arrHorizontal(i, 2) * unitWidth)), RGB(arrColor(arrHorizontal(i, 3), 0), arrColor(arrHorizontal(i, 3), 1), arrColor(arrHorizontal(i, 3), 2))
    Next i
End If
' vertical
If verticalCount > -1 Then
    For i = 0 To verticalCount
        picMap.Line (arrVertical(i, 2) * unitWidth, mapWidth - (arrVertical(i, 0) * unitWidth))-(arrVertical(i, 2) * unitWidth, mapWidth - (arrVertical(i, 1) * unitWidth)), RGB(arrColor(arrVertical(i, 3), 0), arrColor(arrVertical(i, 3), 1), arrColor(arrVertical(i, 3), 2))
    Next i
End If
End If

' Draw the stars
If DISP_showStar = 1 Then
For i = 0 To STAR_MAX
    If arrStar(i, 0) = 1 Then
        shpStar(i).Width = arrStar(i, 6) * unitWidth / 2 * DISP_starMult
        shpStar(i).Height = shpStar(i).Width
        shpStar(i).left = (arrStar(i, 1) * unitWidth) - (shpStar(i).Width / 2)
        shpStar(i).top = mapWidth - (arrStar(i, 2) * unitWidth) - (shpStar(i).Width / 2)
        shpStar(i).BackColor = RGB(arrStar(i, 7), arrStar(i, 8), arrStar(i, 9))
        shpStar(i).Visible = True
    Else
        shpStar(i).Visible = False
    End If
Next i
End If

' Draw the spawn points
If DISP_showSpawn = 1 Then
If spawnCount > -1 Then
    For i = 0 To spawnCount
        If arrSpawn(i, 0) > -1 Then
            lblSpawn(i).Visible = True
            lblSpawn(i).Caption = i + 1
            lblSpawn(i).left = (arrSpawn(i, 0) * unitWidth) - (lblSpawn(i).Width / 2)
            lblSpawn(i).top = mapWidth - (arrSpawn(i, 1) * unitWidth) - (lblSpawn(i).Height / 2)
        End If
    Next i
End If
End If

' draw PUPs
If DISP_showPUP = 1 Then
If pupCount > -1 Then
    For i = 0 To pupCount
        If arrPUP(i, 0) > 0 Then
            picMap.Line ((arrPUP(i, 1) + 15) * unitWidth, mapWidth - ((arrPUP(i, 2) + 15) * unitWidth))-((arrPUP(i, 1) - 15) * unitWidth, mapWidth - ((arrPUP(i, 2) - 15) * unitWidth)), RGB(255, 204, 0), BF

        End If
    Next i
End If
End If

' And finally, we draw appropriate marks around the "selected object"
If selSpawn >= 0 Then
    If arrSpawn(selSpawn, 0) >= 0 Then
        picMap.Circle (arrSpawn(selSpawn, 0) * unitWidth, mapWidth - (arrSpawn(selSpawn, 1) * unitWidth)), secWidth, RGB(255, 255, 255)
    End If
End If

If selZone >= 0 Then
    If arrZone(selZone, 0) > 0 Then
        picMap.Line (arrZone(selZone, 1) * unitWidth, mapWidth - (arrZone(selZone, 2) * unitWidth))-(arrZone(selZone, 3) * unitWidth, mapWidth - (arrZone(selZone, 4) * unitWidth)), RGB(255, 255, 255), B
    End If
End If

If selStar >= 0 Then
    If arrStar(selStar, 0) > 0 Then
        picMap.Circle (arrStar(selStar, 1) * unitWidth, mapWidth - (arrStar(selStar, 2) * unitWidth)), secWidth, RGB(255, 255, 255)
        picMap.Circle (arrStar(selStar, 1) * unitWidth, mapWidth - (arrStar(selStar, 2) * unitWidth)), unitWidth * arrStar(selStar, 4), RGB(255, 255, 255)
    End If
End If

If selHorizontal >= 0 Then
    picMap.Circle (arrHorizontal(selHorizontal, 0) * unitWidth, mapWidth - (arrHorizontal(selHorizontal, 2) * unitWidth)), secWidth, RGB(255, 255, 255)
    picMap.Circle (arrHorizontal(selHorizontal, 1) * unitWidth, mapWidth - (arrHorizontal(selHorizontal, 2) * unitWidth)), secWidth, RGB(255, 255, 255)
End If

If selVertical >= 0 Then
    picMap.Circle (arrVertical(selVertical, 2) * unitWidth, mapWidth - (arrVertical(selVertical, 0) * unitWidth)), secWidth, RGB(255, 255, 255)
    picMap.Circle (arrVertical(selVertical, 2) * unitWidth, mapWidth - (arrVertical(selVertical, 1) * unitWidth)), secWidth, RGB(255, 255, 255)
End If

If selPUP >= 0 Then
    If arrPUP(selPUP, 0) > 0 Then
        picMap.Circle (arrPUP(selPUP, 1) * unitWidth, mapWidth - (arrPUP(selPUP, 2) * unitWidth)), secWidth, RGB(255, 255, 255)
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.mnuMap.Checked = False
OPEN_frmMap = False
End Sub

Private Sub imgMouse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseDown Button, Shift, X + imgMouse.left, Y + imgMouse.top
End Sub

Private Sub imgMouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseMove Button, Shift, X + imgMouse.left, Y + imgMouse.top
End Sub

Private Sub imgMouse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseUp Button, Shift, X + imgMouse.left, Y + imgMouse.top
End Sub

Private Sub lblSpawn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseDown Button, Shift, X + lblSpawn(Index).left, Y + lblSpawn(Index).top
End Sub

Private Sub lblSpawn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseMove Button, Shift, X + lblSpawn(Index).left, Y + lblSpawn(Index).top
End Sub

Private Sub lblSpawn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseUp Button, Shift, X + lblSpawn(Index).left, Y + lblSpawn(Index).top
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    mouseX = X
    mouseY = Y
End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Int(X / secWidth / DISP_gridSnap) * DISP_gridSnap
Y = 32 - Int(Y / secWidth / DISP_gridSnap) * DISP_gridSnap

Me.Caption = "Map sec(" & X & ", " & Y & ") unit(" & (X * 256) & ", " & (Y * 256) & ")"

imgMouse.left = (X * secWidth) - (0.5 * imgMouse.Width)
imgMouse.top = mapWidth - (Y * secWidth) - (0.5 * imgMouse.Width)

End Sub

' this initiates the editing tools which makes this sub very very important
Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim x0 As Single ' These hold the converted mouse values
Dim x1 As Single ' so that they can be easily used for the
Dim y0 As Single ' tools ahead of time.
Dim y1 As Single '

' convert to units on the map:
x0 = Int(mouseX / secWidth / DISP_gridSnap) * DISP_gridSnap * 256
x1 = Int(X / secWidth / DISP_gridSnap) * DISP_gridSnap * 256
y0 = Int((mapWidth - mouseY) / secWidth / DISP_gridSnap) * DISP_gridSnap * 256 + (256 * DISP_gridSnap)
y1 = Int((mapWidth - Y) / secWidth / DISP_gridSnap) * DISP_gridSnap * 256 + (256 * DISP_gridSnap)

' figure out what tool is being used
' if you add you own tool, you would call it here and pass what you need to it
Select Case varTool

    Case "Barrier"
        setUndo varTool
        If Abs(x0 - x1) > Abs(y0 - y1) Then
            ' horizontal barrier
            createHorizontal x0, x1, y0, toolColorIndex, toolXpar, toolPainIndex
        Else
            ' vertical barrier
            createVertical y0, y1, x0, toolColorIndex, toolXpar, toolPainIndex
        End If
    
    Case "Star"
        setUndo varTool
        createStar x1, y1, toolMass, toolReach, toolTemperature, toolRadius, toolColorR, toolColorG, toolColorB
        
    Case "Spawn"
        setUndo varTool
        arrSpawn(toolCurrentSpawn, 0) = x1
        arrSpawn(toolCurrentSpawn, 1) = y1
        arrSpawn(toolCurrentSpawn, 2) = toolHeading
    
    Case "Zone"
        setUndo varTool
        createZone toolZoneStyle, x0, y0, x1, y1, 0, toolTeam, toolPain, toolBullets, toolWaypointID, toolFriction, toolCurrentX, toolCurrentY
    
    Case "PUP"
        setUndo "Powerup"
        createPUP toolPUPStyle, x1, y1, toolRespawn
    
    Case "Select"
        setSelect x1, y1
        If Button = vbRightButton Then deleteSelected
        frmSelector.showProperties
    
    Case "Fill Area"
        setUndo " Fill Area"
        FillArea x0, y0, x1, y1, toolFillFrequency, toolColorIndex, toolXpar, toolPainIndex
    
    Case "Delete Area"
        setUndo "Delete Area"
        DeleteArea x0, y0, x1, y1
        
End Select

varSave = False

drawMap
End Sub
