VERSION 5.00
Begin VB.Form frmDeletearea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Area"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmDeletearea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2970
   Begin VB.CheckBox chkType 
      Caption         =   "Powerups"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Zones"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Stars"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Vertical Barriers"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Horizontal Barriers"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmDeletearea.frx":08CA
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDeletearea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkType_Click(Index As Integer)
Select Case Index
    
    Case 0
        If chkType(Index).Value = 1 Then
            toolDeleteHorizontal = True
        Else
            toolDeleteHorizontal = False
        End If

    Case 1
        If chkType(Index).Value = 1 Then
            toolDeleteVertical = True
        Else
            toolDeleteVertical = False
        End If

    Case 2
        If chkType(Index).Value = 1 Then
            toolDeleteStar = True
        Else
            toolDeleteStar = False
        End If
        
    Case 3
        If chkType(Index).Value = 1 Then
            toolDeleteZone = True
        Else
            toolDeleteZone = False
        End If

    Case 4
        If chkType(Index).Value = 1 Then
            toolDeletePUP = True
        Else
            toolDeletePUP = False
        End If

End Select
End Sub

Private Sub Form_Load()
OPEN_frmDeletearea = True
mdiMain.mnuDeletearea.Checked = True

toolDeleteHorizontal = True
toolDeleteVertical = True
toolDeleteStar = True
toolDeleteZone = True
toolDeletePUP = True

varTool = "Delete Area"

End Sub

Private Sub Form_Unload(Cancel As Integer)
OPEN_frmDeletearea = False
mdiMain.mnuDeletearea.Checked = False

varTool = ""
End Sub
