VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Mosquito's synSpace Map Editor"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10740
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEditmenu 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Map"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuMap 
         Caption         =   "&Map"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Editing &Tools"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "&Color Table"
      End
      Begin VB.Menu mnuPain 
         Caption         =   "Pai&n Table"
      End
      Begin VB.Menu mnuCredit 
         Caption         =   "&Credits"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display &Options"
      End
   End
   Begin VB.Menu mnuSpecial 
      Caption         =   "&Special"
      Begin VB.Menu mnuSelector 
         Caption         =   "Object &Selector"
      End
      Begin VB.Menu mnuFillarea 
         Caption         =   "&Fill Area"
      End
      Begin VB.Menu mnuDeletearea 
         Caption         =   "&Delete Area"
      End
      Begin VB.Menu mnuOptimizer 
         Caption         =   "Optimizer"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDeletearea_Click()
If OPEN_frmDeletearea = False Then
    unloadAllTools
    frmDeletearea.Show
Else
    Unload frmDeletearea
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim yesNo As Integer ' holds what the user said

' check to see if asking is needed
If varSave = True Then
    removeUndoDir
    End
    Else
    yesNo = MsgBox("Save changes before exiting?", vbYesNoCancel, "Close and Save")
    If yesNo = vbNo Then
        removeUndoDir
        End
    End If
    If yesNo = vbYes Then mnuSave_Click: Cancel = True
    If yesNo = vbCancel Then Cancel = True
End If

End Sub

Private Sub mnuColor_Click()
If OPEN_frmColor = False Then
    frmColor.Show
Else
    Unload frmColor
End If
End Sub

Private Sub mnuCopy_Click()
' this sets the clipboard
Clipboard.SetData frmMap.picMap.Image
End Sub

Private Sub mnuCredit_Click()
If OPEN_frmCredit = False Then
    frmCredit.Show
Else
    Unload frmCredit
End If
End Sub

Private Sub mnuDisplay_Click()
If OPEN_frmDisplay = False Then
    frmDisplay.Show
Else
    Unload frmDisplay
End If
End Sub

Private Sub mnuEdit_Click()
If OPEN_frmEdit = False Then
    ' get rid of conflicting tools:
    unloadAllTools
    frmEdit.Show
Else
    Unload frmEdit
End If
End Sub

Private Sub mnuExit_Click()
MDIForm_Unload 0
End Sub

Private Sub mnuFillarea_Click()
If OPEN_frmFillarea = False Then
    unloadAllTools
    frmFillarea.Show
Else
    Unload frmFillarea
End If
End Sub

Private Sub mnuMap_Click()
If OPEN_frmMap = False Then
    frmMap.Show
Else
    Unload frmMap
End If
End Sub

Private Sub mnuNew_Click()
Dim yesNo As Integer ' holds what the user said

yesNo = MsgBox("Are you sure you want to clear this map out and start fresh?", vbYesNo, "New Map")

If yesNo = vbYes Then
resetMap

' open the map and refresh everything
OPEN_frmMap = True
updateDisplay
End If

End Sub

Private Sub mnuOpen_Click()
'Open Save dialog
frmOpenSave.Show
frmOpenSave.Caption = "Open Map"
frmOpenSave.cmdOpen.Visible = True
frmOpenSave.cmdSave.Visible = False
End Sub

Private Sub mnuOptimizer_Click()
If OPEN_frmOptimizer = False Then
    frmOptimizer.Show
Else
    Unload frmOptimizer
End If
End Sub

Private Sub mnuPain_Click()
If OPEN_frmPain = False Then
    frmPain.Show
Else
    Unload frmPain
End If
End Sub

Private Sub mnuSave_Click()
If saveFileName = "" Then Exit Sub
If saveMapINI(pathFileName & "\" & saveFileName) = False Then MsgBox "Error saving map file!"
End Sub

Private Sub mnuSaveas_Click()
'Open Save dialog
frmOpenSave.Show
frmOpenSave.Caption = "Save Map"
frmOpenSave.cmdOpen.Visible = False
frmOpenSave.cmdSave.Visible = True
End Sub

Private Sub mnuSelector_Click()
If OPEN_frmSelector = False Then
    ' get rid of conflicting tools:
    unloadAllTools
    frmSelector.Show
Else
    Unload frmSelector
End If
End Sub

Private Sub mnuUndo_Click()
If useUndo() = False Then MsgBox "Could not undo!"
mnuUndo.Caption = "Undo " & undoName(undoCount)
If undoLeft = 0 Then
    mnuUndo.Caption = "Cannot Undo"
    mnuUndo.Enabled = False
End If
updateDisplay
End Sub
