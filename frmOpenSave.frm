VERSION 5.00
Begin VB.Form frmOpenSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open\Save Map"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmOpenSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Map"
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox cmbPattern 
      Height          =   315
      ItemData        =   "frmOpenSave.frx":08CA
      Left            =   2640
      List            =   "frmOpenSave.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Map"
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   3255
   End
   Begin VB.FileListBox fleSave 
      Height          =   4380
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DriveListBox drvSave 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin VB.DirListBox dirSave 
      Height          =   4365
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmOpenSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbPattern_Click()
If cmbPattern.ListIndex = 0 Then
fleSave.Pattern = "*.ini"
Else
fleSave.Pattern = "*"
End If
End Sub

Private Sub cmdOpen_Click()
Dim yesNo As Integer ' holds the value that indicates user input

If varSave = False Then
yesNo = MsgBox("Open new map without saving current one?", vbYesNo, "Open without saving?")
If yesNo = vbNo Then Exit Sub
End If

' If there was an error, the file does not exist already
FileDNE:

If openMapINI(dirSave.Path & "\" & txtSave.Text) = False Then MsgBox "Error opening map file!"

' map does not need saved:
varSave = True

saveFileName = txtSave.Text
pathFileName = dirSave.Path

If varUpdating = False Then updateDisplay

saveFileName = txtSave.Text
updateMainTitleBar

End Sub

Private Sub cmdSave_Click()
Dim yesNo As Integer ' holds the value that indicates user input

' tack on the .ini if the first filter is selected
If cmbPattern.ListIndex = 0 Then
    If InStr(1, txtSave.Text, ".ini") <= 0 Then
        txtSave.Text = txtSave.Text & ".ini"
    End If
End If

On Error GoTo FileDNE

' check to see of the file exists
Open dirSave.Path & "\" & txtSave.Text For Input As #1
Close #1

yesNo = MsgBox("File already exists. Overwrite?", vbYesNo, "Overwrite File?")
If yesNo = vbNo Then Exit Sub

' If there was an error, the file does not exist already
FileDNE:

If saveMapINI(dirSave.Path & "\" & txtSave.Text) = False Then MsgBox "Error saving map file!"

' map does not need saving anymore so:
varSave = True

saveFileName = txtSave.Text
pathFileName = dirSave.Path
updateMainTitleBar

Unload Me
End Sub

Private Sub dirSave_Change()
fleSave.Path = dirSave.Path
End Sub

Private Sub drvSave_Change()
On Error Resume Next
dirSave.Path = drvSave.Drive
End Sub

Private Sub fleSave_Click()
txtSave.Text = fleSave.filename
End Sub

Private Sub Form_Load()
On Error Resume Next
mdiMain.Enabled = False
dirSave.Path = arcadiaPath
cmbPattern.ListIndex = 0
txtSave.Text = saveFileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
' reactivate everything else
mdiMain.Enabled = True
End Sub
