VERSION 5.00
Begin VB.Form frmOptimizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizer"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmOptimizer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5865
   Begin VB.ListBox lstFix 
      Height          =   1425
      ItemData        =   "frmOptimizer.frx":08CA
      Left            =   120
      List            =   "frmOptimizer.frx":08CC
      TabIndex        =   6
      Top             =   4200
      Width           =   5655
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Optimize"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   600
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   960
      Width           =   4695
      Begin VB.Shape shpWhat 
         FillColor       =   &H0000FF00&
         Height          =   255
         Index           =   4
         Left            =   240
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblWhat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zero area zones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblWhat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zero length lines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   2025
      End
      Begin VB.Shape shpWhat 
         FillColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   240
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape shpWhat 
         FillColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   240
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblWhat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barriers set end-to-end"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Top             =   2160
         Width           =   2805
      End
      Begin VB.Shape shpWhat 
         FillColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   240
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblWhat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incorrectly overlapping zones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   840
         TabIndex        =   3
         Top             =   1680
         Width           =   3525
      End
      Begin VB.Shape shpWhat 
         FillColor       =   &H0000FF00&
         Height          =   255
         Index           =   0
         Left            =   240
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblWhat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overlapping parallel barriers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   840
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmOptimizer.frx":08CE
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblOpt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This tool searches the map for errors and incorrect usage."
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5625
   End
End
Attribute VB_Name = "frmOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBegin_Click()

Dim i As Integer ' loop variable
Dim j As Integer ' loop variable
Dim k As Integer ' loop variable

Dim Xa0 As Single ' These all abbreviate the real names to make this easier to read
Dim Xa1 As Single ' "a" refers to the first line, "b" to the second
Dim Xb0 As Single ' 0 is the first property, 1 is the second
Dim Xb1 As Single '
Dim Ya0 As Single '
Dim Ya1 As Single '
Dim Yb0 As Single '
Dim Yb1 As Single '

Dim listCheck As Integer ' Used to compare the listCount value to look for a change

setUndo "Omptimize"

' clear the list
lstFix.Clear

Do

' see how long the list is now to compare at the end
listCheck = lstFix.ListCount

For i = 0 To shpWhat.Count - 1
    shpWhat(i).FillStyle = 1
Next i

' zero length barriers
shpWhat(0).FillStyle = 0
For i = 0 To horizontalCount
    If arrHorizontal(i, 0) = arrHorizontal(i, 1) Then
        lstFix.AddItem "Zero length barrier was at (" & arrHorizontal(i, 0) & "" & arrHorizontal(i, 2) & ")"
        removeHorizontal i
    End If
Next i
For i = 0 To verticalCount
    If arrVertical(i, 0) = arrVertical(i, 1) Then
        lstFix.AddItem "Zero length barrier was at (" & arrVertical(i, 0) & "" & arrVertical(i, 2) & ")"
        removeVertical i
    End If
Next i

' zero area zones
shpWhat(1).FillStyle = 0
For i = 0 To zoneCount
    If (arrZone(i, 1) = arrZone(i, 3)) Or (arrZone(i, 2) = arrZone(i, 4)) Then
        arrZone(i, 0) = 0
        lstFix.AddItem "Zero area zone found at (" & arrZone(i, 1) & ", " & arrZone(i, 2) & " - " & arrZone(i, 3) & ", " & arrZone(i, 4) & ")"
    End If
Next i

' overlapping barriers and end-to-end
shpWhat(2).FillStyle = 0
shpWhat(4).FillStyle = 0
i = 0
Do
    j = 0
    Do
        If j <> i Then
            Xa0 = arrHorizontal(i, 0)
            Xa1 = arrHorizontal(i, 1)
            Xb0 = arrHorizontal(j, 0)
            Xb1 = arrHorizontal(j, 1)
            Ya0 = arrHorizontal(i, 2)
            Yb0 = arrHorizontal(j, 2)
            If (Ya0 = Yb0) Then
                If (Xa0 > Xb0) And (Xa0 < Xb1) And (Xa1 > Xb1) Then
                    ' left a inside barrier b, but right outside
                    arrHorizontal(j, 1) = arrHorizontal(i, 1)
                    removeHorizontal i
                    lstFix.AddItem "Horizontal barrier overlap at: " & Xa0 & ", " & Xb1 & ", " & Ya0
                End If
                If (Xa0 >= Xb0) And (Xa1 <= Xb1) Then
                    ' a and b are the same, or a is inside b
                    lstFix.AddItem "Horizontal barrier inside another at: " & Xa0 & ", " & Xa1 & ", " & Ya0
                    removeHorizontal i
                End If
                If (Xa0 = Xb1) And (arrHorizontal(i, 3) = arrHorizontal(j, 3)) And (arrHorizontal(i, 4) = arrHorizontal(j, 4)) And (arrHorizontal(i, 5) = arrHorizontal(j, 5)) Then
                    ' a is next to b and both are identical
                    lstFix.AddItem "Identitcal horizontal barriers found end-to-end at (" & Xa0 & ", " & Ya0 & ")"
                    arrHorizontal(j, 1) = arrHorizontal(i, 1)
                    removeHorizontal i
                End If
            End If
        End If
    j = j + 1
    Loop While j <= horizontalCount
i = i + 1
Loop While i <= horizontalCount

i = 0
Do
    j = 0
    Do
        If j <> i Then
            Xa0 = arrVertical(i, 0)
            Xa1 = arrVertical(i, 1)
            Xb0 = arrVertical(j, 0)
            Xb1 = arrVertical(j, 1)
            Ya0 = arrVertical(i, 2)
            Yb0 = arrVertical(j, 2)
            If (Ya0 = Yb0) Then
                If (Xa0 > Xb0) And (Xa0 < Xb1) And (Xa1 > Xb1) Then
                    ' left a inside barrier b, but right outside
                    arrVertical(j, 1) = arrVertical(i, 1)
                    removeVertical i
                    lstFix.AddItem "Vertical barriers overlap at: " & Xa0 & ", " & Xb1 & ", " & Ya0
                End If
                If (Xa0 >= Xb0) And (Xa1 <= Xb1) Then
                    ' a and b are the same, or a is inside b
                    lstFix.AddItem "Vertical barrier inside another at: " & Xa0 & ", " & Xa1 & ", " & Ya0
                    removeVertical i
                End If
                If (Xa0 = Xb1) And (arrVertical(i, 3) = arrVertical(j, 3)) And (arrVertical(i, 4) = arrVertical(j, 4)) And (arrVertical(i, 5) = arrVertical(j, 5)) Then
                    ' a is next to b and both are identical
                    lstFix.AddItem "Identitcal vertical barriers found end-to-end at (" & Xa0 & ", " & Ya0 & ")"
                    arrVertical(j, 1) = arrVertical(i, 1)
                    removeVertical i
                End If
            End If
        End If
    j = j + 1
    Loop While j <= verticalCount
i = i + 1
Loop While i <= verticalCount

' overlapping barriers
shpWhat(3).FillStyle = 0
For i = 0 To zoneCount
    For j = i To zoneCount
        If i <> j And arrZone(i, 0) > 0 And arrZone(j, 0) > 0 Then
            Xa0 = arrZone(i, 1)
            Xa1 = arrZone(i, 3)
            Ya0 = arrZone(i, 2)
            Ya1 = arrZone(i, 4)
            Xb0 = arrZone(j, 1)
            Xb1 = arrZone(j, 3)
            Yb0 = arrZone(j, 2)
            Yb1 = arrZone(j, 4)
            If (Xa0 > Xb0) And (Xa1 < Xb1) And (Ya0 < Yb0) And (Ya1 > Yb1) Then
                lstFix.AddItem "Lower priority zone found within a higher priority zone at (" & arrZone(i, 1) & ", " & arrZone(i, 2) & " - " & arrZone(i, 3) & ", " & arrZone(i, 4) & ")"
                arrZone(i, 0) = 0
            End If
            If (Xa0 < Xb0) And (Xa1 > Xb1) And (Ya0 > Yb0) And (Ya1 < Yb1) Then
                If arrZone(i, 5) = arrZone(i, 5) And arrZone(i, 6) = arrZone(i, 6) And arrZone(i, 7) = arrZone(i, 7) And arrZone(i, 8) = arrZone(i, 8) And arrZone(i, 9) = arrZone(i, 9) And arrZone(i, 10) = arrZone(i, 10) And arrZone(i, 11) = arrZone(i, 11) And arrZone(i, 12) = arrZone(i, 12) Then
                    lstFix.AddItem "Zone found within a larger identical zone at (" & arrZone(i, 1) & ", " & arrZone(i, 2) & " - " & arrZone(i, 3) & ", " & arrZone(i, 4) & ")"
                    arrZone(j, 0) = 0
                End If
            End If
        End If
    Next j
Next i


' the loop will end when nothing was added
Loop Until (lstFix.ListCount - listCheck) = 0

If lstFix.ListCount = 0 Then lstFix.AddItem "No problems found..."

updateDisplay

End Sub

Private Sub Form_Load()
OPEN_frmOptimizer = True
mdiMain.mnuOptimizer.Checked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
OPEN_frmOptimizer = False
mdiMain.mnuOptimizer.Checked = False
End Sub
