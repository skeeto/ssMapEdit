Attribute VB_Name = "modMain"
'*************************************************
'               synSpace Map Editor
'         Originally created by Mosquito
'
'   This program was written to aid in the
' creation of maps for Synthetic-Reality's game
' called synSpace.
'   (see http://www.synthetic-reality.com)
'
'   This is the remake of the originally map
' editor which was created a few years before. It
' was crap.
'
'   If you want to contact Mosquito, the e-mail
' address is mosquito@adelphia.net
'
'        *********************************
'
'                  Terms of Use
'
'   You are free to edit this code and compile
' it in any way you like as long as you meet the
' following criteria:
'
'   1) You must keep the entire code open and
'      you must distribute the code with the
'      compiled program.
'
'   2) The program must be free. You may not
'      ever charge for the use of the program.
'
'*************************************************

'*************************************************
'           Tips for Modding This Code
'
'   Keep your code segregated from the rest of
' the code. For example, if you are adding
' something new, make a new form or function for
' it. That way, if someone else also does
' something, it will be easier to combine your
' additions.
'
'   Indicate what code belongs to you so that
' it is clearer who has done what.
'
'*************************************************

'*************************************************
'           Tips for Creating Forms
'
'   When you are creating a new form, there are
' some things you should do to make it fit right
' in to the rest of the program.
'   - Create an OPEN_frmXxxx variable to track
'     the state of the form.
'   - If your form needs to be kept up to speed
'     with properties that change, make sure you
'     add it to the sub updateDisplay.
'   - Create a menu item and set it up like the
'     others to work with the checkbox.
'   - If it is a tool, add it to the sub
'     unloadAllTools so that it will not conflict
'     with other tools. Be sure to use the sub
'     before your form opens.
'   - When your form changes a property, use
'     updateDisplay sub to get everything else
'     set up right.
'
'*************************************************

'*************************************************
'              sub/function catalog
'
'   Here are a list of the functions that are
' important outside this module:
'
'   saveMapINI(filePath)
'       Pass the path for the filename and it
'       will write all the current data to it.
'       The function returns a boolean True if
'       the save was successful.
'
'   openMapINI(filePath)
'       Pass the path of the filename and it
'       will open the file and read all the
'       map data from it and dump the data
'       into the global arrays. If the open
'       failed, the function will return a
'       boolean False.
'
'   resetMap()
'       This sub resets all the map arrays to
'       0 and reloads the default.
'
'   mdiMain.frmDraw.drawMap
'       This method will cause the map to be
'       refreshed with all the data. If the
'       map is not open, this will open the
'       map and draw it.
'
'   createHorizontal(hLeft, hRight, hTop, hColor, hXpar, hPain)
'       This offers an easy way to create a
'       horizontal barrier. It will even check
'       to make sure that another one can be
'       made. It will even make hLeft the
'       smaller value.
'
'   createVertical(vTop, vBottom, vLeft, vColor, vXpar, vPain)
'       This offers an easy way to create a
'       vertical barrier. It will even check to
'       make sure that another one can be made.
'       It will even make vBottom the smaller
'       value.
'
'   createStar(X, Y, mass, reach, temp, radius, R, G, B)
'       This sub finds the next avaliable spot
'       for a star and puts it there with the
'       given properties.
'
'   createZone(left, top, right, bottom, texture, team, pain, bullets, wpid, friction, cx, cy)
'       This function finds the next avaliable slot
'       for a zone and fills it in with the given
'       information.
'
'   unloadAllTools()
'       This sub will close all open tools in the
'       workspace. Use this command before loading
'       a new tool so that conflicting tools will
'       not pose a problem. If you create a new
'       tool, add it to this sub.
'
'   removeHorizontal(Index)
'   removeVertical(Index)
'       These remove the given array index from the
'       array. Use this to easily and quickly get
'       rid of barriers. Note: to remove other things
'       such as Zones, Stars, or PUPs, just set their
'       0 property to 0. This program is set up to
'       handle the 0 property perfectly.
'
'   deleteSelected()
'       This will deleted the selected object if
'       there is an object selected. Note: it does
'       not refresh any displays.
'
'   updateDisplay()
'       This sub will cause all open windows to refresh
'       their displays. Note: if you create a form, be
'       sure to add it to this sub.
'
'   setUndo(nameUndo)
'       This function sets up an undo. Use this before
'       changing a map and it will be able to undo it
'       by calling the useUndo function. Make sure you
'       pass the name of the undo to the sub.
'
'   useUndo()
'       This function loads the last undo map. If there
'       are no undos to load, it will return a False.
'
'   updateMainTitleBar()
'       Change this sub to change the display in the main
'       title bar.
'
'*************************************************

' I will invoke option explicit to force us all
' to keep things nice and neat. :-)
Option Explicit

' This is the current version number. The basic
' format I am using here is X.XXXAB where X.XXX
' is simply some standard version number while
' the AB are initials for the programmer. As
' Mosquito, I will use "MQ" and so my first
' version is "1.000MQ". Additionally, people
' who mod a mod could just tag on their intials
' like "1.000SKDS". SK for one person and DS for
' antoher. The MQ isn't needed because it can be
' assumed so you won't tag onto my MQ.
Public Const versionNumber = "1.003MQ"

' These handle the Undo function. Undo works by saving the
' map before every change to a temporary file. undoCount
' keeps track of which undo file the program last used.
' undoLeft keeps track of how many undos are left to use
' and UNDO_MAX decides how many undos are possible (which
' can be set by the mapdefs.dat file. undoName() keeps
' track of the name of the undo such as "Barrier" or
' "Zone." This is handled by the setUndo function.
' savingUndo tells the saveMapINI function to not clean up
' arrays because it messes up the order of everything.
Global undoCount As Integer
Global undoLeft As Integer
Global UNDO_MAX As Integer
Global undoName(99) As String
Global savingUndo As Boolean

' This keeps track of the save state. If the map has been
' saved and no changes have been made, it will be True.
Global varSave As Boolean

' This holds the current filename and path to make saving
' easier. It also makes "Save/Save as" work
Global saveFileName
Global pathFileName

' This will tell the save function if it should write in
' the first twenty slots under powerups.
Global disableAllRandomPups As Boolean

' These tell everthing what object has been selected by
' the "Selector". Usually, all values except one are set
' to -1.
Global selSpawn
Global selZone
Global selStar
Global selHorizontal
Global selVertical
Global selPUP

' This keeps track of what windows are open and not open. It
' is up to the form to take car of it's boolean. If you
' create a new form, add a new variable here and change the
' updateDisplay function.
Global OPEN_frmMap As Boolean
Global OPEN_frmEdit As Boolean
Global OPEN_frmCredit As Boolean
Global OPEN_frmColor As Boolean
Global OPEN_frmPain As Boolean
Global OPEN_frmSelector As Boolean
Global OPEN_frmDisplay As Boolean
Global OPEN_frmOptimizer As Boolean
Global OPEN_frmFillarea As Boolean
Global OPEN_frmDeletearea As Boolean

' This last one tells everything that a display is updating.
' This is so when the forms (or whatever) update, they don't
' need to tell everything to update because this is a
' recursive mess.
Global varUpdating As Boolean

' This keeps track of the current tool being used for
' map editing. This is the communication line between
' an editing window and the map window where the actual
' clicking happens.
Global varTool As String

' These pass tool data to the sub that handles the tool.
Global toolXpar As Single
Global toolColorIndex As Single
Global toolPainIndex As Single
Global toolMass As Single
Global toolReach As Single
Global toolRadius As Single
Global toolColorR As Single
Global toolColorG As Single
Global toolColorB As Single
Global toolTemperature As Single
Global toolHeading As Single
Global toolCurrentSpawn As Integer
Global toolFriction As Single
Global toolPain As Single
Global toolCurrentX As Single
Global toolCurrentY As Single
Global toolWaypointID As Integer
Global toolTeam As Integer
Global toolBullets As Integer
Global toolRespawn As Single
Global toolPUPStyle As Integer
Global toolZoneStyle As Integer
Global toolFillFrequency As Single
Global toolFillHorizontal As Boolean
Global toolFillVertical As Boolean
Global toolFillBorder As Boolean
Global toolDeleteHorizontal As Boolean
Global toolDeleteVertical As Boolean
Global toolDeleteZone As Boolean
Global toolDeleteStar As Boolean
Global toolDeletePUP As Boolean

' This holds the xpar values that were obtained from the
' file. These value are painted into the dropdown boxes
' later on.
Global xparBarrier(31, 1) As String
Global xparBarrierCount As Integer
Global stylePUP(63, 1) As String
Global stylePUPCount As Integer

' These are the display option variables that are used
' by anything that displays a map.
Global DISP_showBarriers As Integer
Global DISP_showPUP As Integer
Global DISP_showStar As Integer
Global DISP_showGrid As Integer
Global DISP_showSpawn As Integer
Global DISP_showZone As Integer
Global DISP_gridColor As Single
Global DISP_gridSnap As Single
Global DISP_starMult As Integer
Global DISP_zoneColor(2) As Single

'*************************************************
' These are the global arrays that hold everything
' that will go into the map .ini file. For more
' information on these, see the map0.ini.
'
' For arrays, the first value indicates the onject
' id number. The second number is the property
' number.
'
Global strCredits As String
Global arrHorizontal(99, 5) As Integer
Global arrVertical(99, 5) As Integer
Global arrStar(15, 9) As Single
Global arrSpawn(7, 2) As Integer
Global arrPUP(235, 3) As Integer
Global arrZone(99, 12) As Single
Global arrColor(15, 2) As Integer
Global arrPain(15) As Integer
'
' These hold the values from the mapdefs.dat file
Global HORIZONTAL_MAX
Global VERTICAL_MAX
Global STAR_MAX
Global SPAWN_MAX
Global COLOR_MAX
Global PAIN_MAX
Global PUP_MAX
Global ZONE_MAX
' This last one holds the path to the location of synSpace
Global arcadiaPath As String
'
' And now this keeps track of the number of properties
Public Const horizontalProp = 5
Public Const verticalProp = 5
Public Const starProp = 9
Public Const spawnProp = 2
Public Const colorProp = 2
Public Const pupProp = 3
Public Const zoneProp = 12
'
' This keeps track of how many we have of everything
Global horizontalCount
Global verticalCount
Global starCount
Global spawnCount
Global pupCount
Global zoneCount
Global colorCount
Global painCount
'
'*************************************************

' This is what begins the magic...
Sub Main()

' set up the undo directory
createUndoDir

' initialize some variables
DISP_showBarriers = 1
DISP_showPUP = 1
DISP_showStar = 1
DISP_showGrid = 1
DISP_showSpawn = 1
DISP_showZone = 1
DISP_gridColor = RGB(0, 127, 0)
DISP_gridSnap = 1
DISP_starMult = 3
DISP_zoneColor(0) = RGB(150, 0, 0)
DISP_zoneColor(1) = RGB(0, 150, 150)
DISP_zoneColor(2) = RGB(150, 150, 0)

' Begin by getting all the things that the program needs to know.
'
' Get the map definitions from the mapdefs.dat file but set things
' up first so that the program will still function if the file is
' missing or some values are missing for some reason.
HORIZONTAL_MAX = 99
VERTICAL_MAX = 99
STAR_MAX = 15
SPAWN_MAX = 7
COLOR_MAX = 15
PAIN_MAX = 15
PUP_MAX = 235 ' note, the first 20 powerups are reserved so that 20 less are avaliable
ZONE_MAX = 99
UNDO_MAX = 19
arcadiaPath = "c:\arcadia\toys\toy7\maps\"
If getMapDefs() = False Then MsgBox "Error opening file mapdefs.dat!"

' This is the same as starting a new map from the menu:
resetMap

' Now, we load the xpar and style values from dat files:
If loadBarrierXpar() = False Then MsgBox "Error loading barrier xpar values from file barxpar.dat!"
If loadStylePUP() = False Then MsgBox "Error loading powerup styles from file pupstyle.dat!"

' Now, we begin the visual part of the program.
mdiMain.Show
updateMainTitleBar

' Start the session with a new map window:
frmMap.drawMap

End Sub

Public Sub resetMap()
Dim i As Integer, j As Integer  ' throwaway loop variables

' reset the undo
undoLeft = 0
mdiMain.mnuUndo.Enabled = False

' reset the counters
horizontalCount = -1
verticalCount = -1
starCount = -1
spawnCount = -1
pupCount = -1
zoneCount = -1
colorCount = -1
painCount = -1

' reset some other variables
saveFileName = ""
updateMainTitleBar
varSave = True
varUpdating = False
disableAllRandomPups = False

' turn off select
unselectAll

' spawn
For i = 0 To SPAWN_MAX
    For j = 0 To spawnProp
        arrSpawn(i, j) = 0
    Next j
Next i

' credits
strCredits = ""

' stars
For i = 0 To STAR_MAX
    For j = 0 To starProp
        arrStar(i, j) = 0
    Next j
Next i

' horizontal
For i = 0 To HORIZONTAL_MAX
    For j = 0 To horizontalProp
        arrHorizontal(i, j) = 0
    Next j
Next i

' vertical
For i = 0 To VERTICAL_MAX
    For j = 0 To verticalProp
        arrVertical(i, j) = 0
    Next j
Next i

' zones
For i = 0 To ZONE_MAX
    For j = 0 To zoneProp
        arrZone(i, j) = 0
    Next j
Next i

' PUPs
For i = 0 To PUP_MAX
    For j = 0 To pupProp
        arrPUP(i, j) = 0
    Next j
Next i

' pain
For i = 0 To PAIN_MAX
    arrPain(i) = 0
Next i


' Get the default setup by opening the defaults.dat file as if
' it was a map
If openMapINI("defaults.dat") = False Then MsgBox "Error opening file defaults.dat!"

End Sub

' This opens the mapdefs.dat file and sets the appropriate variables
Private Function getMapDefs()
On Error GoTo ErrorTrap

Dim varFile As String ' a string that holds the whole file
Dim curLine, curVal As String ' these hold the line being evaluated
Dim nextChar, curChar As Integer ' these keep track of the the position of the end of the line
curChar = 1

' open the file
Open "mapdefs.dat" For Input As #1

Do
    varFile = varFile & Input(1, #1)
Loop Until EOF(1)

' close the file
Close #1

Do

nextChar = InStr(curChar, varFile, "=")
' nextChar is 0 if there are no more values in the file (no more equal signs)
If nextChar = 0 Then Exit Do
curLine = Mid(varFile, curChar, nextChar - curChar)
curChar = InStr(nextChar, varFile, Chr(13))
curVal = Mid(varFile, nextChar + 1, curChar - nextChar + 1)

If InStr(1, UCase(curLine), "HORIZONTAL_MAX") > 0 Then HORIZONTAL_MAX = Val(curVal)
If InStr(1, UCase(curLine), "VERTICAL_MAX") > 0 Then VERTICAL_MAX = Val(curVal)
If InStr(1, UCase(curLine), "STAR_MAX") > 0 Then STAR_MAX = Val(curVal)
If InStr(1, UCase(curLine), "SPAWN_MAX") > 0 Then SPAWN_MAX = Val(curVal)
If InStr(1, UCase(curLine), "COLOR_MAX") > 0 Then COLOR_MAX = Val(curVal)
If InStr(1, UCase(curLine), "PAIN_MAX") > 0 Then PAIN_MAX = Val(curVal)
If InStr(1, UCase(curLine), "PUP_MAX") > 0 Then PUP_MAX = Val(curVal)
If InStr(1, UCase(curLine), "ZONE_MAX") > 0 Then ZONE_MAX = Val(curVal)
If InStr(1, UCase(curLine), "UNDO_MAX") > 0 Then UNDO_MAX = Val(curVal)
If InStr(1, UCase(curLine), "ARCADIA") > 0 Then arcadiaPath = removeSpace(curVal)

Loop

getMapDefs = True
Exit Function

ErrorTrap:
' file failed to work properly, so return a False boolean
getMapDefs = False
End Function

' This function returns the given string without header spaces
Private Function removeSpace(curVal)
10
If Mid(curVal, 1, 1) = " " Then curVal = Mid(curVal, 2, Len(curVal)): GoTo 10
removeSpace = curVal
End Function

' This opens the map file - see the catalog above for detail
'
' This function should never have to be fixed even if more properties
' become avaliable because the total number of properties is stored
' in constants
Public Function openMapINI(filePath As String)
On Error GoTo ErrorTrap

' begin by resetting the map
horizontalCount = -1
verticalCount = -1
starCount = -1
spawnCount = -1
pupCount = -1
zoneCount = -1
colorCount = -1
painCount = -1

Dim varFile As String            ' a string that holds the whole file
Dim varMode As String            ' the mode that helps put values where they belong
Dim curLine As String            ' holds the line being evaluated
Dim nextChar As Integer          ' keeps track of the the position of the end of the line
Dim curChar As Integer           ' keeps track of the the position of the end of the line
Dim argVal()                     ' used to hold onto retun values
Dim i As Integer                 ' throwaway loop number


strCredits = ""

curChar = 1

' open the file
Open filePath For Input As #1

Do
    varFile = varFile & Input(1, #1)
Loop Until EOF(1)

' close the file
Close #1


' This section uses a varMode variable to keep track of the current "Mode".
' When opening a certain section (such as [stars] or [pains]), the mode
' is set to the name of the current section being set up.


Do
' get a single line

nextChar = InStr(curChar, varFile, Chr(10))
' nextChar is 0 if there is nothing left
If nextChar = 0 Then Exit Do

curLine = Mid(varFile, curChar, nextChar - curChar)

' check for a comment line
If InStr(1, curLine, ";") = 0 Then
    ' check for a mode change
    If InStr(1, curLine, "[credits]") > 0 Then varMode = "[credits]"
    If InStr(1, curLine, "[stars]") > 0 Then varMode = "[stars]"
    If InStr(1, curLine, "[spawn]") > 0 Then varMode = "[spawn]"
    If InStr(1, curLine, "[horizontal]") > 0 Then varMode = "[horizontal]"
    If InStr(1, curLine, "[vertical]") > 0 Then varMode = "[vertical]"
    If InStr(1, curLine, "[pains]") > 0 Then varMode = "[pains]"
    If InStr(1, curLine, "[powerups]") > 0 Then varMode = "[powerups]"
    If InStr(1, curLine, "[colors]") > 0 Then varMode = "[colors]"
    If InStr(1, curLine, "[zones]") > 0 Then varMode = "[zones]"
    ' without a mode, nothing can happen
    ' there also needs to be an equal sign in the line to be valid
    If InStr(1, curLine, "=") > 0 Then
        Select Case varMode
            Case "[credits]"
                strCredits = strCredits & curLine & vbNewLine
            Case "[zones]"
                zoneCount = zoneCount + 1
                If zoneCount <= ZONE_MAX Then
                    For i = 0 To zoneProp
                        arrZone(zoneCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[stars]"
                starCount = starCount + 1
                If starCount <= STAR_MAX Then
                    For i = 0 To starProp
                        arrStar(starCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[spawn]"
                spawnCount = spawnCount + 1
                If spawnCount <= STAR_MAX Then
                    For i = 0 To spawnProp
                        arrSpawn(spawnCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[horizontal]"
                If horizontalCount <= HORIZONTAL_MAX Then
                    horizontalCount = horizontalCount + 1
                    For i = 0 To horizontalProp
                        arrHorizontal(horizontalCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[vertical]"
                verticalCount = verticalCount + 1
                If verticalCount <= VERTICAL_MAX Then
                    For i = 0 To verticalProp
                        arrVertical(verticalCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[powerups]"
                pupCount = pupCount + 1
                If pupCount <= PUP_MAX Then
                    For i = 0 To pupProp
                        arrPUP(pupCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[colors]"
                colorCount = colorCount + 1
                If colorCount <= COLOR_MAX Then
                    For i = 0 To colorProp
                        arrColor(colorCount, i) = argReturn(curLine, i)
                    Next i
                End If
            Case "[pains]"
                painCount = painCount + 1
                If painCount <= PAIN_MAX Then arrPain(painCount) = argReturn(curLine, 0)
        End Select
    End If
End If

curChar = nextChar + 1
Loop

' Now we have to check for the annoying thing that powerups do.
' If the first 20 slots are blank, the random powerups are disabled
' so we have to keep that in mind.
disableAllRandomPups = True
For i = 0 To 19
    If arrPUP(i, 0) > 0 Then disableAllRandomPups = False
Next i

' This will adjust any value order errors in the file
fixPropertyOrders

openMapINI = True
Exit Function

ErrorTrap:
' file failed to work properly, so return a False boolean
openMapINI = False
End Function

' This returns the requested argument number in the given string
' If a non-existant argument is requested, the function returns 0
Private Function argReturn(lineFeed As String, argNum As Integer)

Dim curChar As Integer ' keeps track of current place
Dim i As Integer       ' throwaway loop number

curChar = InStr(1, lineFeed, "=") + 1

If argNum > 0 Then
For i = 1 To argNum
curChar = InStr(curChar, lineFeed, ",") + 1
If curChar = 1 Then argReturn = 0: Exit Function
Next i
End If

argReturn = Val(Mid(lineFeed, curChar, Len(lineFeed)))

End Function

' This saves the map file - see the catalog above for detail
'
' This function should never have to be fixed even if more properties
' become avaliable because the total number of properties is stored
' in constants
'
' The only thing you would change is the programming credits
Public Function saveMapINI(filePath As String)
On Error GoTo ErrorTrap

Dim i As Integer, j As Integer  ' a throwaway loop variables
Dim buildLine As String         ' line building buffer

' clean up PUP and Zone arrays except if saving for an Undo:
If savingUndo <> True Then
    cleanPUPentry
    cleanZoneEntry
End If

' Open the file
Open filePath For Output As #1

' This is the program credit section.
' If you have modified this editor,
' feel free to do whatever you want
' with this segment:
Print #1, ";***********************************"
Print #1, "; Created using:                    "
Print #1, ";    Mosquito's synSpace Map Editor "
Print #1, ";        version " & versionNumber
Print #1, ";                                   "
Print #1, "; mosquito@adelphia.net             "
Print #1, ";***********************************" & vbNewLine

' credits
Print #1, "[credits]"
Print #1, strCredits


' spawn
Print #1, "[spawn]"

For i = 0 To SPAWN_MAX
    buildLine = (i + 1) & "="
    For j = 0 To spawnProp
        buildLine = buildLine & " " & arrSpawn(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' stars
Print #1, "[stars]"

For i = 0 To STAR_MAX
    buildLine = i & "="
    For j = 0 To starProp
        buildLine = buildLine & " " & arrStar(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' colors
Print #1, "[colors]"

For i = 0 To COLOR_MAX
    buildLine = i & "="
    For j = 0 To colorProp
        buildLine = buildLine & " " & arrColor(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' pains
Print #1, "[pains]"

For i = 0 To PAIN_MAX
    Print #1, i & "=" & arrPain(i)
Next i

Print #1, ""


' horizontal barriers
Print #1, "[horizontal]"

For i = 0 To horizontalCount
    buildLine = i & "="
    For j = 0 To horizontalProp
        buildLine = buildLine & " " & arrHorizontal(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' vertical barriers
Print #1, "[vertical]"

For i = 0 To verticalCount
    buildLine = i & "="
    For j = 0 To verticalProp
        buildLine = buildLine & " " & arrVertical(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' powerups
Print #1, "[powerups]"
If disableAllRandomPups = True Then
    For i = 0 To 19
        Print #1, i & "= 0, 0, 0, 0"
    Next i
End If
For i = 0 To pupCount
    buildLine = (i + 20) & "="
    For j = 0 To pupProp
        buildLine = buildLine & " " & arrPUP(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""

' zones
Print #1, "[zones]"

For i = 0 To zoneCount
    buildLine = i & "="
    For j = 0 To zoneProp
        buildLine = buildLine & " " & arrZone(i, j) & ","
    Next j
' dont write the last character because it is a comma
    Print #1, Mid(buildLine, 1, Len(buildLine) - 1)
Next i

Print #1, ""
Print #1, ""

' Close the file
Close #1

saveMapINI = True

Exit Function

ErrorTrap:
' If the file failed to write, return a false value
saveMapINI = False
' make sure it is closed up
Close #1

End Function

' this function make it easier to create a horizontal barrier
Public Sub createHorizontal(hLeft As Single, hRight As Single, hTop As Single, hColor As Single, hXpar As Single, hPain As Single)

Dim swapVar As Single ' variable used to swap values

' check for boundries
If hLeft < 0 Then hLeft = 0
If hTop < 0 Then hTop = 0
If hRight < 0 Then hRight = 0
If hRight > 8192 Then hRight = 8192
If hTop > 8192 Then hTop = 8192
If hLeft > 8192 Then hLeft = 8192

horizontalCount = horizontalCount + 1
If horizontalCount > HORIZONTAL_MAX Or hLeft = hRight Then Exit Sub

If hLeft > hRight Then
swapVar = hLeft
hLeft = hRight
hRight = swapVar
End If

arrHorizontal(horizontalCount, 0) = hLeft
arrHorizontal(horizontalCount, 1) = hRight
arrHorizontal(horizontalCount, 2) = hTop
arrHorizontal(horizontalCount, 3) = hColor
arrHorizontal(horizontalCount, 4) = hXpar
arrHorizontal(horizontalCount, 5) = hPain

End Sub

' this function make it easier to create a vertical barrier
Public Sub createVertical(vBottom As Single, vTop As Single, vLeft As Single, vColor As Single, vXpar As Single, vPain As Single)

Dim swapVar As Single ' variable used to swap values

' check for boundries
If vBottom < 0 Then vBottom = 0
If vTop < 0 Then vTop = 0
If vLeft < 0 Then vLeft = 0
If vBottom > 8192 Then vBottom = 8192
If vTop > 8192 Then vTop = 8192
If vLeft > 8192 Then vLeft = 8192

verticalCount = verticalCount + 1
If verticalCount > VERTICAL_MAX Or vTop = vBottom Then Exit Sub

If vBottom > vTop Then
swapVar = vBottom
vBottom = vTop
vTop = swapVar
End If

arrVertical(verticalCount, 0) = vBottom
arrVertical(verticalCount, 1) = vTop
arrVertical(verticalCount, 2) = vLeft
arrVertical(verticalCount, 3) = vColor
arrVertical(verticalCount, 4) = vXpar
arrVertical(verticalCount, 5) = vPain

End Sub

' this function loads the barrier xpar values from barxpar.dat
' and dumps them into xparBarrier
Public Function loadBarrierXpar()
On Error GoTo ErrorTrap

xparBarrierCount = -1

Open "barxpar.dat" For Input As #1

Do
    xparBarrierCount = xparBarrierCount + 1
    Input #1, xparBarrier(xparBarrierCount, 0)
    Input #1, xparBarrier(xparBarrierCount, 1)
Loop Until EOF(1)


Close #1

loadBarrierXpar = True

Exit Function

ErrorTrap:
' error occured, so we return a false
loadBarrierXpar = False
End Function

Public Function loadStylePUP()
'On Error GoTo ErrorTrap

stylePUPCount = -1

Open "pupstyle.dat" For Input As #1

Do
    stylePUPCount = stylePUPCount + 1
    Input #1, stylePUP(stylePUPCount, 0)
    Input #1, stylePUP(stylePUPCount, 1)
Loop Until EOF(1)


Close #1

loadStylePUP = True

Exit Function

ErrorTrap:
' error occured, so we return a false
loadStylePUP = False
End Function

' this function finds the next open star slot and puts a star there
Public Sub createStar(X As Single, Y As Single, mass As Single, reach As Single, temp As Single, radius As Single, R As Single, G As Single, B As Single)

Dim i As Integer       ' throwaway loop variable

If X < 0 Then X = 0
If Y < 0 Then Y = 0
If X > 8192 Then X = 8192
If Y > 8192 Then Y = 8192

For i = 0 To STAR_MAX
    If arrStar(i, 0) = 0 Then
        arrStar(i, 0) = 1
        arrStar(i, 1) = X
        arrStar(i, 2) = Y
        arrStar(i, 3) = mass
        arrStar(i, 4) = reach
        arrStar(i, 5) = temp
        arrStar(i, 6) = radius
        arrStar(i, 7) = R
        arrStar(i, 8) = G
        arrStar(i, 9) = B
        Exit Sub
    End If
Next i

MsgBox "No more stars avaliable. You have reached the maximum (" & (STAR_MAX + 1) & ")."

End Sub

Public Sub createPUP(pupStyle As Integer, X As Single, Y As Single, respawn As Single)

Dim i As Integer       ' throwaway loop variable

For i = 0 To PUP_MAX
    If arrPUP(i, 0) = 0 Then
        If i > pupCount Then pupCount = i
        arrPUP(i, 0) = pupStyle
        arrPUP(i, 1) = X
        arrPUP(i, 2) = Y
        arrPUP(i, 3) = respawn
        Exit Sub
    End If
Next i

MsgBox "Sweet Jesus! You just tried to add one more past " & (PUP_MAX + 1) & " which is the limit. Remove a PUP before adding a new one!"

End Sub

' this finds the next zone spot and fills it with the given information
Public Sub createZone(zStyle As Integer, left As Single, top As Single, right As Single, bottom As Single, texture As Single, team As Integer, pain As Single, bullets As Integer, wpid As Integer, friction As Single, cx As Single, cy As Single)

Dim i As Integer       ' throwaway loop variable
Dim swapVar As Single  ' variable used to swap values

' check for boundries
If left < 0 Then left = 0
If right < 0 Then right = 0
If top < 0 Then top = 0
If bottom < 0 Then bottom = 0
If left > 8192 Then left = 8192
If right > 8192 Then right = 8192
If top > 8192 Then top = 8192
If bottom > 8192 Then bottom = 8192

' check for order
If left > right Then
    swapVar = left
    left = right
    right = swapVar
End If
If bottom > top Then
    swapVar = top
    top = bottom
    bottom = swapVar
End If

'check for real area
If left = right Then Exit Sub
If top = bottom Then Exit Sub

For i = 0 To ZONE_MAX
    If arrZone(i, 0) = 0 Then
        If zoneCount < i Then zoneCount = i
        arrZone(i, 0) = zStyle
        arrZone(i, 1) = left
        arrZone(i, 2) = top
        arrZone(i, 3) = right
        arrZone(i, 4) = bottom
        arrZone(i, 5) = texture
        arrZone(i, 6) = team
        arrZone(i, 7) = pain
        arrZone(i, 8) = bullets
        arrZone(i, 9) = wpid
        arrZone(i, 10) = friction
        arrZone(i, 11) = cx
        arrZone(i, 12) = cy
        Exit Sub
    End If
Next i

MsgBox "Holy crap, man! You have just tried to use more than " & (ZONE_MAX + 1) & " zones, which is the maximum allowed!"

End Sub

' This function goes through all the pup entries and removes any with "0" style
Public Sub cleanPUPentry()

Dim i As Integer ' loop variable
Dim j As Integer ' loop variable
Dim k As Integer ' loop variable

If pupCount = -1 Then Exit Sub

i = -1

Do
i = i + 1
If i > pupCount Then Exit Do
    If arrPUP(i, 0) = 0 Then
        For j = i To pupCount
            For k = 0 To pupProp
                arrPUP(j, k) = arrPUP(j + 1, k)
            Next k
        Next j
        pupCount = pupCount - 1
        If pupCount = -1 Then Exit Sub
        i = i - 1
    End If
Loop
End Sub

' This function goes through all the zone entries and removes any with a "0" style
Public Sub cleanZoneEntry()

Dim i As Integer ' loop variable
Dim j As Integer ' loop variable
Dim k As Integer ' loop variable

If zoneCount = -1 Then Exit Sub

i = -1

Do
i = i + 1
If i > zoneCount Then Exit Do
    If arrZone(i, 0) = 0 Then
        For j = i To zoneCount
            For k = 0 To zoneProp
                arrZone(j, k) = arrZone(j + 1, k)
            Next k
        Next j
        zoneCount = zoneCount - 1
        If zoneCount = -1 Then Exit Sub
        i = i - 1
    End If
Loop

End Sub

' This function interprets the given X and Y coordinates into a selection.
Public Sub setSelect(X, Y)

Dim i As Integer ' loop variable
Dim clickAcc As Integer ' sets how accurate the user must be in a click in terms of map units

unselectAll

clickAcc = 64

' We will look at the type with least priority first and the most priority last.
' This way, when objects are piled, we select the one that the user probably wants.

' zones
' These are the easiest to select, so they go last... er... uh, first.
If zoneCount > -1 Then
    For i = 0 To zoneCount
        If arrZone(i, 0) > 0 Then
            If (arrZone(i, 1) < X And arrZone(i, 3) > X) Then
                'if it got to here, the x was within the zone width
                If (arrZone(i, 4) < Y And arrZone(i, 2) > Y) Then
                    unselectAll
                    selZone = i
                End If
            End If
        End If
    Next i

End If

' star
If starCount > -1 Then
    For i = 0 To starCount
        If arrStar(i, 0) > 0 And (arrStar(i, 1) - clickAcc < X And arrStar(i, 1) + clickAcc > X) And (arrStar(i, 2) - clickAcc < Y And arrStar(i, 2) + clickAcc > Y) Then
            unselectAll
            selStar = i
        End If
    Next i
End If

' PUP
If pupCount > -1 Then
    For i = 0 To pupCount
        If arrPUP(i, 0) > 0 And (arrPUP(i, 1) - clickAcc < X And arrPUP(i, 1) + clickAcc > X) And (arrPUP(i, 2) - clickAcc < Y And arrPUP(i, 2) + clickAcc > Y) Then
            unselectAll
            selPUP = i
        End If
    Next i
End If

' barriers
' horizontal
If horizontalCount > -1 Then
    For i = 0 To horizontalCount
        If (arrHorizontal(i, 0) < X And arrHorizontal(i, 1) > X) Or (arrHorizontal(i, 1) < X And arrHorizontal(i, 0) > X) Then
            If (arrHorizontal(i, 2) - clickAcc < Y And arrHorizontal(i, 2) + clickAcc > Y) Then
                unselectAll
                selHorizontal = i
            End If
        End If
    Next i
End If

' vertical
If verticalCount > -1 Then
    For i = 0 To verticalCount
        If (arrVertical(i, 0) < Y And arrVertical(i, 1) > Y) Or (arrVertical(i, 1) < Y And arrVertical(i, 0) > Y) Then
            If (arrVertical(i, 2) - clickAcc < X And arrVertical(i, 2) + clickAcc > X) Then
                unselectAll
                selVertical = i
            End If
        End If
    Next i
End If

End Sub

' this unselects all objects
Public Sub unselectAll()
selSpawn = -1
selZone = -1
selStar = -1
selHorizontal = -1
selVertical = -1
selPUP = -1
End Sub

' This is an important function that clears the workspace of all tools to make room for a new tool.
' This is used so that tools cannot conflict. If you create a tool, add it to the list.
Public Sub unloadAllTools()
Unload frmEdit
Unload frmSelector
Unload frmFillarea
Unload frmDeletearea
End Sub

' This sub removes a given horizontal and shifts all data down
Public Sub removeHorizontal(Index)
Dim i As Integer ' loop variable
Dim j As Integer ' loop variable

If horizontalCount > -1 Then
    horizontalCount = horizontalCount - 1
    For i = Index To horizontalCount
        For j = 0 To horizontalProp
            arrHorizontal(i, j) = arrHorizontal(i + 1, j)
        Next j
    Next i
End If
End Sub

' this sub removes a given vertical and shifts all data down
Public Sub removeVertical(Index)
Dim i As Integer ' loop variable
Dim j As Integer ' loop variable

If verticalCount > -1 Then
    verticalCount = verticalCount - 1
    For i = Index To verticalCount
        For j = 0 To verticalProp
            arrVertical(i, j) = arrVertical(i + 1, j)
        Next j
    Next i
End If
End Sub

' This deletes the currently selected object
Public Sub deleteSelected()

' be able to undo it
setUndo "Delete"

' zone
If selZone > -1 Then
    arrZone(selZone, 0) = 0
    unselectAll
End If

' star
If selStar > -1 Then
    arrStar(selStar, 0) = 0
    unselectAll
End If

' PUP
If selPUP > -1 Then
    arrPUP(selPUP, 0) = 0
    unselectAll
End If

' horizontal
If selHorizontal > -1 Then
    removeHorizontal selHorizontal
    unselectAll
End If

' vertical
If selVertical > -1 Then
    removeVertical selVertical
    unselectAll
End If

End Sub

' This checks to see what windows are open and tells them to refresh it they are.
Public Sub updateDisplay()

If varUpdating = True Then Exit Sub

varUpdating = True

If OPEN_frmMap = True Then
    frmMap.drawMap
End If

If OPEN_frmEdit = True Then
    frmEdit.updateEdit
End If

If OPEN_frmCredit = True Then
    frmCredit.updateCredit
End If

If OPEN_frmColor = True Then
    frmColor.updateColors
End If

If OPEN_frmPain = True Then
    frmPain.updatePain
End If

If OPEN_frmSelector = True Then
    frmSelector.showProperties
End If

If OPEN_frmFillarea = True Then
    frmFillarea.updateFillarea
End If

varUpdating = False

End Sub

' This function sets up an undo. Use it before changing something to be able to undo it.
Public Sub setUndo(nameUndo As String)

' up the counter
undoCount = undoCount + 1
If undoCount > UNDO_MAX Then undoCount = 0

undoLeft = undoLeft + 1
If undoLeft > UNDO_MAX Then undoLeft = UNDO_MAX

' This almost seems sloppy to put these lines here, but this is the best spot for them
mdiMain.mnuUndo.Enabled = True
mdiMain.mnuUndo.Caption = "Undo " & nameUndo

undoName(undoCount) = nameUndo

savingUndo = True
If saveMapINI("undo\temp" & undoCount) = False Then MsgBox "Error setting up Undo filesystem."
savingUndo = False

End Sub

' This function calls back the last undo from the file. It will return false if the undo is not possible.
Public Function useUndo()

useUndo = True

' check to see if undo is possible
If undoLeft <= 0 Then
    useUndo = False
    Exit Function
End If
undoLeft = undoLeft - 1

If openMapINI("undo\temp" & undoCount) = False Then MsgBox "Error loading Undo."

' down the counter
undoCount = undoCount - 1
If undoCount < 0 Then undoCount = UNDO_MAX

End Function

Public Sub createUndoDir()
On Error Resume Next

MkDir "undo\"

End Sub

Public Sub removeUndoDir()
On Error Resume Next

Dim i As Integer ' loop variable

For i = 0 To UNDO_MAX
    Kill "undo\temp" & i
Next i

RmDir "undo\"

End Sub

' This sub checks all the barriers to make sure that the left value is less
' than the right for horizontal barriers and bottom values are less than the
' top for vertical barriers.
Public Sub fixPropertyOrders()

Dim i As Integer ' loop variable
Dim swapVar As Single ' used to swap variable values

If horizontalCount > -1 Then
    For i = 0 To horizontalCount
        If arrHorizontal(i, 0) > arrHorizontal(i, 1) Then
            swapVar = arrHorizontal(i, 0)
            arrHorizontal(i, 0) = arrHorizontal(i, 1)
            arrHorizontal(i, 1) = swapVar
        End If
    Next i
End If

If verticalCount > -1 Then
    For i = 0 To verticalCount
        If arrVertical(i, 0) > arrVertical(i, 1) Then
            swapVar = arrVertical(i, 0)
            arrVertical(i, 0) = arrVertical(i, 1)
            arrVertical(i, 1) = swapVar
        End If
    Next i
End If

If zoneCount > -1 Then
    For i = 0 To zoneCount
        If arrZone(i, 0) > 0 Then
            If arrZone(i, 1) > arrZone(i, 3) Then
                swapVar = arrZone(i, 1)
                arrZone(i, 1) = arrZone(i, 3)
                arrZone(i, 3) = swapVar
            End If
        End If
    Next i
    For i = 0 To zoneCount
        If arrZone(i, 0) > 0 Then
            If arrZone(i, 4) > arrZone(i, 2) Then
                swapVar = arrZone(i, 4)
                arrZone(i, 4) = arrZone(i, 2)
                arrZone(i, 2) = swapVar
            End If
        End If
    Next i
End If

End Sub

' This sub fills the given area with the given property barriers. This is used by the Fill Area tool
Public Sub FillArea(x0 As Single, y0 As Single, x1 As Single, y1 As Single, freq As Single, color As Single, xPar As Single, pain As Single)

Dim swapVar As Single ' variable used to swap values
Dim i As Integer ' loop variable

' check for boundries
If x0 < 0 Then x0 = 0
If x0 > 8192 Then x0 = 8192
If x1 < 0 Then x1 = 0
If x1 > 8192 Then x1 = 8192
If y0 < 0 Then y0 = 0
If y0 > 8192 Then y0 = 8192
If y1 < 0 Then y1 = 0
If y1 > 8192 Then y1 = 8192

' put things in order
If x0 > x1 Then
swapVar = x0
x0 = x1
x1 = swapVar
End If

If y0 > y1 Then
swapVar = y0
y0 = y1
y1 = swapVar
End If

If toolFillBorder = -1 Then
    createHorizontal x0, x1, y0, color, xPar, pain
    createHorizontal x0, x1, y1, color, xPar, pain
    createVertical y0, y1, x0, color, xPar, pain
    createVertical y0, y1, x1, color, xPar, pain
End If

If toolFillHorizontal = True Then
    For i = 1 To Int((y1 - y0) / freq)
        createHorizontal x0, x1, y0 + (freq * i), color, xPar, pain
    Next i
End If

If toolFillVertical = True Then
    For i = 1 To Int((x1 - x0) / freq)
        createVertical y0, y1, x0 + (freq * i), color, xPar, pain
    Next i
End If

If horizontalCount > HORIZONTAL_MAX Then horizontalCount = HORIZONTAL_MAX
If verticalCount > VERTICAL_MAX Then verticalCount = VERTICAL_MAX

End Sub

' This sub deletes eveything in the given area. It used the tool variables to determine what needs to be deleted
Public Sub DeleteArea(x0 As Single, y0 As Single, x1 As Single, y1 As Single)

Dim i As Integer ' loop variable
Dim swapVar As Single ' used to swap variable values

' put things in order
If x0 > x1 Then
swapVar = x0
x0 = x1
x1 = swapVar
End If

If y0 > y1 Then
swapVar = y0
y0 = y1
y1 = swapVar
End If

If toolDeleteHorizontal = True Then
    i = 0
    Do
        If arrHorizontal(i, 0) > x0 And arrHorizontal(i, 0) < x1 And arrHorizontal(i, 1) > x0 And arrHorizontal(i, 1) < x1 And arrHorizontal(i, 2) > y0 And arrHorizontal(i, 2) < y1 Then
            removeHorizontal i
            i = i - 1
        End If
        i = i + 1
    Loop Until i > horizontalCount
End If

If toolDeleteVertical = True Then
    i = 0
    Do
        If arrVertical(i, 0) > y0 And arrVertical(i, 0) < y1 And arrVertical(i, 1) > y0 And arrVertical(i, 1) < y1 And arrVertical(i, 2) > x0 And arrVertical(i, 2) < x1 Then
            removeVertical i
            i = i - 1
        End If
        i = i + 1
    Loop Until i > verticalCount
End If

If toolDeleteZone = True And zoneCount > -1 Then
    For i = 0 To zoneCount
        If arrZone(i, 1) > x0 And arrZone(i, 3) < x1 And arrZone(i, 2) < y1 And arrZone(i, 4) > y0 Then
            arrZone(i, 0) = 0
        End If
    Next i
End If

If toolDeleteStar = True Then
    For i = 0 To STAR_MAX
        If arrStar(i, 0) > 0 And arrStar(i, 1) > x0 And arrStar(i, 1) < x1 And arrStar(i, 2) > y0 And arrStar(i, 2) < y1 Then
            arrStar(i, 0) = 0
        End If
    Next i
End If

If toolDeletePUP = True And pupCount > -1 Then
    For i = 0 To pupCount
        If arrPUP(i, 0) > 0 And arrPUP(i, 1) > x0 And arrPUP(i, 1) < x1 And arrPUP(i, 2) > y0 And arrPUP(i, 2) < y1 Then
            arrPUP(i, 0) = 0
        End If
    Next i
End If

End Sub

' This updates the main title bar. To change the format, all you have to change is this:
Public Sub updateMainTitleBar()

Dim totalName As String ' used to build the name
totalName = ""
mdiMain.mnuSave.Enabled = False

If Len(saveFileName) > 0 Then
    totalName = saveFileName & " - "
    mdiMain.mnuSave.Enabled = True
End If

mdiMain.Caption = totalName & "Mosquito's synSpace Map Editor v " & versionNumber

End Sub

