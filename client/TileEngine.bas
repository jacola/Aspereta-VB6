Attribute VB_Name = "TileEngine"
Option Explicit

'For limiting FPS
Public OKToDraw As Boolean

'********** CONSTANTS ***********
'Heading Constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Object Constants
Public Const MAX_INVENORY_OBJS = 99

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'Sound flag constants
Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

Public Const NumSoundBuffers = 7

'********** TYPES ***********

'Bitmap header
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Bitmap info header
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Holds a local position
Public Type Position
    X As Integer
    Y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 16) As Integer
    Speed As Integer
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Bodies list
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Weapons list
Public Type WeapData
    Weap(1 To 4) As Grh
End Type

'Heads list
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Hold info about a character
Public Type Char
    Active As Byte
    Heading As Byte
    Pos As Position

    Body As BodyData
    Head As HeadData
    Weap As WeapData
    
    Moving As Byte
    MoveOffset As Position
    
    HpPercent As Integer
    
    Spell As Grh
    SpellCount As Integer
    
    SayText As String
    TextTime As Integer
    Dead As Boolean
End Type

'Holds info about a object
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Holds info about each tile position
Public Type MapBlock
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    Gold As Integer
End Type

'Hold info about each map
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type

'********** Public VARS ***********
'Paths
Public GrhPath As String
Public IniPath As String
Public MapPath As String

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'User status vars
Public CurMap As Integer 'Current map loaded
Public UserIndex As Integer
Public UserMoving As Byte
Global UserBody As Integer
Global UserHead As Integer
Public UserPos As Position 'Holds current user pos
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop As Integer
Public MainViewLeft As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

'Map editor variables
Public WalkMode As Boolean
Public DrawGrid As Boolean
Public DrawBlock As Boolean

'Totals
Public NumMaps As Integer 'Number of maps
Public NumBodies As Integer
Public NumHeads As Integer
Public NumWeapons As Integer
Public NumGrhFiles As Integer 'Number of bmps
Public NumGrhs As Integer 'Number of Grhs
Global NumChars As Integer
Global LastChar As Integer

'********** Direct X ***********
Public MainViewRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer
Public BackBufferRect As RECT

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7
Public SurfaceDB() As DirectDrawSurface7
Public Font As DirectDrawSurface7

'Sound
Dim DirectSound As DirectSound
Dim DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Dim LastSoundBufferUsed As Integer

'********** Public ARRAYS ***********
Public GrhData() As GrhData 'Holds all the grh data

Public BodyData() As BodyData
Public HeadData() As HeadData
Public WeapData() As WeapData

Public MapData() As MapBlock 'Holds map data for current map
Public MapInfo As MapInfo 'Holds map info for current map
Public CharList(1 To 10000) As Char 'Holds info about all characters on map


'********** OUTSIDE FUNCTIONS ***********
'Good old BitBlt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Sound stuff
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub WriteText(X As Integer, Y As Integer, Message As String)
Dim i As Integer
Dim CurChar As String
Dim Grh As Grh

If Message = "" Then Exit Sub

Grh.FrameCounter = 1
Grh.Started = 0


For i = 1 To Len(Message)
    CurChar = Mid(Message, i, 3)
    If CurChar >= "!" And CurChar <= "~" Then
        Grh.GrhIndex = 6500 - Asc("!") + Asc(CurChar)
        Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, X + i * 6, Y, 0, 0)
    End If

Next i

End Sub
Public Sub DrawTalkBox(X As Integer, Y As Integer, Message As String)
'1 row = 150 pixels  half = 75
Dim MLen As Integer
Dim Mess1 As String
Dim Mess2 As String
Dim Mess3 As String
Dim Mess4 As String
Dim Down As Integer
Dim Length As Integer
Dim bWidth As Integer

Dim Grh As Grh

Grh.FrameCounter = 1
Grh.Started = 0
Grh.GrhIndex = 5

If Left(Message, 3) = "   " Then Message = Right(Message, Len(Message) - 3)
If Left(Message, 2) = "  " Then Message = Right(Message, Len(Message) - 2)
If Left(Message, 1) = " " Then Message = Right(Message, Len(Message) - 1)

'Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, X - 63, Y - 75, 0, 0)
'(CurMouseX + 300, CurMouseY + 300 - 3, CurMouseX + 312 + Len(UserSpellbook(Y * 5 + loopC + 1).Name) * 6, CurMouseY + 300 + 13, 5, 5)

MLen = Len(Message)

If MLen <= 25 Then
    Down = 3
    Mess4 = Left(Message, MLen)
    Length = Len(Mess4)
    bWidth = Length
    bWidth = 25 - bWidth
    bWidth = bWidth / 3 * 6
    bWidth = bWidth + 10
End If

If MLen > 25 And MLen <= 50 Then
    Down = 2
    Mess3 = Left(Message, 25)
    Mess4 = Mid(Message, 26, MLen - 25)
    Length = Len(Mess3)
    bWidth = 1000
End If

If MLen > 50 And MLen <= 75 Then
    Down = 1
    Mess2 = Left(Message, 25)
    Mess3 = Mid(Message, 26, 25)
    Mess4 = Mid(Message, 51, MLen - 50)
    Length = Len(Mess2)
    bWidth = 1000
End If

If MLen > 75 Then
    Down = 0
    Mess1 = Left(Message, 25)
    Mess2 = Mid(Message, 26, 25)
    Mess3 = Mid(Message, 51, 25)
    Mess4 = Mid(Message, 76, MLen - 75)
    Length = Len(Mess1)
    bWidth = 1000
End If

If bWidth = 1000 Then
    TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
    TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
    Call TileEngine.BackBufferSurface.DrawRoundedBox(X - 63, Y - 75 + (Down * 11), X - 53 + (6 * Length), Y - 27, 3, 3)
Else
    TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
    TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
    Call TileEngine.BackBufferSurface.DrawRoundedBox(X - 63 + bWidth, Y - 75 + (Down * 11), X - 53 + (6 * Length) + bWidth, Y - 27, 3, 3)
End If

Call WriteText(X - 64, Y - 70, Mess1)
Call WriteText(X - 64, Y - 60, Mess2)
Call WriteText(X - 64, Y - 50, Mess3)
If bWidth = 1000 Then
    Call WriteText(X - 64, Y - 40, Mess4)
Else
    Call WriteText(X - 64 + bWidth, Y - 40, Mess4)
End If

End Sub


Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sfile As String) As Boolean

Dim bufferDesc As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX

If frmConnect.chkSound.value = Unchecked Then Exit Function

bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

waveFormat.nFormatTag = WAVE_FORMAT_PCM
waveFormat.nChannels = 2
waveFormat.lSamplesPerSec = 22050
waveFormat.nBitsPerSample = 16
waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
Set DSB = DS.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)

If Err.Number <> 0 Then
    'MsgBox "unable to find " + sfile
    'End
    Exit Function
End If

LoadWavetoDSBuffer = True
    
End Function
Sub LoadHeadData()
'*****************************************************************
'Loads Head.dat
'*****************************************************************

Dim loopC As Integer

'Get Number of heads
NumHeads = Val(GetVar(IniPath & "Head.dat", "INIT", "NumHeads"))

'Resize array
ReDim HeadData(1 To NumHeads) As HeadData

'Fill List
For loopC = 1 To NumHeads
    InitGrh HeadData(loopC).Head(1), Val(GetVar(IniPath & "Head.dat", "Head" & loopC, "Head1")), 0
    InitGrh HeadData(loopC).Head(2), Val(GetVar(IniPath & "Head.dat", "Head" & loopC, "Head2")), 0
    InitGrh HeadData(loopC).Head(3), Val(GetVar(IniPath & "Head.dat", "Head" & loopC, "Head3")), 0
    InitGrh HeadData(loopC).Head(4), Val(GetVar(IniPath & "Head.dat", "Head" & loopC, "Head4")), 0
Next loopC

End Sub

Sub LoadBodyData()
'*****************************************************************
'Loads Body.dat
'*****************************************************************

Dim loopC As Integer

'Get number of bodies
NumBodies = Val(GetVar(IniPath & "Body.dat", "INIT", "NumBodies"))

'Resize array
ReDim BodyData(1 To NumBodies) As BodyData

'Fill list
For loopC = 1 To NumBodies
    InitGrh BodyData(loopC).Walk(1), Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "Walk1")), 0
    InitGrh BodyData(loopC).Walk(2), Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "Walk2")), 0
    InitGrh BodyData(loopC).Walk(3), Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "Walk3")), 0
    InitGrh BodyData(loopC).Walk(4), Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "Walk4")), 0

    BodyData(loopC).HeadOffset.X = Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "HeadOffsetX"))
    BodyData(loopC).HeadOffset.Y = Val(GetVar(IniPath & "Body.dat", "Body" & loopC, "HeadOffsetY"))

Next loopC

End Sub

Sub LoadWeaponData()
'*****************************************************************
'Loads Weapon.dat
'*****************************************************************

Dim loopC As Integer

'Get number of bodies
NumWeapons = Val(GetVar(IniPath & "Weapon.dat", "INIT", "NumWeapons"))

'Resize array
ReDim WeapData(1 To NumWeapons) As WeapData

'Fill list
For loopC = 1 To NumWeapons
    InitGrh WeapData(loopC).Weap(1), Val(GetVar(IniPath & "Weapon.dat", "Weapon" & loopC, "Weap1")), 0
    InitGrh WeapData(loopC).Weap(2), Val(GetVar(IniPath & "Weapon.dat", "Weapon" & loopC, "Weap2")), 0
    InitGrh WeapData(loopC).Weap(3), Val(GetVar(IniPath & "Weapon.dat", "Weapon" & loopC, "Weap3")), 0
    InitGrh WeapData(loopC).Weap(4), Val(GetVar(IniPath & "Weapon.dat", "Weapon" & loopC, "Weap4")), 0

Next loopC

End Sub

Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tx As Integer, ty As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tx = UserPos.X + CX
ty = UserPos.Y + CY

End Sub

Sub ConvertCPtoSTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tx As Integer, ty As Integer)
'******************************************
'Converts where the user clicks in the main window
'to the Screen tile position for mouse cursor
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = 0 '(WindowTileWidth \ 2)
HWindowY = 0 '(WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

'If CX > HWindowX Then
'    CX = (CX - HWindowX)
'
'Else
'    If CX < HWindowX Then
'        CX = (0 - (HWindowX - CX))
'    Else
'        CX = 0
'    End If
'End If
'
'If CY > HWindowY Then
'    CY = (0 - (HWindowY - CY))
'Else
'    If CY < HWindowY Then
'        CY = (CY - HWindowY)
'    Else
'        CY = 0
'    End If
'End If

tx = CX + 9
ty = CY + 9

End Sub



Function DeInitTileEngine() As Boolean
'*****************************************************************
'Shutsdown engine
'*****************************************************************
Dim loopC As Integer

EngineRun = False

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

'Clear GRH memory
For loopC = 1 To NumGrhFiles
    Set SurfaceDB(loopC) = Nothing
Next loopC
Set DirectDraw = Nothing

'Reset any channels that are done
For loopC = 1 To NumSoundBuffers
    Set DSBuffers(loopC) = Nothing
Next loopC
Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub MakeChar(charindex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer, HpPercent As Integer)
'*****************************************************************
'Makes a new character and puts it on the map
'*****************************************************************

'Update LastChar
If charindex > LastChar Then LastChar = charindex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(charindex).Body = BodyData(Body)
CharList(charindex).Head = HeadData(Head)
CharList(charindex).Heading = Heading

'Reset moving stats
CharList(charindex).Moving = 0
CharList(charindex).MoveOffset.X = 0
CharList(charindex).MoveOffset.Y = 0

'Update position
CharList(charindex).Pos.X = X
CharList(charindex).Pos.Y = Y

'Make active
CharList(charindex).Active = 1

'Plot on map
MapData(X, Y).charindex = charindex

'set hp percent
CharList(charindex).HpPercent = HpPercent

CharList(charindex).Dead = False


End Sub



Sub EraseChar(charindex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

'Make un-active
CharList(charindex).Active = 0

'Update lastchar
If charindex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(CharList(charindex).Pos.X, CharList(charindex).Pos.Y).charindex = 0

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub KillChar(charindex As Integer)

'Exit Sub

CharList(charindex).Dead = True
CharList(charindex).HpPercent = 0

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(charindex As Integer, nHeading As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer
Dim change As Boolean

X = CharList(charindex).Pos.X
Y = CharList(charindex).Pos.Y

If X = iTx And Y = iTy Then change = True

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).charindex = charindex
CharList(charindex).Pos.X = nX
CharList(charindex).Pos.Y = nY
MapData(X, Y).charindex = 0

CharList(charindex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(charindex).Moving = 1
CharList(charindex).Heading = nHeading

If change = True Then
    iTx = nX
    iTy = nY
End If


End Sub

Sub MoveCharbyPos(charindex As Integer, nX As Integer, nY As Integer)
'*****************************************************************
'Starts the movement of a character to nX,nY
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte
Dim change As Boolean

X = CharList(charindex).Pos.X
Y = CharList(charindex).Pos.Y

If X = iTx And Y = iTy Then change = True

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).charindex = charindex
CharList(charindex).Pos.X = nX
CharList(charindex).Pos.Y = nY
MapData(X, Y).charindex = 0

CharList(charindex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(charindex).Moving = 1
CharList(charindex).Heading = nHeading

If change = True Then
    iTx = nX
    iTy = nY
End If


End Sub

Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tx As Integer
Dim ty As Integer

'Figure out which way to move
Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

'Fill temp pos
tx = UserPos.X + X
ty = UserPos.Y + Y

'Check to see if its out of bounds
If tx < MinXBorder Or tx > MaxXBorder Or ty < MinYBorder Or ty > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tx
    AddtoUserPos.Y = Y
    UserPos.Y = ty
    UserMoving = 1
End If

End Sub


Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopC As Integer

loopC = 1
Do While CharList(loopC).Active
    loopC = loopC + 1
Loop

NextOpenChar = loopC

End Function

Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************

Dim loopC As Integer

For loopC = 1 To LastChar
    If CharList(loopC).Active = 1 Then
        MapData(CharList(loopC).Pos.X, CharList(loopC).Pos.Y).charindex = loopC
    End If
Next loopC
    
End Sub
Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer

'Get Number of Graphics
GrhPath = GetVar(IniPath & "Grh.ini", "INIT", "Path")
NumGrhs = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhs"))

'Resize arrays
ReDim GrhData(1 To NumGrhs) As GrhData

'Open files
Open IniPath & "Grh.dat" For Binary As #1
Seek #1, 1

'Get Header
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > NumGrhs Then GoTo ErrorHandler
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Check to see if its out of bounds
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).charindex > 0 Then
    LegalPos = False
    Exit Function
End If

LegalPos = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function
Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

With DestRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
    
Surface.GetSurfaceDesc SurfaceDesc

'Draw

If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
    
End If

End Sub

Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

With DestRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

Sub DrawBackBufferSurface()
'*****************************************************************
'Copies backbuffer to primarysurface
'*****************************************************************
Dim SourceRect As RECT

With SourceRect
    .Left = (TilePixelWidth * TileBufferSize) - TilePixelWidth
    .Top = (TilePixelHeight * TileBufferSize) - TilePixelHeight
    .Right = .Left + MainViewWidth
    .Bottom = .Top + MainViewHeight
End With

PrimarySurface.Blt MainViewRect, BackBufferSurface, SourceRect, DDBLT_WAIT
'PrimarySurface.Flip Nothing, DDFLIP_WAIT


End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1

bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight

End Function



Sub DrawGrhtoHdc(DestHdc As Long, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, ROP As Long)
'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim retcode As Long
Dim CurrentGrh As Grh
Dim SourceHdc As Long


'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

SourceHdc = SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).GetDC

retcode = BitBlt(DestHdc, X, Y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, SourceHdc, GrhData(CurrentGrh.GrhIndex).sX, GrhData(CurrentGrh.GrhIndex).sY, ROP)

SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).ReleaseDC SourceHdc

End Sub

Sub PlayWaveDS(File As String)

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
   
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), File) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Sub PlayWaveAPI(File As String)
'*****************************************************************
'Plays a Wave using windows APIs
'*****************************************************************
'Dim rc As Integer

'rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim Y As Integer    'Keeps track of where on map we are
Dim X As Integer
Dim screenminY As Integer 'Start Y pos on current screen
Dim screenmaxY As Integer 'End Y pos on current screen
Dim screenminX As Integer 'Start X pos on current screen
Dim screenmaxX As Integer 'End X pos on current screen
Dim minY As Integer 'Start Y pos on current screen + tilebuffer
Dim maxY As Integer 'End Y pos on current screen
Dim minX As Integer 'Start X pos on current screen
Dim maxX As Integer 'End X pos on current screen
Dim ScreenX As Integer 'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer
Dim Moved As Byte
Dim Grh As Grh 'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim WeapOffX As Integer
Dim WeapOffY As Integer
Dim loopC As Integer
' for targeting
Dim TargetX As Integer
Dim TargetY As Integer


'Will get the screen if it was lost to an alt tab or something
DirectDraw.RestoreAllSurfaces

'Figure out Ends and Starts of screen
screenminY = (TileY - (WindowTileHeight \ 2))
screenmaxY = (TileY + (WindowTileHeight \ 2))
screenminX = (TileX - (WindowTileWidth \ 2))
screenmaxX = (TileX + (WindowTileWidth \ 2))

minY = screenminY - TileBufferSize
maxY = screenmaxY + TileBufferSize
minX = screenminX - TileBufferSize
maxX = screenmaxX + TileBufferSize

'Draw floor layer
ScreenY = 0
For Y = screenminY - 1 To screenmaxY + 1
    ScreenX = 0
    For X = screenminX - 1 To screenmaxX + 1
        'Check to see if in bounds
        If InMapBounds(X, Y) Then
            'Layer 1 **********************************
            
            PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth)
            PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight)
            
            'Draw
            Call DDrawGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 0, 1)
            '**********************************
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex > 0 Then
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            End If
            '**********************************
        Else
            Call DDrawGrhtoSurface(BackBufferSurface, MapData(1, 1).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 0, 1)
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y


'Draw transparent layers
ScreenY = 0
For Y = minY To maxY
    ScreenX = 0
    For X = minX To maxX
        'Check to see if in bounds
        If InMapBounds(X, Y) Then
            'Gold Layer ************************************
            If MapData(X, Y).Gold > 0 Then
                Grh.GrhIndex = 6865
                Grh.FrameCounter = 1
                Grh.Started = 0
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
            
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
            End If
            'Object Layer **********************************
            If MapData(X, Y).ObjGrh.GrhIndex > 0 Then
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            End If
            '**********************************
             'Char layer **********************************
            If MapData(X, Y).charindex > 0 Then
                TempChar = CharList(MapData(X, Y).charindex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
                'If needed, move left and right
                If TempChar.MoveOffset.X <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        TempChar.Head.Head(TempChar.Heading).Started = 1
                        PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                        TempChar.MoveOffset.X = TempChar.MoveOffset.X - (ScrollPixelsPerFrameX * Sgn(TempChar.MoveOffset.X))
                        Moved = 1
                End If
                'If needed, move up and down
                If TempChar.MoveOffset.Y <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        TempChar.Head.Head(TempChar.Heading).Started = 1
                        PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                        TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (ScrollPixelsPerFrameY * Sgn(TempChar.MoveOffset.Y))
                        Moved = 1
                End If
                'If done moving stop animation
                If Moved = 0 And TempChar.Moving = 1 Then
                    TempChar.Moving = 0
                    TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Body.Walk(TempChar.Heading).Started = 0
                    TempChar.Head.Head(TempChar.Heading).FrameCounter = 1
                    TempChar.Head.Head(TempChar.Heading).Started = 0
                End If
                
                'Draw Body
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPosX(ScreenX) + PixelOffsetXTemp), PixelPosY(ScreenY) + PixelOffsetYTemp, 1, 1)
                'Draw Head
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Head.Head(TempChar.Heading), (PixelPosX(ScreenX) + PixelOffsetXTemp), PixelPosY(ScreenY) + PixelOffsetYTemp, 1, 1)
                    
                WeapOffX = 0
                WeapOffY = 0
                
                If TempChar.SpellCount > 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Spell, (PixelPosX(ScreenX) + PixelOffsetXTemp), PixelPosY(ScreenY) + PixelOffsetYTemp, 1, 1)
                    TempChar.SpellCount = TempChar.SpellCount - 1
                    If TempChar.Spell.FrameCounter > 5 Then
                        TempChar.SpellCount = 0
                    End If
                End If
                CharList(MapData(X, Y).charindex) = TempChar
            End If
            '**********************************
            'Layer 3 **********************************
            If MapData(X, Y).Graphic(3).GrhIndex > 0 Then
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            End If
            '**********************************
        End If
        ScreenX = ScreenX + 1
    Next X
    
    ScreenX = 0
    For X = minX To maxX
        'Check to see if in bounds
        If InMapBounds(X, Y) Then
            PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
            PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex > 0 Then
                        
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                
            End If
            '**********************************
        End If
        '*** Target *********************
        If X = iTx And Y = iTy And Targeting = True Then
            'Grh.GrhIndex = 2
            'Grh.FrameCounter = 1
            'Grh.Started = 0
            If MapData(X, Y).charindex > 0 Then
                TargetX = GrhData(CharList(MapData(X, Y).charindex).Body.Walk(CharList(MapData(X, Y).charindex).Heading).GrhIndex).pixelWidth
                TargetY = GrhData(CharList(MapData(X, Y).charindex).Body.Walk(CharList(MapData(X, Y).charindex).Heading).GrhIndex).pixelHeight
                
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                'left, top, bottom, right
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp, PixelOffsetYTemp - TargetY + 32, PixelOffsetXTemp, PixelOffsetYTemp + 32)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp, PixelOffsetYTemp - TargetY + 32, PixelOffsetXTemp + ((TargetX / 2) + 16), PixelOffsetYTemp - TargetY + 32)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp, PixelOffsetYTemp + 32, PixelOffsetXTemp + ((TargetX / 2) + 16) + 1, PixelOffsetYTemp + 32)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + ((TargetX / 2) + 16), PixelOffsetYTemp - TargetY + 32, PixelOffsetXTemp + ((TargetX / 2) + 16), PixelOffsetYTemp + 32)
                
                TileEngine.BackBufferSurface.SetForeColor RGB(0, 100, 255)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 1, PixelOffsetYTemp - TargetY + 32 + 1, PixelOffsetXTemp + 1, PixelOffsetYTemp + 32 - 1)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 1, PixelOffsetYTemp - TargetY + 32 + 1, PixelOffsetXTemp + ((TargetX / 2) + 16) - 1, PixelOffsetYTemp - TargetY + 32 + 1)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 1, PixelOffsetYTemp + 32 - 1, PixelOffsetXTemp + ((TargetX / 2) + 16) - 1 + 1, PixelOffsetYTemp + 32 - 1)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + ((TargetX / 2) + 16) - 1, PixelOffsetYTemp - TargetY + 32 + 1, PixelOffsetXTemp + ((TargetX / 2) + 16) - 1, PixelOffsetYTemp + 32 - 1)
                
                TileEngine.BackBufferSurface.SetForeColor RGB(0, 100, 255)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 2, PixelOffsetYTemp - TargetY + 32 + 2, PixelOffsetXTemp + 2, PixelOffsetYTemp + 32 - 2)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 2, PixelOffsetYTemp - TargetY + 32 + 2, PixelOffsetXTemp + ((TargetX / 2) + 16) - 2, PixelOffsetYTemp - TargetY + 32 + 2)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 2, PixelOffsetYTemp + 32 - 2, PixelOffsetXTemp + ((TargetX / 2) + 16) - 2 + 1, PixelOffsetYTemp + 32 - 2)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + ((TargetX / 2) + 16) - 2, PixelOffsetYTemp - TargetY + 32 + 2, PixelOffsetXTemp + ((TargetX / 2) + 16) - 2, PixelOffsetYTemp + 32 - 2)
                
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 3, PixelOffsetYTemp - TargetY + 32 + 3, PixelOffsetXTemp + 3, PixelOffsetYTemp + 32 - 3)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 3, PixelOffsetYTemp - TargetY + 32 + 3, PixelOffsetXTemp + ((TargetX / 2) + 16) - 3, PixelOffsetYTemp - TargetY + 32 + 3)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + 3, PixelOffsetYTemp + 32 - 3, PixelOffsetXTemp + ((TargetX / 2) + 16) - 3 + 1, PixelOffsetYTemp + 32 - 3)
                Call TileEngine.BackBufferSurface.DrawLine(PixelOffsetXTemp + ((TargetX / 2) + 16) - 3, PixelOffsetYTemp - TargetY + 32 + 3, PixelOffsetXTemp + ((TargetX / 2) + 16) - 3, PixelOffsetYTemp + 32 - 3)
            End If
        End If
        '************
        ScreenX = ScreenX + 1
    Next X

    ScreenX = 0
    For X = minX To maxX
        'Check to see if in bounds
        If InMapBounds(X, Y) Then

            
             'Junk**********************************
            If MapData(X, Y).charindex > 0 Then
            
                TempChar = CharList(MapData(X, Y).charindex)
            
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                
                PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                
                'Draw a vita bar for the character
                If TempChar.Dead = True Then
                    Grh.GrhIndex = 6006
                ElseIf TempChar.HpPercent >= 98 Then
                    Grh.GrhIndex = 6000
                ElseIf TempChar.HpPercent >= 80 And TempChar.HpPercent < 98 Then
                    Grh.GrhIndex = 6001
                ElseIf TempChar.HpPercent >= 60 And TempChar.HpPercent < 80 Then
                    Grh.GrhIndex = 6002
                ElseIf TempChar.HpPercent >= 40 And TempChar.HpPercent < 60 Then
                    Grh.GrhIndex = 6003
                ElseIf TempChar.HpPercent >= 15 And TempChar.HpPercent < 40 Then
                    Grh.GrhIndex = 6004
                ElseIf TempChar.HpPercent < 15 Then
                    Grh.GrhIndex = 6005
                End If
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, (PixelPosX(ScreenX) + PixelOffsetXTemp) + TempChar.Body.HeadOffset.X, PixelPosY(ScreenY) + PixelOffsetYTemp + TempChar.Body.HeadOffset.Y - 5, 0, 0)
                'End drawing vita bar
                
                'text over the head
                If TempChar.TextTime > 0 Then
                    Call DrawTalkBox((PixelPosX(ScreenX) + PixelOffsetXTemp), PixelPosY(ScreenY) + PixelOffsetYTemp - TempChar.Body.HeadOffset.Y, TempChar.SayText)
                    TempChar.TextTime = TempChar.TextTime - 1
                End If
                'End drawing text over head
                
                'Refresh charlist
                CharList(MapData(X, Y).charindex) = TempChar
                
            End If
            '**********************************
        End If
    
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
TileEngine.BackBufferSurface.SetFillColor RGB(5, 5, 5)
If TextBoxOn = True Or TextBoxAlwaysOn = True Then Call TileEngine.BackBufferSurface.DrawRoundedBox(305, 680, 940, 768, 5, 5)
If frmMain.StatusBox.Visible = True Then Call TileEngine.BackBufferSurface.DrawRoundedBox(305, 590, 517, 677, 5, 5)

'** 'Paper doll' ********
If ShowStatus = True Then
    Call TileEngine.BackBufferSurface.DrawRoundedBox(500, 355, 700, 644, 5, 5)
    'Call WriteText(500, 405, "Paper doll")
    Grh.FrameCounter = 1
    Grh.Started = 0
    Grh.GrhIndex = 6
    For X = 0 To 7
        Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 505, 360 + X * 35, 0, 0)
        If UserInventory(31 + X).OBJIndex > 0 Then
            Grh.GrhIndex = UserInventory(31 + X).GrhIndex
            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 505, 360 + X * 35, 0, 0)
            Grh.GrhIndex = 6
        End If
        Call WriteText(540, 367 + X * 35, PaperDollList(X + 1) + ":")
        Call WriteText(545, 380 + X * 35, UserInventory(31 + X).Name)
        'Call WriteText(CurMouseX + 300 - Len(UserInventory(Y * 5 + loopC + 1).Name) * 6, CurMouseY + 300, UserInventory(Y * 5 + loopC + 1).Name)
    Next X
    
    
End If


'** Spell list **********
If ShowSpells = True Then
    For Y = 0 To 5
        For loopC = 0 To 4
            Grh.FrameCounter = 1
            Grh.Started = 0
            Grh.GrhIndex = 6
            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 305 + loopC * 36, 311 + Y * 36, 0, 0)
            
            If Y * 5 + loopC <= 30 Then
                Grh = UserSpellbook(Y * 5 + loopC + 1).Icon
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 306 + loopC * 36, 312 + Y * 36, 0, 0)
            End If
        Next loopC
        For loopC = 0 To 4
            If CurMouseX >= (loopC * 36 + 2) And CurMouseY >= (Y * 36) And CurMouseX <= (loopC * 36 + 36) And CurMouseY <= (Y * 36 + 34) And UserSpellbook(Y * 5 + loopC + 1).Name <> "" Then
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
                Call TileEngine.BackBufferSurface.DrawRoundedBox(CurMouseX + 300, CurMouseY + 300 - 3, CurMouseX + 312 + Len(UserSpellbook(Y * 5 + loopC + 1).Name) * 6, CurMouseY + 300 + 13, 5, 5)
                Call WriteText(CurMouseX + 300, CurMouseY + 300, UserSpellbook(Y * 5 + loopC + 1).Name)
            End If
        Next loopC
    Next Y
    
End If

    


'** Hot items ***********
For loopC = 0 To 9

    Grh.FrameCounter = 1
    Grh.Started = 0
    Grh.GrhIndex = 6
    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 583 + loopC * 36, 311, 0, 0)
    
    If UserHotButtons(loopC + 1) >= 1 And UserHotButtons(loopC + 1) <= 100 Then
        Grh.GrhIndex = UserInventory(UserHotButtons(loopC + 1)).GrhIndex
        Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 583 + loopC * 36, 311, 0, 0)
        Call WriteText(575 + loopC * 36, 314, Str(UserInventory(UserHotButtons(loopC + 1)).Amount))
        If UserInventory(UserHotButtons(loopC + 1)).Equipped = 1 Then
            Call WriteText(603 + loopC * 36, 334, "E")
        End If
    End If
    
    If UserHotButtons(loopC + 1) >= 101 And UserHotButtons(loopC + 1) <= 200 Then
        Grh.GrhIndex = UserSpellbook(UserHotButtons(loopC + 1) - 100).Icon.GrhIndex
        Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 583 + loopC * 36, 311, 0, 0)
    End If
Next loopC
'(names on mouse over for hot items)
For loopC = 0 To 9
    If CurMouseX >= (loopC * 36 + 281) And CurMouseY >= (2) And CurMouseX <= (loopC * 36 + 315) And CurMouseY <= 36 Then
        If UserHotButtons(loopC + 1) >= 1 And UserHotButtons(loopC + 1) <= 100 Then
            'If UserInventory(UserHotButtons(loopC + 1)).Name <> "(None)" Then
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
                Call TileEngine.BackBufferSurface.DrawRoundedBox(CurMouseX + 300 - Len(UserInventory(UserHotButtons(loopC + 1)).Name) * 6, CurMouseY + 300 - 3, CurMouseX + 312, CurMouseY + 300 + 13, 5, 5)
                Call WriteText(CurMouseX + 300 - Len(UserInventory(UserHotButtons(loopC + 1)).Name) * 6, CurMouseY + 300, UserInventory(UserHotButtons(loopC + 1)).Name)
            'End If
        End If
        If UserHotButtons(loopC + 1) >= 101 And UserHotButtons(loopC + 1) <= 200 Then
            If UserSpellbook(UserHotButtons(loopC + 1) - 100).Name <> "" Then
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
                Call TileEngine.BackBufferSurface.DrawRoundedBox(CurMouseX + 300 - Len(UserSpellbook(UserHotButtons(loopC + 1) - 100).Name) * 6, CurMouseY + 300 - 3, CurMouseX + 312, CurMouseY + 300 + 13, 5, 5)
                Call WriteText(CurMouseX + 300 - Len(UserSpellbook(UserHotButtons(loopC + 1) - 100).Name) * 6, CurMouseY + 300, UserSpellbook(UserHotButtons(loopC + 1) - 100).Name)
            End If
        End If
    End If
Next loopC


'** Inventory ***********
If ShowInventory = True Then
    For Y = 0 To 5
        For loopC = 0 To 4
            Grh.FrameCounter = 1
            Grh.Started = 0
            Grh.GrhIndex = 6
            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 763 + loopC * 36, 370 + Y * 36, 0, 0)
            
            If Y * 5 + loopC <= 30 Then
                Grh.GrhIndex = UserInventory(Y * 5 + loopC + 1).GrhIndex
                If UserInventory(Y * 5 + loopC + 1).Amount > 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 763 + loopC * 36, 370 + Y * 36, 0, 0)
                    Call WriteText(755 + loopC * 36, 372 + Y * 36, Str(UserInventory(Y * 5 + loopC + 1).Amount))
                    If UserInventory(Y * 5 + loopC + 1).Equipped = 1 Then
                        Call WriteText(783 + loopC * 36, 393 + Y * 36, "E")
                        
                    End If
                End If
            End If
        Next loopC
        For loopC = 0 To 4
            If CurMouseX >= (loopC * 36 + 460) And CurMouseY >= (Y * 36 + 61) And CurMouseX <= (loopC * 36 + 494) And CurMouseY <= (Y * 36 + 95) And UserInventory(Y * 5 + loopC + 1).Name <> "(None)" Then
                TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255)
                TileEngine.BackBufferSurface.SetFillColor RGB(1, 1, 1)
                Call TileEngine.BackBufferSurface.DrawRoundedBox(CurMouseX + 300 - Len(UserInventory(Y * 5 + loopC + 1).Name) * 6, CurMouseY + 300 - 3, CurMouseX + 312, CurMouseY + 300 + 13, 5, 5)
                Call WriteText(CurMouseX + 300 - Len(UserInventory(Y * 5 + loopC + 1).Name) * 6, CurMouseY + 300, UserInventory(Y * 5 + loopC + 1).Name)
            End If
        Next loopC
    Next Y
End If


'** Chat text ***********
For loopC = 1 To 7

    Call WriteText(305, 671 + loopC * 12, ReadField(1, ChatText(33 + loopC + ChatPos), 126))
Next loopC

'6450
'** OK BOX **************
If ShowOKBox = True Then
    Call TileEngine.BackBufferSurface.DrawRoundedBox(350, 400, 700, 575, 5, 5)
    For loopC = 1 To 10
        Call WriteText(356, 396 + loopC * 12, OKBoxText(loopC + OkBoxPos))
    Next loopC
    
    Grh.FrameCounter = 1
    Grh.Started = 0
    Grh.GrhIndex = 6450
    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 645, 545, 0, 0)
    Grh.GrhIndex = 6452
    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 675, 410, 0, 0)
    Grh.GrhIndex = 6451
    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, 675, 524, 0, 0)
    
    TileEngine.BackBufferSurface.SetForeColor RGB(5, 5, 5)
    TileEngine.BackBufferSurface.DrawText 659, 549, "OK", False
End If
'************************

Grh.GrhIndex = DragIndex
Grh.FrameCounter = 1
Grh.Started = 0
Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, CurMouseX + 300, CurMouseY + 300, 0, 0)

'delete chars that are dead now
For loopC = 1 To LastChar
If (CharList(loopC).Dead = True) And (CharList(loopC).SpellCount <= 0) Then
        EraseChar (loopC)
    End If
Next loopC

'TileEngine.BackBufferSurface.SetForeColor RGB(255, 255, 255) 'white

End Sub

Function PixelPosX(X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

PixelPosX = (TilePixelWidth * X) - TilePixelWidth

End Function

Function PixelPosY(Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

PixelPosY = (TilePixelHeight * Y) - TilePixelHeight

End Function
Sub LoadGraphics()
'*****************************************************************
'Loads all the sprites and tiles from the gif or bmp files
'*****************************************************************
Dim loopC As Integer
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim ddsd As DDSURFACEDESC2

NumGrhFiles = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhFiles"))
ReDim SurfaceDB(1 To NumGrhFiles)



'Load the GRHx.bmps into memory
For loopC = 1 To NumGrhFiles

    If FileExist(App.Path & GrhPath & "" & loopC & ".bmp", vbNormal) Then
        
        With ddsd
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End With
        
        GetBitmapDimensions App.Path & GrhPath & "" & loopC & ".bmp", ddsd.lWidth, ddsd.lHeight
        
        Set SurfaceDB(loopC) = DirectDraw.CreateSurfaceFromFile(App.Path & GrhPath & "" & loopC & ".bmp", ddsd)
        'Set color key
        ddck.low = 0
        ddck.high = 0
        SurfaceDB(loopC).SetColorKey DDCKEY_SRCBLT, ddck
    End If
 
Next loopC

End Sub

Public Sub ChangeRes(TargetFrm As Form, Width As Integer, Height As Integer)

DirectDraw.SetCooperativeLevel TargetFrm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT

DirectDraw.SetDisplayMode Width, Height, 16, 0, 0

End Sub



Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\"

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

'Setup borders
MinXBorder = XMinMapSize + 8 '(WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - 8 '(WindowTileWidth \ 2)
MinYBorder = YMinMapSize + 6 '(WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - 6 '(WindowTileHeight \ 2)

MainViewWidth = TilePixelWidth * WindowTileWidth
MainViewHeight = TilePixelHeight * WindowTileHeight

'Resize mapdata array
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Set scroll pixels per frame
ScrollPixelsPerFrameX = 2
ScrollPixelsPerFrameY = 2

'****** INIT DirectDraw ******
' Create the root DirectDraw objec
Set DirectDraw = DirectX.DirectDrawCreate("")

DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'backbuffer.GetSurfaceDesc ddsd3

'Call ChangeRes(DisplayFormhWnd, 800, 600)

'DirectDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_ALLOWREBOOT Or DDSCL_MULTITHREADED Or DDSCL_EXCLUSIVE

If frmConnect.chkFullScrn = Checked Then
    Call ChangeRes(frmMain, 640, 480)
End If


'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
' Create the surface
Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

'Create Primary Clipper
Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

'Back Buffer Surface
With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))
    .Bottom = TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

' Create surface
Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

'Set color key
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

'Load graphic data into memory
Call LoadGrhData
Call LoadBodyData
Call LoadHeadData
Call LoadMapData
Call LoadGraphics
Call LoadWeaponData

If frmConnect.chkSound.value = Checked Then
    'Wave Sound
    Set DirectSound = DirectX.DirectSoundCreate("")
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
    LastSoundBufferUsed = 1
End If

InitTileEngine = True
EngineRun = True

End Function

Sub LoadMapData()
'*****************************************************************
'Load Map.dat
'*****************************************************************

'Get Number of Maps
NumMaps = Val(GetVar(IniPath & "Map.dat", "INIT", "NumMaps"))
MapPath = GetVar(IniPath & "Map.dat", "INIT", "MapPath")
iTx = 0
iTy = 0

End Sub
Sub ShowNextFrame(DisplayFormTop As Integer, DisplayFormLeft As Integer)
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer

    '****** Set main view rectangle ******
    With MainViewRect
        .Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With

    '***** Check if engine is allowed to run ******
    If EngineRun And OKToDraw = True Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (ScrollPixelsPerFrameX * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            End If

            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (ScrollPixelsPerFrameY * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If

            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            DrawBackBufferSurface
            FramesPerSecCounter = FramesPerSecCounter + 1
            OKToDraw = False
    End If
End Sub

Sub MakeGold(ByVal X As Integer, ByVal Y As Integer)
    MapData(X, Y).Gold = 1
End Sub

Sub KillGold(ByVal X As Integer, ByVal Y As Integer)
    MapData(X, Y).Gold = 0
End Sub

