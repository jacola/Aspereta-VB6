Attribute VB_Name = "Declares"
Option Explicit

'The OK box OK button toggle.
Public ShowOKBox As Boolean
'The Ok Box Position
Public OkBoxPos As Integer

'Stuff for boxes behind text
Public ShowStatus As Boolean
Public TextBoxOn As Boolean
Public TextBoxAlwaysOn As Boolean
Public StatusFilter As Boolean
Public ShowInventory As Boolean
Public PaperDollList(8) As String

Public CurMouseX As Integer
Public CurMouseY As Integer
Public DragIndex As Integer

'Target
Public iTx As Integer
Public iTy As Integer
Public LastPX As Integer
Public LastPY As Integer
Public Targeting As Boolean

Public ClientVer As String

'Mouse
Public pMouseX As Integer
Public pMouseY As Integer

'Object constants
Public Const MAX_INVENTORY_OBJS = 99
Public Const MAX_INVENTORY_SLOTS = 38

'User's inventory
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

'** Spell type **
Public Type tSpellBook
    SpellIndex As Integer
    Name As String
    SpellType As Integer
    GrhIndex As Integer
    Targetable As String
    Icon As Grh
End Type

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Public UserSpellbook(1 To MAX_INVENTORY_SLOTS) As tSpellBook
Public UserHotButtons(10) As Integer

Public UserName As String
Public UserPassword As String
Public UserMaxHP As Long
Public UserCurHP As Long
Public UserMaxMP As Long
Public UserCurMP As Long
Public UserLv As Integer
Public UserExp As Long
Public UserTnl As Long
Public UserTexp As Long
Public UserAC As Integer
Public UserDam As Integer
Public UserPath As Integer

Public UserStr As Integer
Public UserCon As Integer
Public UserInt As Integer
Public UserWis As Integer
Public UserDex As Integer

Public UserGold As Long
Public UserPort As Integer
Public UserServerIP As String

Public UserDir As Integer

Public CurrentNPCShop As Integer

Public CurSpellIndex

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server

'String contants
Public ENDC As String 'Endline character for talking with server
Public ENDL As String 'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

'Music stuff
Public CurMidi As String 'Keeps current MIDI file
Public LoopMidi As Byte 'If 1 current MIDI is looped

Public ChatText(1 To 40) As String
Public OKBoxText(1 To 100) As String

Public ChatPos As Integer
Public ShowSpells As Boolean

'********** OUTSIDE FUNCTIONS ***********

'For Get and Write Var
Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
