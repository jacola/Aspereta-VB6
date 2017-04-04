Attribute VB_Name = "Declares"
Option Explicit

'Reset variables
Public CountDown As Integer

'********** Public CONSTANTS ***********

Public Const TotalRanks = 10

'Constants for Headings
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'Map sizes
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Tile size in pixels
Public Const TileSizeX = 32
Public Const TileSizeY = 32

'Window size in tiles
Public Const XWindow = 17
Public Const YWindow = 13

'Sound constants
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_WARP = 3

'Object constants
Public Const MAX_INVENTORY_OBJS = 200
Public Const MAX_INVENTORY_SLOTS = 38

'Spell constants
Public Const MAX_SPELL_SLOTS = 30

'Weapon constants
Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOR = 3
Public Const OBJTYPE_ACC = 4
Public Const OBJTYPE_HELM = 5

'Attack constants
Public Const FORWARD = 1
Public Const SIDE = 2
Public Const BACK = 3

'Text type constants
'red~green~blue~bold~italic
Public Const FONTTYPE_WHISPER = "~150~255~255~0~0"
Public Const FONTTYPE_SHOUT = "~255~255~255~0~1"
Public Const FONTTYPE_TALK = "~200~200~200~0~0"
Public Const FONTTYPE_FIGHT = "~100~0~0~1~0"
Public Const FONTTYPE_WARNING = "~255~0~0~0~0"
Public Const FONTTYPE_INFO = "~0~200~0~0~0"


'Stat constants
Public Const STAT_MAXLv = 99
Public Const STAT_MAXHP = 999999
Public Const STAT_MAXSTA = 999
Public Const STAT_MaxMP = 999999
Public Const STAT_MAXHIT = 99
Public Const STAT_MAXAC = 99
Public Const STAT_MAXSTAT = 99     'Max for general stats (MET,FIT, ect)
Public Const STAT_METRATE = 50     'How many server ticks to recover some HP
Public Const STAT_FITRATE = 20     'How many server ticks to recover some STA
Public Const STAT_ATTACKWAIT = 15   'How many server ticks a user has to wait till he can attack again
Public Const STAT_SPELLWAIT = 10

'Other constants
Public Const MAX_CHARACTERS = 10000 'Should be max number users + max NPCs + some head room
Public Const MAX_NPCs = 5000 'How many NPCs are allowed in the game all together
Public Const MAX_NPC_TYPES = 250 'for look up table

'********** Public TYPES ***********

Type Position
    x As Integer
    y As Integer
End Type

Type WorldPos
    map As Integer
    x As Integer
    y As Integer
End Type

'Holds data for a user or NPC character
Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    Heading As Byte
    Weapon As Integer
End Type

'** Object types **
Public Type ObjData
    Name As String
    ObjType As Integer
    GRHIndex As Integer
    Graphic As Integer
    
    MinHIT As Integer
    MaxHIT As Integer
    
    MaxHP As Long
    CurHP As Long
    MaxMP As Long
    CurMP As Long
    AC As Integer
    Dam As Integer
    Str As Integer
    Con As Integer
    Int As Integer
    Wis As Integer
    Dex As Integer
    
    MinLv As Integer
    uPath As String
    
    Body As Integer
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'** Spell type **
Public Type tSpellBook
    Spellindex As Integer
    
    Name As String
    SpellType As Integer
    GRHIndex As Integer
    
    TakeBaseHP As Integer
    TakeBaseMP As Integer
    GiveBaseHP As Integer
    GiveBaseMP As Integer
    TakePercentHP As Integer
    TakePercentMP As Integer
    GivePercentHP As Integer
    GivePercentMP As Integer
    
    Icon As Integer
End Type

Public Type uSpellBook
    Spellindex As Integer
End Type


'** Rank Data **
Type RankEntry
    Name As String
    Lv As Integer
    Stats As Long
    Path As String
End Type

'** User Types **
'Stats for a user
Type UserStats
    'MET As Integer
    MaxHP As Long
    CurHP As Long
    MaxMP As Long
    CurMP As Long
    Lv As Integer
    Exp As Long
    Tnl As Long
    Texp As Long
    AC As Integer
    Dam As Integer
    
    Str As Integer 'hit
    Con As Integer 'vita regen
    Int As Integer 'spell power
    Wis As Integer 'mana regen
    Dex As Integer 'grace
    
    Gold As Long
    
    MaxHIT As Integer
    MinHIT As Integer
End Type

'Flags for a user
Type UserFlags
    UserLogged As Byte 'is the user logged in
    SwitchingMaps As Byte
    DownloadingMap As Byte
    ReadyForNextTile As Byte
    StatsChanged As Byte
    Sound As String
    PK As String
End Type

Type UserCounters
    Regen As Integer
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    SendMapCounter As WorldPos
End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

'Holds data for a user
Type User
    Name As String
    modName As String
    Password As String
    Char As Char 'ACines users looks
    Desc As String
    Path As String
    
    Pos As WorldPos 'Current User Postion
    
    IP As String 'User Ip
    ConnID As Integer 'Connection ID
    RDBuffer As String 'Broken Line Buffer

    PoisonCount As Integer
    PoisonDamage As Integer
    
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    SpellBook(1 To MAX_SPELL_SLOTS) As uSpellBook
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    AccEqpObjIndex As Integer
    AccEqpSlot As Byte
    HelmEqpObjIndex As Integer
    HelmEqpSlot As Byte
    
    Counters As UserCounters
    Stats As UserStats
    Flags As UserFlags
    
    PoisonName As String
    
    GIndex As Integer
    
    SendSpell As Integer
End Type

'** NPC Types **
Type NPCStats
    MaxHP As Long
    CurHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    AC As Integer
End Type

Type NPCFlags
    NPCAlive As Byte  'is the NPC visible (plotted on map)
    NPCActive As Byte 'is the NPC being updated
End Type

Type NPCCounters
    RespawnCounter As Long
    Movement As Integer
End Type

Type NPC
    Name As String
    Char As Char 'ACines NPC looks
    Desc As String
    
    Pos As WorldPos 'Current NPC Postion
    StartPos As WorldPos
    
    Movement As Integer
    RespawnWait As Long
    Attackable As Byte
    Hostile As Byte
    
    GiveExp As Long
    GiveGold As Long
    DropGoldChance As Integer
    DropItem As Integer
    DropChance As Integer
    
    ParaCount As Integer
    PoisonCount As Integer
    PoisonDamage As Integer
    
    CurseName As String
    
    Stats As NPCStats
    Flags As NPCFlags
    Counters As NPCCounters
    
    Speed As Integer
    
    Shop As Integer
    'SendSpell As Integer
End Type

'** NPC Shop type **
Type NPCShopSlot
    Func As String
    FuncName As String
    Gold As Long
    Path As String
    Item As Integer
    Level As Integer
End Type
    

Type NPCShop
    SayCaption As String
    Slots(20) As NPCShopSlot
End Type

'** Map Types **
'Tile Data
Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    userindex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    Gold As Long
End Type

'Map info
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type


'Group type
Type GroupSlot
    UIndexes(1 To 5) As Integer
End Type
    
'********** Public VARS ***********

Public ENDL As String
Public ENDC As String

'Paths
Public IniPath As String
Public CharPath As String
Public MapPath As String

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos 'Ressurect pos
Public StartPos As WorldPos 'Starting Pos (Loaded from Server.ini)


Public NumUsers As Integer 'current Number of Users
Public LastUser As Integer 'current Last User index
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public TotalNumSpells As Integer
Public LogData As String

Public AllowMultiLogins As Byte
Public IdleLimit As Long
Public MaxUsers As Integer
Public HideMe As Byte

'********** Public ARRAYS ***********
Public UserList() As User 'Holds data for each user
Public NPCList(1 To MAX_NPCs) As NPC 'Holds data for each NPC
Public NPCData(1 To MAX_NPC_TYPES) As NPC
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public CharList(1 To MAX_CHARACTERS) As Integer
Public ObjData() As ObjData
Public SpellData() As tSpellBook
Public Groups(1 To (MAX_CHARACTERS / 2)) As GroupSlot
Public NpcShops(1 To MAX_NPCs) As NPCShop
Public PlayerRanking(1 To TotalRanks) As RankEntry

'********** EXTERNAL FUNCTIONS ***********
'APIs to write and read inis
Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpACault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
