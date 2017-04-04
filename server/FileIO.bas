Attribute VB_Name = "FileIO"
Option Explicit

Sub LoadOBJData()
'*****************************************************************
'Setup OBJ list
'*****************************************************************
Dim Object As Integer

'Get Number of Objects
NumObjDatas = Val(GetVar(IniPath & "Obj.dat", "INIT", "NumObjs"))
ReDim ObjData(1 To NumObjDatas) As ObjData
  
'Fill Object List
For Object = 1 To NumObjDatas
    
    ObjData(Object).Name = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Name")
    
    ObjData(Object).GRHIndex = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
    ObjData(Object).Graphic = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Graphic"))
    
    ObjData(Object).ObjType = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "ObjType"))

    
    ObjData(Object).MaxHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
    ObjData(Object).MinHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
    
    
    ObjData(Object).MaxHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
    ObjData(Object).CurHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "CurHP"))
    ObjData(Object).MaxMP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxMP"))
    ObjData(Object).CurMP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "CurMP"))
    
    ObjData(Object).AC = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "AC"))
    
    ObjData(Object).Dam = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Dam"))
    
    ObjData(Object).Str = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Str"))
    ObjData(Object).Con = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Con"))
    ObjData(Object).Int = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Int"))
    ObjData(Object).Wis = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Wis"))
    ObjData(Object).Dex = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Dex"))
    
    ObjData(Object).MinLv = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinLv"))
    ObjData(Object).uPath = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Path")
    ObjData(Object).Body = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Body"))
    
    LogData = "Object " & ObjData(Object).Name & " loaded at " & Time$ & FONTTYPE_TALK
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
    
Next Object

End Sub

Sub LoadSpellData()
'*****************************************************************
'Setup OBJ list
'*****************************************************************
Dim Object As Integer

'Get Number of Objects
TotalNumSpells = Val(GetVar(IniPath & "Spell.dat", "INIT", "NumSpells"))
ReDim SpellData(1 To TotalNumSpells) As tSpellBook
  
'Fill Object List
For Object = 1 To TotalNumSpells
    
    SpellData(Object).Name = GetVar(IniPath & "Spell.dat", "OBJ" & Object, "Name")
    
    SpellData(Object).GRHIndex = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "GrhIndex"))
    
    SpellData(Object).SpellType = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "SpellType"))

    SpellData(Object).TakeBaseHP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "TakeBaseHP"))
    SpellData(Object).TakeBaseMP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "TakeBaseMP"))
    SpellData(Object).GiveBaseHP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "GiveBaseHP"))
    SpellData(Object).GiveBaseMP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "GiveBaseMP"))
    SpellData(Object).TakePercentHP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "TakePercentHP"))
    SpellData(Object).TakePercentMP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "TakePercentMP"))
    SpellData(Object).GivePercentHP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "GivePercentHP"))
    SpellData(Object).GivePercentMP = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "GivePercentMP"))
    
    SpellData(Object).Icon = Val(GetVar(IniPath & "Spell.dat", "OBJ" & Object, "Icon"))
    
    LogData = "Spell " & SpellData(Object).Name & " loaded at " & Time$ & FONTTYPE_TALK
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Next Object

End Sub

Sub LoadUserStats(userindex As Integer, UserFile As String)
'*****************************************************************
'Loads a user's stats from a text file
'*****************************************************************
UserList(userindex).Stats.MaxHP = Val(GetVar(UserFile, "STATS", "MaxHP"))
UserList(userindex).Stats.CurHP = Val(GetVar(UserFile, "STATS", "CurHP"))
UserList(userindex).Stats.MaxMP = Val(GetVar(UserFile, "STATS", "MaxMP"))
UserList(userindex).Stats.CurMP = Val(GetVar(UserFile, "STATS", "CurMP"))
UserList(userindex).Stats.Lv = Val(GetVar(UserFile, "STATS", "Lv"))
UserList(userindex).Stats.Exp = Val(GetVar(UserFile, "STATS", "Exp"))
UserList(userindex).Stats.Tnl = Val(GetVar(UserFile, "STATS", "tnl"))
UserList(userindex).Stats.Texp = Val(GetVar(UserFile, "STATS", "Texp"))
UserList(userindex).Stats.AC = Val(GetVar(UserFile, "STATS", "AC"))
UserList(userindex).Stats.Dam = Val(GetVar(UserFile, "STATS", "Dam"))

UserList(userindex).Stats.Str = Val(GetVar(UserFile, "STATS", "Str"))
UserList(userindex).Stats.Con = Val(GetVar(UserFile, "STATS", "Con"))
UserList(userindex).Stats.Int = Val(GetVar(UserFile, "STATS", "Int"))
UserList(userindex).Stats.Wis = Val(GetVar(UserFile, "STATS", "Wis"))
UserList(userindex).Stats.Dex = Val(GetVar(UserFile, "STATS", "Dex"))

UserList(userindex).Stats.Gold = Val(GetVar(UserFile, "STATS", "Gold"))

UserList(userindex).Stats.MaxHIT = Val(GetVar(UserFile, "STATS", "MaxHIT"))
UserList(userindex).Stats.MinHIT = Val(GetVar(UserFile, "STATS", "MinHIT"))

End Sub

Sub LoadUserInit(userindex As Integer, UserFile As String)
'*****************************************************************
'Loads the user's Init stuff
'*****************************************************************

Dim LoopC As Integer
Dim ln As String

'Get INIT
UserList(userindex).Char.Heading = Val(GetVar(UserFile, "INIT", "Heading"))
UserList(userindex).Char.Head = Val(GetVar(UserFile, "INIT", "Head"))
UserList(userindex).Char.Body = Val(GetVar(UserFile, "INIT", "Body"))
UserList(userindex).Char.Weapon = Val(GetVar(UserFile, "INIT", "Weapon"))
UserList(userindex).Desc = GetVar(UserFile, "INIT", "Desc")
UserList(userindex).Path = GetVar(UserFile, "INIT", "Path")

'Get last postion
UserList(userindex).Pos.map = Val(ReadField(1, GetVar(UserFile, "INIT", "Position"), 45))
UserList(userindex).Pos.x = Val(ReadField(2, GetVar(UserFile, "INIT", "Position"), 45))
UserList(userindex).Pos.y = Val(ReadField(3, GetVar(UserFile, "INIT", "Position"), 45))

'Get object list
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(UserFile, "Inventory", "Obj" & LoopC)
    UserList(userindex).Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
    UserList(userindex).Object(LoopC).Amount = Val(ReadField(2, ln, 45))
    UserList(userindex).Object(LoopC).Equipped = Val(ReadField(3, ln, 45))
Next LoopC

'Get spell list
For LoopC = 1 To MAX_SPELL_SLOTS
    ln = GetVar(UserFile, "Spells", "Spl" & LoopC)
    UserList(userindex).SpellBook(LoopC).Spellindex = Val(ReadField(1, ln, 45))
Next LoopC

'Get Weapon objectindex and slot
UserList(userindex).WeaponEqpSlot = Val(GetVar(UserFile, "Inventory", "WeaponEqpSlot"))
If UserList(userindex).WeaponEqpSlot > 0 Then
    UserList(userindex).WeaponEqpObjIndex = UserList(userindex).Object(UserList(userindex).WeaponEqpSlot).ObjIndex
End If

'Get Armour objectindex and slot
UserList(userindex).ArmourEqpSlot = Val(GetVar(UserFile, "Inventory", "ArmourEqpSlot"))
If UserList(userindex).ArmourEqpSlot > 0 Then
    UserList(userindex).ArmourEqpObjIndex = UserList(userindex).Object(UserList(userindex).ArmourEqpSlot).ObjIndex
End If

UserList(userindex).AccEqpSlot = Val(GetVar(UserFile, "Inventory", "AccEqpSlot"))
If UserList(userindex).AccEqpSlot > 0 Then
    UserList(userindex).AccEqpObjIndex = UserList(userindex).Object(UserList(userindex).AccEqpSlot).ObjIndex
End If

UserList(userindex).HelmEqpSlot = Val(GetVar(UserFile, "Inventory", "HelmEqpSlot"))
If UserList(userindex).HelmEqpSlot > 0 Then
    UserList(userindex).HelmEqpObjIndex = UserList(userindex).Object(UserList(userindex).HelmEqpSlot).ObjIndex
End If

UserList(userindex).Flags.Sound = GetVar(UserFile, "FLAGS", "Sound")
UserList(userindex).Flags.PK = GetVar(UserFile, "FLAGS", "PK")

If UserList(userindex).Flags.Sound = "" Then UserList(userindex).Flags.Sound = "on"
If UserList(userindex).Flags.PK = "" Then UserList(userindex).Flags.PK = "off"

End Sub

Function WizCheck(Name As String) As Boolean
'*****************************************************************
'Checks to see if Name is a wizard
'*****************************************************************
Dim NumWizs As Integer
Dim WizNum As Integer

NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "NumWizs"))
For WizNum = 1 To NumWizs
    If UCase(Name) = UCase(GetVar(IniPath & "Server.ini", "WizList", "wiz" & WizNum)) Then
        WizCheck = True
        Exit Function
    End If
Next WizNum

WizCheck = False

End Function

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a variable from a text file
'*****************************************************************
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the ACaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
  
  
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
  
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
  
End Function

Function LoadRankData()

Dim iLoop As Integer
Dim RankFile As String

RankFile = IniPath & "Rank.dat"

For iLoop = 1 To TotalRanks
PlayerRanking(iLoop).Name = GetVar(RankFile, "Player" & iLoop, "Name")
PlayerRanking(iLoop).Path = GetVar(RankFile, "Player" & iLoop, "Path")
PlayerRanking(iLoop).Lv = Val(GetVar(RankFile, "Player" & iLoop, "Level"))
PlayerRanking(iLoop).Stats = Val(GetVar(RankFile, "Player" & iLoop, "Stat"))

LogData = "Player " & PlayerRanking(iLoop).Name & " loaded at " & Time$ & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


Next iLoop

End Function


Function SaveRankData()

Dim iLoop As Integer
Dim RankFile As String

RankFile = IniPath & "Rank.dat"

For iLoop = 1 To TotalRanks

Call WriteVar(RankFile, "Player" & iLoop, "Name", PlayerRanking(iLoop).Name)
Call WriteVar(RankFile, "Player" & iLoop, "Path", PlayerRanking(iLoop).Path)
Call WriteVar(RankFile, "Player" & iLoop, "Level", Str(PlayerRanking(iLoop).Lv))
Call WriteVar(RankFile, "Player" & iLoop, "Stat", Str(PlayerRanking(iLoop).Stats))

Next iLoop


End Function


Function LoadNPCData()
'*****************************************************************
'Loads a NPC data to a data list
'*****************************************************************
Dim NpcIndex As Integer
Dim NPCFile As String
Dim TotalNPCs As Integer

'Set NPC file
NPCFile = IniPath & "NPC.dat"

TotalNPCs = Val(GetVar(NPCFile, "INIT", "NumNPCs"))

For NpcIndex = 1 To TotalNPCs

'Load stats from file
NPCData(NpcIndex).Name = GetVar(NPCFile, "NPC" & NpcIndex, "Name")
NPCData(NpcIndex).Desc = GetVar(NPCFile, "NPC" & NpcIndex, "Desc")
NPCData(NpcIndex).Movement = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Movement"))
NPCData(NpcIndex).RespawnWait = Val(GetVar(NPCFile, "NPC" & NpcIndex, "RespawnWait"))

NPCData(NpcIndex).Char.Body = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Body"))
NPCData(NpcIndex).Char.Head = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Head"))
NPCData(NpcIndex).Char.Heading = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Heading"))

NPCData(NpcIndex).Attackable = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Attackable"))
NPCData(NpcIndex).Hostile = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Hostile"))
NPCData(NpcIndex).GiveExp = Val(GetVar(NPCFile, "NPC" & NpcIndex, "GiveExp"))
NPCData(NpcIndex).GiveGold = Val(GetVar(NPCFile, "NPC" & NpcIndex, "GiveGold"))
NPCData(NpcIndex).DropGoldChance = Val(GetVar(NPCFile, "NPC" & NpcIndex, "DropGoldChance"))

NPCData(NpcIndex).DropItem = Val(GetVar(NPCFile, "NPC" & NpcIndex, "DropItem"))
NPCData(NpcIndex).DropChance = Val(GetVar(NPCFile, "NPC" & NpcIndex, "DropChance"))

NPCData(NpcIndex).Stats.MaxHP = Val(GetVar(NPCFile, "NPC" & NpcIndex, "MaxHP"))
NPCData(NpcIndex).Stats.CurHP = NPCData(NpcIndex).Stats.MaxHP
NPCData(NpcIndex).Stats.MaxHIT = Val(GetVar(NPCFile, "NPC" & NpcIndex, "MaxHIT"))
NPCData(NpcIndex).Stats.MinHIT = Val(GetVar(NPCFile, "NPC" & NpcIndex, "MinHIT"))
NPCData(NpcIndex).Stats.AC = Val(GetVar(NPCFile, "NPC" & NpcIndex, "AC"))

NPCData(NpcIndex).Speed = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Speed"))

'If NPCData(NpcIndex).Hostile = 0 Then
'    NPCData(NpcIndex).Shop = Val(GetVar(NPCFile, "NPC" & NpcIndex, "Shop"))
'    If NPCData(NpcIndex).Shop > 0 Then
'        Call OpenNPCShop(NpcIndex)
'    End If
'End If

LogData = "Npc " & NPCData(NpcIndex).Name & " created at " & Time$ & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


Next NpcIndex



End Function

Sub LoadMapData()
'*****************************************************************
'Loads the MapX.X files
'*****************************************************************
Dim map As Integer
Dim LoopC As Integer
Dim x As Integer
Dim y As Integer
Dim TempInt As Integer

NumMaps = Val(GetVar(IniPath & "Map.dat", "INIT", "NumMaps"))
MapPath = GetVar(IniPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For map = 1 To NumMaps
   
    'Open files
    
    'map
    Open App.Path & MapPath & "Map" & map & ".map" For Binary As #1
    Seek #1, 1
    
    'inf
    Open App.Path & MapPath & "Map" & map & ".inf" For Binary As #2
    Seek #2, 1
    
    'map Header
    Get #1, , MapInfo(map).MapVersion
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt

    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt

    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            '.dat file
            Get #1, , MapData(map, x, y).Blocked
            
            'Get GRH number
            For LoopC = 1 To 4
                Get #1, , MapData(map, x, y).Graphic(LoopC)
            Next LoopC
            
            'Space holder for future Expansion
            Get #1, , TempInt
            Get #1, , TempInt
                                
                                
            '.inf file
            
            'tile exit
            Get #2, , MapData(map, x, y).TileExit.map
            Get #2, , MapData(map, x, y).TileExit.x
            Get #2, , MapData(map, x, y).TileExit.y
            
            'Get and make NPC
            Get #2, , TempInt
            If TempInt > 0 Then
                SpawnNPC OpenNPC(TempInt), map, x, y
            Else
                MapData(map, x, y).NpcIndex = 0
            End If

            'Get and make Object
            Get #2, , MapData(map, x, y).ObjInfo.ObjIndex
            Get #2, , MapData(map, x, y).ObjInfo.Amount

            'Space holder for future Expansion
            Get #2, , TempInt
            Get #2, , TempInt
        
        Next x
    Next y

    'Close files
    Close #1
    Close #2

    'Other Room Data
    MapInfo(map).Name = GetVar(App.Path & MapPath & "Map" & map & ".dat", "Map" & map, "Name")
    MapInfo(map).Music = GetVar(App.Path & MapPath & "Map" & map & ".dat", "Map" & map, "MusicNum")
    MapInfo(map).StartPos.map = Val(ReadField(1, GetVar(App.Path & MapPath & "Map" & map & ".dat", "Map" & map, "StartPos"), 45))
    MapInfo(map).StartPos.x = Val(ReadField(2, GetVar(App.Path & MapPath & "Map" & map & ".dat", "Map" & map, "StartPos"), 45))
    MapInfo(map).StartPos.y = Val(ReadField(3, GetVar(App.Path & MapPath & "Map" & map & ".dat", "Map" & map, "StartPos"), 45))
    
    LogData = "Map " & MapInfo(map).Name & " loaded at " & Time$ & FONTTYPE_TALK
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Next map

End Sub

Sub LoadSini()
'*****************************************************************
'Loads the Server.ini
'*****************************************************************

'Misc
frmMain.txPortNumber.Text = GetVar(IniPath & "Server.ini", "INIT", "StartPort")
HideMe = Val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = Val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = Val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

'Start pos
StartPos.map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.x = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Res pos
ResPos.map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.x = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))

'Ressurect pos
ResPos.map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.x = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
'Max users
MaxUsers = Val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
ReDim UserList(1 To MaxUsers) As User

End Sub

Sub WriteVar(File As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(userindex As Integer, UserFile As String)
'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************
Dim LoopC As Integer

UpdateRankList (userindex)

Call WriteVar(UserFile, "INIT", "Password", UserList(userindex).Password)
Call WriteVar(UserFile, "INIT", "Desc", UserList(userindex).Desc)
Call WriteVar(UserFile, "INIT", "Heading", Str(UserList(userindex).Char.Heading))
Call WriteVar(UserFile, "INIT", "Head", Str(UserList(userindex).Char.Head))
Call WriteVar(UserFile, "INIT", "Body", Str(UserList(userindex).Char.Body))
Call WriteVar(UserFile, "INIT", "Weapon", Str(UserList(userindex).Char.Weapon))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(userindex).IP)
Call WriteVar(UserFile, "INIT", "Position", UserList(userindex).Pos.map & "-" & UserList(userindex).Pos.x & "-" & UserList(userindex).Pos.y)

Call WriteVar(UserFile, "INIT", "Path", UserList(userindex).Path)


'Call WriteVar(UserFile, "INIT", "Path", UserList(UserIndex).Path)
'Call WriteVar(UserFile, "INIT", "Title", UserList(UserIndex).Title)

Call WriteVar(UserFile, "FLAGS", "Sound", UserList(userindex).Flags.Sound)
Call WriteVar(UserFile, "FLAGS", "PK", UserList(userindex).Flags.PK)

Call WriteVar(UserFile, "STATS", "MaxHP", Str(UserList(userindex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "CurHP", Str(UserList(userindex).Stats.CurHP))
Call WriteVar(UserFile, "STATS", "MaxMP", Str(UserList(userindex).Stats.MaxMP))
Call WriteVar(UserFile, "STATS", "CurMP", Str(UserList(userindex).Stats.CurMP))
Call WriteVar(UserFile, "STATS", "Lv", Str(UserList(userindex).Stats.Lv))
Call WriteVar(UserFile, "STATS", "Exp", Str(UserList(userindex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "Tnl", Str(UserList(userindex).Stats.Tnl))
Call WriteVar(UserFile, "STATS", "Texp", Str(UserList(userindex).Stats.Texp))
Call WriteVar(UserFile, "STATS", "AC", Str(UserList(userindex).Stats.AC))
Call WriteVar(UserFile, "STATS", "Dam", Str(UserList(userindex).Stats.Dam))

Call WriteVar(UserFile, "STATS", "Str", Str(UserList(userindex).Stats.Str))
Call WriteVar(UserFile, "STATS", "Con", Str(UserList(userindex).Stats.Con))
Call WriteVar(UserFile, "STATS", "Int", Str(UserList(userindex).Stats.Int))
Call WriteVar(UserFile, "STATS", "Wis", Str(UserList(userindex).Stats.Wis))
Call WriteVar(UserFile, "STATS", "Dex", Str(UserList(userindex).Stats.Dex))

Call WriteVar(UserFile, "STATS", "Gold", Str(UserList(userindex).Stats.Gold))


Call WriteVar(UserFile, "STATS", "MaxHIT", Str(UserList(userindex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", Str(UserList(userindex).Stats.MinHIT))
  
  
'Save Inv
For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(userindex).Object(LoopC).ObjIndex & "-" & UserList(userindex).Object(LoopC).Amount & "-" & UserList(userindex).Object(LoopC).Equipped)
Next

For LoopC = 1 To MAX_SPELL_SLOTS
    Call WriteVar(UserFile, "Spells", "Spl" & LoopC, Str(UserList(userindex).SpellBook(LoopC).Spellindex))
Next

'Write Weapon and Armour slots
Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", Str(UserList(userindex).WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", Str(UserList(userindex).ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "AccEqpSlot", Str(UserList(userindex).AccEqpSlot))
Call WriteVar(UserFile, "Inventory", "HelmEqpSlot", Str(UserList(userindex).HelmEqpSlot))

End Sub



