Attribute VB_Name = "GameLogic"
Option Explicit

Sub UpdateRankList(ByVal userindex As Integer)
Dim jLoop As Integer
Dim iLoop As Integer
Dim found As Boolean
Dim TempRank As RankEntry
found = False

For iLoop = 1 To TotalRanks
    If UCase(UserList(userindex).Name) = UCase(PlayerRanking(iLoop).Name) Then
        'Call SendData(ToIndex, userindex, 0, "#Same." & FONTTYPE_TALK)
        PlayerRanking(iLoop).Name = UserList(userindex).Name
        PlayerRanking(iLoop).Path = UserList(userindex).Path
        PlayerRanking(iLoop).Lv = UserList(userindex).Stats.Lv
        PlayerRanking(iLoop).Stats = UserList(userindex).Stats.MaxHP + UserList(userindex).Stats.MaxMP
        If LCase(UserList(userindex).Name) = "inkey" Or LCase(UserList(userindex).Name) = "miracle" Or LCase(UserList(userindex).Name) = "black" Then
            PlayerRanking(iLoop).Name = "noname"
            PlayerRanking(iLoop).Path = "none"
            PlayerRanking(iLoop).Lv = 1
            PlayerRanking(iLoop).Stats = 1
        End If
    
        found = True
        'Call SendData(ToIndex, userindex, 0, "#Found." & FONTTYPE_TALK)
    End If
Next iLoop


If found = False Then
    If LCase(UserList(userindex).Name) = "inkey" Or LCase(UserList(userindex).Name) = "miracle" Or LCase(UserList(userindex).Name) = "black" Then
        Exit Sub
    End If
    For iLoop = 1 To TotalRanks
        'level 99 passing
        If UserList(userindex).Stats.Lv = 99 Then
            If UserList(userindex).Stats.MaxHP + UserList(userindex).Stats.MaxMP >= PlayerRanking(iLoop).Stats Then
                'Call SendData(ToIndex, userindex, 0, "#You passed someone!." & FONTTYPE_TALK)
                For jLoop = TotalRanks To iLoop Step -1
                    If jLoop = 1 Then
                        PlayerRanking(jLoop).Name = ""
                        PlayerRanking(jLoop).Path = ""
                        PlayerRanking(jLoop).Lv = 0
                        PlayerRanking(jLoop).Stats = 0
                    Else
                        PlayerRanking(jLoop).Name = PlayerRanking(jLoop - 1).Name
                        PlayerRanking(jLoop).Path = PlayerRanking(jLoop - 1).Path
                        PlayerRanking(jLoop).Lv = PlayerRanking(jLoop - 1).Lv
                        PlayerRanking(jLoop).Stats = PlayerRanking(jLoop - 1).Stats
                    End If
                Next jLoop
                PlayerRanking(iLoop).Name = UserList(userindex).Name
                PlayerRanking(iLoop).Path = UserList(userindex).Path
                PlayerRanking(iLoop).Lv = UserList(userindex).Stats.Lv
                PlayerRanking(iLoop).Stats = UserList(userindex).Stats.MaxHP + UserList(userindex).Stats.MaxMP
                Exit Sub
            End If
        
        'non 99 passing
        ElseIf UserList(userindex).Stats.Lv < 99 Then
            If UserList(userindex).Stats.Lv > PlayerRanking(iLoop).Lv Then
                For jLoop = TotalRanks To iLoop Step -1
                    If jLoop = 1 Then
                        PlayerRanking(jLoop).Name = ""
                        PlayerRanking(jLoop).Path = ""
                        PlayerRanking(jLoop).Lv = 0
                        PlayerRanking(jLoop).Stats = 0
                    Else
                        PlayerRanking(jLoop).Name = PlayerRanking(jLoop - 1).Name
                        PlayerRanking(jLoop).Path = PlayerRanking(jLoop - 1).Path
                        PlayerRanking(jLoop).Lv = PlayerRanking(jLoop - 1).Lv
                        PlayerRanking(jLoop).Stats = PlayerRanking(jLoop - 1).Stats
                    End If
                Next jLoop
                PlayerRanking(iLoop).Name = UserList(userindex).Name
                PlayerRanking(iLoop).Path = UserList(userindex).Path
                PlayerRanking(iLoop).Lv = UserList(userindex).Stats.Lv
                PlayerRanking(iLoop).Stats = UserList(userindex).Stats.MaxHP + UserList(userindex).Stats.MaxMP
                'Call SendData(ToIndex, userindex, 0, "#You passed someone!." & FONTTYPE_TALK)
                Exit Sub
            End If
        End If
    Next iLoop
End If

If found = True Then
    For iLoop = 1 To TotalRanks - 1
        If PlayerRanking(iLoop + 1).Lv >= PlayerRanking(iLoop).Lv Then
            If PlayerRanking(iLoop + 1).Lv = 99 Then
                If PlayerRanking(iLoop + 1).Stats > PlayerRanking(iLoop).Stats Then
                    TempRank = PlayerRanking(iLoop + 1)
                    PlayerRanking(iLoop + 1) = PlayerRanking(iLoop)
                    PlayerRanking(iLoop) = TempRank
                End If
            Else
                TempRank = PlayerRanking(iLoop + 1)
                PlayerRanking(iLoop + 1) = PlayerRanking(iLoop)
                PlayerRanking(iLoop) = TempRank
            End If
        End If
    Next iLoop
End If



End Sub

Sub FaceChange(ByVal userindex As Integer, ByVal ShopIndex As Integer, ByVal ItemSlot As Integer)
'*****************************************************************
'To change faces
'*****************************************************************

UserList(userindex).Char.Head = NpcShops(ShopIndex).Slots(ItemSlot).Item
Call ChangeUserChar(ToMap, userindex, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)

End Sub


Sub BuyHitPoints(ByVal userindex As Integer)
'*****************************************************************
'For level 99 players to buy HP from an NPC
'*****************************************************************
Dim StatCost As Long


If UserList(userindex).Stats.Lv < 99 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 99." & FONTTYPE_TALK)
    Exit Sub
End If

StatCost = 1000000 + UserList(userindex).Stats.MaxHP * 100

If UserList(userindex).Stats.Texp < StatCost Then
    Call SendData(ToIndex, userindex, 0, "#You need at least " & StatCost & " exp." & FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).AccEqpObjIndex > 0 Or UserList(userindex).WeaponEqpObjIndex > 0 Or UserList(userindex).ArmourEqpObjIndex > 0 Or UserList(userindex).HelmEqpObjIndex > 0 Then
    Call SendData(ToIndex, userindex, 0, "#Please remove any items before buying stats." & FONTTYPE_TALK)
    Exit Sub
End If

UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 10
UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.Texp = UserList(userindex).Stats.Texp - StatCost
UserList(userindex).Flags.StatsChanged = True

UpdateRankList (userindex)

Call SendData(ToIndex, userindex, 0, "#You gain 10 HP." & FONTTYPE_TALK)


End Sub

Sub BuyMagicPoints(ByVal userindex As Integer)
'*****************************************************************
'For level 99 players to buy MP from an NPC
'*****************************************************************
Dim StatCost As Long


If UserList(userindex).Stats.Lv < 99 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 99." & FONTTYPE_TALK)
    Exit Sub
End If

StatCost = 1000000 + UserList(userindex).Stats.MaxMP * 100

If UserList(userindex).Stats.Texp < StatCost Then
    Call SendData(ToIndex, userindex, 0, "#You need at least " & StatCost & " exp." & FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).AccEqpObjIndex > 0 Or UserList(userindex).WeaponEqpObjIndex > 0 Or UserList(userindex).ArmourEqpObjIndex > 0 Or UserList(userindex).HelmEqpObjIndex > 0 Then
    Call SendData(ToIndex, userindex, 0, "#Please remove any items before buying stats." & FONTTYPE_TALK)
    Exit Sub
End If

UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP + 10
UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
UserList(userindex).Stats.Texp = UserList(userindex).Stats.Texp - StatCost
UserList(userindex).Flags.StatsChanged = True

UpdateRankList (userindex)

Call SendData(ToIndex, userindex, 0, "#You gain 10 MP." & FONTTYPE_TALK)


End Sub

Sub UngroupUser(ByVal userindex As Integer)
'*****************************************************************
'Ungroup a user by index
'*****************************************************************
Dim TmpI As Integer
Dim TmpJ As Integer


If UserList(userindex).GIndex > 0 Then
    For TmpI = 1 To 5
        If Groups(UserList(userindex).GIndex).UIndexes(TmpI) = userindex Then
            For TmpJ = 1 To 5
                If Groups(UserList(userindex).GIndex).UIndexes(TmpJ) > 0 Then
                    Call SendData(ToIndex, Groups(UserList(userindex).GIndex).UIndexes(TmpJ), 0, "#" & UserList(userindex).Name & " is leaving this group." & FONTTYPE_TALK)
                End If
            Next TmpJ
            Groups(UserList(userindex).GIndex).UIndexes(TmpI) = 0
            UserList(userindex).GIndex = 0
            Call LegalGroups
            Exit Sub
        End If
    Next TmpI
Else
    UserList(userindex).GIndex = 0
    Call SendData(ToIndex, userindex, 0, "#You leave a group composed only of yourself." & FONTTYPE_TALK)
End If

End Sub

Sub LegalGroups()

Dim LoopB As Integer
Dim LoopC As Integer
Dim i As Integer

For LoopB = 1 To MaxUsers / 2
    i = 0
    For LoopC = 1 To 5
        If Groups(LoopB).UIndexes(LoopC) Then i = i + 1
    Next LoopC
    
    If i = 1 Then
        For LoopC = 1 To 5
            If Groups(LoopB).UIndexes(LoopC) > 0 Then
                UserList(Groups(LoopB).UIndexes(LoopC)).GIndex = 0
                Groups(LoopB).UIndexes(LoopC) = 0
                LogData = "Illegal group found and destoryed." & Time$ & FONTTYPE_TALK
                AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
            End If
        Next LoopC
    End If

Next LoopB

'For LoopB = 1 To MaxUsers / 2
    'For LoopC = 1 To 5
    '    If Groups(LoopB).UIndexes(LoopC) > LastUser Then
    '        Groups(LoopB).UIndexes(LoopC) = 0
    '        LogData = "Illegal group entry found and destoryed." & Time$ & FONTTYPE_TALK
    '        AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
    '    End If
        'If Groups(LoopB).UIndexes(LoopC) > 0 And Groups(LoopB).UIndexes(LoopC) <= LastUser Then
        '    If UserList(Groups(UserList(LoopB).GIndex).UIndexes(LoopC)).Flags.UserLogged = False Then
        '        Groups(LoopB).UIndexes(LoopC) = 0
        '        LogData = "Illegal group entry found and destoryed." & Time$ & FONTTYPE_TALK
        '        AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
        '    End If
        'End If
    'Next LoopC
'Next LoopB

End Sub

Sub GroupUser(ByVal userindex As Integer, ByVal rData As String)
'*****************************************************************
'Group a user by name with a user's index.
'*****************************************************************
Dim AddIndex As Integer
Dim TmpI As Integer
Dim TmpJ As Integer
    
For TmpI = 1 To MaxUsers
    If UCase(UserList(TmpI).Name) = UCase(rData) Then
        AddIndex = TmpI
        TmpI = 5000
    End If
Next TmpI
        
    
If AddIndex = 0 Then
    Call SendData(ToIndex, userindex, 0, "#User not found..." & FONTTYPE_TALK)
    Exit Sub
End If
    
If UserList(AddIndex).Name = "" Then
    Call SendData(ToIndex, userindex, 0, "#User not found..." & FONTTYPE_TALK)
    Exit Sub
End If
    
If AddIndex = userindex Then
    Call SendData(ToIndex, userindex, 0, "#Form a group with yourself?" & FONTTYPE_TALK)
    Exit Sub
End If
    
If UserList(AddIndex).GIndex = 0 Then
    If UserList(AddIndex).GIndex = 0 And UserList(userindex).GIndex = 0 Then
        For TmpI = 1 To MaxUsers / 2
            If Groups(TmpI).UIndexes(1) = 0 & Groups(TmpI).UIndexes(2) = 0 & Groups(TmpI).UIndexes(3) = 0 & Groups(TmpI).UIndexes(4) = 0 & Groups(TmpI).UIndexes(5) = 0 Then
                Groups(TmpI).UIndexes(1) = userindex
                Groups(TmpI).UIndexes(2) = AddIndex
                UserList(userindex).GIndex = TmpI
                UserList(AddIndex).GIndex = TmpI
                Call SendData(ToIndex, userindex, 0, "#Group formed (" & TmpI & ")." & FONTTYPE_TALK)
                Call SendData(ToIndex, userindex, 0, "#" & UserList(userindex).Name & " is joining this group." & FONTTYPE_TALK)
                Call SendData(ToIndex, userindex, 0, "#" & UserList(AddIndex).Name & " is joining this group." & FONTTYPE_TALK)
                Call SendData(ToIndex, AddIndex, 0, "#" & UserList(userindex).Name & " is joining this group." & FONTTYPE_TALK)
                Call SendData(ToIndex, AddIndex, 0, "#" & UserList(AddIndex).Name & " is joining this group." & FONTTYPE_TALK)
                Exit Sub
            End If
        Next TmpI
    Else
        For TmpI = 1 To 5
            If Groups(UserList(userindex).GIndex).UIndexes(TmpI) = 0 Then
                Groups(UserList(userindex).GIndex).UIndexes(TmpI) = AddIndex
                UserList(AddIndex).GIndex = UserList(userindex).GIndex
                For TmpJ = 1 To 5
                    If Groups(UserList(userindex).GIndex).UIndexes(TmpJ) > 0 Then
                        Call SendData(ToIndex, Groups(UserList(userindex).GIndex).UIndexes(TmpJ), 0, "#" & UserList(AddIndex).Name & " is joining this group." & FONTTYPE_TALK)
                    End If
                Next TmpJ
                Exit Sub
            End If
        Next TmpI
        Call SendData(ToIndex, Groups(UserList(userindex).GIndex).UIndexes(TmpJ), 0, "#This group already has 5 members." & FONTTYPE_TALK)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "#That person is already in a group." & FONTTYPE_TALK)
    Exit Sub
End If
    
End Sub


Sub ChangePeasantJob(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ItemSlot As Integer)
'*****************************************************************
'Change the user's path at level 5.
'*****************************************************************

If UserList(userindex).Path <> "Peasant" Then
    Call SendData(ToIndex, userindex, 0, "#You have already found a path." & FONTTYPE_INFO)
    Exit Sub
End If
If UserList(userindex).Stats.Lv <> 5 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 5." & FONTTYPE_INFO)
    Exit Sub
End If
        
If UserList(userindex).Stats.Gold >= 200 Then
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold - 200
    UserList(userindex).Path = NpcShops(NpcIndex).Slots(ItemSlot).Path
    Call SendData(ToAll, userindex, 0, "@" & UserList(userindex).Name & " has become a " & NpcShops(NpcIndex).Slots(ItemSlot).Path & "!" & FONTTYPE_WARNING)
    Call SendData(ToIndex, userindex, 0, "@Seek out training at your guild." & FONTTYPE_INFO)
    Exit Sub
End If

End Sub


Sub UserForgetSpell(ByVal userindex As Integer, ByVal Spellindex As Integer)

UserList(userindex).SpellBook(Spellindex).Spellindex = 0

Call UpdateUserSpell(True, userindex, Spellindex)

Call SendData(ToIndex, userindex, 0, "#You forget slot A" & FONTTYPE_INFO)

End Sub


Sub TeachSpell(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ItemSlot As Integer)
'*****************************************************************
'Teach User a Spell
'*****************************************************************
Dim iFreeSpellSlot As Integer
Dim LoopC As Integer


'Make sure the user is a path that can learn the spell
If UCase(NpcShops(NpcIndex).Slots(ItemSlot).Path) <> "ALL" Then
    If UserList(userindex).Path <> NpcShops(NpcIndex).Slots(ItemSlot).Path Then
        Call SendData(ToIndex, userindex, 0, "#You do not walk the right path." & FONTTYPE_INFO)
        Exit Sub
    End If
End If

'Make sure they are strong enough
If UserList(userindex).Stats.Lv < NpcShops(NpcIndex).Slots(ItemSlot).Level Then
    Call SendData(ToIndex, userindex, 0, "#You must be level " & NpcShops(NpcIndex).Slots(ItemSlot).Level & "." & FONTTYPE_INFO)
    Exit Sub
End If

'Make sure they have enough gold
If UserList(userindex).Stats.Gold < NpcShops(NpcIndex).Slots(ItemSlot).Gold Then
    Call SendData(ToIndex, userindex, 0, "#You do not have enough gold." & FONTTYPE_INFO)
    Exit Sub
End If

'    UserList(UserIndex).Stats.Gold = UserList(UserIndex).Stats.Gold - 200
'    UserList(UserIndex).Path = NpcShops(NpcIndex).Slots(ItemSlot).Path
'    Call SendData(ToAll, UserIndex, 0, "@" & UserList(UserIndex).Name & " has become a " & NpcShops(NpcIndex).Slots(ItemSlot).Path & "!" & FONTTYPE_WARNING)

'    Exit Sub
'End If

'Find a free slot
For iFreeSpellSlot = 1 To 26
    If UserList(userindex).SpellBook(iFreeSpellSlot).Spellindex = 0 Then
        Exit For
    End If
    If UserList(userindex).SpellBook(iFreeSpellSlot).Spellindex = NpcShops(NpcIndex).Slots(ItemSlot).Item Then
        Call SendData(ToIndex, userindex, 0, "#You already know this spell." & FONTTYPE_INFO)
        Exit Sub
    End If
Next iFreeSpellSlot

'Make sure their book is not full.
If iFreeSpellSlot = 27 Then
    If UserList(userindex).SpellBook(iFreeSpellSlot).Spellindex > 0 Then
        Call SendData(ToIndex, userindex, 0, "#Your mind is too full, you already know too many secrets." & FONTTYPE_INFO)
        Exit Sub
    End If
End If


'Take the user's gold
UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold - NpcShops(NpcIndex).Slots(ItemSlot).Gold

'Teach the spell
UserList(userindex).SpellBook(iFreeSpellSlot).Spellindex = NpcShops(NpcIndex).Slots(ItemSlot).Item

Call UpdateUserSpell(True, userindex, iFreeSpellSlot)

Call SendData(ToIndex, userindex, 0, "#Your mind expands as you learn " & SpellData(NpcShops(NpcIndex).Slots(ItemSlot).Item).Name & FONTTYPE_INFO) '". (slot " & iFreeSpellSlot & ")" &
 
End Sub

Sub SendNextMapTile(ByVal userindex As Integer)
'*****************************************************************
'Send a map tile to a user
'*****************************************************************
Dim LoopC As Integer
Dim ln As String
Dim TempInt As Integer
              
If UserList(userindex).Counters.SendMapCounter.y > YMaxMapSize Then
    SendData ToIndex, userindex, 0, "EMT" & UserList(userindex).Counters.SendMapCounter.map
    UserList(userindex).Flags.DownloadingMap = 0
    UserList(userindex).Counters.SendMapCounter.x = 0
    UserList(userindex).Counters.SendMapCounter.y = 0
    UserList(userindex).Counters.SendMapCounter.map = 0
Else
    
    ln = UserList(userindex).Counters.SendMapCounter.x & "," & UserList(userindex).Counters.SendMapCounter.y & "," & MapData(UserList(userindex).Counters.SendMapCounter.map, UserList(userindex).Counters.SendMapCounter.x, UserList(userindex).Counters.SendMapCounter.y).Blocked
    For LoopC = 1 To 4
        TempInt = MapData(UserList(userindex).Counters.SendMapCounter.map, UserList(userindex).Counters.SendMapCounter.x, UserList(userindex).Counters.SendMapCounter.y).Graphic(LoopC)
        If TempInt > 0 Then
            ln = ln & "," & LoopC & TempInt
        End If
    Next LoopC
                
    SendData ToIndex, userindex, 0, "CMT" & ln
    
    UserList(userindex).Counters.SendMapCounter.x = UserList(userindex).Counters.SendMapCounter.x + 1
    If UserList(userindex).Counters.SendMapCounter.x > XMaxMapSize Then
        UserList(userindex).Counters.SendMapCounter.x = XMinMapSize
        UserList(userindex).Counters.SendMapCounter.y = UserList(userindex).Counters.SendMapCounter.y + 1
    End If
    
    UserList(userindex).Flags.ReadyForNextTile = 0
    
End If

End Sub

Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, Body As Integer, Head As Integer, Heading As Byte, Weapon As Integer)
'*****************************************************************
'Changes a user char's head,body and heading
'*****************************************************************

Dim HpPerc As Integer
HpPerc = UserList(userindex).Stats.CurHP / UserList(userindex).Stats.MaxHP * 100

UserList(userindex).Char.Body = Body
UserList(userindex).Char.Head = Head
UserList(userindex).Char.Heading = Heading
UserList(userindex).Char.Weapon = Weapon

Call SendData(sndRoute, sndIndex, sndMap, "CHC" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & HpPerc & "," & Weapon)

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, NpcIndex As Integer, Body As Integer, Head As Integer, Heading As Byte)
'*****************************************************************
'Changes a NPC char's head,body and heading
'*****************************************************************

Dim HpPerc As Integer

NPCList(NpcIndex).Char.Body = Body
NPCList(NpcIndex).Char.Head = Head
NPCList(NpcIndex).Char.Heading = Heading

HpPerc = NPCList(NpcIndex).Stats.CurHP / NPCList(NpcIndex).Stats.MaxHP * 100
Call SendData(sndRoute, sndIndex, sndMap, "CHC" & NPCList(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & HpPerc)

End Sub

Sub CheckUserLevel(ByVal userindex As Integer)
'*****************************************************************
'Checks user's Exp and levels user up
'*****************************************************************

'Make sure user hasn't reached max level
If UserList(userindex).Stats.Lv = STAT_MAXLv Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.Tnl = 0
    Exit Sub
End If

'If UserList(UserIndex).Stats.Lv = 1 Then
'    UserList(UserIndex).Path = 1
'    UserList(UserIndex).strTitle = ""
'End If

'Update the ranklist whenever something is killed I guess.
UpdateRankList (userindex)

'If Exp >= then tnl then level up user
If UserList(userindex).Stats.Exp >= UserList(userindex).Stats.Tnl Then
    If UserList(userindex).Path = "Peasant" Then
        If UserList(userindex).Stats.Lv < 5 Then
            SendData ToIndex, userindex, 0, "#Level up!" & FONTTYPE_TALK
            
            AddtoVar UserList(userindex).Stats.MaxHP, 7, STAT_MAXHP
            AddtoVar UserList(userindex).Stats.MaxMP, 7, STAT_MaxMP

            UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - 1

            AddtoVar UserList(userindex).Stats.Int, 1, 200
            AddtoVar UserList(userindex).Stats.Wis, 1, 200
            AddtoVar UserList(userindex).Stats.Con, 1, 200
            AddtoVar UserList(userindex).Stats.Str, 1, 200
            AddtoVar UserList(userindex).Stats.Dex, 1, 200
            
            UserList(userindex).Stats.Lv = UserList(userindex).Stats.Lv + 1
            UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.Tnl
            Call GetTNLExp(UserList(userindex).Stats.Lv, userindex)
            
            SendUserStatsBox userindex
        Else
            SendData ToIndex, userindex, 0, "@Please choose a path before continuing." & FONTTYPE_INFO
            SendData ToIndex, userindex, 0, "#Please choose a path before continuing." & FONTTYPE_INFO
            SendUserStatsBox userindex
        End If
    End If
    
    If UserList(userindex).Path = "Warrior" Then
        SendData ToIndex, userindex, 0, "#Level up!" & FONTTYPE_TALK
            
        AddtoVar UserList(userindex).Stats.MaxHP, 9, STAT_MAXHP
        AddtoVar UserList(userindex).Stats.MaxMP, 1, STAT_MaxMP

        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - 1

        'AddtoVar UserList(UserIndex).Stats.Int, 1, 200
        'AddtoVar UserList(UserIndex).Stats.Wis, 1, 200
        AddtoVar UserList(userindex).Stats.Con, 1, 200
        AddtoVar UserList(userindex).Stats.Str, 1, 200
        AddtoVar UserList(userindex).Stats.Dex, 1, 200
    
        UserList(userindex).Stats.Lv = UserList(userindex).Stats.Lv + 1
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.Tnl
        Call GetTNLExp(UserList(userindex).Stats.Lv, userindex)
            
        SendUserStatsBox userindex
    End If
    
    If UserList(userindex).Path = "Wizard" Then
        SendData ToIndex, userindex, 0, "#Level up!" & FONTTYPE_TALK
            
        AddtoVar UserList(userindex).Stats.MaxHP, 5, STAT_MAXHP
        AddtoVar UserList(userindex).Stats.MaxMP, 6, STAT_MaxMP

        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - 1

        AddtoVar UserList(userindex).Stats.Int, 1, 200
        AddtoVar UserList(userindex).Stats.Wis, 1, 200
        'AddtoVar UserList(UserIndex).Stats.Con, 1, 200
        'AddtoVar UserList(UserIndex).Stats.Str, 1, 200
        AddtoVar UserList(userindex).Stats.Dex, 1, 200

        UserList(userindex).Stats.Lv = UserList(userindex).Stats.Lv + 1
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.Tnl
        Call GetTNLExp(UserList(userindex).Stats.Lv, userindex)
            
        SendUserStatsBox userindex
    End If
    
    If UserList(userindex).Path = "Cleric" Then
        SendData ToIndex, userindex, 0, "#Level up!" & FONTTYPE_TALK
            
        AddtoVar UserList(userindex).Stats.MaxHP, 6, STAT_MAXHP
        AddtoVar UserList(userindex).Stats.MaxMP, 5, STAT_MaxMP

        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - 1

        AddtoVar UserList(userindex).Stats.Int, 1, 200
        AddtoVar UserList(userindex).Stats.Wis, 1, 200
        'AddtoVar UserList(UserIndex).Stats.Con, 1, 200
        'AddtoVar UserList(UserIndex).Stats.Str, 1, 200
        AddtoVar UserList(userindex).Stats.Dex, 1, 200

        UserList(userindex).Stats.Lv = UserList(userindex).Stats.Lv + 1
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.Tnl
        Call GetTNLExp(UserList(userindex).Stats.Lv, userindex)
            
        SendUserStatsBox userindex
    End If
    
    UpdateRankList (userindex)
    Call SendData(ToIndex, userindex, 0, "TNL" & UserList(userindex).Stats.Exp & "," & UserList(userindex).Stats.Tnl)
    
    LogData = UserList(userindex).Name & " is now level " & UserList(userindex).Stats.Lv & ". (" & Time$ & ")" & FONTTYPE_SHOUT
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "Lv UP " & LogData
    Close #5
End If

End Sub

Sub GetTNLExp(ByVal Lv As Integer, ByVal userindex As Integer)
    If Lv = 2 Then UserList(userindex).Stats.Tnl = 400
    If Lv = 3 Then UserList(userindex).Stats.Tnl = 550
    If Lv = 4 Then UserList(userindex).Stats.Tnl = 750
    If Lv = 5 Then UserList(userindex).Stats.Tnl = 1000
    If Lv = 6 Then UserList(userindex).Stats.Tnl = 1250
    If Lv = 7 Then UserList(userindex).Stats.Tnl = 1600
    If Lv = 8 Then UserList(userindex).Stats.Tnl = 2000
    If Lv = 9 Then UserList(userindex).Stats.Tnl = 2500
    If Lv = 10 Then UserList(userindex).Stats.Tnl = 3000
    If Lv = 11 Then UserList(userindex).Stats.Tnl = 3500
    If Lv = 12 Then UserList(userindex).Stats.Tnl = 4100
    If Lv = 13 Then UserList(userindex).Stats.Tnl = 4700
    If Lv = 14 Then UserList(userindex).Stats.Tnl = 5300
    If Lv = 15 Then UserList(userindex).Stats.Tnl = 6000
    If Lv = 16 Then UserList(userindex).Stats.Tnl = 6700
    If Lv = 17 Then UserList(userindex).Stats.Tnl = 7400
    If Lv = 18 Then UserList(userindex).Stats.Tnl = 8100
    If Lv = 19 Then UserList(userindex).Stats.Tnl = 9000
    If Lv = 20 Then UserList(userindex).Stats.Tnl = 9900
    If Lv = 21 Then UserList(userindex).Stats.Tnl = 11000
    If Lv = 22 Then UserList(userindex).Stats.Tnl = 12100
    If Lv = 23 Then UserList(userindex).Stats.Tnl = 13300
    If Lv = 24 Then UserList(userindex).Stats.Tnl = 15000
    If Lv = 25 Then UserList(userindex).Stats.Tnl = 17000
    If Lv = 26 Then UserList(userindex).Stats.Tnl = 19000
    If Lv = 27 Then UserList(userindex).Stats.Tnl = 22000
    If Lv = 28 Then UserList(userindex).Stats.Tnl = 25000
    If Lv = 29 Then UserList(userindex).Stats.Tnl = 29000
    If Lv = 30 Then UserList(userindex).Stats.Tnl = 34000
    If Lv = 31 Then UserList(userindex).Stats.Tnl = 39000
    If Lv = 32 Then UserList(userindex).Stats.Tnl = 45000
    If Lv = 33 Then UserList(userindex).Stats.Tnl = 51000
    If Lv = 34 Then UserList(userindex).Stats.Tnl = 57000
    If Lv = 35 Then UserList(userindex).Stats.Tnl = 64000
    If Lv = 36 Then UserList(userindex).Stats.Tnl = 71000
    If Lv = 37 Then UserList(userindex).Stats.Tnl = 79000
    If Lv = 38 Then UserList(userindex).Stats.Tnl = 87000
    If Lv = 39 Then UserList(userindex).Stats.Tnl = 96000
    If Lv = 40 Then UserList(userindex).Stats.Tnl = 105000
    If Lv = 41 Then UserList(userindex).Stats.Tnl = 115000
    If Lv = 42 Then UserList(userindex).Stats.Tnl = 125000
    If Lv = 43 Then UserList(userindex).Stats.Tnl = 136000
    If Lv = 44 Then UserList(userindex).Stats.Tnl = 147000
    If Lv = 45 Then UserList(userindex).Stats.Tnl = 158000
    If Lv = 46 Then UserList(userindex).Stats.Tnl = 169000
    If Lv = 47 Then UserList(userindex).Stats.Tnl = 180000
    If Lv = 48 Then UserList(userindex).Stats.Tnl = 191000
    If Lv = 49 Then UserList(userindex).Stats.Tnl = 205000
    If Lv = 50 Then UserList(userindex).Stats.Tnl = 220000
    If Lv = 51 Then UserList(userindex).Stats.Tnl = 235000
    If Lv = 52 Then UserList(userindex).Stats.Tnl = 250000
    If Lv = 53 Then UserList(userindex).Stats.Tnl = 265000
    If Lv = 54 Then UserList(userindex).Stats.Tnl = 280000
    If Lv = 55 Then UserList(userindex).Stats.Tnl = 300000
    If Lv = 56 Then UserList(userindex).Stats.Tnl = 320000
    If Lv = 57 Then UserList(userindex).Stats.Tnl = 340000
    If Lv = 58 Then UserList(userindex).Stats.Tnl = 370000
    If Lv = 59 Then UserList(userindex).Stats.Tnl = 400000
    If Lv = 60 Then UserList(userindex).Stats.Tnl = 430000
    If Lv = 61 Then UserList(userindex).Stats.Tnl = 460000
    If Lv = 62 Then UserList(userindex).Stats.Tnl = 500000
    If Lv = 63 Then UserList(userindex).Stats.Tnl = 540000
    If Lv = 64 Then UserList(userindex).Stats.Tnl = 580000
    If Lv = 65 Then UserList(userindex).Stats.Tnl = 620000
    If Lv = 66 Then UserList(userindex).Stats.Tnl = 670000
    If Lv = 67 Then UserList(userindex).Stats.Tnl = 720000
    If Lv = 68 Then UserList(userindex).Stats.Tnl = 770000
    If Lv = 69 Then UserList(userindex).Stats.Tnl = 820000
    If Lv = 70 Then UserList(userindex).Stats.Tnl = 850000
    If Lv = 71 Then UserList(userindex).Stats.Tnl = 870000
    If Lv = 72 Then UserList(userindex).Stats.Tnl = 890000
    If Lv = 73 Then UserList(userindex).Stats.Tnl = 950000
    If Lv = 74 Then UserList(userindex).Stats.Tnl = 1010000
    If Lv = 75 Then UserList(userindex).Stats.Tnl = 1070000
    If Lv = 76 Then UserList(userindex).Stats.Tnl = 1130000
    If Lv = 77 Then UserList(userindex).Stats.Tnl = 1200000
    If Lv = 78 Then UserList(userindex).Stats.Tnl = 1270000
    If Lv = 79 Then UserList(userindex).Stats.Tnl = 1340000
    If Lv = 80 Then UserList(userindex).Stats.Tnl = 1410000
    If Lv = 81 Then UserList(userindex).Stats.Tnl = 1480000
    If Lv = 82 Then UserList(userindex).Stats.Tnl = 1560000
    If Lv = 83 Then UserList(userindex).Stats.Tnl = 1640000
    If Lv = 84 Then UserList(userindex).Stats.Tnl = 1720000
    If Lv = 85 Then UserList(userindex).Stats.Tnl = 1800000
    If Lv = 86 Then UserList(userindex).Stats.Tnl = 1900000
    If Lv = 87 Then UserList(userindex).Stats.Tnl = 2000000
    If Lv = 88 Then UserList(userindex).Stats.Tnl = 2100000
    If Lv = 89 Then UserList(userindex).Stats.Tnl = 2200000
    If Lv = 90 Then UserList(userindex).Stats.Tnl = 2300000
    If Lv = 91 Then UserList(userindex).Stats.Tnl = 2450000
    If Lv = 92 Then UserList(userindex).Stats.Tnl = 2600000
    If Lv = 93 Then UserList(userindex).Stats.Tnl = 2750000
    If Lv = 94 Then UserList(userindex).Stats.Tnl = 2900000
    If Lv = 95 Then UserList(userindex).Stats.Tnl = 4000000
    If Lv = 96 Then UserList(userindex).Stats.Tnl = 5000000
    If Lv = 97 Then UserList(userindex).Stats.Tnl = 6000000
    If Lv = 98 Then UserList(userindex).Stats.Tnl = 8000000
    
    UserList(userindex).Stats.Tnl = UserList(userindex).Stats.Tnl * (1 + UserList(userindex).Stats.Lv / 100)
    
End Sub

Sub DoTileEvents(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Do any events on a tile
'*****************************************************************

'Check for tile exit
If MapData(map, x, y).TileExit.map > 0 Then
    If LegalPos(MapData(map, x, y).TileExit.map, MapData(map, x, y).TileExit.x, MapData(map, x, y).TileExit.y) Then
        Call WarpUserChar(userindex, MapData(map, x, y).TileExit.map, MapData(map, x, y).TileExit.x, MapData(map, x, y).TileExit.y)
    End If
End If

End Sub

Function InMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub KillNPC(ByVal NpcIndex As Integer)
'*****************************************************************
'Kill a NPC
'*****************************************************************

'Set health back to 100%
NPCList(NpcIndex).Stats.CurHP = NPCList(NpcIndex).Stats.MaxHP
NPCList(NpcIndex).Stats.AC = 0
NPCList(NpcIndex).CurseName = ""
NPCList(NpcIndex).ParaCount = 0
NPCList(NpcIndex).PoisonCount = 0
NPCList(NpcIndex).PoisonDamage = 0

'Erase it from map
KillNPCChar ToMap, 0, NPCList(NpcIndex).Pos.map, NpcIndex

'Set respawn wait
NPCList(NpcIndex).Counters.RespawnCounter = NPCList(NpcIndex).RespawnWait

End Sub

Sub SpawnNPC(ByVal NpcIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Places a NPC that has been Opened
'*****************************************************************

Dim TempPos As WorldPos

NPCList(NpcIndex).Pos.map = map
NPCList(NpcIndex).Pos.x = x
NPCList(NpcIndex).Pos.y = y

'Find a place to put npc
Call ClosestLegalPos(NPCList(NpcIndex).Pos, TempPos)
If LegalPos(TempPos.map, TempPos.x, TempPos.y) = False Then
    Exit Sub
End If

'Set vars
NPCList(NpcIndex).Pos = TempPos
NPCList(NpcIndex).StartPos = TempPos

'Make NPC Char
Call MakeNPCChar(ToMap, 0, TempPos.map, NpcIndex, TempPos.map, TempPos.x, TempPos.y)

End Sub

Sub KillUser(ByVal userindex As Integer)
'*****************************************************************
'Kill a user
'*****************************************************************
Dim TempPos As WorldPos

'Set user health back to full
UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
UserList(userindex).PoisonName = ""
UserList(userindex).PoisonCount = 0
UserList(userindex).PoisonDamage = 0


'Find a place to put user
Call ClosestLegalPos(ResPos, TempPos)
If LegalPos(TempPos.map, TempPos.x, TempPos.y) = False Then
    Call SendData(ToIndex, userindex, 0, "!!No legal position found: Please try again.")
    CloseUser (userindex)
    Exit Sub
End If

'Warp him there
WarpUserChar userindex, TempPos.map, TempPos.x, TempPos.y

End Sub

Sub UseInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Use/Equip a inventory item
'*****************************************************************
Dim Obj As ObjData
Dim Pos As WorldPos
Dim NPos As WorldPos
Dim OldMap As Integer

Obj = ObjData(UserList(userindex).Object(Slot).ObjIndex)

Select Case Obj.ObjType

    Case OBJTYPE_USEONCE
    
        'use item
        'AddtoVar UserList(UserIndex).Stats.MaxHP, Obj.MaxHP, STAT_MAXHP
        AddtoVar UserList(userindex).Stats.CurHP, Obj.CurHP, UserList(userindex).Stats.MaxHP
        AddtoVar UserList(userindex).Stats.CurMP, Obj.CurMP, UserList(userindex).Stats.MaxMP
        
        'Say they used the damned item
        SendData ToIndex, userindex, 0, "#You used " & Obj.Name & "." & FONTTYPE_TALK
        
        '* Scroll **************
        If Obj.CurHP = 0 And Obj.CurMP = 0 Then
            Pos.x = Obj.Str 'x
            Pos.y = Obj.Con 'y
            Pos.map = Obj.Int
            
            ClosestLegalPos Pos, NPos
            If LegalPos(NPos.map, NPos.x, NPos.y) Then
                
                OldMap = UserList(userindex).Pos.map
                
                Call EraseUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex)
                
                UserList(userindex).Pos.x = NPos.x
                UserList(userindex).Pos.y = NPos.y
                UserList(userindex).Pos.map = NPos.map
                
                If OldMap <> NPos.map Then
                    'Set switchingmap flag
                    UserList(userindex).Flags.SwitchingMaps = 1
                    
                    'Tell client to try switching maps
                    Call SendData(ToIndex, userindex, 0, "SCM" & NPos.map & "," & MapInfo(NPos.map).MapVersion)
                
                    'Update new Map Users
                    MapInfo(NPos.map).NumUsers = MapInfo(NPos.map).NumUsers + 1
                    'Update old Map Users
                    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
                    If MapInfo(OldMap).NumUsers < 0 Then
                        MapInfo(OldMap).NumUsers = 0
                    End If
                    
                    'Show Character to others
                    Call MakeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
                    
                Else
                    
                    Call MakeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
                    Call SendData(ToIndex, userindex, 0, "SUC" & UserList(userindex).Char.CharIndex)
                
                End If
            End If
        End If
        '* Scroll end *****
        
        'Remove from inventory
        UserList(userindex).Object(Slot).Amount = UserList(userindex).Object(Slot).Amount - 1
        If UserList(userindex).Object(Slot).Amount <= 0 Then
            UserList(userindex).Object(Slot).ObjIndex = 0
        End If

    Case OBJTYPE_WEAPON
        
        'If currently equipped remove instead
        If UserList(userindex).Object(Slot).Equipped Then
            RemoveInvItem userindex, Slot
            Exit Sub
        End If
        
        'Remove old item if exists
        If UserList(userindex).WeaponEqpObjIndex > 0 Then
            RemoveInvItem userindex, UserList(userindex).WeaponEqpSlot
        End If

        'Equip
        If UserList(userindex).Stats.Lv >= Obj.MinLv Then
            UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + Obj.MaxHIT
            UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + Obj.MinHIT
        
            UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + Obj.MaxHP
            UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP + Obj.MaxMP
            UserList(userindex).Stats.AC = UserList(userindex).Stats.AC + Obj.AC
            UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam + Obj.Dam
            UserList(userindex).Stats.Str = UserList(userindex).Stats.Str + Obj.Str
            UserList(userindex).Stats.Con = UserList(userindex).Stats.Con + Obj.Con
            UserList(userindex).Stats.Int = UserList(userindex).Stats.Int + Obj.Int
            UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis + Obj.Wis
            UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex + Obj.Dex
    
            UserList(userindex).Object(Slot).Equipped = 1
            UserList(userindex).WeaponEqpObjIndex = UserList(userindex).Object(Slot).ObjIndex
            UserList(userindex).WeaponEqpSlot = Slot
            
            UserList(userindex).Char.Weapon = Obj.Graphic
            Call ChangeUserChar(ToMap, userindex, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
            
            SendData ToIndex, userindex, 0, "#Weapon: " & Obj.Name & FONTTYPE_TALK
        Else
            SendData ToIndex, userindex, 0, "#You need experience to lift this." & FONTTYPE_TALK
        End If

    Case OBJTYPE_ARMOR

        'If currently equipped remove instead
        If UserList(userindex).Object(Slot).Equipped Then
            RemoveInvItem userindex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(userindex).ArmourEqpObjIndex > 0 Then
            RemoveInvItem userindex, UserList(userindex).ArmourEqpSlot
        End If

        'Equip
        If UserList(userindex).Stats.Lv >= Obj.MinLv Then
            UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + Obj.MaxHP
            UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP + Obj.MaxMP
            UserList(userindex).Stats.AC = UserList(userindex).Stats.AC + Obj.AC
            UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam + Obj.Dam
            UserList(userindex).Stats.Str = UserList(userindex).Stats.Str + Obj.Str
            UserList(userindex).Stats.Con = UserList(userindex).Stats.Con + Obj.Con
            UserList(userindex).Stats.Int = UserList(userindex).Stats.Int + Obj.Int
            UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis + Obj.Wis
            UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex + Obj.Dex
            
            UserList(userindex).Object(Slot).Equipped = 1
            UserList(userindex).ArmourEqpObjIndex = UserList(userindex).Object(Slot).ObjIndex
            UserList(userindex).ArmourEqpSlot = Slot
            
            UserList(userindex).Char.Body = Obj.Body
            Call ChangeUserChar(ToMap, userindex, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
            
            SendData ToIndex, userindex, 0, "#Armor: " & Obj.Name & FONTTYPE_TALK
        Else
            SendData ToIndex, userindex, 0, "#You are too young." & FONTTYPE_TALK
        End If
    
    Case OBJTYPE_ACC

        'If currently equipped remove instead
        If UserList(userindex).Object(Slot).Equipped Then
            RemoveInvItem userindex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(userindex).AccEqpObjIndex > 0 Then
            RemoveInvItem userindex, UserList(userindex).AccEqpSlot
        End If
        
        If Obj.uPath <> UserList(userindex).Path And Obj.uPath <> "All" Then
            SendData ToIndex, userindex, 0, "#Your path forbids such an item." & FONTTYPE_TALK
            SendData ToIndex, userindex, 0, "#Path: " & Obj.uPath & FONTTYPE_TALK
            
            Exit Sub
        End If
        
        'Equip
        If UserList(userindex).Stats.Lv >= Obj.MinLv Then
            UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + Obj.MaxHP
            UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP + Obj.MaxMP
            UserList(userindex).Stats.AC = UserList(userindex).Stats.AC + Obj.AC
            UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam + Obj.Dam
            UserList(userindex).Stats.Str = UserList(userindex).Stats.Str + Obj.Str
            UserList(userindex).Stats.Con = UserList(userindex).Stats.Con + Obj.Con
            UserList(userindex).Stats.Int = UserList(userindex).Stats.Int + Obj.Int
            UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis + Obj.Wis
            UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex + Obj.Dex
            
            UserList(userindex).Object(Slot).Equipped = 1
            UserList(userindex).AccEqpObjIndex = UserList(userindex).Object(Slot).ObjIndex
            UserList(userindex).AccEqpSlot = Slot
            SendData ToIndex, userindex, 0, "#Accessory: " & Obj.Name & FONTTYPE_TALK
        Else
            SendData ToIndex, userindex, 0, "#You are too young." & FONTTYPE_TALK
        End If
    
    Case OBJTYPE_HELM
        'If currently equipped remove instead
        If UserList(userindex).Object(Slot).Equipped Then
            RemoveInvItem userindex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(userindex).HelmEqpObjIndex > 0 Then
            RemoveInvItem userindex, UserList(userindex).HelmEqpSlot
        End If
        
        'If Obj.uPath <> UserList(UserIndex).Path And Obj.uPath <> "All" Then
        '    SendData ToIndex, UserIndex, 0, "#Your path forbids such an item." & FONTTYPE_TALK
        '    SendData ToIndex, UserIndex, 0, "#Path: " & Obj.uPath & FONTTYPE_TALK
        '
        '    Exit Sub
        'End If
        
        'Equip
        If UserList(userindex).Stats.Lv >= Obj.MinLv Then
            UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + Obj.MaxHP
            UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP + Obj.MaxMP
            UserList(userindex).Stats.AC = UserList(userindex).Stats.AC + Obj.AC
            UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam + Obj.Dam
            UserList(userindex).Stats.Str = UserList(userindex).Stats.Str + Obj.Str
            UserList(userindex).Stats.Con = UserList(userindex).Stats.Con + Obj.Con
            UserList(userindex).Stats.Int = UserList(userindex).Stats.Int + Obj.Int
            UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis + Obj.Wis
            UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex + Obj.Dex
            
            UserList(userindex).Object(Slot).Equipped = 1
            UserList(userindex).HelmEqpObjIndex = UserList(userindex).Object(Slot).ObjIndex
            UserList(userindex).HelmEqpSlot = Slot
            SendData ToIndex, userindex, 0, "#Helm: " & Obj.Name & FONTTYPE_TALK
        Else
            SendData ToIndex, userindex, 0, "#You are too young." & FONTTYPE_TALK
        End If

End Select

'Update user's stats and inventory
SendUserStatsBox userindex
UpdateUserInv True, userindex, 0

End Sub

Sub AddtoVar(Var As Variant, ByVal Addon As Variant, ByVal Max As Variant)
'*****************************************************************
'Adds a value to a variable respecting a max value
'*****************************************************************

If Var >= Max Then
    Var = Max
    Exit Sub
End If

Var = Var + Addon
If Var > Max Then
    Var = Max
End If

End Sub

Sub RemoveInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Unequip a inventory item
'*****************************************************************

Dim Obj As ObjData

If Slot = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(userindex).Object(Slot).ObjIndex)


Select Case Obj.ObjType


    Case OBJTYPE_WEAPON

        UserList(userindex).Stats.MaxHIT = 3
        UserList(userindex).Stats.MinHIT = 1
        
        UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP - Obj.MaxHP
        UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP - Obj.MaxMP
        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - Obj.AC
        UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam - Obj.Dam
        UserList(userindex).Stats.Str = UserList(userindex).Stats.Str - Obj.Str
        UserList(userindex).Stats.Con = UserList(userindex).Stats.Con - Obj.Con
        UserList(userindex).Stats.Int = UserList(userindex).Stats.Int - Obj.Int
        UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis - Obj.Wis
        UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex - Obj.Dex

        UserList(userindex).Object(Slot).Equipped = 0
        UserList(userindex).WeaponEqpObjIndex = 0
        UserList(userindex).WeaponEqpSlot = 0
        
        UserList(userindex).Char.Weapon = 0
        Call ChangeUserChar(ToMap, userindex, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
        
        If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
        If UserList(userindex).Stats.CurMP > UserList(userindex).Stats.MaxMP Then UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
        SendData ToIndex, userindex, 0, "#You remove " & Obj.Name & "." & FONTTYPE_TALK

    Case OBJTYPE_ARMOR
        
        UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP - Obj.MaxHP
        UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP - Obj.MaxMP
        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - Obj.AC
        UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam - Obj.Dam
        UserList(userindex).Stats.Str = UserList(userindex).Stats.Str - Obj.Str
        UserList(userindex).Stats.Con = UserList(userindex).Stats.Con - Obj.Con
        UserList(userindex).Stats.Int = UserList(userindex).Stats.Int - Obj.Int
        UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis - Obj.Wis
        UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex - Obj.Dex

        UserList(userindex).Object(Slot).Equipped = 0
        UserList(userindex).ArmourEqpObjIndex = 0
        UserList(userindex).ArmourEqpSlot = 0
        
        UserList(userindex).Char.Body = 1
        Call ChangeUserChar(ToMap, userindex, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
        
        If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
        If UserList(userindex).Stats.CurMP > UserList(userindex).Stats.MaxMP Then UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
        SendData ToIndex, userindex, 0, "#You remove " & Obj.Name & "." & FONTTYPE_TALK
    
    Case OBJTYPE_ACC

        UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP - Obj.MaxHP
        UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP - Obj.MaxMP
        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - Obj.AC
        UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam - Obj.Dam
        UserList(userindex).Stats.Str = UserList(userindex).Stats.Str - Obj.Str
        UserList(userindex).Stats.Con = UserList(userindex).Stats.Con - Obj.Con
        UserList(userindex).Stats.Int = UserList(userindex).Stats.Int - Obj.Int
        UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis - Obj.Wis
        UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex - Obj.Dex

        UserList(userindex).Object(Slot).Equipped = 0
        UserList(userindex).AccEqpObjIndex = 0
        UserList(userindex).AccEqpSlot = 0
        
        If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
        If UserList(userindex).Stats.CurMP > UserList(userindex).Stats.MaxMP Then UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
        SendData ToIndex, userindex, 0, "#You remove " & Obj.Name & "." & FONTTYPE_TALK

    Case OBJTYPE_HELM

        UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP - Obj.MaxHP
        UserList(userindex).Stats.MaxMP = UserList(userindex).Stats.MaxMP - Obj.MaxMP
        UserList(userindex).Stats.AC = UserList(userindex).Stats.AC - Obj.AC
        UserList(userindex).Stats.Dam = UserList(userindex).Stats.Dam - Obj.Dam
        UserList(userindex).Stats.Str = UserList(userindex).Stats.Str - Obj.Str
        UserList(userindex).Stats.Con = UserList(userindex).Stats.Con - Obj.Con
        UserList(userindex).Stats.Int = UserList(userindex).Stats.Int - Obj.Int
        UserList(userindex).Stats.Wis = UserList(userindex).Stats.Wis - Obj.Wis
        UserList(userindex).Stats.Dex = UserList(userindex).Stats.Dex - Obj.Dex

        UserList(userindex).Object(Slot).Equipped = 0
        UserList(userindex).HelmEqpObjIndex = 0
        UserList(userindex).HelmEqpSlot = 0
        
        If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
        If UserList(userindex).Stats.CurMP > UserList(userindex).Stats.MaxMP Then UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
        SendData ToIndex, userindex, 0, "#You remove " & Obj.Name & "." & FONTTYPE_TALK


End Select

SendUserStatsBox userindex
UpdateUserInv True, userindex, 0

End Sub

Function NextOpenCharIndex() As Integer
'*****************************************************************
'Finds the next open CharIndex in Charlist
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer
'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).Flags.UserLogged = 0
    LoopC = LoopC + 1
Loop
  
NextOpenUser = LoopC

End Function

Function NextOpenNPC() As Integer
'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until NPCList(LoopC).Flags.NPCActive = 0
    LoopC = LoopC + 1
Loop
  
NextOpenNPC = LoopC

End Function

Sub ClosestLegalPos(Pos As WorldPos, NPos As WorldPos)
'*****************************************************************
'Finds the closest legal tile to Pos and stores it in nPos
'*****************************************************************
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

NPos.map = Pos.map

Do While LegalPos(Pos.map, NPos.x, NPos.y) = False
    
    If LoopC > 10 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.y - LoopC To Pos.y + LoopC
        For tX = Pos.x - LoopC To Pos.x + LoopC
        
            If LegalPos(NPos.map, tX, tY) = True Then
                'Check to see if its an exit
                If MapData(NPos.map, tX, tY).TileExit.map = 0 Then
                    NPos.x = tX
                    NPos.y = tY
                    tX = Pos.x + LoopC
                    tY = Pos.y + LoopC
                End If
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    NPos.x = 0
    NPos.y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
'*****************************************************************
'Searches userlist for a name and return userindex
'*****************************************************************
Dim userindex As Integer
  
'check for bad name
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
  
userindex = 1
Do Until UCase(Left$(UserList(userindex).Name, Len(Name))) = UCase(Name)
    
    userindex = userindex + 1
    
    If userindex > LastUser Then
        userindex = 0
        Exit Do
    End If
    
Loop
  
NameIndex = userindex

End Function

Sub NPCAI(ByVal NpcIndex As Integer)
'*****************************************************************
'Moves NPC based on it's .movement value
'*****************************************************************
Dim NPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer

'Look for someone to attack if hostile
If NPCList(NpcIndex).Hostile Then

    'Check in all directions
    For HeadingLoop = NORTH To WEST
        NPos = NPCList(NpcIndex).Pos
        HeadtoPos HeadingLoop, NPos
        
        'if a legal pos and a user is found attack
        If InMapBounds(NPos.map, NPos.x, NPos.y) Then
            If MapData(NPos.map, NPos.x, NPos.y).userindex > 0 Then
                'Face NPC to target
                ChangeNPCChar ToMap, 0, NPos.map, NpcIndex, NPCList(NpcIndex).Char.Body, NPCList(NpcIndex).Char.Head, HeadingLoop
                'Attack
                NPCAttackUser NpcIndex, MapData(NPos.map, NPos.x, NPos.y).userindex
                'Don't move if fighting
                Exit Sub
            End If
        End If
        
    Next HeadingLoop
End If


'Movement
Select Case NPCList(NpcIndex).Movement

    'Stand
    Case 1
        'Do nothing
        
    'Move randomly
    Case 2
        Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))

    'Go towards any nearby Users
    Case 3
        For y = NPCList(NpcIndex).Pos.y - 6 To NPCList(NpcIndex).Pos.y + 6    'Makes a loop that looks at
            For x = NPCList(NpcIndex).Pos.x - 8 To NPCList(NpcIndex).Pos.x + 8   '6x8 tiles in every direction

                'Make sure tile is legal
                If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                
                    'look for a user
                    If MapData(NPCList(NpcIndex).Pos.map, x, y).userindex > 0 Then
                        'Move towards user
                        tHeading = FindDirection(NPCList(NpcIndex).Pos, UserList(MapData(NPCList(NpcIndex).Pos.map, x, y).userindex).Pos)
                        
                        MoveNPCChar NpcIndex, tHeading
                        'Leave sub
                        Exit Sub
                    End If
                    
                End If
                     
            Next x
        Next y

End Select

End Sub

Function OpenNPC(ByVal NPCNumber As Integer) As Integer
'*****************************************************************
'Loads a NPC from the npc Data and returns its index
'*****************************************************************
Dim NpcIndex As Integer
Dim NPCFile As String

'Set NPC file
NPCFile = IniPath & "NPC.dat"

'Find next open NPCindex
NpcIndex = NextOpenNPC

NPCList(NpcIndex) = NPCData(NPCNumber)
NPCList(NpcIndex).Counters.Movement = Int(Rnd * NPCList(NpcIndex).Speed)
If NPCList(NpcIndex).Hostile = 0 Then
    NPCList(NpcIndex).Shop = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Shop"))
    If NPCList(NpcIndex).Shop > 0 Then
        Call OpenNPCShop(NpcIndex)
    End If
End If


'Setup NPC
NPCList(NpcIndex).Flags.NPCActive = 1

'Update NPC counters
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1

'Return new NPCIndex
OpenNPC = NpcIndex

'LogData = "Npc " & NPCList(NpcIndex).Name & " created at " & Time$ & FONTTYPE_TALK
'AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


End Function

Function OpenNPCShop(ByVal NpcIndex As Integer)
Dim ShopFile As String
Dim TotalItems As Integer
Dim LoopI As Integer

ShopFile = IniPath & "Shop" & NPCList(NpcIndex).Shop & ".dat"

NpcShops(NpcIndex).SayCaption = GetVar(ShopFile, "init", "Caption")
TotalItems = Val(GetVar(ShopFile, "init", "ShopSlots"))

For LoopI = 1 To TotalItems
    NpcShops(NpcIndex).Slots(LoopI).Func = GetVar(ShopFile, "slot" & LoopI, "Func")
    NpcShops(NpcIndex).Slots(LoopI).FuncName = GetVar(ShopFile, "slot" & LoopI, "FuncName")
    NpcShops(NpcIndex).Slots(LoopI).Gold = Val(GetVar(ShopFile, "slot" & LoopI, "Gold"))
    NpcShops(NpcIndex).Slots(LoopI).Item = Val(GetVar(ShopFile, "slot" & LoopI, "Item"))
    NpcShops(NpcIndex).Slots(LoopI).Path = GetVar(ShopFile, "slot" & LoopI, "Path")
    NpcShops(NpcIndex).Slots(LoopI).Level = Val(GetVar(ShopFile, "slot" & LoopI, "Level"))
    
Next LoopI

'LogData = "Npc Shop " & NPCList(NpcIndex).Shop & " loaded at " & Time$ & FONTTYPE_TALK
'AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


End Function


Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal Num As Integer, ByVal map As Byte, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Erase a object
'*****************************************************************

MapData(map, x, y).ObjInfo.Amount = MapData(map, x, y).ObjInfo.Amount - Num

If MapData(map, x, y).ObjInfo.Amount <= 0 Then
    MapData(map, x, y).ObjInfo.ObjIndex = 0
    MapData(map, x, y).ObjInfo.Amount = 0
    Call SendData(sndRoute, sndIndex, sndMap, "EOB" & x & "," & y)
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'For dropping
'*****************************************************************

If MapData(map, x, y).ObjInfo.ObjIndex = Obj.ObjIndex Then
    MapData(map, x, y).ObjInfo.Amount = MapData(map, x, y).ObjInfo.Amount + Obj.Amount
    Call SendData(sndRoute, sndIndex, sndMap, "MOB" & ObjData(Obj.ObjIndex).GRHIndex & "," & x & "," & y)
Else
    MapData(map, x, y).ObjInfo = Obj
    Call SendData(sndRoute, sndIndex, sndMap, "MOB" & ObjData(Obj.ObjIndex).GRHIndex & "," & x & "," & y)
End If

End Sub

Sub DropGold(ByVal userindex As Integer, ByVal Amount As Long)
    'SendData ToIndex, userindex, 0, "#You used " & Amount & "." & FONTTYPE_TALK
    If UserList(userindex).Stats.Gold < Amount Then
        Exit Sub
    End If
    
    If Amount < 1 Then
        Exit Sub
    End If
    
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold - Amount
    
    If MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).Gold > 0 Then
        MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).Gold = MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).Gold + Amount
    Else
        MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).Gold = Amount
    End If
    Call SendData(ToMap, userindex, UserList(userindex).Pos.map, "MAG" & "*" & "," & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)
    SendData ToIndex, userindex, 0, "#You drop " & Amount & " gold." & FONTTYPE_TALK
End Sub

Sub MakeGold(ByVal Amount As Long, ByVal x As Integer, ByVal y As Integer, ByVal map As Integer, ByVal NpcIndex As Integer)
    If MapData(map, x, y).Gold > 0 Then
        MapData(map, x, y).Gold = MapData(map, x, y).Gold + Amount
    Else
        MapData(map, x, y).Gold = Amount
    End If
    Call SendData(ToMap, NpcIndex, map, "MAG" & "*" & "," & x & "," & y)
End Sub

Sub MapMakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'For the server so we don't dupe items
'*****************************************************************

MapData(map, x, y).ObjInfo = Obj
Call SendData(sndRoute, sndIndex, sndMap, "MOB" & ObjData(Obj.ObjIndex).GRHIndex & "," & x & "," & y)

End Sub

Sub GetObj(ByVal userindex As Integer)
'*****************************************************************
'Puts a object in a User's slot from the current User's position
'*****************************************************************

Dim x As Integer
Dim y As Integer
Dim Slot As Byte
Dim PickUp As Long

x = UserList(userindex).Pos.x
y = UserList(userindex).Pos.y

If MapData(UserList(userindex).Pos.map, x, y).Gold > 0 Then
    PickUp = MapData(UserList(userindex).Pos.map, x, y).Gold
    MapData(UserList(userindex).Pos.map, x, y).Gold = 0
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold + PickUp
    Call SendData(ToMap, userindex, UserList(userindex).Pos.map, "MAG" & "x" & "," & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)
End If

'Check for object on ground
If MapData(UserList(userindex).Pos.map, x, y).ObjInfo.ObjIndex <= 0 Then
    'Call SendData(ToIndex, UserIndex, 0, "#Nothing there." & FONTTYPE_INFO)
    Exit Sub
End If

'Check to see if User already has object type
Slot = 1
Do Until UserList(userindex).Object(Slot).ObjIndex = MapData(UserList(userindex).Pos.map, x, y).ObjInfo.ObjIndex
    Slot = Slot + 1

    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Else check if there is a empty slot
If Slot > MAX_INVENTORY_SLOTS Then
    Slot = 1
    Do Until UserList(userindex).Object(Slot).ObjIndex = 0
        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
            Call SendData(ToIndex, userindex, 0, "#Can't Hold anymore." & FONTTYPE_TALK)
            Exit Sub
            Exit Do
        End If
    Loop
End If

'Fill object slot
If UserList(userindex).Object(Slot).Amount + MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount <= MAX_INVENTORY_OBJS Then
    'Under MAX_INV_OBJS
    UserList(userindex).Object(Slot).ObjIndex = MapData(UserList(userindex).Pos.map, x, y).ObjInfo.ObjIndex
    UserList(userindex).Object(Slot).Amount = UserList(userindex).Object(Slot).Amount + MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount
    Call EraseObj(ToMap, 0, UserList(userindex).Pos.map, MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
Else
    'Over MAX_INV_OBJS
    If MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount < UserList(userindex).Object(Slot).Amount Then
        MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount = Abs(MAX_INVENTORY_OBJS - (UserList(userindex).Object(Slot).Amount + MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount))
    Else
        MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount = Abs((MAX_INVENTORY_OBJS + UserList(userindex).Object(Slot).Amount) - MapData(UserList(userindex).Pos.map, x, y).ObjInfo.Amount)
    End If
    UserList(userindex).Object(Slot).Amount = MAX_INVENTORY_OBJS
    Call SendData(ToIndex, userindex, 0, "#You can't carry more." & FONTTYPE_TALK)
End If

Call UpdateUserInv(False, userindex, Slot)

End Sub

Sub SellUserItem(ByVal userindex As Integer, ByVal ShopIndex As Integer, ByVal xItem As Integer)
Dim Slot As Byte
Dim ItemCost As Long

ItemCost = NpcShops(ShopIndex).Slots(xItem).Gold

If UserList(userindex).Stats.Gold < ItemCost Then
    Call SendData(ToIndex, userindex, 0, "#You can't afford this." & FONTTYPE_TALK)
    Exit Sub
End If

'Check to see if User already has object type
Slot = 1
Do Until UserList(userindex).Object(Slot).ObjIndex = NpcShops(ShopIndex).Slots(xItem).Item
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(userindex).Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, userindex, 0, "#Your Inventory is full." & FONTTYPE_TALK)
                Exit Sub
                Exit Do
            End If
        Loop
End If

'Fill object slot
If UserList(userindex).Object(Slot).Amount + 1 <= MAX_INVENTORY_OBJS Then
    'Under MAX_INV_OBJS
    UserList(userindex).Object(Slot).ObjIndex = NpcShops(ShopIndex).Slots(xItem).Item
    UserList(userindex).Object(Slot).Amount = UserList(userindex).Object(Slot).Amount + 1
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold - ItemCost
    Call SendData(ToIndex, userindex, 0, "#" & ObjData(UserList(userindex).Object(Slot).ObjIndex).Name & " (" & UserList(userindex).Object(Slot).Amount & ")" & FONTTYPE_TALK)
Else
    Call SendData(ToIndex, userindex, 0, "#You can't hold more." & FONTTYPE_TALK)
End If

Call UpdateUserInv(False, userindex, Slot)

End Sub

Sub SellNPCItem(ByVal userindex As Integer, ByVal ShopIndex As Integer, ByVal xItem As Integer)
Dim Slot As Byte
Dim ItemCost As Integer

ItemCost = NpcShops(ShopIndex).Slots(xItem).Gold

'Check to see if User already has object type
Slot = 1
Do Until UserList(userindex).Object(Slot).ObjIndex = NpcShops(ShopIndex).Slots(xItem).Item
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

If Slot > MAX_INVENTORY_SLOTS Then
    Call SendData(ToIndex, userindex, 0, "#You don't have that item." & FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).Object(Slot).Equipped = 1 Then
    Call SendData(ToIndex, userindex, 0, "#I can't buy the clothes off your body." & FONTTYPE_TALK)
    Exit Sub
End If

'Remove item
If UserList(userindex).Object(Slot).Amount - 1 >= 1 Then
    UserList(userindex).Object(Slot).ObjIndex = NpcShops(ShopIndex).Slots(xItem).Item
    UserList(userindex).Object(Slot).Amount = UserList(userindex).Object(Slot).Amount - 1
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold + ItemCost
    Call SendData(ToIndex, userindex, 0, "#" & ObjData(UserList(userindex).Object(Slot).ObjIndex).Name & "(" & UserList(userindex).Object(Slot).Amount & ")" & FONTTYPE_TALK)
Else
    UserList(userindex).Object(Slot).ObjIndex = 0
    UserList(userindex).Object(Slot).Amount = 0
    UserList(userindex).Stats.Gold = UserList(userindex).Stats.Gold + ItemCost
End If

UserList(userindex).Flags.StatsChanged = True
Call UpdateUserInv(False, userindex, Slot)

End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Updates a User's inventory
'*****************************************************************
Dim NullObj As UserOBJ
Dim LoopC As Byte

'Update one slot
If UpdateAll = False Then

    'Update User inventory
    If UserList(userindex).Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(userindex, Slot, UserList(userindex).Object(Slot))
    Else
        Call ChangeUserInv(userindex, Slot, NullObj)
    End If

Else

'Update every slot
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Update User invetory
        If UserList(userindex).Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(userindex, LoopC, UserList(userindex).Object(LoopC))
        Else
            Call ChangeUserInv(userindex, LoopC, NullObj)
        End If

    Next LoopC

End If

End Sub

Sub UpdateUserSpell(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Updates a User's spells
'*****************************************************************
Dim NullObj As uSpellBook
Dim LoopC As Byte

'Update one slot
If UpdateAll = False Then

    'Update User inventory
    If UserList(userindex).Object(Slot).ObjIndex > 0 Then
        Call ChangeUserSpell(userindex, Slot, UserList(userindex).SpellBook(Slot))
    Else
        Call ChangeUserSpell(userindex, Slot, NullObj)
    End If

Else

'Update every slot
    For LoopC = 1 To MAX_SPELL_SLOTS

        'Update User invetory
        If UserList(userindex).SpellBook(LoopC).Spellindex > 0 Then
            Call ChangeUserSpell(userindex, LoopC, UserList(userindex).SpellBook(LoopC))
        Else
            Call ChangeUserSpell(userindex, LoopC, NullObj)
        End If

    Next LoopC

End If


'Sub ChangeUserSpell(ByVal UserIndex As Integer, ByVal Slot As Byte, Object As UserOBJ)
End Sub

Sub ChangeUserInv(ByVal userindex As Integer, ByVal Slot As Byte, Object As UserOBJ)
'*****************************************************************
'Changes a user's inventory
'*****************************************************************

UserList(userindex).Object(Slot) = Object

If Object.ObjIndex > 0 Then

    If UserList(userindex).AccEqpSlot = Slot Then
        Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex & "," & "ACC")
    ElseIf UserList(userindex).ArmourEqpSlot = Slot Then
        Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex & "," & "ARMOR")
    ElseIf UserList(userindex).HelmEqpSlot = Slot Then
        Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex & "," & "HELM")
    ElseIf UserList(userindex).WeaponEqpSlot = Slot Then
        Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex & "," & "WEAPON")
    Else
        Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex & "," & "*")
    End If

Else

    Call SendData(ToIndex, userindex, 0, "SIS" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If
End Sub

Sub ChangeUserSpell(ByVal userindex As Integer, ByVal Slot As Byte, Object As uSpellBook)
'*****************************************************************
'Changes a user's spells
'*****************************************************************
Dim SType As Integer

UserList(userindex).SpellBook(Slot) = Object

If Object.Spellindex > 0 Then

    SType = SpellData(Object.Spellindex).SpellType
    If SType = 2 Or SType = 3 Or SType = 4 Or SType = 7 Or SType = 8 Or SType = 9 Then  ' 234789
        Call SendData(ToIndex, userindex, 0, "SSS" & Slot & "," & SpellData(Object.Spellindex).Name & "," & SpellData(Object.Spellindex).GRHIndex & "," & Object.Spellindex & "," & "T" & "," & SpellData(Object.Spellindex).Icon)
    Else
        Call SendData(ToIndex, userindex, 0, "SSS" & Slot & "," & SpellData(Object.Spellindex).Name & "," & SpellData(Object.Spellindex).GRHIndex & "," & Object.Spellindex & "," & "X" & "," & SpellData(Object.Spellindex).Icon)
    End If

Else

    Call SendData(ToIndex, userindex, 0, "SSS" & Slot & "," & "" & "," & "0" & "," & "X" & "," & "0")

End If


End Sub

Sub DropObj(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Drops a object from a User's slot
'*****************************************************************
Dim Obj As Obj

'Check amount
If Num <= 0 Then
    Exit Sub
End If

If Num > UserList(userindex).Object(Slot).Amount Then
    Num = UserList(userindex).Object(Slot).Amount
End If

'Check for object on gorund
If MapData(UserList(userindex).Pos.map, x, y).ObjInfo.ObjIndex <> 0 Then
    If MapData(UserList(userindex).Pos.map, x, y).ObjInfo.ObjIndex <> UserList(userindex).Object(Slot).ObjIndex Then
        Call SendData(ToIndex, userindex, 0, "#No room on ground." & FONTTYPE_TALK)
        Exit Sub
    End If
End If

Obj.ObjIndex = UserList(userindex).Object(Slot).ObjIndex
Obj.Amount = Num
Call MakeObj(ToMap, 0, map, Obj, map, x, y)

'Remove object
UserList(userindex).Object(Slot).Amount = UserList(userindex).Object(Slot).Amount - Num
If UserList(userindex).Object(Slot).Amount <= 0 Then
    
    'Unequip is the object is currently equipped
    If UserList(userindex).Object(Slot).Equipped = 1 Then
        Call RemoveInvItem(userindex, Slot)
    End If
    
    UserList(userindex).Object(Slot).ObjIndex = 0
    UserList(userindex).Object(Slot).Amount = 0
    UserList(userindex).Object(Slot).Equipped = 0
End If

Call UpdateUserInv(False, userindex, Slot)

End Sub

Sub CloseNPC(ByVal NpcIndex As Integer)
'*****************************************************************
'Closes a NPC
'*****************************************************************

NPCList(NpcIndex).Flags.NPCActive = 0

'update last npc
If NpcIndex = LastNPC Then
    Do Until NPCList(LastNPC).Flags.NPCActive = 1
        LastNPC = LastNPC - 1
        If LastNPC = 0 Then Exit Do
    Loop
End If
  
'update number of users
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If

End Sub

Sub UserAttackNPC(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal SType As Integer)
'*****************************************************************
'Have a User attack a NPC
'*****************************************************************
Dim Hit As Integer
Dim MinHitVal As Integer
Dim MaxHitVal As Integer


'Calculate hit
MinHitVal = (UserList(userindex).Stats.Str * 2) + UserList(userindex).Stats.MinHIT
MaxHitVal = (UserList(userindex).Stats.Str * 2) + UserList(userindex).Stats.MaxHIT

Hit = Int(RandomNumber(MinHitVal, MaxHitVal))
Hit = Hit + (Hit * (UserList(userindex).Stats.Dam * 0.1))
Hit = Hit * ((NPCList(NpcIndex).Stats.AC + 110) / 200)

If SType = SIDE Then
    If UserList(userindex).Path = "Warrior" Then
        'If UserList(userindex).Stats.Lv <= 25 Then Hit = Hit * 0.25
        'If UserList(userindex).Stats.Lv > 25 And UserList(userindex).Stats.Lv <= 50 Then Hit = Hit * 0.5
        'If UserList(userindex).Stats.Lv > 50 And UserList(userindex).Stats.Lv < 99 Then Hit = Hit * 0.75
        'If UserList(userindex).Stats.Lv = 99 Then
        Hit = Hit * 0.8
    Else
        Exit Sub
    End If
End If
If SType = BACK Then
    If UserList(userindex).Path = "Warrior" Then
        'If UserList(userindex).Stats.Lv <= 25 Then Hit = Hit * 0.1
        'If UserList(userindex).Stats.Lv > 25 And UserList(userindex).Stats.Lv <= 50 Then Hit = Hit * 0.25
        'If UserList(userindex).Stats.Lv > 50 And UserList(userindex).Stats.Lv < 99 Then Hit = Hit * 0.5
        'If UserList(userindex).Stats.Lv = 99 Then
        Hit = Hit * 0.6
    Else
        Exit Sub
    End If
End If

If Hit < 1 Then Hit = 1

'Hit NPC
'SendData ToIndex, UserIndex, 0, "#Damage: " & Hit & FONTTYPE_TALK
NPCList(NpcIndex).Stats.CurHP = NPCList(NpcIndex).Stats.CurHP - Hit

'NPC Die
If NPCList(NpcIndex).Stats.CurHP <= 0 Then
    Call GiveExp(userindex, True, NPCList(NpcIndex).GiveExp)
    NpcDrops (NpcIndex)
    KillNPC NpcIndex
End If

End Sub

Sub NpcDrops(ByVal NpcIndex As Integer)
'Drops any items the npc has
Dim DropItem As Obj
Dim Chance As Integer

If NPCList(NpcIndex).DropItem > 0 Then
    Chance = Rnd * 100
    If Chance <= NPCList(NpcIndex).DropChance Then
        DropItem.Amount = 1
        DropItem.ObjIndex = NPCList(NpcIndex).DropItem
        Call MakeObj(ToMap, 0, NPCList(NpcIndex).Pos.map, DropItem, NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y)
    End If
End If

If NPCList(NpcIndex).GiveGold > 0 Then
    Chance = Rnd * 100
    If Chance <= NPCList(NpcIndex).DropGoldChance Then
        Call MakeGold(NPCList(NpcIndex).GiveGold, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y, NPCList(NpcIndex).Pos.map, NpcIndex)
    End If
End If

End Sub

Sub GiveExp(ByVal userindex As Integer, ByVal togroup As Boolean, ByVal GiveExp As Long)
'Gives exp to player(s)
Dim TmpI As Integer
Dim Gtotal As Integer
Dim DividedExp As Long
Dim ScaledExp As Long
Dim tIndex As Integer
Dim tLevel As Integer

    
'Give Exp
If UserList(userindex).GIndex = 0 Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + GiveExp
    UserList(userindex).Stats.Texp = UserList(userindex).Stats.Texp + GiveExp
    Call SendData(ToIndex, userindex, 0, "TNL" & UserList(userindex).Stats.Exp & "," & UserList(userindex).Stats.Tnl)
    SendData ToIndex, userindex, 0, "#" & GiveExp & " experience!" & FONTTYPE_TALK
    CheckUserLevel userindex
    UserList(userindex).Flags.StatsChanged = 1
    Exit Sub
End If
If togroup = False Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + GiveExp
    UserList(userindex).Stats.Texp = UserList(userindex).Stats.Texp + GiveExp
    Call SendData(ToIndex, userindex, 0, "TNL" & UserList(userindex).Stats.Exp & "," & UserList(userindex).Stats.Tnl)
    SendData ToIndex, userindex, 0, "#" & GiveExp & " experience!" & FONTTYPE_TALK
    CheckUserLevel userindex
    UserList(userindex).Flags.StatsChanged = 1
    Exit Sub
Else
    For TmpI = 1 To 5
        If Groups(UserList(userindex).GIndex).UIndexes(TmpI) > 0 Then
            Gtotal = Gtotal + 1
            If UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Lv > tLevel Then
                tIndex = Groups(UserList(userindex).GIndex).UIndexes(TmpI)
                tLevel = UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Lv
            End If
        End If
    Next TmpI
    DividedExp = GiveExp - (GiveExp * Gtotal * 0.05)
    For TmpI = 1 To 5
        If Groups(UserList(userindex).GIndex).UIndexes(TmpI) > 0 Then
            If UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Pos.map = UserList(userindex).Pos.map Then
                ScaledExp = DividedExp * (UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Lv / tLevel)
                UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Exp = UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Exp + ScaledExp
                UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Texp = UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Texp + ScaledExp
                Call SendData(ToIndex, Groups(UserList(userindex).GIndex).UIndexes(TmpI), 0, "TNL" & UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Exp & "," & UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Stats.Tnl)
                SendData ToIndex, Groups(UserList(userindex).GIndex).UIndexes(TmpI), 0, "#" & Str$(ScaledExp) & " experience!" & FONTTYPE_TALK
                CheckUserLevel Groups(UserList(userindex).GIndex).UIndexes(TmpI)
                UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Flags.StatsChanged = 1
                CheckUserLevel Groups(UserList(userindex).GIndex).UIndexes(TmpI)
            End If
        End If
    Next TmpI
End If

End Sub


Sub NPCAttackUser(ByVal NpcIndex As Integer, ByVal userindex As Integer)
'*****************************************************************
'Have a NPC attack a User
'*****************************************************************
Dim Hit As Integer
Dim MinHitVal As Integer
Dim MaxHitVal As Integer

'Don't allow if switchingmaps maps
If UserList(userindex).Flags.SwitchingMaps Then
    Exit Sub
End If

'Calculate hit
MinHitVal = NPCList(NpcIndex).Stats.MinHIT
MaxHitVal = NPCList(NpcIndex).Stats.MaxHIT

Hit = Int(RandomNumber(MinHitVal, MaxHitVal))
Hit = Hit * ((UserList(userindex).Stats.AC + 110) / 200)

If UserList(userindex).Path = "Warrior" Then
    Hit = Hit * 0.75
End If

If Hit < 1 Then Hit = 1


'Hit user
'SendData ToIndex, UserIndex, 0, "#" & NPCList(NPCIndex).Name & " " & Hit & "!" & FONTTYPE_TALK
UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP - Hit
SendData ToPCArea, userindex, UserList(userindex).Pos.map, "PLW" & 1
'User Die
If UserList(userindex).Stats.CurHP <= 0 Then
    
    'Kill user
    'SendData ToIndex, UserIndex, 0, "@The " & NPCList(NPCIndex).Name & " kills you!" & FONTTYPE_FIGHT
    If UserList(userindex).Stats.Lv < 99 Then
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp * 0.95
        Call SendData(ToIndex, userindex, 0, "TNL" & UserList(userindex).Stats.Exp & "," & UserList(userindex).Stats.Tnl)
    End If
    KillUser userindex

End If

'Set update stats flag
UserList(userindex).Flags.StatsChanged = 1

End Sub

Sub UserAttackUser(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'*****************************************************************
'Have a user attack a user
'*****************************************************************
Dim Hit As Integer
Dim MinHitVal As Integer
Dim MaxHitVal As Integer

'Don't allow if switchingmaps maps
If UserList(VictimIndex).Flags.SwitchingMaps Then
    Exit Sub
End If

If UserList(AttackerIndex).Flags.PK = "off" Then Exit Sub
If UserList(VictimIndex).Flags.PK = "off" Then Exit Sub

'Calculate hit
MinHitVal = (UserList(AttackerIndex).Stats.Str * 2) + UserList(AttackerIndex).Stats.MinHIT
MaxHitVal = (UserList(AttackerIndex).Stats.Str * 2) + UserList(AttackerIndex).Stats.MaxHIT
Hit = Int(RandomNumber(MinHitVal, MaxHitVal))
Hit = Hit + (Hit * (UserList(AttackerIndex).Stats.Dam * 0.1))
Hit = Hit * ((UserList(VictimIndex).Stats.AC + 110) / 200)

If Hit < 1 Then Hit = 1

'Hit User
'SendData ToIndex, AttackerIndex, 0, "#You hit " & UserList(VictimIndex).Name & "!" & FONTTYPE_TALK
'SendData ToIndex, VictimIndex, 0, "#" & UserList(AttackerIndex).Name & " attacks you!" & FONTTYPE_TALK
UserList(VictimIndex).Stats.CurHP = UserList(VictimIndex).Stats.CurHP - Hit

'User Die
If UserList(VictimIndex).Stats.CurHP <= 0 Then
    
    'Give Exp and gold
    'UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + (UserList(VictimIndex).Stats.Lv * 20)

    'Kill user
    SendData ToIndex, AttackerIndex, 0, "@You kill " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT
    SendData ToIndex, VictimIndex, 0, "@" & UserList(AttackerIndex).Name & " kills you!" & FONTTYPE_FIGHT
    KillUser VictimIndex

End If

'update users level and stats

CheckUserLevel AttackerIndex
'Set update stats flag
UserList(AttackerIndex).Flags.StatsChanged = 1

CheckUserLevel VictimIndex
'Set update stats flag
UserList(VictimIndex).Flags.StatsChanged = 1

End Sub

Sub UserAttack(ByVal userindex As Integer)
'*****************************************************************
'Begin a user attack sequence
'*****************************************************************
Dim AttackPos As WorldPos
Dim HitMonster As Integer

'Check switching maps
If UserList(userindex).Flags.SwitchingMaps Then
    Exit Sub
End If

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_ATTACKWAIT
'UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - 1

'Get tile user is attacking
AttackPos = UserList(userindex).Pos
'HeadtoPos UserList(UserIndex).Char.Heading, AttackPos

'Exit if not legal
'If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
'    SendData ToPCArea, UserIndex, AttackPos.Map, "PLW" & 2
'    Exit Sub
'End If

'Look for user
If UserList(userindex).Char.Heading = NORTH Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).userindex > 0 Then
        HitMonster = HitMonster + 1
        UserAttackUser userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).userindex
        Exit Sub
    End If
End If
If UserList(userindex).Char.Heading = SOUTH Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).userindex > 0 Then
        UserAttackUser userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).userindex
        Exit Sub
    End If
End If
If UserList(userindex).Char.Heading = EAST Then
    If MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).userindex > 0 Then
        HitMonster = HitMonster + 1
        UserAttackUser userindex, MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).userindex
        Exit Sub
    End If
End If
If UserList(userindex).Char.Heading = WEST Then
    If MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).userindex > 0 Then
        HitMonster = HitMonster + 1
        UserAttackUser userindex, MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).userindex
        Exit Sub
    End If
End If

'NPC code
If UserList(userindex).Char.Heading = NORTH Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex, FORWARD
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex, BACK
        End If
    End If
End If
If UserList(userindex).Char.Heading = SOUTH Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex, BACK
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex, FORWARD
        End If
    End If
End If
If UserList(userindex).Char.Heading = EAST Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex, FORWARD
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex, BACK
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex, SIDE
        End If
    End If
End If
If UserList(userindex).Char.Heading = WEST Then
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y - 1).NpcIndex, SIDE
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x + 1, AttackPos.y).NpcIndex, BACK
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x - 1, AttackPos.y).NpcIndex, FORWARD
        End If
    End If
    If MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex > 0 Then
        If NPCList(MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex).Attackable Then
            HitMonster = HitMonster + 1
            UserAttackNPC userindex, MapData(AttackPos.map, AttackPos.x, AttackPos.y + 1).NpcIndex, SIDE
        End If
    End If
End If

If HitMonster = 0 Then
    SendData ToPCArea, userindex, AttackPos.map, "PLW" & 2
Else
    SendData ToPCArea, userindex, AttackPos.map, "PLW" & 1
End If

End Sub

Function userindex(ByVal SocketId As Integer) As Integer
'*****************************************************************
'Finds the User with a certain SocketID
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        userindex = 0
        Exit Function
    End If
    
Loop
  
userindex = LoopC

End Function

Function CheckForSameIP(ByVal userindex As Integer, ByVal UserIP As String) As Boolean
'*****************************************************************
'Checks for a user with the same IP
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastUser

    If UserList(LoopC).Flags.UserLogged = 1 Then
        If UserList(LoopC).IP = UserIP And userindex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If

Next LoopC

CheckForSameIP = False

End Function

Function CheckForSameName(ByVal userindex As Integer, ByVal Name As String) As Boolean
'*****************************************************************
'Checks for a user with the same Name
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastUser

    If UserList(LoopC).Flags.UserLogged = 1 Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) And userindex <> LoopC Then
            CheckForSameName = True
            Exit Function
        End If
    End If

Next LoopC

CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Takes Pos and ad moves it in heading direction
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

x = Pos.x
y = Pos.y

If Head = NORTH Then
    nX = x
    nY = y - 1
End If

If Head = SOUTH Then
    nX = x
    nY = y + 1
End If

If Head = EAST Then
    nX = x + 1
    nY = y
End If

If Head = WEST Then
    nX = x - 1
    nY = y
End If

'return values
Pos.x = nX
Pos.y = nY

End Sub


Sub UpdateUserMap(ByVal userindex As Integer)
'*****************************************************************
'Updates a user with the place of all chars in the Map
'*****************************************************************
Dim map As Integer
Dim x As Integer
Dim y As Integer

map = UserList(userindex).Pos.map

'Place chars
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        If MapData(map, x, y).userindex > 0 Then
            Call MakeUserChar(ToIndex, userindex, 0, MapData(map, x, y).userindex, map, x, y)
        End If

        If MapData(map, x, y).NpcIndex > 0 Then
            Call MakeNPCChar(ToIndex, userindex, 0, MapData(map, x, y).NpcIndex, map, x, y)
        End If

        If MapData(map, x, y).ObjInfo.ObjIndex > 0 Then
            Call MapMakeObj(ToIndex, userindex, 0, MapData(map, x, y).ObjInfo, map, x, y)
        End If
        
        If MapData(map, x, y).Gold > 1 Then
            Call SendData(ToMap, userindex, map, "MAG" & "*" & "," & x & "," & y)
        End If


    Next x
Next y

End Sub

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As Byte)
'*****************************************************************
'Moves a User from one tile to another
'*****************************************************************
Dim NPos As WorldPos

'Move
NPos = UserList(userindex).Pos
Call HeadtoPos(nHeading, NPos)

'Move if legal pos
If LegalPos(UserList(userindex).Pos.map, NPos.x, NPos.y) = True Then
    Call SendData(ToMapButIndex, userindex, UserList(userindex).Pos.map, "MOC" & UserList(userindex).Char.CharIndex & "," & NPos.x & "," & NPos.y)

    'Update map and user pos
    MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).userindex = 0
    UserList(userindex).Pos = NPos
    UserList(userindex).Char.Heading = nHeading
    MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).userindex = userindex
Else
    'else correct user's pos
    Call SendData(ToIndex, userindex, 0, "SUP" & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)
End If

End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
'*****************************************************************
'Moves a NPC from one tile to another
'*****************************************************************
Dim NPos As WorldPos
Dim map As Integer
Dim x As Integer
Dim y As Integer
Dim RandDir As Integer
'Move
NPos = NPCList(NpcIndex).Pos
Call HeadtoPos(nHeading, NPos)

map = NPCList(NpcIndex).Pos.map
x = NPos.x
y = NPos.y

'Move if legal pos
If LegalPos(NPCList(NpcIndex).Pos.map, NPos.x, NPos.y) = True And MapData(map, x, y).TileExit.map = 0 Then
    Call SendData(ToMap, 0, NPCList(NpcIndex).Pos.map, "MOC" & NPCList(NpcIndex).Char.CharIndex & "," & NPos.x & "," & NPos.y)

    'Update map and user pos
    MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = 0
    NPCList(NpcIndex).Pos = NPos
    NPCList(NpcIndex).Char.Heading = nHeading
    MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = NpcIndex
Else
    map = NPCList(NpcIndex).Pos.map
    x = NPCList(NpcIndex).Pos.x
    y = NPCList(NpcIndex).Pos.y
    NPCList(NpcIndex).Char.Heading = nHeading
    If NPCList(NpcIndex).Char.Heading = NORTH Then
        If MapData(map, x + 1, y).NpcIndex = 0 And MapData(map, x + 1, y).userindex = 0 And MapData(map, x + 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, EAST)
        ElseIf MapData(map, x - 1, y).NpcIndex = 0 And MapData(map, x - 1, y).userindex = 0 And MapData(map, x - 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, WEST)
        ElseIf MapData(map, x, y + 1).NpcIndex = 0 And MapData(map, x, y + 1).userindex = 0 And MapData(map, x, y + 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, SOUTH)
        End If
    End If
    
    If NPCList(NpcIndex).Char.Heading = SOUTH Then
        If MapData(map, x + 1, y).NpcIndex = 0 And MapData(map, x + 1, y).userindex = 0 And MapData(map, x + 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, EAST)
        ElseIf MapData(map, x - 1, y).NpcIndex = 0 And MapData(map, x - 1, y).userindex = 0 And MapData(map, x - 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, WEST)
        ElseIf MapData(map, x, y - 1).NpcIndex = 0 And MapData(map, x, y - 1).userindex = 0 And MapData(map, x, y - 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, NORTH)
        End If
    End If
    
    If NPCList(NpcIndex).Char.Heading = EAST Then
        If MapData(map, x, y - 1).NpcIndex = 0 And MapData(map, x, y - 1).userindex = 0 And MapData(map, x, y - 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, NORTH)
        ElseIf MapData(map, x - 1, y).NpcIndex = 0 And MapData(map, x - 1, y).userindex = 0 And MapData(map, x - 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, WEST)
        ElseIf MapData(map, x, y + 1).NpcIndex = 0 And MapData(map, x + 1, y).userindex = 0 And MapData(map, x, y + 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, SOUTH)
        End If
    End If
    
    If NPCList(NpcIndex).Char.Heading = WEST Then
        If MapData(map, x, y - 1).NpcIndex = 0 And MapData(map, x, y - 1).userindex = 0 And MapData(map, x, y - 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, NORTH)
        ElseIf MapData(map, x + 1, y).NpcIndex = 0 And MapData(map, x + 1, y).userindex = 0 And MapData(map, x + 1, y).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, EAST)
        ElseIf MapData(map, x, y + 1).NpcIndex = 0 And MapData(map, x + 1, y).userindex = 0 And MapData(map, x, y + 1).Blocked = 0 And MapData(map, x, y).TileExit.map = 0 Then
            Call MoveNPCChar2(NpcIndex, SOUTH)
        End If
    End If
End If

End Sub

Sub MoveNPCChar2(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
'*****************************************************************
'Moves a NPC from one tile to another on 2nd try
'*****************************************************************
Dim NPos As WorldPos
Dim map As Integer
Dim x As Integer
Dim y As Integer
Dim RandDir As Integer
'Move


NPos = NPCList(NpcIndex).Pos
Call HeadtoPos(nHeading, NPos)

map = NPCList(NpcIndex).Pos.map
x = NPos.x
y = NPos.y

'Move if legal pos
If LegalPos(NPCList(NpcIndex).Pos.map, NPos.x, NPos.y) = True And MapData(map, x, y).TileExit.map = 0 Then
    Call SendData(ToMap, 0, NPCList(NpcIndex).Pos.map, "MOC" & NPCList(NpcIndex).Char.CharIndex & "," & NPos.x & "," & NPos.y)

    'Update map and user pos
    MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = 0
    NPCList(NpcIndex).Pos = NPos
    NPCList(NpcIndex).Char.Heading = nHeading
    MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = NpcIndex
End If

End Sub

Sub MakeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Makes and places a user's character
'*****************************************************************
Dim CharIndex As Integer
Dim HpPerc As Integer

'If needed make a new character in list
If UserList(userindex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    UserList(userindex).Char.CharIndex = CharIndex
    CharList(CharIndex) = userindex
End If

'Place character on map
MapData(map, x, y).userindex = userindex

HpPerc = UserList(userindex).Stats.CurHP / UserList(userindex).Stats.MaxHP * 100

'Send make character command to clients
Call SendData(sndRoute, sndIndex, sndMap, "MAC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & x & "," & y & "," & HpPerc)

End Sub

Sub MakeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Makes and places a NPC character
'*****************************************************************
Dim CharIndex As Integer
Dim HpPerc As Integer

'If needed make a new character in list
If NPCList(NpcIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    NPCList(NpcIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = NpcIndex
End If

'Place character on map
MapData(map, x, y).NpcIndex = NpcIndex

'Set alive flag
NPCList(NpcIndex).Flags.NPCAlive = 1

HpPerc = (NPCList(NpcIndex).Stats.CurHP / NPCList(NpcIndex).Stats.MaxHP) * 100

'Send make character command to clients
Call SendData(sndRoute, sndIndex, sndMap, "MAC" & NPCList(NpcIndex).Char.Body & "," & NPCList(NpcIndex).Char.Head & "," & NPCList(NpcIndex).Char.Heading & "," & NPCList(NpcIndex).Char.CharIndex & "," & x & "," & y & "," & HpPerc)

End Sub

Function LegalPos(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Make sure it's a legal map
If map <= 0 Or map > NumMaps Then
    LegalPos = False
    Exit Function
End If

'Check to see if its out of bounds
If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(map, x, y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'User
If MapData(map, x, y).userindex > 0 Then
    LegalPos = False
    Exit Function
End If

'NPC
If MapData(map, x, y).NpcIndex > 0 Then
    LegalPos = False
    Exit Function
End If

LegalPos = True

End Function

Sub SendHelp(ByVal userindex As Integer)
'*****************************************************************
'Sends help strings to Index
'*****************************************************************
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = Val(GetVar(IniPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(ToIndex, userindex, 0, "@" & GetVar(IniPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Sub EraseUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer)
'*****************************************************************
'Erase a character
'*****************************************************************

'Remove from list
CharList(UserList(userindex).Char.CharIndex) = 0

'Update LsstChar
If UserList(userindex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y).userindex = 0

'Send erase command to clients
Call SendData(ToMap, 0, UserList(userindex).Pos.map, "ERC" & UserList(userindex).Char.CharIndex)

'Update userlist
UserList(userindex).Char.CharIndex = 0

'update NumChars
NumChars = NumChars - 1

End Sub

Sub EraseNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer)
'*****************************************************************
'Erase a character
'*****************************************************************

'Remove from list
CharList(NPCList(NpcIndex).Char.CharIndex) = 0

'Update LastChar
If NPCList(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = 0

'Send erase command to clients
Call SendData(ToMap, 0, NPCList(NpcIndex).Pos.map, "ERC" & NPCList(NpcIndex).Char.CharIndex)

'Update npclist
NPCList(NpcIndex).Char.CharIndex = 0

'Set alive flag
NPCList(NpcIndex).Flags.NPCAlive = 0

'update NumChars
NumChars = NumChars - 1

End Sub

Sub KillNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer)
'*****************************************************************
''kills' a character
'*****************************************************************

'Remove from list
CharList(NPCList(NpcIndex).Char.CharIndex) = 0

'Update LastChar
If NPCList(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(NPCList(NpcIndex).Pos.map, NPCList(NpcIndex).Pos.x, NPCList(NpcIndex).Pos.y).NpcIndex = 0

'Send erase command to clients
Call SendData(ToMap, 0, NPCList(NpcIndex).Pos.map, "KIL" & NPCList(NpcIndex).Char.CharIndex)

'Update npclist
NPCList(NpcIndex).Char.CharIndex = 0

'Set alive flag
NPCList(NpcIndex).Flags.NPCAlive = 0

'update NumChars
NumChars = NumChars - 1

End Sub


Sub LookatTile(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Responds to the user clicking on a square
'*****************************************************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim LoopC As Integer

'Check if legal
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'*** Check for Characters ***
If y + 1 <= YMaxMapSize Then
    If MapData(map, x, y + 1).userindex > 0 Then
        TempCharIndex = MapData(map, x, y + 1).userindex
        FoundChar = 1
    End If
    If MapData(map, x, y + 1).NpcIndex > 0 Then
        TempCharIndex = MapData(map, x, y + 1).NpcIndex
        FoundChar = 2
    End If
End If
'Check for Character
If FoundChar = 0 Then
    If MapData(map, x, y).userindex > 0 Then
        TempCharIndex = MapData(map, x, y).userindex
        FoundChar = 1
    End If
    If MapData(map, x, y).NpcIndex > 0 Then
        TempCharIndex = MapData(map, x, y).NpcIndex
        FoundChar = 2
    End If
End If
'React to character
If FoundChar = 1 Then
        'Call SendData(ToIndex, userindex, 0, "#" & UserList(TempCharIndex).Path & " " & UserList(TempCharIndex).Name & FONTTYPE_TALK)
        Call SendData(ToIndex, userindex, 0, "*CLEAR")
        Call SendData(ToIndex, userindex, 0, "*" & UserList(TempCharIndex).Path & " " & UserList(TempCharIndex).Name & FONTTYPE_SHOUT)
        Call SendData(ToIndex, userindex, 0, "*")
        If UserList(TempCharIndex).WeaponEqpObjIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "*Weapon: " & ObjData(UserList(TempCharIndex).WeaponEqpObjIndex).Name & FONTTYPE_TALK)
        End If
        If UserList(TempCharIndex).ArmourEqpObjIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "*Armor: " & ObjData(UserList(TempCharIndex).ArmourEqpObjIndex).Name & FONTTYPE_TALK)
        End If
        If UserList(TempCharIndex).HelmEqpObjIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "*Helm: " & ObjData(UserList(TempCharIndex).HelmEqpObjIndex).Name & FONTTYPE_TALK)
        End If
        If UserList(TempCharIndex).AccEqpObjIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "*Accessory: " & ObjData(UserList(TempCharIndex).AccEqpObjIndex).Name & FONTTYPE_TALK)
        End If
        FoundSomething = 1
        
End If
If FoundChar = 2 Then
        If NPCList(TempCharIndex).Hostile > 0 Then
            Call SendData(ToIndex, userindex, 0, "#" & NPCList(TempCharIndex).Name & FONTTYPE_TALK)
        Else
            If NPCList(TempCharIndex).Shop > 0 Then
                Call SendData(ToIndex, userindex, 0, "NPCSHOP" & NPCList(TempCharIndex).Name & "," & NpcShops(TempCharIndex).SayCaption & "," & TempCharIndex)
                For LoopC = 1 To 20
                    If NpcShops(TempCharIndex).Slots(LoopC).Func <> "" Then
                        Call SendData(ToIndex, userindex, 0, "NPCFUNC" & NpcShops(TempCharIndex).Slots(LoopC).FuncName)
                    End If
                Next LoopC
                Exit Sub
            End If
        End If
        FoundSomething = 1
End If

'*** Check for gold ***
If MapData(map, x, y).Gold > 0 Then
    Call SendData(ToIndex, userindex, 0, "#" & MapData(map, x, y).Gold & " coins." & FONTTYPE_TALK)
    FoundSomething = 1
End If

'*** Check for object ***
If MapData(map, x, y).ObjInfo.ObjIndex > 0 Then
    Call SendData(ToIndex, userindex, 0, "#" & ObjData(MapData(map, x, y).ObjInfo.ObjIndex).Name & FONTTYPE_TALK)
    FoundSomething = 1
End If

End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'*****************************************************************
'Warps user to another spot
'*****************************************************************
Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer
Dim BouncePos As WorldPos


'fox
If map >= 29 And map <= 33 And UserList(userindex).Stats.Lv < 7 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 7." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'goblin
If map >= 34 And map <= 36 And UserList(userindex).Stats.Lv < 15 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 15." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'snake
If map >= 37 And map <= 39 And UserList(userindex).Stats.Lv < 20 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 20." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'spider
If map >= 40 And map <= 43 And UserList(userindex).Stats.Lv < 32 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 32." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'skell
If map >= 44 And map <= 48 And UserList(userindex).Stats.Lv < 40 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 40." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'b skell
If map >= 49 And map <= 53 And UserList(userindex).Stats.Lv < 45 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 45." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If
'ghost 1
If map >= 54 And map <= 54 And UserList(userindex).Stats.Lv < 60 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 60." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If

'ghost 2 -- 70
If map >= 58 And map <= 63 And UserList(userindex).Stats.Lv < 70 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 70." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If

'ghost 3 -- 80
If map >= 64 And map <= 66 And UserList(userindex).Stats.Lv < 80 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 80." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If

'werewolf -- 95
If map >= 67 And map <= 70 And UserList(userindex).Stats.Lv < 95 Then
    Call SendData(ToIndex, userindex, 0, "#You must be level 95." & FONTTYPE_FIGHT)
    Call ClosestLegalPos(UserList(userindex).Pos, BouncePos)
    Call WarpUserChar(userindex, BouncePos.map, BouncePos.x, BouncePos.y)
    Exit Sub
End If

'skelmages -- 90

'green dragons -- 99

'red dragons 5k stats (combined?)

OldMap = UserList(userindex).Pos.map
OldX = UserList(userindex).Pos.x
OldY = UserList(userindex).Pos.y

Call EraseUserChar(ToMap, 0, OldMap, userindex)

UserList(userindex).Pos.x = x
UserList(userindex).Pos.y = y
UserList(userindex).Pos.map = map

If OldMap <> map Then
    'Set switchingmap flag
    UserList(userindex).Flags.SwitchingMaps = 1
    
    'Tell client to try switching maps
    Call SendData(ToIndex, userindex, 0, "SCM" & map & "," & MapInfo(map).MapVersion)

    'Update new Map Users
    MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If
    
    'Show Character to others
    Call MakeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
    
Else
    
    Call MakeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
    Call SendData(ToIndex, userindex, 0, "SUC" & UserList(userindex).Char.CharIndex)

End If

End Sub

Sub SendUserStatsBox(ByVal userindex As Integer)
'*****************************************************************
'Updates a User's stat box
'*****************************************************************
'EDIT LOG IN TOO!
Call SendData(ToIndex, userindex, 0, "SST" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.CurHP & "," & UserList(userindex).Stats.MaxMP & "," & UserList(userindex).Stats.CurMP & "," & UserList(userindex).Stats.Gold & "," & UserList(userindex).Stats.Lv & "," & UserList(userindex).Stats.Texp & "," & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)


End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Returns the direction in which the Target is from the Pos, 0 if equal
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim angle As Integer

x = Pos.x - Target.x
y = Pos.y - Target.y

'NE
If Sgn(x) = -1 And Sgn(y) = 1 Then
    angle = Int(Rnd * 2)
    If angle = 1 Then
        FindDirection = NORTH
    Else
        FindDirection = EAST
    End If
    Exit Function
End If

'NW
If Sgn(x) = 1 And Sgn(y) = 1 Then
    angle = Int(Rnd * 2)
    If angle = 1 Then
        FindDirection = WEST
    Else
        FindDirection = NORTH
    End If
    Exit Function
End If

'SW
If Sgn(x) = 1 And Sgn(y) = -1 Then
    angle = Int(Rnd * 2)
    If angle = 1 Then
        FindDirection = WEST
    Else
        FindDirection = SOUTH
    End If
    Exit Function
End If

'SE
If Sgn(x) = -1 And Sgn(y) = -1 Then
    angle = Int(Rnd * 2)
    If angle = 1 Then
        FindDirection = SOUTH
    Else
        FindDirection = EAST
    End If
    Exit Function
End If

'South
If Sgn(x) = 0 And Sgn(y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'north
If Sgn(x) = 0 And Sgn(y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'West
If Sgn(x) = 1 And Sgn(y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'East
If Sgn(x) = -1 And Sgn(y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'Same spot
If Sgn(x) = 0 And Sgn(y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function



