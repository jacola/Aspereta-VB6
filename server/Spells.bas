Attribute VB_Name = "Spells"
Sub CastSpell1(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 1   *
'********************
Dim FoundChar As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)
If SpellDamage < 0 Then
    SpellDamage = SpellDamage - UserList(userindex).Stats.Int - UserList(userindex).Stats.Wis
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

'Self nontargeted healing
'take or give mana
UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
'take or give user's HP
UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP + SpellDamage

'Make sure the char doesn't have too much health
If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'If the spell was ranged make sure the target wasn't given too much health
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP > UserList(TempCharIndex).Stats.MaxHP Then
        UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.MaxHP
    End If
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
    UserList(TempCharIndex).Flags.StatsChanged = True
End If

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)
If FoundChar = 1 Then
    If userindex <> TempCharIndex Then
        Call SendData(ToIndex, TempCharIndex, 0, "#" & UserList(userindex).Name & " casts " & SpellData(iSpell).Name & " spell on you." & FONTTYPE_TALK)
    End If
ElseIf FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP > NPCList(TempCharIndex).Stats.MaxHP Then
            NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.MaxHP
        End If
End If



'User attacked death
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP <= 0 Then
        KillUser TempCharIndex
    End If
End If

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub

Sub CastSpell2(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 2   *
'*************************************************
'* Ranged healing and non armor affected attacks *
'*************************************************
Dim FoundChar As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)
If SpellDamage < 0 Then
    SpellDamage = SpellDamage - UserList(userindex).Stats.Int - UserList(userindex).Stats.Wis
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If
If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 1
End If
If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'player
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'deal spell healing
    UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.CurHP + SpellDamage
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

If FoundChar = 2 Then  'monster
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'exit if they try to attack an npc that they shouldn't
    If NPCList(TempCharIndex).Attackable = False Then
        SendData ToIndex, userindex, 0, "#Fizzle." & FONTTYPE_FIGHT
        Exit Sub
    End If
    NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.CurHP + SpellDamage
    Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

'Make sure the char doesn't have too much health
If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'If the spell was ranged make sure the target wasn't given too much health
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP > UserList(TempCharIndex).Stats.MaxHP Then
        UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.MaxHP
    End If
    UserList(TempCharIndex).Flags.StatsChanged = True
End If

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)
If FoundChar = 1 Then
    If userindex <> TempCharIndex Then
        Call SendData(ToIndex, TempCharIndex, 0, "#" & UserList(userindex).Name & " casts " & SpellData(iSpell).Name & " spell on you." & FONTTYPE_TALK)
    End If
ElseIf FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP > NPCList(TempCharIndex).Stats.MaxHP Then
            NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.MaxHP
        End If
End If

'User attacked death
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP <= 0 Then
        KillUser TempCharIndex
    End If
End If

'NPC death
If FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP <= 0 Then
        'Give Exp and gold
        Call GiveExp(userindex, True, NPCList(TempCharIndex).GiveExp)
        NpcDrops TempCharIndex
        KillNPC TempCharIndex
    End If
End If

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub

Sub CastSpell3(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 3   *
'*********************************************************
'* Armor affected ranged damage spells and maybe healing *
'*********************************************************
Dim FoundChar As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT

'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)
If SpellDamage < 0 Then
    SpellDamage = SpellDamage ' - UserList(userindex).Stats.Wis '- UserList(userindex).Stats.Int
End If

'use DAM modifier??
'SpellDamage = SpellDamage * (1 + (UserList(UserIndex).Stats.Dam * 0.5))
'SpellDamage = SpellDamage * ((UserList(TempCharIndex).Stats.AC + 110) / 200)

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 1
End If
If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'player
    'take or give mana
    If SpellData(iSpell).TakePercentMP = 0 Then
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    Else
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * (SpellData(iSpell).TakePercentMP / 100))
    End If
    
    'If the user has PK off
    If UserList(userindex).Flags.PK = "off" Then
        Call SendData(ToIndex, userindex, 0, "#PK off." & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    'If the user has PK off
    If UserList(TempCharIndex).Flags.PK = "off" Then
        Call SendData(ToIndex, userindex, 0, "#Deflects." & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    'recalculate the damage for the spell with armor class
    SpellDamage = SpellDamage * ((UserList(TempCharIndex).Stats.AC + 110) / 200)
    'deal spell damage
    UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.CurHP + SpellDamage
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

If FoundChar = 2 Then  'monster
    'take or give mana
    If SpellData(iSpell).TakePercentMP = 0 Then
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    Else
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * (SpellData(iSpell).TakePercentMP / 100))
    End If
    'exit if they try to attack an npc that they shouldn't
    If NPCList(TempCharIndex).Attackable = False Then
        SendData ToIndex, userindex, 0, "#Fizzle." & FONTTYPE_FIGHT
        Exit Sub
    End If
    'recalculate the damage for the spell with armor class
    SpellDamage = SpellDamage * ((NPCList(TempCharIndex).Stats.AC + 110) / 200)
    'Deal spell damage
    NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.CurHP + SpellDamage
    'SendData ToIndex, UserIndex, 0, "#Damage: " & SpellDamage & FONTTYPE_TALK
    Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

'Make sure the char doesn't have too much health
If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'If the spell was ranged make sure the target wasn't given too much health
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP > UserList(TempCharIndex).Stats.MaxHP Then
        UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.MaxHP
    End If
    UserList(TempCharIndex).Flags.StatsChanged = True
End If

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)
If FoundChar = 1 Then
    If userindex <> TempCharIndex Then
        Call SendData(ToIndex, TempCharIndex, 0, "#" & UserList(userindex).Name & " casts " & SpellData(iSpell).Name & " spell on you." & FONTTYPE_TALK)
    End If
ElseIf FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP > NPCList(TempCharIndex).Stats.MaxHP Then
            NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.MaxHP
        End If
End If

'User attacked death
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP <= 0 Then
        KillUser TempCharIndex
    End If
End If

'NPC death
If FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP <= 0 Then
        'Give Exp and gold
        Call GiveExp(userindex, True, NPCList(TempCharIndex).GiveExp)
        NpcDrops TempCharIndex
        KillNPC TempCharIndex
    End If
End If

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub

Sub CastSpell4(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 4   *
'***********************************
'* Ranged curses on monsters only. *
'* Sets target's ac to GiveBaseHP. *
'***********************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)
'If SpellDamage < 0 Then
'    SpellDamage = SpellDamage - UserList(userindex).Stats.Int - UserList(userindex).Stats.Wis
'End If

'use DAM modifier??
'SpellDamage = SpellDamage * (1 + (UserList(UserIndex).Stats.Dam * 0.5))
'SpellDamage = SpellDamage * ((UserList(TempCharIndex).Stats.AC + 110) / 200)

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If
If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 1
End If
If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'player
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'Tell the player they can't curse another player
    Call SendData(ToIndex, userindex, 0, "#Your magic deflects." & FONTTYPE_WARNING)
End If

If FoundChar = 2 Then  'monster
    If NPCList(TempCharIndex).CurseName <> "" Then
        SendData ToIndex, userindex, 0, "#[" & NPCList(TempCharIndex).CurseName & "] already afflicts this monster." & FONTTYPE_FIGHT
        Exit Sub
    End If
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'exit if they try to attack an npc that they shouldn't
    If NPCList(TempCharIndex).Attackable = False Then
        SendData ToIndex, userindex, 0, "#Your magic deflects and fizzles." & FONTTYPE_FIGHT
        Exit Sub
    End If
    NPCList(TempCharIndex).Stats.AC = SpellData(iSpell).GiveBaseHP
    NPCList(TempCharIndex).CurseName = SpellData(iSpell).Name
    Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub

Sub CastSpell5(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 5   *
'*********************************************************
'* Armor affected vita attacks                           *
'*********************************************************
Dim FoundChar As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

Call SendData(ToPCArea, userindex, UserList(userindex).Pos.map, "CTXT" & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y & ",  " & SpellData(iSpell).Name & "-!")

x = UserList(userindex).Pos.x
y = UserList(userindex).Pos.y

If UserList(userindex).Char.Heading = EAST Then x = x + 1
If UserList(userindex).Char.Heading = WEST Then x = x - 1
If UserList(userindex).Char.Heading = SOUTH Then y = y + 1
If UserList(userindex).Char.Heading = NORTH Then y = y - 1

If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 1
End If
If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'player
    'take or give mana
    If SpellData(iSpell).TakePercentMP = 0 Then
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    Else
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * (SpellData(iSpell).TakePercentMP / 100))
    End If
    
    'If the user has PK off
    If UserList(userindex).Flags.PK = "off" Then
        Call SendData(ToIndex, userindex, 0, "#PK off." & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    If UserList(TempCharIndex).Flags.PK = "off" Then
        Call SendData(ToIndex, userindex, 0, "#Deflects." & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    If SpellData(iSpell).TakePercentHP <> 0 Then
        UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP - (UserList(userindex).Stats.CurHP * (SpellData(iSpell).TakePercentHP / 100))
    End If
    
    'recalculate the damage for the spell with armor class
    SpellDamage = SpellDamage * ((UserList(TempCharIndex).Stats.AC + 110) / 200)
    'deal spell damage
    UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.CurHP + SpellDamage
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

If FoundChar = 2 Then  'monster
    'take or give mana
    If SpellData(iSpell).TakePercentMP = 0 Then
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    Else
        UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * (SpellData(iSpell).TakePercentMP / 100))
    End If
    
    If SpellData(iSpell).TakePercentHP <> 0 Then
        UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP - (UserList(userindex).Stats.CurHP * (SpellData(iSpell).TakePercentHP / 100))
    End If
    
    'exit if they try to attack an npc that they shouldn't
    If NPCList(TempCharIndex).Attackable = False Then
        SendData ToIndex, userindex, 0, "#Miss." & FONTTYPE_FIGHT
        Exit Sub
    End If
    'recalculate the damage for the spell with armor class
    SpellDamage = SpellDamage * ((NPCList(TempCharIndex).Stats.AC + 110) / 200)
    'Deal spell damage
    NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.CurHP + SpellDamage
    'SendData ToIndex, UserIndex, 0, "#Damage: " & SpellDamage & FONTTYPE_TALK
    Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

'Make sure the char doesn't have too much health
If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'If the spell was ranged make sure the target wasn't given too much health
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP > UserList(TempCharIndex).Stats.MaxHP Then
        UserList(TempCharIndex).Stats.CurHP = UserList(TempCharIndex).Stats.MaxHP
    End If
    UserList(TempCharIndex).Flags.StatsChanged = True
End If

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)
If FoundChar = 1 Then
    If userindex <> TempCharIndex Then
        Call SendData(ToIndex, TempCharIndex, 0, "#" & UserList(userindex).Name & " casts " & SpellData(iSpell).Name & " spell on you." & FONTTYPE_TALK)
    End If
ElseIf FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP > NPCList(TempCharIndex).Stats.MaxHP Then
            NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.MaxHP
        End If
End If

'User attacked death
If FoundChar = 1 Then
    If UserList(TempCharIndex).Stats.CurHP <= 0 Then
        KillUser TempCharIndex
    End If
End If

'NPC death
If FoundChar = 2 Then
    If NPCList(TempCharIndex).Stats.CurHP <= 0 Then
        'Give Exp and gold
        Call GiveExp(userindex, True, NPCList(TempCharIndex).GiveExp)
        NpcDrops TempCharIndex
        KillNPC TempCharIndex
    End If
End If

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub

Sub CastSpell6(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 6   *
'*********************************************************
'* Screen attacks                                        *
'*********************************************************
Dim FoundChar As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT + 5

'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Calculate Healing Damage
'Calculates base given hp and gives mp and hp by percent of user's cur hp/mp
SpellDamage = SpellData(iSpell).GiveBaseHP
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentMP / 100) * UserList(userindex).Stats.CurMP)
SpellDamage = SpellDamage + ((SpellData(iSpell).GivePercentHP / 100) * UserList(userindex).Stats.CurHP)

If SpellDamage < 0 Then
    SpellDamage = SpellDamage - UserList(userindex).Stats.Int / 2  ' - UserList(userindex).Stats.Wis
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If


'take or give mana
If SpellData(iSpell).TakePercentMP = 0 Then
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
Else
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * (SpellData(iSpell).TakePercentMP / 100))
End If

'Take vita
If SpellData(iSpell).TakePercentHP <> 0 Then
    UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP - (UserList(userindex).Stats.CurHP * (SpellData(iSpell).TakePercentHP / 100))
    If UserList(userindex).Stats.CurHP < 1 Then UserList(userindex).Stats.CurHP = 1
End If

For x = UserList(userindex).Pos.x - 5 To UserList(userindex).Pos.x + 5
    For y = UserList(userindex).Pos.y - 6 To UserList(userindex).Pos.y + 6
    
        If MapData(map, x, y).NpcIndex > 0 Then
            TempCharIndex = MapData(map, x, y).NpcIndex
            FoundChar = 2
        End If
        
        If FoundChar = 2 Then  'monster
            'exit if they try to attack an npc that they shouldn't
            If NPCList(TempCharIndex).Attackable = False Then
                SendData ToIndex, userindex, 0, "#Fizzle." & FONTTYPE_FIGHT
            Else
                
                'Deal spell damage
                NPCList(TempCharIndex).Stats.CurHP = NPCList(TempCharIndex).Stats.CurHP + (SpellDamage * ((NPCList(TempCharIndex).Stats.AC + 110) / 200))
                'SendData ToIndex, UserIndex, 0, "#Damage: " & SpellDamage & FONTTYPE_TALK
                Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
                
                'NPC death
                If NPCList(TempCharIndex).Stats.CurHP <= 0 Then
                    'Give Exp and gold
                    Call GiveExp(userindex, True, NPCList(TempCharIndex).GiveExp)
                    NpcDrops TempCharIndex
                    KillNPC TempCharIndex
                End If
            End If
            FoundChar = 0
        End If

    Next y
Next x

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True


'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

'Check if the user leveled
Call CheckUserLevel(userindex)

End Sub


Sub CastSpell7(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 7   *
'***********************************
'* Paralyze spell code             *
'***********************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 1
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'monster
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'If the monster is para'd tell them they can't cast it again.
    If NPCList(TempCharIndex).ParaCount > 0 Then
        SendData ToIndex, userindex, 0, "#This magic already afflicts the target." & FONTTYPE_FIGHT
        Exit Sub
    Else
        If NPCList(TempCharIndex).Attackable = False Then
            SendData ToIndex, userindex, 0, "#Your magic deflects and fizzles." & FONTTYPE_FIGHT
            Exit Sub
        End If
        NPCList(TempCharIndex).ParaCount = SpellData(iSpell).GiveBaseHP
        Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
    End If
End If

'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

End Sub

Sub CastSpell8(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 8   *
'***********************************
'* Poison style spell code         *
'***********************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 1
End If

If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'monster
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'If the monster is poisoned tell them they can't cast it again.
    If NPCList(TempCharIndex).PoisonCount > 0 Then
        SendData ToIndex, userindex, 0, "#Poison already afflicts this monster." & FONTTYPE_FIGHT
        Exit Sub
    Else
        If NPCList(TempCharIndex).Attackable = False Then
            SendData ToIndex, userindex, 0, "#Your magic deflects and fizzles." & FONTTYPE_FIGHT
            Exit Sub
        End If
        NPCList(TempCharIndex).PoisonCount = SpellData(iSpell).GiveBaseHP
        NPCList(TempCharIndex).PoisonDamage = SpellData(iSpell).GiveBaseMP
        Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
    End If
End If

If FoundChar = 2 Then  'player
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    
    If SpellData(iSpell).GiveBaseMP < 0 Then
        'If the user has PK off
        If UserList(userindex).Flags.PK = "off" Then
            Call SendData(ToIndex, userindex, 0, "#PK off." & FONTTYPE_WARNING)
            Exit Sub
        End If
        'If the target has PK off
        If UserList(TempCharIndex).Flags.PK = "off" Then
            Call SendData(ToIndex, userindex, 0, "#Deflects." & FONTTYPE_WARNING)
            Exit Sub
        End If
    End If

    'If the player them they can't cast it again.
    If UserList(TempCharIndex).PoisonCount > 0 Then
        SendData ToIndex, userindex, 0, "#[" & UserList(TempCharIndex).PoisonName & "] is already in effect." & FONTTYPE_FIGHT
        Exit Sub
    Else
        UserList(TempCharIndex).PoisonName = SpellData(iSpell).Name
        UserList(TempCharIndex).PoisonCount = SpellData(iSpell).GiveBaseHP
        UserList(TempCharIndex).PoisonDamage = SpellData(iSpell).GiveBaseMP
    End If
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If


'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

End Sub


Sub CastSpell9(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 9   *
'**************************************
'* Cure spells, givebasehp makes type *
'**************************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Chance As Integer
Dim DropItem As Obj
Dim SpellDamage As Long
Dim AttackPos As WorldPos


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 1
End If

If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 2
End If

'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then  'monster
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'If the monster is poisoned tell them they can't cast it again.
    If SpellData(iSpell).GiveBaseHP = 1 Then
        NPCList(TempCharIndex).ParaCount = 0
    End If
    If SpellData(iSpell).GiveBaseHP = 2 Then
        NPCList(TempCharIndex).PoisonCount = 0
        NPCList(TempCharIndex).PoisonDamage = 0
    End If
    Call SendData(ToPCArea, userindex, NPCList(TempCharIndex).Pos.map, "SP" & NPCList(TempCharIndex).Pos.x & "," & NPCList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If

If FoundChar = 2 Then  'player
    'take or give mana
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)
    'If the player them they can't cast it again.
    If SpellData(iSpell).GiveBaseHP = 2 Then
        UserList(TempCharIndex).PoisonCount = 0
        UserList(TempCharIndex).PoisonDamage = 0
        If userindex <> TempCharIndex Then
            Call SendData(ToIndex, TempCharIndex, 0, "#" & UserList(userindex).Name & " casts " & SpellData(iSpell).Name & " spell on you." & FONTTYPE_TALK)
        End If
    End If
    Call SendData(ToPCArea, TempCharIndex, UserList(TempCharIndex).Pos.map, "SP" & UserList(TempCharIndex).Pos.x & "," & UserList(TempCharIndex).Pos.y & "," & SpellData(iSpell).GRHIndex)
End If


'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

End Sub

Sub CastSpell10(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal iSpell As Integer)
'********************
'*   SPELL TYPE 10  *
'**************************************
'* Push spell code                    *
'**************************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim AttackPos As WorldPos
Dim UserFace As Integer
Dim Direction As Integer

'Check attacker counter
If UserList(userindex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'update counters
UserList(userindex).Counters.AttackCounter = STAT_SPELLWAIT


'Check if legal no matter what the spell is
If InMapBounds(map, x, y) = False Then
    Exit Sub
End If

'Make sure they are casting a real spell
If iSpell > TotalNumSpells Then
    Exit Sub
End If

'Make sure the spell is a legal one
If iSpell = 0 Then
    Exit Sub
End If

'Enough Mana?
If UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP < 0 Then
    Call SendData(ToIndex, userindex, 0, "#Insufficient mana." & FONTTYPE_WARNING)
    Exit Sub
End If

UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + SpellData(iSpell).TakeBaseMP
UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP - (UserList(userindex).Stats.CurMP * SpellData(iSpell).TakePercentMP)

x = UserList(userindex).Pos.x
y = UserList(userindex).Pos.y

Select Case UserList(userindex).Char.Heading
    Case NORTH
        x = UserList(userindex).Pos.x
        y = UserList(userindex).Pos.y - 1
        Direction = NORTH
    Case EAST
        x = UserList(userindex).Pos.x + 1
        y = UserList(userindex).Pos.y
        Direction = EAST
    Case SOUTH
        x = UserList(userindex).Pos.x
        y = UserList(userindex).Pos.y + 1
        Direction = SOUTH
    Case WEST
        x = UserList(userindex).Pos.x - 1
        y = UserList(userindex).Pos.y
        Direction = WEST
End Select

If MapData(map, x, y).NpcIndex > 0 Then
    TempCharIndex = MapData(map, x, y).NpcIndex
    FoundChar = 1
End If

If MapData(map, x, y).userindex > 0 Then
    TempCharIndex = MapData(map, x, y).userindex
    FoundChar = 2
End If

Call SendData(ToIndex, userindex, 0, "#" & FoundChar & FONTTYPE_TALK)
'If nothing was found don't cast.
If FoundChar = 0 Then Exit Sub

If FoundChar = 1 Then
    Call MoveNPCChar(TempCharIndex, Direction)
End If

If FoundChar = 2 Then  'player
    Call MoveUserChar(TempCharIndex, Direction)
End If


'Say the user's stats changed since all spells cost mana
UserList(userindex).Flags.StatsChanged = True

'Tell them the spell they cast
Call SendData(ToIndex, userindex, 0, "#You cast " & SpellData(iSpell).Name & " spell." & FONTTYPE_TALK)

End Sub

