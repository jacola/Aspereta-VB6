Attribute VB_Name = "TCP"
Option Explicit

Sub HandleData(rData As String)
'*********************************************
'Handle all data from server
'*********************************************
Dim retVal As Variant
Dim X As Integer
Dim Y As Integer
Dim charindex As Integer
Dim ServerHandle As Integer
Dim TempInt As Integer
Dim TempStr As String
Dim Slot As Integer
Dim iLoop As Integer


'**************** Npc Shop stuff *********************
If Left$(rData, 7) = "NPCSHOP" Then
    rData = Right$(rData, Len(rData) - 7)
    CurrentNPCShop = ReadField(3, rData, 44)
    frmMain.NpcFrame.Visible = True
    'frmMain.SpellLst.Visible = False
    'frmMain.CmdNpcSelectx10.Visible = True
    frmMain.NpcTalk = ReadField(1, rData, 44) & ": " & ReadField(2, rData, 44)
    frmMain.NpcFrame.Caption = ReadField(1, rData, 44) & " (" & CurrentNPCShop & ")"
    frmMain.NpcList.Clear
    Exit Sub
End If


If Left$(rData, 7) = "NPCFUNC" Then
    rData = Right$(rData, Len(rData) - 7)
    frmMain.NpcList.AddItem rData
    Exit Sub
End If

'**************** Communication stuff ****************

'For the text box stuff
If Left$(rData, 4) = "CTXT" Then
    rData = Right$(rData, Len(rData) - 4)
    X = ReadField(1, rData, 44)
    Y = ReadField(2, rData, 44)
    
    If MapData(X, Y).charindex > 0 Then
        CharList(MapData(X, Y).charindex).SayText = Right$(rData, Len(rData) - 6)
        CharList(MapData(X, Y).charindex).TextTime = (Len(rData) * 3) + 10
    End If
    
    'CharList(UserIndex).TextTime = 60
End If

'Send to Rectxt
If Left$(rData, 1) = "@" Then
    rData = Right$(rData, Len(rData) - 1)
    
    'AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
    AddToTalk (rData)
    Exit Sub
End If

If Left$(rData, 1) = "#" Then
    rData = Right$(rData, Len(rData) - 1)
    
    If StatusFilter = True Then
        AddtoRichTextBox frmMain.StatusBox, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
    Else
        AddToTalk (rData)
    End If
    
    Exit Sub
End If

If Left$(rData, 1) = "*" Then
    rData = Right$(rData, Len(rData) - 1)
    If Left$(rData, 5) = "CLEAR" Then
        frmMain.StatBox.Text = ""
        frmMain.StatBox.Visible = True
    Else
        AddtoRichTextBox frmMain.StatBox, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
    End If
    Exit Sub
End If

'The OK box
If Left$(rData, 5) = "OKBOX" Then
    OkBoxPos = 0
    rData = Right(rData, Len(rData) - 5)
    SetOKText (rData)
    ShowOKBox = True
End If

'Urgant MsgBox
If Left$(rData, 2) = "!!" Then
    rData = Right$(rData, Len(rData) - 2)
    frmMain.svrAlert.Visible = True
    frmMain.svrMsg = rData
    'MsgBox rData, vbApplicationModal
    Exit Sub
End If

'MsgBox
If Left$(rData, 1) = "!" Then
    rData = Right$(rData, Len(rData) - 1)
    'MsgBox rData
    frmMain.svrAlert.Visible = True
    frmMain.svrMsg = rData
    Exit Sub
End If

If rData = "KILL" Then
    frmMain.Hide
    DeInitTileEngine
End If

'**************** Intitialization stuff ****************

'Get UserServerIndex
If Left$(rData, 3) = "SUI" Then
    rData = Right$(rData, Len(rData) - 3)
    UserIndex = (Val(rData))
    Exit Sub
End If

'Get UserCharIndex
If Left$(rData, 3) = "SUC" Then
    rData = Right$(rData, Len(rData) - 3)
    UserCharIndex = (Val(rData))
    UserPos = CharList(UserCharIndex).Pos
    Exit Sub
End If

'Set user's screen pos
If Left$(rData, 3) = "SSP" Then
    rData = Right$(rData, Len(rData) - 3)
    UserPos.X = ReadField(1, rData, 44)
    UserPos.Y = ReadField(2, rData, 44)
    Exit Sub
End If

'Set user position
If Left$(rData, 3) = "SUP" Then
    rData = Right$(rData, Len(rData) - 3)
    
    X = ReadField(1, rData, 44)
    Y = ReadField(2, rData, 44)
    
    MapData(UserPos.X, UserPos.Y).charindex = 0
    MapData(X, Y).charindex = UserCharIndex
    
    UserPos.X = X
    UserPos.Y = Y
    CharList(UserCharIndex).Pos = UserPos
    
    Exit Sub
End If

'**************** Map stuff ****************

'Load map
If Left$(rData, 3) = "SCM" Then
    rData = Right$(rData, Len(rData) - 3)
    
    'Stop engine
    EngineRun = False
    
    'Set switching map flag
    DownloadingMap = True
    
    'Get Version Num
    If FileExist(App.Path & MapPath & "Map" & ReadField(1, rData, 44) & ".map", vbNormal) Then
        Open App.Path & MapPath & "Map" & ReadField(1, rData, 44) & ".map" For Binary As #1
        Seek #1, 1
        Get #1, , TempInt
        Close #1
        If TempInt = Val(ReadField(2, rData, 44)) Then
            'Correct Version
            SwitchMap ReadField(1, rData, 44)
            SendData "DLM" 'Tell the server we are done loading map
        Else
            'Not correct version
            SendData "RMU" & ReadField(1, rData, 44)
        End If
    Else
        'Didn't find map
        SendData "RMU" & ReadField(1, rData, 44)
    End If
    
    Exit Sub
End If

'Start Map Transfer
If Left$(rData, 3) = "SMT" Then
    rData = Right$(rData, Len(rData) - 3)
    
    MapInfo.MapVersion = Val(rData)
    
    ClearMapArray
    frmMain.MapLoadFrame.Visible = True
    
    Exit Sub
End If

'Set Map Tile
If Left$(rData, 3) = "CMT" Then
    rData = Right$(rData, Len(rData) - 3)
    
    ReadMapTileStr rData
    
    SendData "RNT"
    
    Exit Sub
    
End If

'End Map Transfer
If Left$(rData, 3) = "EMT" Then
    rData = Right$(rData, Len(rData) - 3)
    If Val(rData) > NumMaps Then NumMaps = Val(rData)
    SaveMapData Val(rData)
    SwitchMap Val(rData)
    frmMain.MapLoadFrame.Visible = False
    SendData "DLM" 'Tell the server we are done loading map
End If

'Done switching maps
If rData = "DSM" Then
    DownloadingMap = False
    EngineRun = True
    Exit Sub
End If

'Change map name
If Left$(rData, 3) = "SMN" Then
    MapInfo.Name = Right$(rData, Len(rData) - 3)
    frmMain.MapNameLbl.Caption = MapInfo.Name
    Exit Sub
End If

'**************** Character and object stuff ****************

'Ignore this stuff if downloading a map
If DownloadingMap = False Then


    'Make Char
    If Left$(rData, 3) = "MAC" Then
        rData = Right$(rData, Len(rData) - 3)
    
        charindex = ReadField(4, rData, 44)
        X = ReadField(5, rData, 44)
        Y = ReadField(6, rData, 44)
    
        Call MakeChar(charindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), X, Y, ReadField(7, rData, 44))
        'edit make char
        Exit Sub
    End If

    'Erase Char
    If Left$(rData, 3) = "ERC" Then
        rData = Right$(rData, Len(rData) - 3)
        'MapData(CharList(Val(rData)).Pos.x, CharList(Val(rData)).Pos.y).Spell = CharList(Val(rData)).Spell
        'MapData(CharList(Val(rData)).Pos.x, CharList(Val(rData)).Pos.y).Counter = 30

        Call EraseChar(Val(rData))

        Exit Sub
    End If
    
    'Kill char
    If Left$(rData, 3) = "KIL" Then
        rData = Right$(rData, Len(rData) - 3)

        Call EraseChar(Val(rData))
        'Call KillChar(Val(rData))

        Exit Sub
    End If
    
    'Move Char
    If Left$(rData, 3) = "MOC" Then
        rData = Right$(rData, Len(rData) - 3)

        charindex = Val(ReadField(1, rData, 44))
        
        Call MoveCharbyPos(charindex, ReadField(2, rData, 44), ReadField(3, rData, 44))

        Exit Sub
    End If

    'Change Char
    If Left$(rData, 3) = "CHC" Then
        rData = Right$(rData, Len(rData) - 3)

        charindex = Val(ReadField(1, rData, 44))
        If charindex = 0 Then Exit Sub

        CharList(charindex).Body = BodyData(Val(ReadField(2, rData, 44)))
        CharList(charindex).Head = HeadData(Val(ReadField(3, rData, 44)))
        CharList(charindex).Heading = Val(ReadField(4, rData, 44))
        CharList(charindex).HpPercent = Val(ReadField(5, rData, 44))
        If Val(ReadField(6, rData, 44)) <> 0 Then
            CharList(charindex).Weap = WeapData(Val(ReadField(6, rData, 44)))
        Else
            CharList(charindex).Weap.Weap(1).GrhIndex = 0
        End If

        Exit Sub
    End If

    'Vita percent update
    If Left$(rData, 2) = "VC" Then
        rData = Right$(rData, Len(rData) - 2)

        charindex = Val(ReadField(1, rData, 44))

        CharList(charindex).HpPercent = Val(ReadField(2, rData, 44))

        Exit Sub
    End If
    
    'Spell update
    If Left$(rData, 2) = "SP" Then
        rData = Right$(rData, Len(rData) - 2)

        X = ReadField(1, rData, 44)
        Y = ReadField(2, rData, 44)
        
        If MapData(X, Y).charindex > 0 Then
            CharList(MapData(X, Y).charindex).Spell.GrhIndex = ReadField(3, rData, 44)
            CharList(MapData(X, Y).charindex).Spell.FrameCounter = 1
            CharList(MapData(X, Y).charindex).Spell.Started = 1
            CharList(MapData(X, Y).charindex).Spell.SpeedCounter = 7
            CharList(MapData(X, Y).charindex).SpellCount = 30
        End If
        

        Exit Sub
    End If
    

    'Make Obj layer
    If Left$(rData, 3) = "MOB" Then
        rData = Right$(rData, Len(rData) - 3)
        X = Val(ReadField(2, rData, 44))
        Y = Val(ReadField(3, rData, 44))
        MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, rData, 44))
        InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
        Exit Sub
    End If

    'Erase Obj layer
    If Left$(rData, 3) = "EOB" Then
        rData = Right$(rData, Len(rData) - 3)
        X = Val(ReadField(1, rData, 44))
        Y = Val(ReadField(2, rData, 44))
        MapData(X, Y).ObjGrh.GrhIndex = 0
        Exit Sub
    End If

    If Left$(rData, 3) = "MAG" Then
        rData = Right$(rData, Len(rData) - 3)
        X = Val(ReadField(2, rData, 44))
        Y = Val(ReadField(3, rData, 44))
        If ReadField(1, rData, 44) = "x" Then
            Call KillGold(X, Y)
        Else
            Call MakeGold(X, Y)
        End If
        
        Exit Sub
    End If

    
End If

'**************** Status stuff ****************

'Update Main Stats
If Left$(rData, 3) = "SST" Then
    rData = Right$(rData, Len(rData) - 3)
    
    'frmMain.Label1.Caption = rData
    Dim tmpX As Integer
    Dim tmpY As Integer
    
    UserMaxHP = Val(ReadField(1, rData, 44))
    UserCurHP = Val(ReadField(2, rData, 44))
    UserMaxMP = Val(ReadField(3, rData, 44))
    UserCurMP = Val(ReadField(4, rData, 44))
    UserGold = Val(ReadField(5, rData, 44))
    UserLv = Val(ReadField(6, rData, 44))
    UserTexp = Val(ReadField(7, rData, 44))
    tmpX = Val(ReadField(8, rData, 44))
    tmpY = Val(ReadField(9, rData, 44))
    If UserCurHP <= 0 Then
        frmMain.HPshp.Width = 0
    Else
        frmMain.HPshp.Width = (((UserCurHP / 100) / (UserMaxHP / 100)) * 2250)
    End If

    If UserCurMP <= 0 Then
        frmMain.MANShp.Width = 0
    Else
        frmMain.MANShp.Width = (((UserCurMP / 100) / (UserMaxMP / 100)) * 2250)
    End If

    'frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 150)

    frmMain.LblName.Caption = UserName
    frmMain.LblGold.Caption = UserGold
    frmMain.LblLvl.Caption = UserLv
    frmMain.LblHP.Caption = UserCurHP
    frmMain.LblMP.Caption = UserCurMP
    frmMain.LblTexp.Caption = UserTexp
    frmMain.LblXY.Caption = Str$(tmpX) + " / " + Str$(tmpY)
    
    If iTx = 0 And iTy = 0 Then
        iTx = tmpX
        iTy = tmpY
    End If

    Exit Sub
End If

'Call SendData(ToIndex, UserIndex, 0, "TNL" & UserList(UserIndex).Stats.Exp & "," & UserList(UserIndex).Stats.Tnl)
'Update TNL bar
If Left$(rData, 3) = "TNL" Then
    rData = Right$(rData, Len(rData) - 3)
    
    UserExp = Val(ReadField(1, rData, 44))
    UserTnl = Val(ReadField(2, rData, 44))
    If UserTnl > 0 Then
        frmMain.ExpProg.Width = (((UserExp / 100) / (UserTnl / 100)) * 150)
        frmMain.ExpLbl.Caption = Str$(UserExp) + " /" + Str$(UserTnl)
    Else
        frmMain.ExpProg.Width = 150
        frmMain.ExpLbl.Caption = Str$(0) + " /" + Str$(0)
    End If
    Exit Sub
End If


'Set Inventory Slot
If Left$(rData, 3) = "SIS" Then
    rData = Right$(rData, Len(rData) - 3)

    Slot = ReadField(1, rData, 44)
    UserInventory(Slot).OBJIndex = ReadField(2, rData, 44)
    UserInventory(Slot).Name = ReadField(3, rData, 44)
    UserInventory(Slot).Amount = ReadField(4, rData, 44)
    UserInventory(Slot).Equipped = ReadField(5, rData, 44)
    UserInventory(Slot).GrhIndex = Val(ReadField(6, rData, 44))
    rData = ReadField(7, rData, 44)
    
    TempStr = ""
    TempStr = TempStr & Chr$(64 + Slot) & ": "
    
    'If UserInventory(Slot).Amount > 0 Then
    '    TempStr = TempStr & UserInventory(Slot).Name & " [" & UserInventory(Slot).Amount & "]"
    'End If
    
    'If UserInventory(Slot).Equipped = 1 Then
    '    If rData = "ACC" Then
    '        TempStr = TempStr & "   <Acc>"
    '    ElseIf rData = "ARMOR" Then
    '        TempStr = TempStr & "   <Armor>"
    '    ElseIf rData = "HELM" Then
    '        TempStr = TempStr & "   <Helm>"
    '    ElseIf rData = "WEAPON" Then
    '        TempStr = TempStr & "   <Weapon>"
    '    Else
    '        TempStr = TempStr + rData
    '    End If
    'End If

    
    Exit Sub
End If

If Left$(rData, 3) = "SSS" Then
    rData = Right$(rData, Len(rData) - 3)

    Slot = ReadField(1, rData, 44)
    UserSpellbook(Slot).Name = ReadField(2, rData, 44)
    UserSpellbook(Slot).GrhIndex = Val(ReadField(3, rData, 44))
    UserSpellbook(Slot).SpellIndex = Val(ReadField(4, rData, 44))
    UserSpellbook(Slot).Targetable = ReadField(5, rData, 44)
    UserSpellbook(Slot).Icon.GrhIndex = Val(ReadField(6, rData, 44))
    UserSpellbook(Slot).Icon.FrameCounter = 1
    UserSpellbook(Slot).Icon.Started = 0
    'Grh.FrameCounter = 1
    'Grh.Started = 0
    'Grh.GrhIndex = 6450
    
    TempStr = ""
    TempStr = TempStr & Chr$(64 + Slot) & ": "
    
    TempStr = TempStr & UserSpellbook(Slot).Name
    'UserSpellbook(Slot).Name = TempStr
    'frmMain.SpellLst.List(Slot - 1) = TempStr
    
    Exit Sub
End If

'**************** Sound stuff ****************

'Play midi
If Left$(rData, 3) = "PLM" Then
    If frmConnect.chkSound.value = Checked Then
        rData = Right$(rData, Len(rData) - 3)
        
        CurMidi = IniPath & "Mus" & Val(ReadField(1, rData, 45)) & ".mid"
        LoopMidi = Val(ReadField(2, rData, 45))
        Call PlayMidi(CurMidi)
    End If
    
    Exit Sub
End If

'Play Wave
If Left$(rData, 3) = "PLW" Then
    If frmConnect.chkSound.value = Checked Then
        rData = Right$(rData, Len(rData) - 3)
        Call PlayWaveDS(IniPath & "Snd" & rData & ".wav")
    End If
    
    Exit Sub
End If

End Sub

Sub SendData(sdData As String)
'*********************************************
'Attach a ENDC to a string and send to server
'*********************************************
Dim retcode

sdData = sdData & ENDC

'To avoid spam set a limit
If Len(sdData) > 300 Then
    Exit Sub
End If

retcode = frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub Login()
'*********************************************
'Send login strings
'*********************************************

'Pre-saved character
If SendNewChar = False Then
    SendData ("LOGIN" & UserName & "," & UserPassword & "," & ClientVer & "," & App.Major & "." & App.Minor & "." & App.Revision)
End If

'New character
If SendNewChar = True Then
    SendData ("NLOGIN" & UserName & "," & UserPassword & "," & UserBody & "," & UserHead & "," & App.Major & "." & App.Minor & "." & App.Revision)
End If

End Sub


