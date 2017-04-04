Attribute VB_Name = "General"
Option Explicit


' Adds text to the list of things said.
Sub AddToTalk(ByVal sTalk As String)
Dim i As Integer

For i = 1 To 39
ChatText(i) = ChatText(i + 1)
Next i
ChatText(40) = sTalk

End Sub

'Set the OK box's text
Sub SetOKText(ByVal Info As String)
Dim i As Integer

For i = 1 To 100
    OKBoxText(i) = ReadField(i, Info, Asc("|"))
    If OKBoxText(i) = "" Then OKBoxText(i) = " "
Next i


End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*****************************************************************
'Find a Random number between a range
'*****************************************************************

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

End Function

Function StringChecker(ByVal S As String)
Dim CurChar As String
Dim i As Integer

For i = 1 To Len(S)
    CurChar = Mid(S, i, 1)
    If CurChar < "A" Or CurChar > "z" Then
        frmConnect.NameTxt.Text = ""
    End If
Next i

S = S + " "

If Len(S) > 3 Then
    For i = 1 To Len(S) - 3
        CurChar = Mid(S, i, 3)
        If UCase(CurChar) = "ASS" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SHT" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "FCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "HOE" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "CUM" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SEX" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "FAG" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "GAY" Then
            frmConnect.NameTxt.Text = ""
        End If
    Next i
End If

If Len(S) > 4 Then
    For i = 1 To Len(S) - 4
        CurChar = Mid(S, i, 4)
        If UCase(CurChar) = "FUCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "CUNT" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SHIT" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "FOCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SUCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "CRAP" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SLUT" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "CLIT" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "SUCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "COCK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "DICK" Then
            frmConnect.NameTxt.Text = ""
        End If
         If UCase(CurChar) = "ANAL" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "LICK" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "HOMO" Then
            frmConnect.NameTxt.Text = ""
        End If
    Next i
End If

If Len(S) > 5 Then
    For i = 1 To Len(S) - 5
        CurChar = Mid(S, i, 5)
        If UCase(CurChar) = "BITCH" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "WHORE" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "AHOLE" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "PENIS" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "PUSSY" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "CHOAD" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "PENIS" Then
            frmConnect.NameTxt.Text = ""
        End If
        If UCase(CurChar) = "LESBO" Then
            frmConnect.NameTxt.Text = ""
        End If
    Next i
End If

If Len(S) > 6 Then
    For i = 1 To Len(S) - 6
        CurChar = Mid(S, i, 6)
        If UCase(CurChar) = "NYMPHO" Then
            frmConnect.NameTxt.Text = ""
        End If
    Next i
End If


End Function


Sub ReadMapTileStr(TileString As String)
'*****************************************************************
'Takes a tile packet from server, decodes it puts it into the map array
'*****************************************************************
Dim loopC As Integer
Dim AcumStr As String
Dim TempStr As String
Dim X As Integer, Y As Integer
Dim FieldCounter As Integer

For loopC = 1 To Len(TileString)
    TempStr = Mid(TileString, loopC, 1)
    
    If loopC = Len(TileString) Then
        AcumStr = AcumStr & TempStr
        TempStr = Chr(44)
    End If
    
    If Asc(TempStr) = 44 Then
        Select Case FieldCounter
            Case 0
                X = Val(AcumStr)
            Case 1
                Y = Val(AcumStr)
            Case 2
                MapData(X, Y).Blocked = Val(AcumStr)
            Case Is > 2
                MapData(X, Y).Graphic(Val(Left(AcumStr, 1))).GrhIndex = Val(Right(AcumStr, Len(AcumStr) - 1))
        End Select
        FieldCounter = FieldCounter + 1
        AcumStr = ""
    Else
        AcumStr = AcumStr & TempStr
    End If
    
Next loopC

If DownloadingMap Then
    frmMain.LoadForward.Width = ((Y / YMaxMapSize) * 1559) + 1
End If

End Sub

Sub ClearMapArray()
'*****************************************************************
'Clears all layers
'*****************************************************************

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        'Change blockes status
        MapData(X, Y).Blocked = 0

        'Erase layer 1 and 4
        MapData(X, Y).Graphic(1).GrhIndex = 0
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase characters
        If MapData(X, Y).charindex > 0 Then
            Call EraseChar(MapData(X, Y).charindex)
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

End Sub
Sub PlayMidi(File As String)
'*****************************************************************
'Plays a Midi using the MCIControl
'*****************************************************************

'frmMain.MidiPlayer.Command = "Close"

'frmMain.MidiPlayer.FileName = File
    
'frmMain.MidiPlayer.Command = "Open"

'frmMain.MidiPlayer.Command = "Play"

End Sub




Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, RED As Byte, GREEN As Byte, BLUE As Byte, Bold As Byte, Italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

RichTextBox.SelStart = Len(RichTextBox.Text)
RichTextBox.SelLength = 0
RichTextBox.SelColor = RGB(RED, GREEN, BLUE)

If Bold Then
    RichTextBox.SelBold = True
Else
    RichTextBox.SelBold = False
End If

If Italic Then
    RichTextBox.SelItalic = True
Else
    RichTextBox.SelItalic = False
End If

RichTextBox.SelText = Chr(13) & Chr(10) & Text

End Sub
Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0


TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub


Sub SaveGameini()
'******************************************
'Saves Game.ini
'******************************************

'update Game.ini
Call WriteVar(IniPath & "Game.ini", "INIT", "Name", UserName)
Call WriteVar(IniPath & "Game.ini", "INIT", "Password", UserPassword)
Call WriteVar(IniPath & "Game.ini", "INIT", "Port", Str(UserPort))
Call WriteVar(IniPath & "Game.ini", "INIT", "IP", frmConnect.IPTxt.Text)

End Sub

Function CheckUserData() As Boolean
'*****************************************************************
'Checks all user data for mistakes and reports them.
'*****************************************************************

Dim loopC As Integer
Dim CharAscii As Integer

'IP
If UserServerIP = "" Then
    MsgBox ("Server IP box is empty.")
    Exit Function
End If

'Port
If Str(UserPort) = "" Then
    MsgBox ("Port box is empty.")
    Exit Function
End If

'Password
If UserPassword = "" Then
    MsgBox ("Password box is empty.")
    Exit Function
End If

If Len(UserPassword) > 10 Then
    MsgBox ("Password must be 10 characters or less.")
    Exit Function
End If

For loopC = 1 To Len(UserPassword)

    CharAscii = Asc(Mid$(UserPassword, loopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Invalid Password.")
        Exit Function
    End If
    
Next loopC

'Name
If UserName = "" Then
    MsgBox ("Name box is empty.")
    Exit Function
End If

If Len(UserName) > 30 Then
    MsgBox ("Name must be 30 characters or less.")
    Exit Function
End If

For loopC = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, loopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Invalid Name.")
        Exit Function
    End If
    
Next loopC

'If all good send true
CheckUserData = True

End Function

Sub UnloadAllForms()
'*****************************************************************
'Unloads all forms
'*****************************************************************

On Error Resume Next

Unload frmConnect
Unload frmMain

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

'Set Connected
Connected = True

'Save Game.ini
If frmConnect.SavePassChk.value = 0 Then
    UserPassword = ""
End If
Call SaveGameini

'Unload the connect form
Unload frmConnect

'Load main form
frmMain.Visible = True

End Sub

Sub TargetingKeys()
'*****************************************************************
'Targeting key code
'*****************************************************************
Static KeyTimer As Integer
Static HeldKey As Integer
Dim infloop As Integer


If UserSpellbook(CurSpellIndex).Targetable = "X" Then
    KeyTimer = 0
    Targeting = False
    SendData "CAST" & UserSpellbook(CurSpellIndex).SpellIndex & "," & iTx & "," & iTy
End If


'Makes sure keys aren't being pressed to fast
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If

LastPX = CharList(UserCharIndex).Pos.X
LastPY = CharList(UserCharIndex).Pos.Y


If UserSpellbook(CurSpellIndex).SpellIndex = 0 Then
    KeyTimer = 10
    Targeting = False
    Exit Sub
End If

If GetKeyState(vbKeyHome) < 0 Then
    iTx = LastPX
    iTy = LastPY
    Exit Sub
End If

If GetKeyState(vbKeyEscape) < 0 Then
    Targeting = False
    Exit Sub
End If

If GetKeyState(vbKeyReturn) < 0 Then
    If HeldKey > 1 Then
        Exit Sub
    End If
    HeldKey = HeldKey + 1
    Targeting = False
    If CurSpellIndex > 0 Then
        If UserSpellbook(CurSpellIndex).Name <> "" Then
            SendData "CAST" & UserSpellbook(CurSpellIndex).SpellIndex & "," & iTx & "," & iTy
        End If
    End If
    Exit Sub
End If
HeldKey = 0

If GetKeyState(vbKeyRight) < 0 Or GetKeyState(vbKeyDown) < 0 Then
    While 1 = 1
        iTx = iTx + 1
        If MapData(iTx, iTy).charindex > 0 Or MapData(iTx, iTy).NPCIndex > 0 Then
            KeyTimer = 10
            Exit Sub
        End If
        
        If iTx > LastPX + 8 Then
            iTx = LastPX - 9
            iTy = iTy + 1
        End If
        
        If iTy > LastPY + 5 Then
            iTy = LastPY - 6
        End If
        infloop = infloop + 1
        If infloop = 200 Then
            SetTarget
            Exit Sub
        End If
    Wend
End If

If GetKeyState(vbKeyLeft) < 0 Or GetKeyState(vbKeyUp) < 0 Then
    While 1 = 1
        iTx = iTx - 1
        If MapData(iTx, iTy).charindex > 0 Or MapData(iTx, iTy).NPCIndex > 0 Then
            KeyTimer = 10
            Exit Sub
        End If
        
        If iTx < LastPX - 7 Then
            iTx = LastPX + 9
            iTy = iTy - 1
        End If
        
        If iTy < LastPY - 5 Then
            iTy = LastPY + 6
        End If
        infloop = infloop + 1
        If infloop = 200 Then
            SetTarget
            Exit Sub
        End If
    Wend
End If

End Sub

Sub SetTarget()
'*****************************************************************
'If the target is off screen, target the player
'*****************************************************************
If iTx < LastPX - 7 Or iTx > LastPX + 7 Or iTy < LastPY - 6 Or iTy > LastPY + 6 Then
    iTx = CharList(UserCharIndex).Pos.X
    iTy = CharList(UserCharIndex).Pos.Y
End If
If MapData(iTx, iTy).NPCIndex = 0 And MapData(iTx, iTy).charindex = 0 Then
    iTx = CharList(UserCharIndex).Pos.X
    iTy = CharList(UserCharIndex).Pos.Y
End If

End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
Static KeyTimer As Integer
Dim rData As String
Static HeldKey As Integer
Dim SpellUpdate As Integer

If Targeting = True Then
    KeyTimer = 0
    Call TargetingKeys
    Exit Sub
End If


If GetKeyState(vbKeyF1) < 0 Then
    frmMain.WindowState = 1
End If

'Makes sure keys aren't being pressed to fast
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If

' Close the chat entry box
If GetKeyState(vbKeyEscape) < 0 And frmMain.SendTxt.Visible = True Then
    frmMain.SendTxt.Visible = False
    frmMain.SendTxt.Text = ""
    frmMain.TxtCatch.SetFocus
    Exit Sub
    KeyTimer = 2
End If

If ShowOKBox = True Then
    Exit Sub
End If

If frmMain.NpcFrame.Visible = True Then
    Exit Sub
End If

'Forum stuff
If frmMain.frmForum.Visible = True Then
    Exit Sub
End If
'************

If GetKeyState(vbKeyF2) < 0 Then
    SendData "REFRESH"
    KeyTimer = 10
End If

'Don't allow any these keys during movement..
If UserMoving = 0 Then

If frmMain.SendTxt.Visible = False Then
    If frmMain.TxtCatch.Text = "," Or frmMain.TxtCatch.Text = ",," Or frmMain.TxtCatch.Text = ",,," Then
        SendData "GET"
        'frmMain.TxtCatch.Text = ""
    End If
    
    If GetKeyState(222) < 0 Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.Text = ""
        frmMain.SendTxt.SetFocus
        Exit Sub
    End If
    
    If frmMain.TxtCatch.Text = "/" Or frmMain.TxtCatch.Text = "//" Or frmMain.TxtCatch.Text = "///" Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.Text = "/"
        frmMain.SendTxt.SetFocus
        Exit Sub
    End If
    
    'Scrolling the chat box
    If GetKeyState(vbKeyPageUp) < 0 Then
        ChatPos = ChatPos - 1
        If ChatPos < -33 Then ChatPos = -33
        KeyTimer = 3
    End If
    
    If GetKeyState(vbKeyPageDown) < 0 Then
        ChatPos = ChatPos + 1
        If ChatPos > 0 Then ChatPos = 0
        KeyTimer = 3
    End If
    
    'Toggle text background always on
    If GetKeyState(vbKeyF3) < 0 Then
        If TextBoxAlwaysOn = True Then
            TextBoxAlwaysOn = False
            AddToTalk ("Text background -> Mouse over" + "~255~0~0~")
        Else
            TextBoxAlwaysOn = True
            AddToTalk ("Text background -> Always on" + "~255~0~0~")
        End If
        KeyTimer = 10
    End If
    
    'change status filtering
    If GetKeyState(vbKeyF4) < 0 Then
        If StatusFilter = True Then
            StatusFilter = False
            AddToTalk ("Status -> Chat Messages" + "~255~0~0~")
        Else
            StatusFilter = True
            AddToTalk ("Status -> Status Messages" + "~255~0~0~")
        End If
        KeyTimer = 10
    End If
    
    'change the status box on and off
    If GetKeyState(vbKeyF5) < 0 Then
        If frmMain.StatusBox.Visible = True Then
            frmMain.StatusBox.Visible = False
            AddToTalk ("Status box -> Off" + "~255~0~0~")
        Else
            frmMain.StatusBox.Visible = True
            AddToTalk ("Status box -> On" + "~255~0~0~")
        End If
        KeyTimer = 10
    End If
    
    'This toggles the hp/mp bars
    If GetKeyState(vbKeyTab) < 0 Then
        If frmMain.HPFrame.Visible = True Then
            frmMain.HPFrame.Visible = False
            frmMain.MPFrame.Visible = False
            AddToTalk ("HP/MP bars -> Off" + "~255~0~0~")
        Else
            frmMain.HPFrame.Visible = True
            frmMain.MPFrame.Visible = True
            AddToTalk ("HP/MP bars -> On" + "~255~0~0~")
        End If
        
        KeyTimer = 10
        Exit Sub
    End If
    
    'Show the forum
    If GetKeyState(vbKeyF) < 0 Then
        frmMain.frmForum.Visible = True
        Call frmMain.wbForum.Navigate2("http://www.coolbm.com/aspbbs/", Null, frmMain.wbForum, Null, Null)
        KeyTimer = 10
        Exit Sub
    End If
    

    If GetKeyState(vbKey1) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(1) <= 100 Then
            SendData "USE" & UserHotButtons(1)
            Exit Sub
        ElseIf UserHotButtons(1) >= 101 And UserHotButtons(1) <= 200 Then
            CurSpellIndex = UserHotButtons(1) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey2) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(2) <= 100 Then
            SendData "USE" & UserHotButtons(2)
            Exit Sub
        ElseIf UserHotButtons(2) >= 101 And UserHotButtons(2) <= 200 Then
            CurSpellIndex = UserHotButtons(2) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey3) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(3) <= 100 Then
            SendData "USE" & UserHotButtons(3)
            Exit Sub
        ElseIf UserHotButtons(3) >= 101 And UserHotButtons(3) <= 200 Then
            CurSpellIndex = UserHotButtons(3) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey4) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(4) <= 100 Then
            SendData "USE" & UserHotButtons(4)
            Exit Sub
        ElseIf UserHotButtons(4) >= 101 And UserHotButtons(4) <= 200 Then
            CurSpellIndex = UserHotButtons(4) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey5) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(5) <= 100 Then
            SendData "USE" & UserHotButtons(5)
            Exit Sub
        ElseIf UserHotButtons(5) >= 101 And UserHotButtons(5) <= 200 Then
            CurSpellIndex = UserHotButtons(5) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey6) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(6) <= 100 Then
            SendData "USE" & UserHotButtons(6)
            Exit Sub
        ElseIf UserHotButtons(6) >= 101 And UserHotButtons(6) <= 200 Then
            CurSpellIndex = UserHotButtons(6) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey7) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(7) <= 100 Then
            SendData "USE" & UserHotButtons(7)
            Exit Sub
        ElseIf UserHotButtons(7) >= 101 And UserHotButtons(7) <= 200 Then
            CurSpellIndex = UserHotButtons(7) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey8) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(8) <= 100 Then
            SendData "USE" & UserHotButtons(8)
            Exit Sub
        ElseIf UserHotButtons(8) >= 101 And UserHotButtons(8) <= 200 Then
            CurSpellIndex = UserHotButtons(8) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey9) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(9) <= 100 Then
            SendData "USE" & UserHotButtons(9)
            Exit Sub
        ElseIf UserHotButtons(9) >= 101 And UserHotButtons(9) <= 200 Then
            CurSpellIndex = UserHotButtons(9) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    If GetKeyState(vbKey0) < 0 Then
        If HeldKey > 1 Then
            Exit Sub
        End If
        HeldKey = HeldKey + 1
        If UserHotButtons(10) <= 100 Then
            SendData "USE" & UserHotButtons(10)
            Exit Sub
        ElseIf UserHotButtons(10) >= 101 And UserHotButtons(10) <= 200 Then
            CurSpellIndex = UserHotButtons(10) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
    End If
    
    HeldKey = 1

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        If frmMain.SendTxt.Text = "" Then
            frmMain.NpcFrame.Visible = False
            If UserDir = NORTH Then
                If LegalPos(UserPos.X, UserPos.Y - 1) Then
                    Call SendData("M" & NORTH)
                    MoveCharbyHead UserCharIndex, NORTH
                    MoveScreen NORTH
                Else
                    KeyTimer = 10
                End If
            Else
                Call SendData("FNORTH")
                UserDir = NORTH
                KeyTimer = 10
            End If
        End If
        Exit Sub
    End If

    'Move Right
    If GetKeyState(vbKeyRight) < 0 Then
        If frmMain.SendTxt.Text = "" Then
            frmMain.NpcFrame.Visible = False
            If UserDir = EAST Then
                If LegalPos(UserPos.X + 1, UserPos.Y) Then
                    Call SendData("M" & EAST)
                    MoveCharbyHead UserCharIndex, EAST
                    MoveScreen EAST
                    UserDir = EAST
                Else
                    KeyTimer = 10
                End If
            Else
                Call SendData("FEAST")
                UserDir = EAST
                KeyTimer = 10
            End If
        End If
        Exit Sub
    End If

    'Move down
    If GetKeyState(vbKeyDown) < 0 Then
        If frmMain.SendTxt.Text = "" Then
            frmMain.NpcFrame.Visible = False
            If UserDir = SOUTH Then
                If LegalPos(UserPos.X, UserPos.Y + 1) Then
                    Call SendData("M" & SOUTH)
                    MoveCharbyHead UserCharIndex, SOUTH
                    MoveScreen SOUTH
                    UserDir = SOUTH
                Else
                    KeyTimer = 10
                End If
            Else
                Call SendData("FSOUTH")
                UserDir = SOUTH
                KeyTimer = 10
            End If
        End If
        Exit Sub
    End If

    'Move left
    If GetKeyState(vbKeyLeft) < 0 Then
        If frmMain.SendTxt.Text = "" Then
            frmMain.NpcFrame.Visible = False
            If UserDir = WEST Then
                If LegalPos(UserPos.X - 1, UserPos.Y) Then
                    Call SendData("M" & WEST)
                    MoveCharbyHead UserCharIndex, WEST
                    MoveScreen WEST
                    UserDir = WEST
                Else
                    KeyTimer = 10
                End If
            Else
                Call SendData("FWEST")
                UserDir = WEST
                KeyTimer = 10
            End If
        End If

        Exit Sub
    End If
    
    If GetKeyState(vbKeySpace) < 0 Then
        If frmMain.SendTxt.Text = "" Then
            frmMain.NpcFrame.Visible = False
            SendData ("ATT")
            KeyTimer = 10
        End If
        Exit Sub
    End If
    
    If GetKeyState(vbKeyI) < 0 Then
        If ShowInventory = True Then
            'frmMain.ObjLst.Visible = False
            ShowInventory = False
            KeyTimer = 10
            Exit Sub
        End If
        ShowInventory = True
        KeyTimer = 10
        Exit Sub
    End If
    
    If GetKeyState(vbKeyZ) < 0 Then
        If ShowSpells = True Then
            ShowSpells = False
            KeyTimer = 10
            Exit Sub
        End If
        frmMain.StatBox.Visible = False
        ShowSpells = True
        frmMain.NpcFrame.Visible = False
        KeyTimer = 10
        Exit Sub
    End If
    
    If GetKeyState(vbKeyS) < 0 Then
        If ShowStatus = True Then
            ShowStatus = False
            KeyTimer = 10
            Exit Sub
        Else
            ShowStatus = True
            SendData ("/STATS")
            KeyTimer = 10
            Exit Sub
        End If
    End If
End If

End If

End Sub




Sub SwitchMap(Map As Integer)
'*****************************************************************
'Loads and switches to a new map
'*****************************************************************

Dim loopC As Integer
Dim Y As Integer
Dim X As Integer
Dim TempInt As Integer
Dim Blank As MapBlock
Dim tx As Integer
Dim ty As Integer

'Open files
Open App.Path & MapPath & "Map" & Map & ".map" For Binary As #1
Seek #1, 1
        
'map Header
Get #1, , MapInfo.MapVersion
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
        
'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        '.dat file
        Get #1, , MapData(X, Y).Blocked
        For loopC = 1 To 4
            Get #1, , MapData(X, Y).Graphic(loopC).GrhIndex
            
            'Set up GRH
            If MapData(X, Y).Graphic(loopC).GrhIndex > 0 Then
                InitGrh MapData(X, Y).Graphic(loopC), MapData(X, Y).Graphic(loopC).GrhIndex
            End If
            
        Next loopC
        'Empty place holders for future expansion
        Get #1, , TempInt
        Get #1, , TempInt
        
        'Erase NPCs
        If MapData(X, Y).charindex > 0 Then
            Call EraseChar(MapData(X, Y).charindex)
            MapData(X, Y).charindex = 0
        End If
        
        'Erase OBJs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0
        'and gold!
        MapData(X, Y).Gold = 0

    Next X
Next Y

Close #1

'Clear out old mapinfo variables
MapInfo.Name = ""
MapInfo.Music = ""

'Set current map
CurMap = Map

End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Sub Main()
'*****************************************************************
'Main
'*****************************************************************
Dim loopC As Integer

'***************************************************
'Start up
'***************************************************
'****** Init vars ******
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

'Init Engine
'InitTileEngine frmMain.hWnd, 16, 8, 32, 32, 13, 17, 10 '152 7
InitTileEngine frmMain.hWnd, -21, -15, 32, 32, 15, 21, 10

'****** Display connect window ******
'frmConnect.Visible = True

'****** MidiPlayer INIT ******
'frmMain.MidiPlayer.Notify = False
'frmMain.MidiPlayer.Wait = False
'frmMain.MidiPlayer.Shareable = False
'frmMain.MidiPlayer.TimeFormat = mciFormatMilliseconds
'frmMain.MidiPlayer.DeviceType = "Sequencer"

'***************************************************
'Main Loop
'***************************************************
prgRun = True
Do While prgRun

    '****** Check Request position timer ******
    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Request position Update
            Call SendData("RPU")
        End If
    End If
    'Call SendData("RPU")
    '****** Refesh characters on map ******
    Call RefreshAllChars

    '****** Show Next Frame ******

    'Don't draw frame is window is minimized or there is no map loaded
    If frmMain.WindowState <> 1 And CurMap > 0 Then
        
        ShowNextFrame frmMain.Top, frmMain.Left

        '****** Check keys ******
        If DownloadingMap = False Then
            CheckKeys
        End If
    
    End If

    '****** Go do other events ******
    DoEvents

Loop
    

'*****************************************************************
'Close Down
'*****************************************************************

'****** Stop any midis ******
'mciSendString "close all", 0, 0, 0

'****** Stop Engine ******
DeInitTileEngine

'****** Unload forms and end******
Call UnloadAllForms
End

End Sub

Sub SaveMapData(SaveAs As Integer)
'*****************************************************************
'Saves map data to file
'*****************************************************************

Dim loopC As Integer
Dim Y As Integer
Dim X As Integer
Dim TempInt As Integer

If FileExist(App.Path & MapPath & "Map" & SaveAs & ".map", vbNormal) = True Then
    Kill App.Path & MapPath & "Map" & SaveAs & ".map"
End If

'Write header info on Map.dat
Call WriteVar(IniPath & "Map.dat", "INIT", "NumMaps", Str(NumMaps))

'Open .map file
Open App.Path & MapPath & "Map" & SaveAs & ".map" For Binary As #1
Seek #1, 1

'map Header
Put #1, , MapInfo.MapVersion
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'Write .map file
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(X, Y).Blocked
        For loopC = 1 To 4
            Put #1, , MapData(X, Y).Graphic(loopC).GrhIndex
        Next loopC
        'Empty place holders for future expansion
        Put #1, , TempInt
        Put #1, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

End Sub

Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function





