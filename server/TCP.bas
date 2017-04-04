Attribute VB_Name = "TCP"
Option Explicit

Public Const SOCKET_BUFFER_SIZE = 20480 'Buffer in bytes for each socket
Public Const COMMAND_BUFFER_SIZE = 1000 'How many commands the server can store from each client

'Constants used in the SendData sub
Public Const ToIndex = 0 'Send data to a single User index
Public Const ToAll = 1 'Send it to all User indexa
Public Const ToMap = 2 'Send it to all users in a map
Public Const ToPCArea = 3 'Send to all users in a user's area
Public Const ToNone = 4 'Send to none
Public Const ToAllButIndex = 5 'Send to all but the index
Public Const ToMapButIndex = 6 'Send to all on a map but the index

' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocketWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Socket Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Socket Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addresses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.255.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Response
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Codes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500

Dim TmpI As Integer
Dim TmpJ As Integer



Sub ConnectNewUser(ByVal userindex As Integer, ByVal Name As String, ByVal Password As String, ByVal Body As Integer, ByVal Head As Integer)
'*****************************************************************
'Opens a new user. Loads ACault vars, saves then calls connectuser
'*****************************************************************
Dim LoopC As Integer
  
Dim Blank As User
  
'Check for Character file
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, userindex, 0, "!!Character already exist.")
    CloseSocket (userindex)
    Exit Sub
End If
  
UserList(userindex) = Blank
  
'create file
UserList(userindex).Name = Name
UserList(userindex).Password = Password
UserList(userindex).Char.Heading = SOUTH
UserList(userindex).Char.Head = Head
UserList(userindex).Char.Body = Body
UserList(userindex).Path = "Peasant"

UserList(userindex).Stats.MaxHP = 20
UserList(userindex).Stats.CurHP = 20
UserList(userindex).Stats.MaxMP = 20
UserList(userindex).Stats.CurMP = 10
UserList(userindex).Stats.Lv = 1
UserList(userindex).Stats.Exp = 1
UserList(userindex).Stats.Tnl = 250
UserList(userindex).Stats.Texp = 1
UserList(userindex).Stats.AC = 100
UserList(userindex).Stats.Dam = 0

UserList(userindex).Stats.Str = 1
UserList(userindex).Stats.Con = 1
UserList(userindex).Stats.Int = 1
UserList(userindex).Stats.Wis = 1
UserList(userindex).Stats.Dex = 1

UserList(userindex).Stats.Gold = 50

UserList(userindex).Stats.MaxHIT = 3
UserList(userindex).Stats.MinHIT = 1

For LoopC = 1 To 26
    UserList(userindex).Object(LoopC).ObjIndex = 0
    UserList(userindex).Object(LoopC).Amount = 0
    UserList(userindex).SpellBook(LoopC).Spellindex = 0
Next LoopC


UserList(userindex).Object(1).ObjIndex = 2
UserList(userindex).Object(1).Amount = 1

UserList(userindex).Object(2).ObjIndex = 17
UserList(userindex).Object(2).Amount = 1

UserList(userindex).Object(3).ObjIndex = 1
UserList(userindex).Object(3).Amount = 5

Call SaveUser(userindex, CharPath & UCase(Name) & ".chr")

Call SendData(ToIndex, userindex, 0, "!!Character created, click continue to play.")

Call CloseSocket(userindex)
  
'Open User
'Call ConnectUser(userindex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal userindex As Integer)
'*****************************************************************
'Close the users socket
'*****************************************************************
On Error Resume Next
  
Dim i As Integer
  
If userindex > 0 Then

    frmMain.Socket2(userindex).Disconnect

    If UserList(userindex).Flags.UserLogged = 1 Then
        Call CloseUser(userindex)
    End If
    'If UserList(userindex).GIndex > 0 Then
    '    For i = 1 To 5
    '        If Groups(UserList(userindex).GIndex).UIndexes(i) = userindex Then
    '            Groups(UserList(userindex).GIndex).UIndexes(i) = 0
    '        End If
    '    Next i
    'End If
    UserList(userindex).GIndex = 0
    UserList(userindex).ConnID = -1
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)

End If

End Sub


Sub SendData(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
'*****************************************************************
'Sends data to sendRoute
'*****************************************************************
Dim LoopC As Integer
Dim x As Integer
Dim y As Integer

'Add End character
sndData = sndData & ENDC
  
'send NONE
If sndRoute = ToNone Then
    Exit Sub
End If
  
  
'Send to All
If sndRoute = ToAll Then
    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged Then
            If Left(sndData, 3) = "PLW" Then
                If UserList(LoopC).Flags.Sound = "off" Then Exit Sub
            End If
            frmMain.Socket2(LoopC).Write sndData, Len(sndData)
        End If
    Next LoopC
    Exit Sub
End If

'Send to everyone but the sndindex
If sndRoute = ToAllButIndex Then
    For LoopC = 1 To LastUser
              
      If UserList(LoopC).Flags.UserLogged And LoopC <> sndIndex Then
            If Left(sndData, 3) = "PLW" Then
                If UserList(LoopC).Flags.Sound = "off" Then Exit Sub
            End If
            frmMain.Socket2(LoopC).Write sndData, Len(sndData)
      End If
      
    Next LoopC
    Exit Sub
End If

'Send to Map
If sndRoute = ToMap Then

    For LoopC = 1 To LastUser

        If UserList(LoopC).Flags.UserLogged Then
            If UserList(LoopC).Pos.map = sndMap Then
                If Left(sndData, 3) = "PLW" Then
                    If UserList(LoopC).Flags.Sound = "off" Then Exit Sub
                End If
                frmMain.Socket2(LoopC).Write sndData, Len(sndData)
            End If
        End If
      
    Next LoopC
    
    Exit Sub
End If

'Send to everone on map but sndIndex
If sndRoute = ToMapButIndex Then

    For LoopC = 1 To LastUser

        If UserList(LoopC).Flags.UserLogged And LoopC <> sndIndex Then
            If UserList(LoopC).Pos.map = sndMap Then
                If Left(sndData, 3) = "PLW" Then
                    If UserList(LoopC).Flags.Sound = "off" Then Exit Sub
                End If
                frmMain.Socket2(LoopC).Write sndData, Len(sndData)
             End If
        End If
  
    Next LoopC
    
    Exit Sub
End If

'Send to PC Area
If sndRoute = ToPCArea Then
    
    For y = UserList(sndIndex).Pos.y - MinYBorder + 1 To UserList(sndIndex).Pos.y + MinYBorder - 1
        For x = UserList(sndIndex).Pos.x - MinXBorder + 1 To UserList(sndIndex).Pos.x + MinXBorder - 1

            If MapData(sndMap, x, y).userindex > 0 Then
                If Left(sndData, 3) = "PLW" Then
                    If UserList(MapData(sndMap, x, y).userindex).Flags.Sound = "off" Then Exit Sub
                End If
                frmMain.Socket2(MapData(sndMap, x, y).userindex).Write sndData, Len(sndData)
            End If
        
        Next x
    Next y
    
    Exit Sub
End If

'Send to the UserIndex
If sndRoute = ToIndex Then
    If Left(sndData, 3) = "PLW" Then
        If UserList(ToIndex).Flags.Sound = "off" Then Exit Sub
    End If
    frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
    Exit Sub
End If

End Sub


Sub ConnectUser(ByVal userindex As Integer, ByVal Name As String, ByVal Password As String, ByVal ClientVer As String)
'*****************************************************************
'Reads the users .chr file and loads into Userlist array
'*****************************************************************
Dim Blank As User

'clent ver
If ClientVer <> "APLHA080" Then
    Call SendData(ToIndex, userindex, 0, "!!Your client is outdated.  Please download the newest client at http://inkey.angelcities.com/aspereta/.")
    CloseSocket (userindex)
    Exit Sub
End If

'Check for max users
If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, userindex, 0, "!!Too many users logged on. Try again later.")
    CloseSocket (userindex)
    Exit Sub
End If
  
'Check to see is user already logged with IP
If AllowMultiLogins = 0 Then
    If CheckForSameIP(userindex, frmMain.Socket2(userindex).PeerAddress) = True Then
        Call SendData(ToIndex, userindex, 0, "!!Multiple logging of the same IP is not allowed..")
        CloseSocket (userindex)
        Exit Sub
    End If
End If

'Check to see is user already logged with Name
If CheckForSameName(userindex, Name) = True Then
    Call SendData(ToIndex, userindex, 0, "!!That user is already logged on.")
    CloseSocket (userindex)
    Exit Sub
End If

'Check for Character file
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, userindex, 0, "!!Character does not exist.")
    CloseSocket (userindex)
    Exit Sub
End If

'Check Password
If Password <> GetVar(CharPath & UCase(Name) & ".chr", "INIT", "Password") Then
    Call SendData(ToIndex, userindex, 0, "!!Password incorrect.")
    CloseSocket (userindex)
    Exit Sub
End If

UserList(userindex) = Blank

'Load init vars from file
Call LoadUserInit(userindex, CharPath & UCase(Name) & ".chr")
Call LoadUserStats(userindex, CharPath & UCase(Name) & ".chr")

'Figure out where to put user
If UserList(userindex).Pos.map > 0 Then
    If MapInfo(UserList(userindex).Pos.map).StartPos.map > 0 Then
        UserList(userindex).Pos = MapInfo(UserList(userindex).Pos.map).StartPos
    End If
Else
    UserList(userindex).Pos = StartPos
End If

'Get closest legal pos
Call ClosestLegalPos(UserList(userindex).Pos, UserList(userindex).Pos)
If LegalPos(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y) = False Then
    Call SendData(ToIndex, userindex, 0, "!!No legal position found: Please try again.")
    CloseUser (userindex)
    Exit Sub
End If

'Get mod name
UserList(userindex).Name = Name
If WizCheck(UserList(userindex).Name) Then
    UserList(userindex).modName = Name & "*"
Else
    UserList(userindex).modName = Name
End If

'************** Initialize variables
UserList(userindex).Password = Password
UserList(userindex).IP = frmMain.Socket2(userindex).PeerAddress
UserList(userindex).Flags.UserLogged = 1

'Set switching map flag
UserList(userindex).Flags.SwitchingMaps = 1
'Send User index
Call SendData(ToIndex, userindex, 0, "SUI" & userindex)
'Tell client to try switching maps
Call SendData(ToIndex, userindex, 0, "SCM" & UserList(userindex).Pos.map & "," & MapInfo(UserList(userindex).Pos.map).MapVersion)
'Welcome message
Call SendData(ToIndex, userindex, 0, "#Welcome to Aspereta!" & FONTTYPE_INFO)

'Update inventory
Call UpdateUserInv(True, userindex, 0)
Call UpdateUserSpell(True, userindex, 0)

'update Num of Users
If userindex > LastUser Then LastUser = userindex
NumUsers = NumUsers + 1
frmMain.TxStatus.Text = "Total Users= " & NumUsers
MapInfo(UserList(userindex).Pos.map).NumUsers = MapInfo(UserList(userindex).Pos.map).NumUsers + 1

Call SendData(ToIndex, userindex, 0, "#Players: " & NumUsers & FONTTYPE_INFO)


'Update info
Call SendData(ToIndex, userindex, 0, "OKBOXWelcome to the Aspereta Alpha test.  Please make|sure to keep your client and yourself updated|at http://inkey.angelcities.com/aspereta/.||The Aspereta server is being ported to Linux at|the moment, so we are very busy.||You can now access the Aspereta Forum by pressing F!||     -Aspereta Team|0|1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|34|35|36|37|38|39|40|41|42|43|44|45|46|47|48|49|50|51|52|53|54|55|56|57|58|59|60|61|62|63|64|65|66|67|68|69|70|71|72|73|74|75|76")

'Show Character to others
Call MakeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)

'Refresh list box and send log on string
Call RefreshUserListBox

'Fix any typos that were in the save files.
If LCase(UserList(userindex).Path) = "wizzard" Then UserList(userindex).Path = "Wizard"
If LCase(UserList(userindex).Path) = "peasent" Then UserList(userindex).Path = "Peasant"
  
'Send flags
If UserList(userindex).Flags.Sound = "on" Then
    Call SendData(ToIndex, userindex, 0, "#Sound: on" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, userindex, 0, "#Sound: off" & FONTTYPE_INFO)
End If

If UserList(userindex).Flags.PK = "on" Then
    Call SendData(ToIndex, userindex, 0, "#PK: on" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, userindex, 0, "#PK: off" & FONTTYPE_INFO)
End If

'Send login phrase
Call SendData(ToAll, 0, 0, "#" & UserList(userindex).Name & " has logged in." & FONTTYPE_INFO)

'Log it
Open App.Path & "\Connect.log" For Append Shared As #5
Print #5, UserList(userindex).Name & " logged in. UserIndex:" & userindex & " " & Time & " " & Date
Close #5

Call SendData(ToIndex, userindex, 0, "SST" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.CurHP & "," & UserList(userindex).Stats.MaxMP & "," & UserList(userindex).Stats.CurMP & "," & UserList(userindex).Stats.Gold & "," & UserList(userindex).Stats.Lv & "," & UserList(userindex).Stats.Texp & "," & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)
Call SendData(ToIndex, userindex, 0, "TNL" & UserList(userindex).Stats.Exp & "," & UserList(userindex).Stats.Tnl)

End Sub

Sub CloseUser(ByVal userindex As Integer)
'*****************************************************************
'save user then reset user's slot
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim LoopC As Integer
Dim map As Integer
Dim Name As String


Call UngroupUser(userindex)

'Save temps
map = (UserList(userindex).Pos.map)
x = UserList(userindex).Pos.x
y = UserList(userindex).Pos.y
Name = UserList(userindex).Name

'Set logged to false
UserList(userindex).Flags.UserLogged = 0

'Save user
Call SaveUser(userindex, CharPath & UCase(Name) & ".chr")

'Erase user's character
UserList(userindex).Char.Body = 0
UserList(userindex).Char.Head = 0
UserList(userindex).Char.Heading = 0

If UserList(userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMap, 0, map, userindex)
End If

'Clear main vars
UserList(userindex).Name = ""
UserList(userindex).modName = ""
UserList(userindex).Password = ""
UserList(userindex).Pos.map = 0
UserList(userindex).Pos.x = 0
UserList(userindex).Pos.y = 0
UserList(userindex).IP = ""
UserList(userindex).RDBuffer = ""

'Clear Counters
UserList(userindex).Counters.IdleCount = 0
UserList(userindex).Counters.AttackCounter = 0
UserList(userindex).Counters.SendMapCounter.map = 0
UserList(userindex).Counters.SendMapCounter.x = 0
UserList(userindex).Counters.SendMapCounter.y = 0
UserList(userindex).Counters.HPCounter = 0
UserList(userindex).Counters.STACounter = 0

'Clear Flags
UserList(userindex).Flags.DownloadingMap = 0
UserList(userindex).Flags.SwitchingMaps = 0
UserList(userindex).Flags.StatsChanged = 0
UserList(userindex).Flags.ReadyForNextTile = 0

'update last user
If userindex = LastUser Then
    Do Until UserList(LastUser).Flags.UserLogged = 1
        LastUser = LastUser - 1
        If LastUser = 0 Then Exit Do
    Loop
End If
  
'update number of users
If NumUsers <> 0 Then
    NumUsers = NumUsers - 1
End If
frmMain.TxStatus.Text = "Total Users= " & NumUsers
Call RefreshUserListBox

'Update Map Users
MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1
If MapInfo(map).NumUsers < 0 Then
    MapInfo(map).NumUsers = 0
End If

'Send log off phrase
Call SendData(ToAll, 0, 0, "#" & Name & " disconnected." & FONTTYPE_INFO)

'Log it
Open App.Path & "\Connect.log" For Append Shared As #5
Print #5, Name & " logged off. " & "User Index:" & userindex & " " & Time & " " & Date
Close #5

LogData = Name & " logged off. " & "User Index:" & userindex & " " & Time & " " & Date & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
  
End Sub

Sub HandleData(ByVal userindex As Integer, ByVal rData As String)
'*****************************************************************
'Handles all data from the clients
'*****************************************************************
Dim sndData As String
Dim LoopC As Integer
Dim NPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim NpcIndex
Dim ItemSlot

'Check to see if user has a valid UserIndex
If userindex < 0 Then
    Exit Sub
End If
    
'Reset Idle
UserList(userindex).Counters.IdleCount = 0
    
'******************* Login Commands ****************************
    
'Logon on existing character
If Left$(rData, 5) = "LOGIN" Then
    rData = Right$(rData, Len(rData) - 5)
    
    Call ConnectUser(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44))
    
    LogData = UserList(userindex).Name & " connected at " & Time$ & "  IP: " & UserList(userindex).IP & FONTTYPE_TALK
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

    Exit Sub
End If
  
'Make a new character
If Left$(rData, 6) = "NLOGIN" Then
    rData = Right$(rData, Len(rData) - 6)
    
    Call ConnectNewUser(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44), Val(ReadField(3, rData, 44)), ReadField(4, rData, 44))
    
    LogData = UserList(userindex).Name & " created at " & Time$ & "  IP: " & UserList(userindex).IP & FONTTYPE_TALK
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
    
    Exit Sub
End If
  
'If not trying to log on must not be a client so log it off
If UserList(userindex).Flags.UserLogged = 0 Then
    CloseSocket (userindex)
    Exit Sub
End If
  
'******************* Flags *********************************************
If Left$(rData, 8) = "TOGSOUND" Then
    If UserList(userindex).Flags.Sound = "on" Then
        UserList(userindex).Flags.Sound = "off"
        Call SendData(ToIndex, userindex, 0, "#Sound: off" & FONTTYPE_INFO)
    Else
        UserList(userindex).Flags.Sound = "on"
        Call SendData(ToIndex, userindex, 0, "#Sound: on" & FONTTYPE_INFO)
    End If
End If

If Left$(rData, 5) = "TOGPK" Then
    If UserList(userindex).Flags.PK = "on" Then
        UserList(userindex).Flags.PK = "off"
        Call SendData(ToIndex, userindex, 0, "#PK: off" & FONTTYPE_INFO)
    Else
        UserList(userindex).Flags.PK = "on"
        Call SendData(ToIndex, userindex, 0, "#PK: on" & FONTTYPE_INFO)
    End If
End If

'******************* Communication Commands ****************************
'Say
If Left$(rData, 1) = ";" Then
    rData = Right$(rData, Len(rData) - 1)
    Call SendData(ToPCArea, userindex, UserList(userindex).Pos.map, "@" & UserList(userindex).Name & ": " & rData & FONTTYPE_TALK)
    Call SendData(ToPCArea, userindex, UserList(userindex).Pos.map, "CTXT" & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y & ",  " & UserList(userindex).Name & ": " & rData)
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "SAY " & UserList(userindex).Name & ": " & rData
    Close #5
    Exit Sub
End If

'Broadcast
If Left$(rData, 1) = "'" Then
    rData = Right$(rData, Len(rData) - 1)
    
    Call SendData(ToAll, 0, UserList(userindex).Pos.map, "@" & UserList(userindex).Name & "[world] " & rData & FONTTYPE_SHOUT)
    
    LogData = UserList(userindex).Name & ": " & rData & "  " & Time$ & "  IP: " & UserList(userindex).IP & FONTTYPE_SHOUT
    AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "SAGE " & UserList(userindex).Name & ": " & rData
    Close #5
    Exit Sub
    
    Exit Sub
End If
  
'Shout
If Left$(rData, 1) = "-" Then
    rData = Right$(rData, Len(rData) - 1)
    Call SendData(ToMap, 0, UserList(userindex).Pos.map, "@" & UserList(userindex).Name & "!  " & rData & FONTTYPE_SHOUT)
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "SHOUT " & UserList(userindex).Name & ": " & rData
    Close #5
    Exit Sub
    
    Exit Sub
End If
  
'Emote
'If Left$(rData, 1) = ":" Then
'    rData = Right$(rData, Len(rData) - 1)
'    Call SendData(ToPCArea, userindex, UserList(userindex).Pos.map, "@" & UserList(userindex).Name & " " & rData & FONTTYPE_TALK)
'    Exit Sub
'End If
  
'Whisper
If Left$(rData, 1) = "\" Then
    rData = Right$(rData, Len(rData) - 1)
    
    tName = ReadField(1, rData, 32)
    tIndex = NameIndex(tName)
    
    If tIndex <> 0 Then
    
        If Len(rData) <> Len(tName) Then
            tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
        Else
            tMessage = " "
        End If
        
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(userindex).Name & Chr(34) & " " & tMessage & FONTTYPE_WHISPER)
        Call SendData(ToIndex, userindex, 0, "@" & UserList(tIndex).Name & ">" & tMessage & FONTTYPE_WHISPER)
        LogData = UserList(userindex).Name & "->" & UserList(tIndex).Name & ": " & tMessage & FONTTYPE_WHISPER
        AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
        'Log it
        Open App.Path & "\Main.log" For Append Shared As #5
        Print #5, "WHISP " & LogData
        Close #5
        Exit Sub
        
        Exit Sub
    End If
    
    Call SendData(ToIndex, userindex, 0, "@User not online. " & FONTTYPE_INFO)
    Exit Sub
End If

'Who
If UCase$(rData) = "/WHO" Then
    Call SendData(ToIndex, userindex, 0, "#Total Users: " & NumUsers & FONTTYPE_INFO)
    
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") Then
            tStr = tStr & UserList(LoopC).modName & ", "
        End If
    Next LoopC
    tStr = Left$(tStr, Len(tStr) - 2)
    
    Call SendData(ToIndex, userindex, 0, "#" & tStr & FONTTYPE_INFO)
    
    Exit Sub
End If


'Ranking
If UCase$(rData) = "/RANK" Then
    Call SendData(ToIndex, userindex, 0, "@ **** Top " & TotalRanks & " listing ****" & FONTTYPE_INFO)
    tMessage = ""
    For LoopC = 1 To TotalRanks Step 2
        If LoopC + 1 <= TotalRanks Then
            tMessage = Str(LoopC) & ") " & PlayerRanking(LoopC).Path & " " & PlayerRanking(LoopC).Name & " (Lv" & Str(PlayerRanking(LoopC).Lv) & ")          "
            tMessage = tMessage + Str(LoopC + 1) & ") " & PlayerRanking(LoopC + 1).Path & " " & PlayerRanking(LoopC + 1).Name & " (Lv" & Str(PlayerRanking(LoopC + 1).Lv) & ")"
            Call SendData(ToIndex, userindex, 0, "@" & tMessage & FONTTYPE_INFO)
        End If
    Next LoopC
    'Call SendData(ToIndex, userindex, 0, "@" & LoopC & ") " & PlayerRanking(LoopC).Path & " " & PlayerRanking(LoopC).Name & " (Lv" & PlayerRanking(LoopC).Lv & ")" & FONTTYPE_INFO)
    'Call SendData(ToIndex, userindex, 0, "@" & tMessage & FONTTYPE_INFO)
    Exit Sub
End If


'******************* Npc Shops ***********************************
If Left$(rData, 4) = "SHOP" Then
    rData = Right$(rData, Len(rData) - 4)
    NpcIndex = Val(ReadField(1, rData, 44))
    ItemSlot = Val(ReadField(2, rData, 44))
    
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "BUY" Then Call SellUserItem(userindex, NpcIndex, ItemSlot)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "SELL" Then Call SellNPCItem(userindex, NpcIndex, ItemSlot)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "PATH" Then Call ChangePeasantJob(userindex, NpcIndex, ItemSlot)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "TSPELL" Then Call TeachSpell(userindex, NpcIndex, ItemSlot)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "BHP" Then Call BuyHitPoints(userindex)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "BMP" Then Call BuyMagicPoints(userindex)
    If NpcShops(NpcIndex).Slots(ItemSlot).Func = "CFACE" Then Call FaceChange(userindex, NpcIndex, ItemSlot)
    Exit Sub
End If


'******************* General Commands ****************************
'Refresh
If UCase(rData) = "REFRESH" Then
    Call WarpUserChar(userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
    Exit Sub
End If

'Move
If Left$(rData, 1) = "M" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    rData = Right$(rData, Len(rData) - 1)
    Call MoveUserChar(userindex, Val(rData))
    Exit Sub
Else
    UserList(userindex).Flags.StatsChanged = 1
    'Exit Sub
End If

'*************Turning**************
If rData = "FWEST" Then
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(userindex).Char.Heading = WEST
    Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
    Exit Sub
End If

If rData = "FEAST" Then
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(userindex).Char.Heading = EAST
    Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
    Exit Sub
End If

If rData = "FSOUTH" Then
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(userindex).Char.Heading = SOUTH
    Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
    Exit Sub
End If

If rData = "FNORTH" Then
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(userindex).Char.Heading = NORTH
    Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
    Exit Sub
End If

'************END Turning***********

'Attack
If rData = "ATT" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserAttack userindex
    Exit Sub
End If

'Left Click
If Left$(rData, 2) = "LC" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, userindex, 0, "#User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    Exit Sub
End If

'Right Click
If Left$(rData, 2) = "RC" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, userindex, 0, "#User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    rData = Right$(rData, Len(rData) - 2)
    Call LookatTile(userindex, UserList(userindex).Pos.map, ReadField(1, rData, 44), ReadField(2, rData, 44))
    Exit Sub
End If

'HELP
If UCase$(rData) = "/HELP" Then
    Call SendHelp(userindex)
    Exit Sub
End If

'Character Info
If UCase$(rData) = "/STATS" Then
    
    LogData = UserList(userindex).Path & " " & UserList(userindex).Name & "  Lv: " & UserList(userindex).Stats.Lv & FONTTYPE_SHOUT
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "" & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Vita: " & UserList(userindex).Stats.CurHP & "/" & UserList(userindex).Stats.MaxHP & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Mana: " & UserList(userindex).Stats.CurMP & "/" & UserList(userindex).Stats.MaxMP & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "" & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Str: " & UserList(userindex).Stats.Str & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Con: " & UserList(userindex).Stats.Con & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Int: " & UserList(userindex).Stats.Int & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Wis: " & UserList(userindex).Stats.Wis & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Dex: " & UserList(userindex).Stats.Dex & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "" & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "AC: " & UserList(userindex).Stats.AC & "   Dam: " & UserList(userindex).Stats.Dam & "     Hit: " & UserList(userindex).Stats.MinHIT & "-" & UserList(userindex).Stats.MaxHIT & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Exp: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.Tnl & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Total Exp: " & UserList(userindex).Stats.Texp & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    LogData = "Gold: " & UserList(userindex).Stats.Gold & FONTTYPE_TALK
    Call SendData(ToIndex, userindex, 0, "*" & LogData)
    If UserList(userindex).PoisonCount > 0 Then
        LogData = "" & FONTTYPE_TALK
        Call SendData(ToIndex, userindex, 0, "*" & LogData)
        LogData = UserList(userindex).PoisonName & "     " & UserList(userindex).PoisonCount / 20 & "s" & FONTTYPE_TALK
        Call SendData(ToIndex, userindex, 0, "*" & LogData)
        'LogData = "--------------------" & FONTTYPE_TALK
        'Call SendData(ToIndex, userindex, 0, "*" & LogData)
    End If
    Exit Sub
End If

'Quit
If UCase$(rData) = "/QUIT" Then
    Call UngroupUser(userindex)
    Call CloseSocket(userindex)
    Exit Sub
End If

'******************* Group Commands **************************
'Print out group
If UCase(Left$(rData, 6)) = "/GROUP" Then
    'Call LegalGroups
    'Call SendData(ToIndex, userindex, 0, "#" & UserList(userindex).GIndex & FONTTYPE_TALK)
    
    If UserList(userindex).GIndex = 0 Then
         Call SendData(ToIndex, userindex, 0, "#You are not in a group" & FONTTYPE_TALK)
    End If
    If UserList(userindex).GIndex > 0 Then
        For TmpI = 1 To 5
            If Groups(UserList(userindex).GIndex).UIndexes(TmpI) > 0 Then
                Call SendData(ToIndex, userindex, 0, "#" & UserList(Groups(UserList(userindex).GIndex).UIndexes(TmpI)).Name & FONTTYPE_TALK)
            End If
        Next TmpI
    End If
    Exit Sub
End If

'If UCase(Left$(rData, 6)) = "/GETID" Then
'    Call SendData(ToIndex, UserIndex, 0, "#Grouping ID: " & UserIndex & FONTTYPE_TALK)
'End If


'add a member
If UCase(Left$(rData, 2)) = "/G" Then
    rData = Right$(rData, Len(rData) - 2)
    Call GroupUser(userindex, rData)
    Exit Sub
End If



'Quit your group
If UCase(Left$(rData, 4)) = "/UNG" Then
    Call UngroupUser(userindex)
    Exit Sub
End If
'******************* Map Commands ****************************

'Request Map Update
If Left$(rData, 3) = "RMU" Then
    rData = Right$(rData, Len(rData) - 3)
    
    UserList(userindex).Flags.DownloadingMap = 1
    UserList(userindex).Flags.ReadyForNextTile = 1
    UserList(userindex).Counters.SendMapCounter.map = Val(rData)
    UserList(userindex).Counters.SendMapCounter.x = XMinMapSize
    UserList(userindex).Counters.SendMapCounter.y = YMinMapSize
    
    Call SendData(ToIndex, userindex, 0, "SMT" & MapInfo(Val(rData)).MapVersion)

End If

'Request Pos update
If rData = "RPU" Then
    Call SendData(ToIndex, userindex, 0, "SUP" & UserList(userindex).Pos.x & "," & UserList(userindex).Pos.y)
    UserList(userindex).Flags.StatsChanged = 1
    Exit Sub
End If

'Ready for next tile
If rData = "RNT" Then
    UserList(userindex).Flags.ReadyForNextTile = 1
    Exit Sub
End If

'Done Loading Map
If rData = "DLM" Then
    UserList(userindex).Flags.SwitchingMaps = 0
    Call SendData(ToIndex, userindex, 0, "SMN" & MapInfo(UserList(userindex).Pos.map).Name)
    Call SendData(ToIndex, userindex, 0, "PLM" & MapInfo(UserList(userindex).Pos.map).Music)
    
    Call SendData(ToIndex, userindex, 0, "DSM") 'Tell client to start drawing
    
    Call UpdateUserMap(userindex) 'Fill in all the characters and objects
    Call SendData(ToIndex, userindex, 0, "SUC" & UserList(userindex).Char.CharIndex)
    
    Exit Sub
End If

'******************* Object Commands ****************************

'Get
If rData = "GET" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    Call GetObj(userindex)
    Exit Sub
End If
  
'Drop
If Left$(rData, 3) = "DRP" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    rData = Right$(rData, Len(rData) - 3)
    If UserList(userindex).Object(ReadField(1, rData, 44)).ObjIndex = 0 Then
        Exit Sub
    End If
    Call DropObj(userindex, Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)), UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
    Exit Sub
End If

'USE
If Left$(rData, 3) = "USE" Then
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    rData = Right$(rData, Len(rData) - 3)
    If UserList(userindex).Object(Val(rData)).ObjIndex = 0 Then
        Exit Sub
    End If
    Call UseInvItem(userindex, Val(rData))
    Exit Sub
End If

'Drop gold
If UCase$(Left$(rData, 5)) = "/DROP" Then
    Dim Amount As Long
    rData = Right$(rData, Len(rData) - 5)
    If Len(rData) > 6 Then Exit Sub
    For Amount = 1 To Len(rData)
        If Mid(rData, Amount, 1) <> "1" And Mid(rData, Amount, 1) <> "2" And Mid(rData, Amount, 1) <> "3" And Mid(rData, Amount, 1) <> "4" And Mid(rData, Amount, 1) <> "5" And Mid(rData, Amount, 1) <> "6" And Mid(rData, Amount, 1) <> "7" And Mid(rData, Amount, 1) <> "8" And Mid(rData, Amount, 1) <> "9" And Mid(rData, Amount, 1) <> "0" Then Exit Sub
    Next Amount
    If rData = "" Then
        Amount = 1
    Else
        Amount = Int(rData)
    End If
    Call DropGold(userindex, Amount)
    Exit Sub
End If

'MOVE INV ITEMS
If UCase$(Left$(rData, 6)) = "CHANGE" Then
    rData = Right$(rData, Len(rData) - 6)
    Dim slot1 As Integer
    Dim slot2 As Integer
    Dim Obj As UserOBJ
    
    slot1 = Val(ReadField(1, rData, 44))
    slot2 = Val(ReadField(2, rData, 44))
    'Call SendData(ToIndex, userindex, 0, "#" + Str(Slot1) + ", " + Str(Slot2))
    If slot1 >= 1 And slot1 <= 38 And slot2 >= 1 And slot2 <= 38 Then
        If UserList(userindex).Object(slot1).Equipped <> 0 Or UserList(userindex).Object(slot2).Equipped <> 0 Then
            Call SendData(ToIndex, userindex, 0, "#You must remove the item(s) first." & FONTTYPE_TALK)
            Exit Sub
        End If
        Obj = UserList(userindex).Object(slot1)
        UserList(userindex).Object(slot1) = UserList(userindex).Object(slot2)
        UserList(userindex).Object(slot2) = Obj
        Call UpdateUserInv(True, userindex, slot1)
    End If
End If

'******************* SPELL RELATED COMMANDS *********************

'CAST
If Left$(rData, 4) = "CAST" Then
    Dim ClickX As Integer
    Dim ClickY As Integer
    Dim iSpell As Integer
    
    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    rData = Right$(rData, Len(rData) - 4)
    
    ClickX = ReadField(2, rData, 44)
    ClickY = ReadField(3, rData, 44)
    iSpell = ReadField(1, rData, 44)
    
    Select Case SpellData(iSpell).SpellType
        Case 1: Call CastSpell1(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 2: Call CastSpell2(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 3: Call CastSpell3(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 4: Call CastSpell4(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 5: Call CastSpell5(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 6: Call CastSpell6(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 7: Call CastSpell7(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 8: Call CastSpell8(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 9: Call CastSpell9(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
        Case 10: Call CastSpell10(userindex, UserList(userindex).Pos.map, ClickX, ClickY, iSpell)
    End Select
    Exit Sub
End If

'Forget a spell
If UCase$(Left$(rData, 7)) = "/FORGET" Then
    Dim Index As Integer
    rData = Right$(rData, Len(rData) - 7)
    If Len(rData) > 1 Then Exit Sub
    Index = Asc(UCase$(rData)) - 64
    If Index >= 0 And Index <= (Asc("Z") - 64) Then
        Call UserForgetSpell(userindex, Index)
    End If
End If

If UCase$(Left$(rData, 4)) = "SWAP" Then
    rData = Right$(rData, Len(rData) - 4)
    Dim Index1 As Integer
    Dim Index2 As Integer
    Dim TmpHld As Integer
    Index1 = Val(ReadField(1, rData, 44))
    Index2 = Val(ReadField(2, rData, 44))
    If Index1 >= 1 And Index1 <= 30 And Index2 >= 1 And Index2 <= 30 Then
        TmpHld = UserList(userindex).SpellBook(Index1).Spellindex
        UserList(userindex).SpellBook(Index1).Spellindex = UserList(userindex).SpellBook(Index2).Spellindex
        UserList(userindex).SpellBook(Index2).Spellindex = TmpHld
        Call UpdateUserSpell(True, userindex, Index1)
    End If
End If


'******************* Stat Commands ****************************
'Save
If UCase$(rData) = "/SAVE" Then
    Call SaveUser(userindex, CharPath & UCase(UserList(userindex).Name) & ".chr")
    Call SendData(ToIndex, userindex, 0, "#Character saved." & FONTTYPE_FIGHT)
    Exit Sub
End If

'Change Desc
'If UCase$(Left$(rData, 6)) = "/DESC " Then
'    rData = Right$(rData, Len(rData) - 6)
'    UserList(UserIndex).Desc = rData
'    Call SendData(ToIndex, UserIndex, 0, "#Description changed." & FONTTYPE_FIGHT)
'    Exit Sub
'End If

'*************** Wizard commands *****************************
If WizCheck(UserList(userindex).Name) = False Then
    Exit Sub
End If

'Reset Server
If UCase$(rData) = "/RESET" Then
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "!Reset started by " & UserList(userindex).Name & ". " & Time & " " & Date
    Close #5
    
    Call Restart
    Exit Sub
End If
  
'Shutdown server
If UCase$(rData) = "/SHUTDOWN" Then
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "!Shutdown started by " & UserList(userindex).Name & ". " & Time & " " & Date
    Close #5
    
    Unload frmMain
    Exit Sub
End If

'System Message
If UCase$(Left$(rData, 6)) = "/SMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    
    If rData <> "" Then
        Call SendData(ToAll, 0, 0, "!" & rData)
    End If
    
    Exit Sub
End If

'Spoof
If UCase$(Left$(rData, 7)) = "/SPOOF " Then
    rData = Right$(rData, Len(rData) - 7)
    
    tIndex = NameIndex(ReadField(1, rData, 32))
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToPCArea, tIndex, UserList(tIndex).Pos.map, "@" & rData & FONTTYPE_TALK)
    
    Exit Sub
End If
  
'Emergency System Message
If UCase$(Left$(rData, 6)) = "/EMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    
    If rData <> "" Then
        Call SendData(ToAll, 0, 0, "!!" & rData)
    End If
    
    Exit Sub
End If

'Control Code (send a command to all the clients)
If UCase$(Left$(rData, 4)) = "/CC " Then
    rData = Right$(rData, Len(rData) - 4)
    
    If rData <> "" Then
        Call SendData(ToAll, 0, 0, rData)
    End If
    
    Exit Sub
End If

'RP Message
If UCase$(Left$(rData, 6)) = "/RMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    
    If rData <> "" Then
        Call SendData(ToAll, 0, 0, "@" & rData & FONTTYPE_TALK)
    End If
    
    Exit Sub
End If

'Time
If UCase$(Left$(rData, 5)) = "/TIME" Then
    rData = Right$(rData, Len(rData) - 5)
    
        Call SendData(ToAll, 0, 0, "@At the tone, the server time will be: " & Time & " " & Date & FONTTYPE_INFO)
    
    Exit Sub
End If

'Where is
If UCase$(Left$(rData, 9)) = "/WHEREIS " Then
    rData = Right$(rData, Len(rData) - 9)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, userindex, 0, "@Loc for " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.map & ", " & UserList(tIndex).Pos.x & ", " & UserList(tIndex).Pos.y & "." & FONTTYPE_INFO)
    
    Exit Sub
End If

'Approach
If UCase$(Left$(rData, 6)) = "/APPR " Then
    rData = Right$(rData, Len(rData) - 6)

    'Don't allow if switching maps
    If UserList(userindex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, userindex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If

    'See if user online
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Find closest legal position and warp there
    ClosestLegalPos UserList(tIndex).Pos, NPos
    If LegalPos(NPos.map, NPos.x, NPos.y) Then
        Call WarpUserChar(userindex, NPos.map, NPos.x, NPos.y)
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(userindex).Name & " approached you." & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If

'Summon
If UCase$(Left$(rData, 5)) = "/SUM " Then
    rData = Right$(rData, Len(rData) - 5)
    
    'See if user online
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Don't allow if switching maps
    If UserList(tIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, userindex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Find closest legal position and warp there
    ClosestLegalPos UserList(userindex).Pos, NPos
    If LegalPos(NPos.map, NPos.x, NPos.y) Then
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(userindex).Name & " has summoned you." & FONTTYPE_INFO)
        Call WarpUserChar(tIndex, NPos.map, NPos.x, NPos.y)
    End If
    
    Exit Sub
End If

'GM Message
If UCase$(Left$(rData, 6)) = "/GMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    
    If rData <> "" Then
        Call SendData(ToAll, 0, 0, "@" & UserList(userindex).Name & "!>" & rData & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If

'Boot user
If UCase$(Left$(rData, 6)) = "/BOOT " Then
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "" & UserList(userindex).Name & " booted " & UserList(tIndex).Name & ". " & Time & " " & Date
    Close #5
    
    Call SendData(ToAll, 0, 0, "@" & UserList(userindex).Name & " booted " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    CloseSocket (tIndex)
    
    Exit Sub
End If

'Character modify
If UCase(Left(rData, 9)) = "/CHARMOD " Then
    rData = Right$(rData, Len(rData) - 9)
    
    tIndex = NameIndex(ReadField(1, rData, 32))
    Arg1 = ReadField(2, rData, 32)
    Arg2 = ReadField(3, rData, 32)
    Arg3 = ReadField(4, rData, 32)
    Arg4 = ReadField(5, rData, 32)

    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Don't allow if switching maps maps
    If UserList(tIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, userindex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Select Case UCase(Arg1)
    
        Case "Gold"
            UserList(tIndex).Stats.Gold = Val(Arg2)
            Call SendUserStatsBox(tIndex)

        Case "LVL"
            UserList(tIndex).Stats.Lv = Val(Arg2)
            Call SendUserStatsBox(tIndex)
    
        Case "BODY"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.map, tIndex, Val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(userindex).Char.Weapon)

        Case "HEAD"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.map, tIndex, UserList(tIndex).Char.Body, Val(Arg2), UserList(tIndex).Char.Heading, UserList(userindex).Char.Weapon)
        
        Case "WARP"
            If LegalPos(Val(Arg2), Val(Arg3), Val(Arg4)) Then
                Call WarpUserChar(tIndex, Val(Arg2), Val(Arg3), Val(Arg4))
            Else
                Call SendData(ToIndex, userindex, 0, "@Not a legal position." & FONTTYPE_INFO)
            End If
        
        Case Else
            Call SendData(ToIndex, userindex, 0, "@Not a charmod command." & FONTTYPE_INFO)
    
    End Select

    Exit Sub
End If

'**************************************

End Sub



