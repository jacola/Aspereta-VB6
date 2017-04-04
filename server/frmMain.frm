VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   Caption         =   "Aspereta Server 1.0"
   ClientHeight    =   6480
   ClientLeft      =   8850
   ClientTop       =   6240
   ClientWidth     =   13170
   FillColor       =   &H0000FF00&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   13170
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   2160
      Top             =   5400
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   1680
      Top             =   5400
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer MonsterTimer 
      Index           =   10
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   9
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   8
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   7
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   6
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   5
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   4
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   3
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   2
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer MonsterTimer 
      Index           =   1
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer PlayerTimer 
      Interval        =   50
      Left            =   4800
      Top             =   5400
   End
   Begin VB.TextBox txtResetTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Text            =   "1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdSvrStart 
      Caption         =   "Start Server"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Timer ResetTimer 
      Interval        =   600
      Left            =   4920
      Top             =   1680
   End
   Begin VB.CommandButton cmdUsrInfo 
      Caption         =   "Info"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox ServerLog 
      Height          =   6255
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer MonsterTimer 
      Index           =   0
      Interval        =   50
      Left            =   4320
      Top             =   5400
   End
   Begin VB.Timer AutoMaptimer 
      Interval        =   5
      Left            =   2640
      Top             =   5400
   End
   Begin VB.TextBox TxStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   4095
   End
   Begin VB.ListBox Userslst 
      BackColor       =   &H00000000&
      Columns         =   4
      ForeColor       =   &H0000FF00&
      Height          =   2475
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox txPortNumber 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   690
   End
   Begin VB.TextBox LocalAdd 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2130
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Users online:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "frmMain.frx":008B
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Server IP:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AutoMaptimer_Timer()
'*****************************************************************
'Send out map updates if needed
'*****************************************************************
Dim userindex As Integer
Dim LoopC As Integer

For userindex = 1 To LastUser
    If UserList(userindex).Flags.UserLogged = 1 Then
    
        'Send map chunk
        If UserList(userindex).Flags.DownloadingMap = 1 Then
            For LoopC = 1 To 15
                If UserList(userindex).Flags.DownloadingMap Then
                    SendNextMapTile userindex
                End If
            Next LoopC
        End If
        
    End If
Next userindex

End Sub

Private Sub cmdReset_Click()
'*******************************
'*** Starts a server reset  ****
'*******************************
' Ten Second wait

CountDown = Val(txtResetTime.Text)

Call SendData(ToAll, 0, 0, "@   *** RESET IN " & CountDown & " SECONDS ***" & FONTTYPE_WARNING)
LogData = "   ***  Reset in " & CountDown & " seconds." & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
cmdReset.Enabled = False

End Sub

Private Sub CmdSvrStart_Click()

'*****************************************************************
'Load up server
'*****************************************************************
Dim LoopC As Integer
Dim StartTime As String

CmdSvrStart.Enabled = False

'*** Init vars ***
frmMain.Caption = frmMain.Caption '& " V." & App.Major & "." & App.Minor & "." & App.Revision
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

StartTime = Time$
LogData = "Server loading started at " & StartTime & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'Setup Map borders
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)

'Reset User connections
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next LoopC
LogData = "User connections cleared at " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


'*** Load data ***

Call LoadSini
LogData = "Server initialization data loaded at " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

LogData = "   ***  Loading NPC types *** " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Call LoadNPCData

LogData = "   ***  Loading Maps *** " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Call LoadMapData
LogData = "   ***  Loading Objects *** " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Call LoadOBJData
LogData = "   *** Loading Spells *** " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Call LoadSpellData

LogData = "   *** Loading Rank Data *** " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
Call LoadRankData

'*** Setup sockets ***
LogData = "Sockets set at " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = SOCKET_BUFFER_SIZE

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = SOCKET_BUFFER_SIZE

'*** Listen ***
frmMain.Socket1.LocalPort = Val(frmMain.txPortNumber.Text)
frmMain.Socket1.Listen
  
'*** Misc ***
'Hide
If HideMe = 1 Then
    frmMain.Hide
End If

'Show status
frmMain.TxStatus.Text = "Listening for connection ..."
Call RefreshUserListBox

'Show local IP
frmMain.LocalAdd.Text = frmMain.Socket1.LocalAddress

'Log it
Open App.Path & "\Main.log" For Append Shared As #5
Print #5, "Server started. " & Time & " " & Date
Close #5

LogData = "Total NPCs " & LastNPC & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

LogData = "Server started at " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

LogData = "Started on IP: " & LocalAdd.Text & "  Port: " & txPortNumber.Text & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

LogData = StartTime & " -> " & Time$ & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

CmdSvrStart.Visible = False
cmdReset.Visible = True
txtResetTime.Visible = True

End Sub

Private Sub cmdUsrInfo_Click()

Dim userindex As Integer
userindex = Userslst.ListIndex + 1

If userindex = 0 Then Exit Sub
    
'name path level
LogData = UserList(userindex).Path & " " & UserList(userindex).Name & "  Level: " & UserList(userindex).Stats.Lv & " (" & Time$ & ")" & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'hp and mp
LogData = "Vita: " & UserList(userindex).Stats.CurHP & "/" & UserList(userindex).Stats.MaxHP & "     Mana: " & UserList(userindex).Stats.CurMP & "/" & UserList(userindex).Stats.MaxMP & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'stats
LogData = "Str: " & UserList(userindex).Stats.Str & " | Con: " & UserList(userindex).Stats.Con & " | Int: " & UserList(userindex).Stats.Int & " | Wis: " & UserList(userindex).Stats.Wis & " | Dex: " & UserList(userindex).Stats.Dex & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'ac dam etc
LogData = "AC: " & UserList(userindex).Stats.AC & "   Dam: " & UserList(userindex).Stats.Dam & "   AC: " & UserList(userindex).Stats.AC & "     Hit: " & UserList(userindex).Stats.MinHIT & "-" & UserList(userindex).Stats.MaxHIT & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'exp
LogData = "Exp: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.Tnl & "    Total Exp: " & UserList(userindex).Stats.Texp & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'gold
LogData = "Gold: " & UserList(userindex).Stats.Gold & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


End Sub


Private Sub Form_Load()

CountDown = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim LoopC As Integer

'ensure that the sockets are closed, ignore any errors
On Error Resume Next

Socket1.Cleanup

For LoopC = 1 To MaxUsers
    CloseSocket (LoopC)
Next

'Log it
Open App.Path & "\Main.log" For Append Shared As #5
Print #5, "Server unloaded. " & Time & " " & Date
Close #5

Call SaveRankData

End

End Sub







Private Sub MonsterTimer_Timer(Index As Integer)

'*****************************************************************
'update world
'*****************************************************************
Dim userindex As Integer
Dim NpcIndex As Integer
Dim TempPos As WorldPos
Dim tmpperc
Dim RegenVal As Integer

'Update NPCs
For NpcIndex = Index * 50 To Index * 50 + 49

'LastNPC
If NpcIndex > 0 And NpcIndex < LastNPC Then
    'make sure NPC is active
    If NPCList(NpcIndex).Flags.NPCActive = 1 Then
        
        If NPCList(NpcIndex).ParaCount > 0 Then
            NPCList(NpcIndex).ParaCount = NPCList(NpcIndex).ParaCount - 1
        End If
        
        If NPCList(NpcIndex).PoisonCount > 0 Then
            NPCList(NpcIndex).PoisonCount = NPCList(NpcIndex).PoisonCount - 1
            NPCList(NpcIndex).Stats.CurHP = NPCList(NpcIndex).Stats.CurHP + NPCList(NpcIndex).PoisonDamage
            If NPCList(NpcIndex).Stats.CurHP < 1 Then NPCList(NpcIndex).Stats.CurHP = 1
            If NPCList(NpcIndex).Stats.CurHP > NPCList(NpcIndex).Stats.MaxHP Then NPCList(NpcIndex).Stats.CurHP = NPCList(NpcIndex).Stats.MaxHP
        End If
        
        If MapInfo(NPCList(NpcIndex).StartPos.map).NumUsers > 0 Then
            If NPCList(NpcIndex).Flags.NPCAlive Then
                If NPCList(NpcIndex).ParaCount <= 0 Then
                    NPCList(NpcIndex).Counters.Movement = NPCList(NpcIndex).Counters.Movement + 1
                    If NPCList(NpcIndex).Counters.Movement >= NPCList(NpcIndex).Speed Then
                        NPCList(NpcIndex).Counters.Movement = 0
                        Call NPCAI(NpcIndex)
                    End If
                End If
            End If
            Call ChangeNPCChar(ToMap, 0, NPCList(NpcIndex).Pos.map, NpcIndex, NPCList(NpcIndex).Char.Body, NPCList(NpcIndex).Char.Head, NPCList(NpcIndex).Char.Heading)
        End If
        If NPCList(NpcIndex).Flags.NPCAlive = False Then
            If NPCList(NpcIndex).RespawnWait > 0 Then
                NPCList(NpcIndex).Counters.RespawnCounter = NPCList(NpcIndex).Counters.RespawnCounter - 1
                If NPCList(NpcIndex).Counters.RespawnCounter <= 0 Then
                    SpawnNPC NpcIndex, NPCList(NpcIndex).StartPos.map, NPCList(NpcIndex).StartPos.x, NPCList(NpcIndex).StartPos.y
                End If
            End If
        End If
    End If
End If

Next NpcIndex


End Sub

Private Sub PlayerTimer_Timer()

'*****************************************************************
'update world
'*****************************************************************
Dim userindex As Integer
Dim NpcIndex As Integer
Dim TempPos As WorldPos
Dim tmpperc
Dim RegenVal As Integer

'Update Users
For userindex = 1 To LastUser

    'make sure user is logged on
    If UserList(userindex).Flags.UserLogged = 1 Then
        
        'Do special tile events
        Call DoTileEvents(userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.y)
        
        'Update stats
        UserList(userindex).Counters.Regen = UserList(userindex).Counters.Regen + 1
        
        If UserList(userindex).Counters.Regen >= 20 Then
            If UserList(userindex).Stats.CurHP < UserList(userindex).Stats.MaxHP Then
                If Int(Rnd * 4) = 1 Then
                    If UserList(userindex).Stats.MaxHP > 100 Then
                        RegenVal = UserList(userindex).Stats.MaxHP * 0.01
                        If RegenVal > 5 Then RegenVal = 5
                        UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP + RegenVal
                        UserList(userindex).Flags.StatsChanged = 1
                    Else
                        UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP + 1
                        UserList(userindex).Flags.StatsChanged = 1
                    End If
                End If
            End If
            If UserList(userindex).Stats.CurMP < UserList(userindex).Stats.MaxMP Then
                UserList(userindex).Stats.CurMP = UserList(userindex).Stats.CurMP + (UserList(userindex).Stats.MaxMP * 0.05)
                If UserList(userindex).Stats.CurMP > UserList(userindex).Stats.MaxMP Then
                    UserList(userindex).Stats.CurMP = UserList(userindex).Stats.MaxMP
                End If
                UserList(userindex).Flags.StatsChanged = 1
            End If
            UserList(userindex).Counters.Regen = 0
        End If
        
        If UserList(userindex).PoisonCount > 0 Then
            UserList(userindex).PoisonCount = UserList(userindex).PoisonCount - 1
            UserList(userindex).Stats.CurHP = UserList(userindex).Stats.CurHP + UserList(userindex).PoisonDamage
            If UserList(userindex).Stats.CurHP < 1 Then UserList(userindex).Stats.CurHP = 1
            If UserList(userindex).Stats.CurHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.CurHP = UserList(userindex).Stats.MaxHP
        End If
        
        
        'Update attack counter
        If UserList(userindex).Counters.AttackCounter > 0 Then
            UserList(userindex).Counters.AttackCounter = UserList(userindex).Counters.AttackCounter - 1
        End If
        
        'Update Stats box if need be
        If UserList(userindex).Flags.StatsChanged Then
            SendUserStatsBox userindex
            tmpperc = UserList(userindex).Stats.CurHP / UserList(userindex).Stats.MaxHP * 100
            Call SendData(ToIndex, userindex, UserList(userindex).Pos.map, "VC" & userindex & "," & tmpperc)
            
            If UserList(userindex).SendSpell > 0 Then
                Call SendData(ToIndex, userindex, UserList(userindex).Pos.map, "SP" & userindex & "," & UserList(userindex).SendSpell)
                UserList(userindex).SendSpell = 0
            End If
            
            UserList(userindex).Flags.StatsChanged = 0
            Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.Weapon)
        End If
            
        'Update idle counter
        UserList(userindex).Counters.IdleCount = UserList(userindex).Counters.IdleCount + 1
        If UserList(userindex).Counters.IdleCount >= IdleLimit Then
            Call SendData(ToIndex, userindex, 0, "!!Sorry you have been idle to long. Disconnected..")
            Call CloseSocket(userindex)
        End If
            
    End If

Next userindex

End Sub

Private Sub ResetTimer_Timer()

If CountDown > 0 Then
    CountDown = CountDown - 1
    If CountDown = 180 Or CountDown = 120 Or CountDown = 60 Or CountDown = 30 Or CountDown = 20 Or CountDown <= 10 Then
        LogData = "   ***  Reset in " & CountDown & " seconds." & FONTTYPE_SHOUT
        AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))
        Call SendData(ToAll, 0, 0, "#Reset in " & CountDown & " seconds." & FONTTYPE_WARNING)
    End If
End If

If CountDown = 0 Then
    Call Restart
    cmdReset.Enabled = True
    CountDown = -1
End If

End Sub

Sub Socket1_Accept(SocketId As Integer)
'*********************************************
'Accepts new user and assigns an open Index
'*********************************************
Dim Index As Integer

Index = NextOpenUser

If UserList(Index).ConnID >= 0 Then
    'Close down user socket
    Call CloseSocket(Index)
End If

UserList(Index).ConnID = SocketId
Load Socket2(Index)

Socket2(Index).AddressFamily = AF_INET
Socket2(Index).Protocol = IPPROTO_IP
Socket2(Index).SocketType = SOCK_STREAM
Socket2(Index).Binary = False
Socket2(Index).BufferSize = SOCKET_BUFFER_SIZE
Socket2(Index).Blocking = False

Socket2(Index).Accept = SocketId

End Sub

Sub Socket2_Disconnect(Index As Integer)
'*********************************************
'Begins close procedure
'*********************************************

CloseSocket (Index)

End Sub


Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
'*********************************************
'Seperate lines by ENDC and send each to HandleData()
'*********************************************
Dim LoopC As Integer
Dim RD As String
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer

Socket2(Index).Read RD, DataLength

'Check for previous broken data and add to current data
If UserList(Index).RDBuffer <> "" Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

'Check for more than one line
sChar = 1
For LoopC = 1 To Len(RD)

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

'Check for broken line and save for next time
If Len(RD) - (sChar - 1) <> 0 Then
    UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))
End If

'Send buffer to Handle data
For LoopC = 1 To CR
    Call HandleData(Index, rBuffer(LoopC))
Next LoopC

End Sub



