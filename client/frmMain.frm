VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Aspereta"
   ClientHeight    =   11160
   ClientLeft      =   -255
   ClientTop       =   -240
   ClientWidth     =   17100
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   744
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1140
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   7080
      Top             =   240
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
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
   Begin VB.Timer GFXTimer 
      Interval        =   20
      Left            =   4200
      Top             =   8160
   End
   Begin VB.Frame frmForum 
      Caption         =   "Frame1"
      Height          =   4665
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   9375
      Begin SHDocVwCtl.WebBrowser wbForum 
         Height          =   4095
         Left            =   0
         TabIndex        =   30
         Top             =   280
         Width           =   9375
         ExtentX         =   16536
         ExtentY         =   7223
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Image ForumClose 
         Height          =   300
         Left            =   -230
         Picture         =   "frmMain.frx":030A
         Top             =   4380
         Width           =   9600
      End
      Begin VB.Image FourmTitle 
         Height          =   300
         Left            =   0
         Picture         =   "frmMain.frx":994C
         Top             =   0
         Width           =   9600
      End
   End
   Begin VB.Frame MPFrame 
      Caption         =   "Frame1"
      Height          =   225
      Left            =   7200
      TabIndex        =   26
      Top             =   5295
      Width           =   2250
      Begin VB.Label LblMP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   -45
         Width           =   120
      End
      Begin VB.Image MANShp 
         Height          =   225
         Left            =   0
         Picture         =   "frmMain.frx":12F8E
         Top             =   0
         Width           =   2250
      End
      Begin VB.Image imgHPBack 
         Height          =   225
         Left            =   0
         Picture         =   "frmMain.frx":14A4C
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.TextBox txtMousePos 
      Height          =   375
      Left            =   9720
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   390
      Left            =   0
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6885
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.Frame MapLoadFrame 
      Caption         =   "Loading"
      Height          =   207
      Left            =   10440
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   1561
      Begin VB.Image LoadForward 
         Height          =   210
         Left            =   0
         Picture         =   "frmMain.frx":1650A
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image LoadBack 
         Height          =   210
         Left            =   0
         Picture         =   "frmMain.frx":1765C
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.Frame svrAlert 
      Caption         =   "Frame1"
      Height          =   2981
      Left            =   600
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   7464
      Begin VB.TextBox svrMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C000&
         Height          =   1455
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Image SvrOK 
         Height          =   615
         Left            =   6240
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Image SvrMsgImage 
         Appearance      =   0  'Flat
         Height          =   3000
         Left            =   0
         Picture         =   "frmMain.frx":187AE
         Top             =   0
         Width           =   7500
      End
   End
   Begin VB.TextBox Macros 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   10440
      TabIndex        =   11
      Text            =   "1,2,3,4,5,6,7,8,9,0"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Frame NpcFrame 
      BackColor       =   &H00000000&
      Caption         =   "NPC"
      ForeColor       =   &H0000C000&
      Height          =   3933
      Left            =   1800
      TabIndex        =   6
      Top             =   1650
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton CmdNpcSelectx10 
         Caption         =   "Select x10"
         Height          =   345
         Left            =   3240
         TabIndex        =   12
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton CmdNpcSelect 
         Caption         =   "Select"
         Height          =   345
         Left            =   4320
         TabIndex        =   10
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton cmdNpcLeave 
         Caption         =   "Done"
         Height          =   345
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   735
      End
      Begin VB.ListBox NpcList 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1605
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label NpcTalk 
         BackColor       =   &H00000000&
         Caption         =   "Npc Talking"
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4815
      End
      Begin VB.Image Image4 
         Height          =   3930
         Left            =   0
         Picture         =   "frmMain.frx":61BD0
         Top             =   0
         Width           =   5295
      End
   End
   Begin RichTextLib.RichTextBox StatusBox 
      Height          =   1125
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1984
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":A58EA
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
   Begin VB.TextBox TxtCatch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox DrpAmountTxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11400
      TabIndex        =   1
      Text            =   "1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer FPSTimer 
      Interval        =   1000
      Left            =   7560
      Top             =   240
   End
   Begin RichTextLib.RichTextBox StatBox 
      Height          =   4080
      Left            =   12720
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7197
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":A596C
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
   Begin VB.Frame HPFrame 
      Caption         =   "Frame1"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   7200
      TabIndex        =   17
      Top             =   5040
      Width           =   2250
      Begin VB.Shape ExpBack 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   105
         Top             =   2025
         Width           =   2280
      End
      Begin VB.Label LblHP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   -40
         Width           =   120
      End
      Begin VB.Image HPshp 
         Height          =   225
         Left            =   0
         Picture         =   "frmMain.frx":A59EE
         Top             =   0
         Width           =   2250
      End
      Begin VB.Image imgMPBack 
         Height          =   225
         Left            =   0
         Picture         =   "frmMain.frx":A74AC
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.Frame TNLFrame 
      Caption         =   "Frame1"
      Height          =   120
      Left            =   4800
      TabIndex        =   31
      Top             =   5400
      Width           =   2295
      Begin VB.Image ExpProg 
         Height          =   105
         Left            =   0
         Picture         =   "frmMain.frx":A8F6A
         Top             =   0
         Width           =   2250
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Label MapNameLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(no map loaded)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10800
      TabIndex        =   29
      Top             =   5160
      Width           =   1725
   End
   Begin VB.Label lblFPS 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Image CmdSpells 
      Height          =   240
      Left            =   9000
      Picture         =   "frmMain.frx":A9C08
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image CmdItems 
      Height          =   240
      Left            =   8760
      Picture         =   "frmMain.frx":A9F4A
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image CmdPvP 
      Height          =   240
      Left            =   8520
      Picture         =   "frmMain.frx":AA28C
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image CmdSound 
      Height          =   240
      Left            =   8280
      Picture         =   "frmMain.frx":AA5CE
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image CmdHelp 
      Height          =   240
      Left            =   8040
      Picture         =   "frmMain.frx":AA910
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   7800
      Picture         =   "frmMain.frx":AAC52
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image CmdExit 
      Height          =   240
      Left            =   9360
      Picture         =   "frmMain.frx":AAF94
      Top             =   6960
      Width           =   240
   End
   Begin VB.Image GetCmd 
      Height          =   240
      Left            =   7560
      Picture         =   "frmMain.frx":AB2D6
      Top             =   6960
      Width           =   240
   End
   Begin VB.Label LblTexp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   12960
      TabIndex        =   23
      Top             =   480
      Width           =   120
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "noname"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   22
      Top             =   360
      Width           =   975
   End
   Begin VB.Label LblGold 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12960
      TabIndex        =   21
      Top             =   840
      Width           =   120
   End
   Begin VB.Label LblLvl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12960
      TabIndex        =   20
      Top             =   240
      Width           =   120
   End
   Begin VB.Label ExpLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12840
      TabIndex        =   19
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label CurSpell 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spell: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   10560
      TabIndex        =   5
      Top             =   2400
      Width           =   2385
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   6885
      Width           =   9615
   End
   Begin VB.Label LblXY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10800
      TabIndex        =   2
      Top             =   2880
      Width           =   345
   End
   Begin VB.Shape MainViewShp 
      Height          =   7200
      Left            =   0
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseDownSlot As Integer
Dim MouseButton As Integer

Private Sub CmdExit_Click()
    Dim loopC As Integer
    
    SendData "/ung"
    
    SendData "/SAVE"
    
    Open IniPath & "\user\" & frmConnect.NameTxt.Text + ".ini" For Output As #1
    For loopC = 1 To 10
        Write #1, UserHotButtons(loopC)
    Next loopC
    Close #1
    
    If TextBoxAlwaysOn = True Then
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "TextBoxAlwaysOn", "true")
    Else
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "TextBoxAlwaysOn", "false")
    End If
    If StatusFilter = True Then
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "StatusFilter", "true")
    Else
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "StatusFilter", "false")
    End If
    If StatusBox.Visible = True Then
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "StatusBox", "true")
    Else
        Call WriteVar(IniPath & "Game.ini", "SETTINGS", "StatusBox", "false")
    End If
    
    prgRun = False
End Sub

Private Sub CmdHelp_Click()
    SendData "/HELP"
End Sub

Private Sub CmdItems_Click()

ShowInventory = True

End Sub

Private Sub cmdNpcLeave_Click()
    TxtCatch.SetFocus
    NpcFrame.Visible = False
End Sub

Private Sub CmdNpcSelect_Click()
    TxtCatch.SetFocus
    SendData "SHOP" & CurrentNPCShop & "," & NpcList.ListIndex + 1
End Sub

Private Sub CmdNpcSelectx10_Click()
    Dim iLoop
    TxtCatch.SetFocus
    For iLoop = 1 To 10
        SendData "SHOP" & CurrentNPCShop & "," & NpcList.ListIndex + 1
    Next iLoop
End Sub

Private Sub CmdPvP_Click()
    SendData "TOGPK"
End Sub


Private Sub CmdSound_Click()
    SendData "TOGSOUND"
End Sub

Private Sub CmdSpells_Click()

If ShowSpells = False Then
    ShowSpells = True
Else
    ShowSpells = False
End If

End Sub



Private Sub DrpAmountTxt_Change()

'Make sure amount is legal
If DrpAmountTxt.Text < 1 Then
    DrpAmountTxt.Text = MAX_INVENTORY_OBJS
End If

If DrpAmountTxt.Text > MAX_INVENTORY_OBJS Then
    DrpAmountTxt.Text = 1
End If

End Sub

Private Sub Form_DblClick()
Dim Y As Integer
Dim X As Integer

If MouseButton <> 1 Then Exit Sub

'Double click hot button
For X = 0 To 9
    If CurMouseX >= (X * 36 + 281) And CurMouseY >= (2) And CurMouseX <= (X * 36 + 315) And CurMouseY <= 36 Then
        If UserHotButtons(X + 1) <= 100 Then
            SendData "USE" & UserHotButtons(X + 1)
            Exit Sub
        End If
        If UserHotButtons(X + 1) >= 101 And UserHotButtons(X + 1) <= 200 Then
            CurSpellIndex = UserHotButtons(X + 1) - 100
            Targeting = True
            Call SetTarget
            Exit Sub
        End If
        
        Exit Sub
    End If
Next X

'Double click to cast spell
If ShowSpells = True Then
    For Y = 0 To 5
        For X = 0 To 4
            If CurMouseX >= (X * 36 + 2) And CurMouseY >= (Y * 36) And CurMouseX <= (X * 36 + 36) And CurMouseY <= (Y * 36 + 34) Then
                CurSpellIndex = Y * 5 + X + 1
                Targeting = True
                Call SetTarget
                Exit Sub
            End If
        Next X
    Next Y
End If

'Double click to use inventory
If ShowInventory = True Then
    For Y = 0 To 5
        For X = 0 To 4
            If CurMouseX >= (X * 36 + 460) And CurMouseY >= (Y * 36 + 61) And CurMouseX <= (X * 36 + 494) And CurMouseY <= (Y * 36 + 95) Then
                SendData "USE" & Y * 5 + X + 1
                Exit Sub
            End If
        Next X
    Next Y
    Exit Sub
End If

'For the paper doll/status
If ShowStatus = True Then
    For Y = 0 To 7
        If CurMouseX >= 202 And CurMouseY >= (51 + Y * 36) And CurMouseX <= 238 And CurMouseY <= (51 + Y * 36 + 36) Then
            SendData "USE" & (31 + Y)
            Exit Sub
        End If
    Next Y
End If


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If frmMain.NpcFrame.Visible = True Then
    Exit Sub
End If

If Len(TxtCatch.Text) > 2 Then
    TxtCatch.Text = ""
End If

If TxtCatch.Text <> "," Or TxtCatch.Text <> " " Then
    TxtCatch.Text = ""
End If
'SendTxt.Text = Str$(KeyCode)


End Sub


Private Sub Form_Load()

    frmMain.Left = 0
    frmMain.Top = 0
    frmMain.Width = 12000
    frmMain.Height = 9000
    
    ChatPos = 0
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tx As Integer
Dim ty As Integer
MouseButton = Button

If ShowSpells = True And Button = 1 Then
    For ty = 0 To 5
        For tx = 0 To 4
            If CurMouseX >= (tx * 36 + 2) And CurMouseY >= (ty * 36) And CurMouseX <= (tx * 36 + 36) And CurMouseY <= (ty * 36 + 34) Then
                MouseDownSlot = ty * 5 + tx + 1 + 100
                DragIndex = UserSpellbook(ty * 5 + tx + 1).Icon.GrhIndex
                Exit Sub
            End If
        Next tx
    Next ty
End If

If ShowInventory = True And Button = 1 Then
    For ty = 0 To 5
        For tx = 0 To 4
            If CurMouseX >= (tx * 36 + 460) And CurMouseY >= (ty * 36 + 61) And CurMouseX <= (tx * 36 + 494) And CurMouseY <= (ty * 36 + 95) Then
                MouseDownSlot = ty * 5 + tx + 1
                DragIndex = UserInventory(ty * 5 + tx + 1).GrhIndex
                Exit Sub
            End If
        Next tx
    Next ty
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CurMouseX = X
CurMouseY = Y



If X > MainViewShp.Left And X < MainViewShp.Left + MainViewShp.Width And Y > MainViewShp.Top And Y < MainViewShp.Top + MainViewShp.Height Then
    ConvertCPtoSTP MainViewShp.Left, MainViewShp.Top, X, Y, pMouseX, pMouseY
    'frmMain.MousePointer = 5
'Else
    'frmMain.MousePointer = 1
End If

txtMousePos.Text = "(" + Str(X) + ", " + Str(Y) + ")"

If X > 10 And Y > 370 And X < 635 And Y < 455 Then
    TextBoxOn = True
Else
    TextBoxOn = False
End If


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************************
'See if user is clicking in the view window then send
'the tile click position to the server
'*****************************************************************
Dim tx As Integer
Dim ty As Integer
Dim A As String
Dim B As String

'Make sure engine is running
If EngineRun = False Then Exit Sub

'Don't do if downloading map
If DownloadingMap = True Then
    Exit Sub
End If

'*** if an item is dragged put it in hot buttons or drop it if it isn't or move it to new spot
If DragIndex <> 0 Then
    ' Hot button
    For tx = 0 To 9
        If CurMouseX >= (tx * 36 + 281) And CurMouseY >= (2) And CurMouseX <= (tx * 36 + 315) And CurMouseY <= 36 Then
            UserHotButtons(tx + 1) = MouseDownSlot
            MouseDownSlot = 0
            DragIndex = 0
            Exit Sub
        End If
    Next tx
    
    ' Move
    If MouseDownSlot <= 100 Then
        For ty = 0 To 5
            For tx = 0 To 4
                If CurMouseX >= (tx * 36 + 460) And CurMouseY >= (ty * 36 + 61) And CurMouseX <= (tx * 36 + 494) And CurMouseY <= (ty * 36 + 95) Then
                    A = MouseDownSlot
                    B = ty * 5 + tx + 1
                    If A <> B Then
                        SendData ("CHANGE" + A + "," + B)
                        'AddToTalk ("CHANGE" + A + "," + B)
                    End If
                    MouseDownSlot = 0
                    DragIndex = 0
                    Exit Sub
                End If
            Next tx
        Next ty
        ' Paper doll
        'If ShowStatus = True Then
        '    For ty = 0 To 7
        '        If CurMouseX >= 202 And CurMouseY >= (51 + ty * 36) And CurMouseX <= 238 And CurMouseY <= (51 + ty * 36 + 36) Then
        '            A = MouseDownSlot
        '            B = ty + 31
        '            SendData ("CHANGE" + A + "," + B)
        '            MouseDownSlot = 0
        '            DragIndex = 0
        '            Exit Sub
        '        End If
        '    Next ty
        'End If
    End If
    
    'move a spell
    If MouseDownSlot >= 101 And MouseDownSlot <= 200 Then
        For ty = 0 To 5
            For tx = 0 To 4
                If CurMouseX >= (tx * 36 + 2) And CurMouseY >= (ty * 36) And CurMouseX <= (tx * 36 + 36) And CurMouseY <= (ty * 36 + 34) Then
                    A = MouseDownSlot - 100
                    B = ty * 5 + tx + 1
                    SendData ("SWAP" + A + "," + B)
                    'AddToTalk ("CHANGE" + A + "," + B)
                    MouseDownSlot = 0
                    DragIndex = 0
                    Exit Sub
                End If
            Next tx
        Next ty
    End If
    
    'Drop onto paper doll (use)
    If CurMouseX >= 197 And CurMouseY >= 46 And CurMouseX <= 397 And CurMouseY <= 336 And ShowStatus = True Then
        If MouseDownSlot < 100 Then SendData "USE" & MouseDownSlot
        MouseDownSlot = 0
        DragIndex = 0
        Exit Sub
    End If
    
    ' Drop
    If MouseDownSlot < 100 Then
        SendData "DRP" & MouseDownSlot & "," & 1
        MouseDownSlot = 0
        DragIndex = 0
    Else
        MouseDownSlot = 0
        DragIndex = 0
    End If
End If

'*** Clear out the macro'd items by right clickng
If Button = 2 Then
    For tx = 0 To 9
       If CurMouseX >= (tx * 36 + 281) And CurMouseY >= (2) And CurMouseX <= (tx * 36 + 315) And CurMouseY <= 36 Then
           UserHotButtons(tx + 1) = 0
       End If
    Next tx
End If

If ShowOKBox = True Then
    If X >= 343 And Y >= 236 And X <= 390 And Y <= 259 Then ShowOKBox = False
    '372, 215
    '386, 230
    If Button = 1 And X >= 372 And Y >= 215 And X <= 386 And Y <= 230 Then OkBoxPos = OkBoxPos + 1
    If Button = 1 And X >= 372 And Y >= 101 And X <= 386 And Y <= 116 Then OkBoxPos = OkBoxPos - 1
    If Button = 2 And X >= 372 And Y >= 215 And X <= 386 And Y <= 230 Then OkBoxPos = OkBoxPos + 10
    If Button = 2 And X >= 372 And Y >= 101 And X <= 386 And Y <= 116 Then OkBoxPos = OkBoxPos - 10
    If OkBoxPos < 0 Then OkBoxPos = 0
    If OkBoxPos > 90 Then OkBoxPos = 90
    Exit Sub
End If


If ShowInventory = True And ShowOKBox = False Then
    For ty = 0 To 5
        For tx = 0 To 4
    
        If CurMouseX >= (tx * 36 + 460) And CurMouseY >= (ty * 36 + 61) And CurMouseX <= (tx * 36 + 494) And CurMouseY <= (ty * 36 + 95) And Button = 2 Then
            SendData "DRP" & ty * 5 + tx + 1 & "," & 1
            'SendData "DROP" & tY * 5 + tX + 1
            'SendData "USE" & frmMain.ObjLst.ListIndex + 1
            Exit Sub
        End If
        
        Next tx
    Next ty
    Exit Sub
End If

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tx, ty

If Button = vbLeftButton Then
    'SendData "LC" & tX & "," & tY
    'Send use command
    If CurSpellIndex > 0 Then
        If UserSpellbook(CurSpellIndex).Name <> "" Then
            SendData "CAST" & UserSpellbook(CurSpellIndex).SpellIndex & "," & tx & "," & ty
        End If
    End If
    TxtCatch.SetFocus
Else
    SendData "RC" & tx & "," & ty

End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Allow the MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub


Private Sub ForumClose_Click()

frmForum.Visible = False

End Sub

Private Sub FourmTitle_Click()

Call frmMain.wbForum.Navigate2("http://www.coolbm.com/aspbbs/", Null, frmMain.wbForum, Null, Null)

End Sub

Private Sub FPSTimer_Timer()

'Display and reset FPS
FramesPerSec = FramesPerSecCounter
FramesPerSecCounter = 0
lblFPS.Caption = "FPS: " + Str(FramesPerSec)

End Sub

Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub






Private Sub GetCmd_Click()

'Send the get command
SendData "GET"
TxtCatch.SetFocus

End Sub

Private Sub MidiPlayer_StatusUpdate()

'LSee if MIDI is done
'If MidiPlayer.Length = MidiPlayer.Position Then
        
    'Loop if needed
'    If LoopMidi Then
'        Call PlayMidi(CurMidi)
'    End If

'End If

End Sub



Private Sub GFXTimer_Timer()
'Since the engine will run at about 111 FPS
'on a P3 1.6ghz with 256 ram, this lets the
'client show the next frame 50 times/second
'limiting the FPS to 50.
OKToDraw = True
End Sub

Private Sub npcList_Click()
    TxtCatch.SetFocus
End Sub

Private Sub RecTxt_Change()

If frmMain.SendTxt.Visible = False Then TxtCatch.SetFocus
If frmMain.SendTxt.Visible = True Then SendTxt.SetFocus

End Sub

Private Sub RecTxt_GotFocus()

If frmMain.SendTxt.Visible = False Then TxtCatch.SetFocus
If frmMain.SendTxt.Visible = True Then SendTxt.SetFocus

End Sub

Private Sub SendTxt_Change()

stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

'BackSpace
If KeyAscii = 8 Then
    Exit Sub
End If

'Every other letter
If KeyAscii >= 32 And KeyAscii <= 126 Then
    Exit Sub
End If

KeyAscii = 0

End Sub


Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

Dim retcode As Integer
Dim i As Integer

'Send text
If KeyCode = vbKeyReturn Then
    For i = 1 To Len(stxtbuffer)
        If Mid(stxtbuffer, i, 1) = "~" Then
            Mid(stxtbuffer, i, 1) = "-"
        End If
    Next i
    'Command
    If UCase(stxtbuffer) = "/MIDIOFF" Then
        retcode = mciSendString("close all", 0, 0, 0)
        LoopMidi = 0
    ElseIf UCase(stxtbuffer) = "/STATS" Then
        frmMain.StatBox.Text = ""
        frmMain.StatBox.Visible = True
        SendData (stxtbuffer)
        
    ElseIf Left$(stxtbuffer, 1) = "/" Then
        SendData (stxtbuffer)

    'yell
    ElseIf Left$(stxtbuffer, 1) = "'" Then
        SendData ("'" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
    
    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        SendData ("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Whisper
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        SendData ("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Emote
    'ElseIf Left$(stxtbuffer, 1) = ":" Then
    '    SendData (":" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Say
    ElseIf stxtbuffer <> "" Then
        SendData (";" & stxtbuffer)

    End If

    stxtbuffer = ""
    SendTxt.Text = ""
    KeyCode = 0
    SendTxt.Visible = False
    Exit Sub

End If

End Sub

Private Sub Socket1_Connect()

Call Login
Call SetConnected


End Sub


Private Sub Socket1_Disconnect()

Connected = False

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
'*********************************************
'Handle socket errors
'*********************************************

Select Case (ErrorCode)

Case 24065
    frmMain.Hide
    DeInitTileEngine
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    End
    
Case 24061
    frmMain.Hide
    DeInitTileEngine
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    End
    
Case 24064
    frmMain.Hide
    DeInitTileEngine
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    End
    
Case Else
    frmMain.Hide
    DeInitTileEngine
    MsgBox (ErrorString)
    frmConnect.MousePointer = 1
    End

End Select


End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
'*********************************************
'Seperate lines by ENDC and send each to HandleData()
'*********************************************

Dim loopC As Integer

Dim RD As String
Dim rBuffer(1 To 500) As String
Static TempString As String

Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer

Socket1.Read RD, DataLength

'Check for previous broken data and add to current data
If TempString <> "" Then
    RD = TempString & RD
    TempString = ""
End If

'Check for more than one line
sChar = 1
For loopC = 1 To Len(RD)

    tChar = Mid$(RD, loopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = loopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = loopC + 1
    End If
    
Next loopC

'Check for broken line and save for next time
If Len(RD) - (sChar - 1) <> 0 Then
    TempString = Mid$(RD, sChar, Len(RD))
End If

'Send buffer to Handle data
For loopC = 1 To CR
    Call HandleData(rBuffer(loopC))
Next loopC

End Sub


Private Sub srvmsgexit_Click()
End Sub

Private Sub StatusBox_Change()
If frmMain.SendTxt.Visible = False Then TxtCatch.SetFocus
If frmMain.SendTxt.Visible = True Then SendTxt.SetFocus
End Sub


Private Sub StatusBox_GotFocus()
If frmMain.SendTxt.Visible = False Then TxtCatch.SetFocus
If frmMain.SendTxt.Visible = True Then SendTxt.SetFocus
End Sub

Private Sub SvrOK_Click()
    prgRun = False
End Sub

Private Sub TxtCatch_Change()
    'Call CheckKeys
    If Len(TxtCatch.Text) > 1 Then
        TxtCatch.Text = ""
    End If
End Sub
