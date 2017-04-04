VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Server"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   7185
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSound 
      BackColor       =   &H00000000&
      Caption         =   "Sound Driver"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   240
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkFullScrn 
      BackColor       =   &H00000000&
      Caption         =   "Full Screen (800x600)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   2055
      Left            =   5520
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox NameTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         Height          =   465
         Left            =   840
         MaxLength       =   12
         TabIndex        =   10
         Top             =   240
         Width           =   2595
      End
      Begin VB.TextBox PasswordTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   840
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox SavePassChk 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Password"
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
         Height          =   345
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   1590
      End
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   6600
      Width           =   2235
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   465
      Left            =   2400
      TabIndex        =   1
      Text            =   "7777"
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   3870
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblSndWarn 
      BackColor       =   &H00000000&
      Caption         =   "If you do not initialize sound you MUST have your sound shut off or you will crash."
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   7320
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image ImgNew 
      Height          =   855
      Left            =   6600
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Image imgContinue 
      Height          =   855
      Left            =   6480
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Image ImgExit 
      Height          =   855
      Left            =   6480
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP &&"
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
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
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
      Height          =   345
      Left            =   2400
      TabIndex        =   2
      Top             =   6240
      Width           =   465
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSound_Click()
    If chkSound.value = Unchecked Then
        lblSndWarn.Visible = True
    Else
        lblSndWarn.Visible = False
    End If
End Sub



Private Sub Command1_Click()
'*****************************************************************
'Makes sure user data is ok then trys to connect to server
'*****************************************************************

Dim loopC As Integer

On Error Resume Next

If frmConnect.MousePointer = 11 Then
    Exit Sub
End If

'update user info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

If CheckUserData = True Then
       
    'FrmMain.Socket1.Close
    frmMain.Socket1.HostName = UserServerIP
    frmMain.Socket1.RemotePort = UserPort

    SendNewChar = False
    frmConnect.MousePointer = 11
    frmMain.Socket1.Connect
    
    Call SaveGameini

    Call Main
    
End If

End Sub

Private Sub Command2_Click()
'*****************************************************************
'Makes sure user data is ok then begins new character process
'*****************************************************************

StringChecker (NameTxt.Text)

If Len(NameTxt.Text) < 2 Then Exit Sub


'update user info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

UserBody = 1
UserHead = 1

If CheckUserData = True Then
    frmMain.Visible = False
    'FrmMain.Socket1.Close
    frmMain.Socket1.HostName = UserServerIP
    frmMain.Socket1.RemotePort = UserPort

    SendNewChar = True
    frmConnect.MousePointer = 11
    frmMain.BorderStyle = 0
    frmMain.Socket1.Connect
    Call SaveGameini
    Frame1.Visible = False
    Command2.Visible = False
    imgContinue.Visible = False
    ImgNew.Visible = False
    Call Main

End If


End Sub


Private Sub Command3_Click()

'update info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

Call SaveGameini

frmConnect.MousePointer = 1
frmMain.MousePointer = 1

'End program
prgRun = False
End

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    Label4.Visible = True
    
    'Server IP
    IPTxt.Text = "localhost"
    IPTxt.Visible = True
    Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()

ClientVer = "APLHA080"
Dim temp As String
Dim loopC As Integer

For loopC = 1 To 5
    UserHotButtons(loopC) = 0
Next loopC

PaperDollList(1) = "Head"
PaperDollList(2) = "Chest"
PaperDollList(3) = "Legs"
PaperDollList(4) = "Feet"
PaperDollList(5) = "Arms"
PaperDollList(6) = "Back"
PaperDollList(7) = "Left Hand"
PaperDollList(8) = "Right Hand"


'Get Game.ini Data
If FileExist(IniPath & "Game.ini", vbNormal) = True Then
    NameTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Name")
    PasswordTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Password")
    PortTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Port")
    IPTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "IP")
    If IPTxt.Text = "" Then IPTxt.Text = "localhost"
End If

If FileExist(IniPath & "\user\" & NameTxt.Text + ".ini", vbNormal) = True Then
    Open IniPath & "\user\" & NameTxt.Text + ".ini" For Input As #1
        For loopC = 1 To 10
            Input #1, UserHotButtons(loopC)
        Next loopC
    Close #1
End If

For loopC = 1 To 20
    ChatText(loopC) = " "
Next loopC

If LCase(GetVar(IniPath & "Game.ini", "SETTINGS", "TextBoxAlwaysOn")) = "true" Then
    TextBoxAlwaysOn = True
    TextBoxOn = True
Else
    TextBoxAlwaysOn = False
    TextBoxOn = False
End If

If LCase(GetVar(IniPath & "Game.ini", "SETTINGS", "StatusFilter")) = "true" Then
    StatusFilter = True
Else
    StatusFilter = False
End If
If LCase(GetVar(IniPath & "Game.ini", "SETTINGS", "StatusBox")) = "true" Then
    frmMain.StatusBox.Visible = True
Else
    frmMain.StatusBox.Visible = False
End If

ShowSpells = False

End Sub


Private Sub imgContinue_Click()
    Frame1.Visible = True
    Command2.Visible = False
    imgContinue.Visible = False
    ImgNew.Visible = False
End Sub

Private Sub ImgExit_Click()


'update info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

Call SaveGameini

frmConnect.MousePointer = 1
frmMain.MousePointer = 1

'End program
prgRun = False
End


End Sub

Private Sub ImgNew_Click()
    NameTxt.Text = ""
    PasswordTxt.Text = ""
    Frame1.Visible = True
    Command1.Visible = False
    imgContinue.Visible = False
    ImgNew.Visible = False
End Sub


Private Sub NameTxt_Change()
Dim CurChar As String

If Len(NameTxt.Text) >= 1 Then
    'CurChar = Right(NameTxt.Text, Len(NameTxt.Text) - (Len(NameTxt.Text) - 1))
    'If CurChar < "A" Or CurChar > "z" Then
    'NameTxt.Text = Left(NameTxt.Text, Len(NameTxt.Text) - 1)
    'End If
    StringChecker (NameTxt.Text)
End If

End Sub
