VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Aspereta (Create Character)"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCity 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   2840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Make me!"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3480
      Width           =   3615
   End
   Begin VB.ListBox lstSex 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   2350
      Width           =   1575
   End
   Begin VB.ListBox lstPaths 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      ForeColor       =   &H00E0E0E0&
      Height          =   3550
      Left            =   1730
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()

If lstPaths.ListIndex < 0 Then
    Call MsgBox("You must choose a class before you create your character.", vbOKOnly)
    Exit Sub
End If
    
If lstSex.ListIndex < 0 Then
    Call MsgBox("You must be either male or female.", vbOKOnly)
    Exit Sub
End If

If lstCity.ListIndex < 0 Then
    Call MsgBox("You must choose a place to be born.", vbOKOnly)
    Exit Sub
End If

End Sub

Private Sub Form_Load()

Randomize Timer

lstPaths.Clear
lstPaths.AddItem "Warrior", 0
lstPaths.AddItem "Rogue", 1
lstPaths.AddItem "Monk", 2
lstPaths.AddItem "Cleric", 3
lstPaths.AddItem "Enchanter", 4
lstPaths.AddItem "Shaman", 5
lstPaths.AddItem "Wizard", 6

lstSex.Clear
lstSex.AddItem "Male", 0
lstSex.AddItem "Female", 1

lstCity.Clear
lstCity.AddItem "Minita", 0

lstPaths.ListIndex = Int(Rnd * 7)
lstSex.ListIndex = Int(Rnd * 2)
lstCity.ListIndex = 0

End Sub


Private Sub lstPaths_Click()

If lstPaths.ListIndex = 0 Then lblDesc.Caption = "Warrior -- They hit stuff!"
If lstPaths.ListIndex = 1 Then lblDesc.Caption = "Rogues -- Swift little pussies that can't take hits."
If lstPaths.ListIndex = 2 Then lblDesc.Caption = "Monks -- Like a monkey, but utilizes different spelling!"
If lstPaths.ListIndex = 3 Then lblDesc.Caption = "Cleric -- Me likes to heal!"
If lstPaths.ListIndex = 4 Then lblDesc.Caption = "Enchanter -- Obviously a melee tank....or not."
If lstPaths.ListIndex = 5 Then lblDesc.Caption = "Shaman -- Buff me please, buff me!"
If lstPaths.ListIndex = 6 Then lblDesc.Caption = "Wizard -- Uses magical attacks."

End Sub

