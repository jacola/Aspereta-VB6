Attribute VB_Name = "General"
Option Explicit

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

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double
'*****************************************************************
'Finds the distance between two points
'*****************************************************************

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*****************************************************************
'Find a Random number between a range
'*****************************************************************

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

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
Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
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


Sub RefreshUserListBox()
'*****************************************************************
'Refreshes the User list box on the frmMain
'*****************************************************************
Dim LoopC As Integer
  
If LastUser < 0 Then
    frmMain.Userslst.Clear
    Exit Sub
End If
  
frmMain.Userslst.Clear
For LoopC = 1 To LastUser
    If UserList(LoopC).Name <> "" Then
        frmMain.Userslst.AddItem UserList(LoopC).Name
    End If
Next LoopC

End Sub

Sub Restart()
'*****************************************************************
'Restarts the server
'*****************************************************************

LogData = "Reset started at " & Time$ & "." & FONTTYPE_SHOUT
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))

'ensure that the sockets are closed, ignore any errors
On Error Resume Next

Dim LoopC As Integer

'*** Clear vars ***

frmMain.Socket1.Cleanup
frmMain.Socket1.Startup
  
frmMain.Socket2(0).Cleanup
frmMain.Socket2(0).Startup

'Clear users
For LoopC = 1 To MaxUsers
    CloseSocket (LoopC)
Next

'Reset User connections
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next LoopC

'Clear NPCs
For LoopC = 1 To LastNPC
    If NPCList(LoopC).Flags.NPCActive Then
        If NPCList(LoopC).Flags.NPCAlive Then
            KillNPC LoopC
        End If
        CloseNPC LoopC
    End If
Next LoopC

'Clear char list
For LoopC = 1 To MAX_CHARACTERS
    CharList(LoopC) = 0
Next LoopC

'Init vars
LastUser = 0
NumUsers = 0
LastChar = 0
LastNPC = 0

'*** Load data ***

Call LoadSini
Call LoadMapData
Call LoadOBJData

'*** Setup sockets ***

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
Print #5, "**** Server restarted. " & Time & " " & Date
Close #5

LogData = "Reset complete!  (" & Time$ & ")" & FONTTYPE_TALK
AddtoRichTextBox frmMain.ServerLog, ReadField(1, LogData, 126), Val(ReadField(2, LogData, 126)), Val(ReadField(3, LogData, 126)), Val(ReadField(4, LogData, 126)), Val(ReadField(5, LogData, 126)), Val(ReadField(6, LogData, 126))


End Sub

