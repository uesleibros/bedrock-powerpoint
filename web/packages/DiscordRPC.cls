VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiscordRPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DiscordRPC - Experimental build
' Bedrock (c) 2024 // (UesleiDev)
' References:
'   - https://stackoverflow.com/questions/72187794/ipc-between-vba-and-python
'   - https://discord.com/developers/docs/topics/rpc
'   - https://github.com/discord/discord-rpc/blob/master/documentation/hard-mode.md
'   - https://robins.one/notes/discord-rpc-documentation.html

Option Explicit

Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const PIPE_ACCESS_INBOUND As Long = &H80000000
Private Const PIPE_ACCESS_OUTBOUND As Long = &H40000000
Private Const BUFFER_SIZE As Long = 1024
Private Const DISCORD_PIPE As String = "\\.\pipe\discord-ipc-"
Private WithEvents App As application
Attribute App.VB_VarHelpID = -1
Private hPipe As LongPtr
Private pid As Long
Private timestamp As Dictionary

Public Sub Connect(client_id As String, application As application)
    Dim i As Long
    
    Set App = application
    If hPipe > 0 Then Disconnect
    
    For i = 0 To 9
        hPipe = CreateFile(DISCORD_PIPE & i, PIPE_ACCESS_OUTBOUND Or PIPE_ACCESS_INBOUND, 0, 0, 3, 0, 0)
        If hPipe <> -1 Then Exit For
    Next i
    
    pid = GetCurrentProcessId

    If hPipe = -1 Then
        CloseHandle (hPipe)
        Exit Sub
    End If
    
    Set timestamp = New Dictionary
    timestamp.Add "start", 0

    DoHandshake (client_id)
    AuthorizeClient (client_id)
End Sub

Public Sub Update(activity As String, Optional preserve_time As Boolean = True, Optional state As String = "", Optional details As String = "", _
Optional party_size As Integer = -1, Optional party_max As Integer = -1)
    Dim presence_data As Dictionary
    Dim cmd_former As New Dictionary, args_cmd_former As New Dictionary
    Set presence_data = Parse(ReadTextFile(ActivePresentation.path & "/rpc/activities/" & activity & ".json"))
    
    If Len(state) > 0 Then presence_data("state") = state
    If Len(details) > 0 Then presence_data("details") = details
    
    If timestamp("start") = 0 And preserve_time Then
        timestamp("start") = GetCurrentUnixTimestampUTC
    End If

    If preserve_time Then
        presence_data.Add "timestamps", timestamp
    Else
        timestamp("start") = 0
    End If

    If presence_data.Exists("party") Then
        If party_size > -1 Then presence_data("party")("size")(0) = party_size
        If party_max > -1 Then presence_data("party")("size")(0) = party_max
    End If

    args_cmd_former.Add "pid", pid
    args_cmd_former.Add "activity", presence_data
    
    cmd_former.Add "cmd", "SET_ACTIVITY"
    cmd_former.Add "nonce", GenerateNonce
    cmd_former.Add "args", args_cmd_former

    If Not SendPacket(1, Stringify(cmd_former)) Then Exit Sub
End Sub

Private Function GetCurrentUnixTimestampUTC() As Long
    Dim currentTime As Date
    Dim utcTime As Date

    currentTime = Now
    utcTime = DateAdd("h", 3, currentTime)

    GetCurrentUnixTimestampUTC = DateDiff("s", "1/1/1970 00:00:00", utcTime)
End Function

Public Sub Disconnect()
    ClearActivity
    CloseHandle (hPipe)
End Sub

Private Sub DoHandshake(client_id As String)
    Dim cmd_former As New Dictionary
    
    cmd_former.Add "v", 1
    cmd_former.Add "client_id", client_id
    
    If Not SendPacket(0, Stringify(cmd_former)) Then Exit Sub
End Sub

Private Sub ClearActivity()
    Dim cmd_former As New Dictionary, args_cmd_former As New Dictionary
    
    args_cmd_former.Add "pid", pid
    
    cmd_former.Add "cmd", "SET_ACTIVITY"
    cmd_former.Add "nonce", GenerateNonce
    cmd_former.Add "args", args_cmd_former

    If Not SendPacket(1, Stringify(cmd_former)) Then Exit Sub
End Sub

Private Sub AuthorizeClient(client_id As String)
    Dim cmd_former As New Dictionary, args_cmd_former As New Dictionary
    Dim args_scopes As New List
    
    args_scopes.Add "rpc"
    
    args_cmd_former.Add "client_id", client_id
    args_cmd_former.Add "scopes", args_scopes
    
    cmd_former.Add "nonce", GenerateNonce
    cmd_former.Add "args", args_cmd_former
    cmd_former.Add "cmd", "AUTHORIZE"
    
    If Not SendPacket(1, Stringify(cmd_former)) Then Exit Sub
End Sub

' Only for debug.
Private Sub ReadResponse()
    Dim result As Long
    Dim buffer(BUFFER_SIZE) As Byte
    Dim bytesRead As Long

    result = ReadFile(hPipe, buffer(0), UBound(buffer) + 1, bytesRead, 0)

    If bytesRead > 0 Then
        Dim jsonResponse As String
        Dim jsonStart As Long
        Dim jsonLength As Long

        jsonStart = 8
        jsonLength = bytesRead - jsonStart

        jsonResponse = StrConv(MidB(buffer, jsonStart + 1, jsonLength), vbUnicode)
        Debug.Print "Resposta do Discord: " & jsonResponse
    End If
End Sub

Private Function SendPacket(opcode As Long, jsonData As String) As Boolean
    Dim packet() As Byte
    Dim length As Long
    Dim bytesWritten As Long
    Dim result As Long
    
    jsonData = StrConv(jsonData, vbFromUnicode)
    length = LenB(jsonData)
    
    ReDim packet(7 + length)

    PutLongToByteArray packet, 0, opcode
    PutLongToByteArray packet, 4, length

    Dim i As Long
    For i = 0 To length - 1
        packet(8 + i) = AscB(MidB(jsonData, i + 1, 1))
    Next i
    
    result = WriteFile(hPipe, packet(0), UBound(packet) + 1, bytesWritten, 0)
    SendPacket = (result <> 0)
End Function

Private Sub PutLongToByteArray(Arr() As Byte, Pos As Long, Value As Long)
    Arr(Pos) = Value And &HFF
    Arr(Pos + 1) = (Value \ &H100) And &HFF
    Arr(Pos + 2) = (Value \ &H10000) And &HFF
    Arr(Pos + 3) = (Value \ &H1000000) And &HFF
End Sub

Private Function GenerateNonce() As String
    Dim guid As String
    Dim i As Integer
    
    Randomize
    
    For i = 1 To 32
        guid = guid & Hex(Int(Rnd() * 16))
    Next i
    
    GenerateNonce = Mid(guid, 1, 8) & "-" & _
                    Mid(guid, 9, 4) & "-" & _
                    "4" & Mid(guid, 14, 3) & "-" & _
                    Mid(guid, 17, 4) & "-" & _
                    Mid(guid, 21, 12)
End Function

Private Function ReadTextFile(path As String, Optional Charset As String = "utf-8") As String
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .LoadFromFile path
        .Type = 2
        .Charset = Charset
        ReadTextFile = .ReadText(-1)
        .Close
    End With
End Function

Private Sub App_PresentationClose(ByVal Pres As Presentation)
    Disconnect
End Sub
