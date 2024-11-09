Attribute VB_Name = "KailleraMessages"
'---------------------------------------------------------------------------------
'Protocol
'---------------------------------------------------------------------------------
'//Logging in State
'Client: HELLO0.83
'Server: Port notify HELLOD00D#\0 [# = new port number. eg. HELLOD00D7159]
'        or TOO\0 [if server is full]
'Client: User Login Information [0x03]
'Server: Server to Client ACK [0x05]
'Client: Client to Server ACK [0x06]
'Server: Server to Client ACK [0x05]
'Client: Client to Server ACK [0x06]
'Server: Server to Client ACK [0x05]
'Client: Client to Server ACK [0x06]
'Server: Server to Client ACK [0x05]
'Client: Client to Server ACK [0x06]
'Server: Server Status [0x04]
'Server: User Joined [0x02]
'Server: Server Information Message [0x17]
'
'//Global Chat State
'Client: Global Chat Request [0x07]
'Server: Global Chat Notification [0x07]
'
'//Game Chat State
'Client: Game Chat Request [0x08]
'Server: Global Chat Notification [0x08]
'
'//Create Game State
'Client: Create Game Request [0x0A]
'Server: Create Game Notification [0x0A]
'Server: Update Game Status [0x0A]----------->why?
'Server: Player Information [0x0D]----------->why?
'Server: Join Game Notification [0x0C]
'Server: Server Information Message [0x17]
'
'//Join Game State
'Client: Join Game Request [0x0C]
'Server: Update Game Status [0x0E]
'Server: Player Information [0x0D]
'Server: Join Game Notification [0x0C]
'
'//Quit Game State
'Client: Quit Game Request [0x0B]
'Server: Update Game Status [0x0E]
'Server: Quit Game Notification [0x0B]
'
'//Close Game State
'Client: Quit Game Request [0x0B]
'Server: Close Game Notification [0x10]
'Server: Quit Game Notification [0x0B]
'
'//Start Game State
'Client: Start Game Request [0x11]
'Server: Update Game Status [0x0E]
'Server: Start Game Notification [0x11]
'Client: *Netsync Mode* Wait for all players to give: Ready to Play Signal [0x15]
'Server: Update Game Status [0x0E]
'Server: *Playing Mode* All players are ready to start: Ready to Play Signal Notification [0x15]
'Client(s): Exchange Data. Game Data Send [0x12] or Game Cache Send [0x13]
'Server: Sends data accordingly. Game Data Notify [0x12] or Game Cache Notify [0x13]
'Note: Client is ALWAYS sending data. If no input is happening, it sends 0's. Try for 60 fps!
'
'//Drop Game State
'Client: Drop Game Request [0x14]
'Server: Update Game Status [0x0E]
'Server: Drop Game Notification [0x14]
'
'//Kick Player State
'Client: Kick Request [0x0F]
'Server: Quit Game Notification [0x0B]
'Server: Update Game Status [0x0E]
'
'//User Quit State
'Client: User Quit Request [0x01]
'Server: User Quit Notification [0x01]
'
'---------------------------------------------------------------------------------


'---------------------------------------------------------------------------------
'Server List - http://www.kaillera.com/raw_server_list2.php
'              http://master.anti3d.com/raw_server_list2.php
'---------------------------------------------------------------------------------
'Note: LF = LineFeed = 10
'
'Format: serverName[LF]ipAddress:port;users/maxusers;gameCount;version;location[LF]
'
'Example:
'
'Unknown Server1
'111.111.111.1111:27888;0/25;0.86;USA
'Unknown Server2
'222.222.222.2222:27888;0/50;0.86;Canada
'---------------------------------------------------------------------------------


'---------------------------------------------------------------------------------
'Packet format: multi bytes in little endian format [1st_Byte * 256 + 2nd_Byte]
'---------------------------------------------------------------------------------
'//Inital byte
'1B - byte - Number of messages in packet [generally you will always have n-3 messages. _
'                                         During gameplay, it may be neccessary to increase.]
'
'//5 byte header before each message
'2B - word - The number of this message; every message sent is increased by 1.
'2B - word - Length of this message; includes messsage type and size of data being sent
'1B - byte - Message type ex). 0x03 User Login
'
'//Message
'Data
'
'[Repeats 5 byte header and Data for each message]
'
'eg). (1B), 2B,2B,1B,DataA [Repeats} 2B,2B,1B,DataB [Repeats] 2B,2B,1B,DataC and etc.
'eg). User Quit Request [0x01]:  (01),    01,00,    04,00,  01,   00,00,FF
'                                initial  message#  Length  Type  Data
'                                ^
'                                Note: We are only sending one message! If it were _
'                                      2 messages this would be 02.
'The packet to send would look like this: 01,01,00,04,00,01,00,00,FF
'---------------------------------------------------------------------------------


'---------------------------------------------------------------------------------
'User Info
'---------------------------------------------------------------------------------
':USERINFO=userID1,userIP1,userAccessLevel1;userID2,userIP2,userAccessLevel2;
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
'Message Types [all multi-bytes are in little endian format]
'---------------------------------------------------------------------------------
'     0x01 = User Quit Request(00,FF,FF,Message\0)/Notification
'            nick\0
'            2B : userId
'            message\0
'
'     0x02 = User joined
'            nick\0
'            2B : userId
'            4B : ping
'            1B : connection (6=Bad...1=LAN)
'
'     0x03 = User Login Information
'            nick\0
'            emulator\0
'            1B : connection (6=Bad...1=LAN)
'
'     0x04 = Server status
'            1B : 00
'            4B : Num users in server (not including you)
'            4B : Num games in server
'            NB : List of users
'                   nick\0
'                   4B : ping
'                   1B : connection (3 = Good)
'                   2B : userId
'                   1B : status (0=Playing, 1=Idle)
'            NB : List of games
'                   game\0
'                   4B : gameid
'                   emulator\0
'                   owner\0
'                   number of players/maximum players\0
'                   1B:status (0=Waiting, 1=Playing, 2=Netsync)
'
'     0x05 = Server to client ACK
'            1B : 00
'            4B : 00
'            4B : 01
'            4B : 02
'            4B : 03
'
'     0x06 = Client to server ACK
'            1B : 00
'            4B : 00
'            4B : 01
'            4B : 02
'            4B : 03
'
'     0x07 = Global Chat Request(00,message\0)/Notification
'            nick\0
'            message\0
'
'     0x08 = Game Chat Request(00,message\0)/Notification
'            nick\0
'            message\0
'
'     0x09 = Client Keep Alive
'            1B : 00
'
'     0x0A = Create Game Request(00,game\0,00,FF,FF,FF,FF)/Notification
'            nick\0
'            game\0
'            emulator\0
'            2B : gameId
'            1B : status (0=Idle, 1=Playing) [since it was just created, it's initally 0]
'
'     0x0B = Quit Game Request(00,FF,FF)/Notification
'            nick\0
'            2B : userId
'
'     0x0C = Join Game Request(00,gameId,00,00,00,00,00,00,00,FF,FF,connection)/Notification
'            1B : 00
'            2B : gameId
'            1B : 00
'            1B : 00
'            nick\0
'            4B : ping
'            2B : userId
'            1B : connection (6=Bad...1=LAN)
'
'
'     0x0D = Player Information
'            1B : 00
'            4B : number of users
'            username\0
'            4B : ping
'            2B : userId
'            1B : connection (6=Bad...1=LAN)
'
'     0x0E = Update Game Status
'            1B : 00
'            2B : gameId
'            1B : 00
'            1B : 00
'            1B : status (0=Waiting, 1=Playing, 2=Netsync)
'            1B : number of players
'            1B : maximum players
'
'     0x0F = Kick request
'            1B : 00
'            2B : userId
'
'     0x10 = Close game
'            1B : 00
'            2B : gameId
'            2B : 00
'
'     0x11 = Start Game Request(00,FF,FF,FF,FF)/Notification
'            1B : 00
'            2B : multiplier (eg. (CT * (multiplier + 1) <-Block on that frame
'            1B : your player number (eg. if you're player 1 or 2...)
'            1B : total players
'
'     0x12 = Game Data Send(00, 2B, NB)/Notification
'            1B : 00
'            2B : length of data
'            NB : data (depends on connection type: 3 = good = 3 sets/packet)
'            eg). MAME32K 0.64 = 2 bytes/set
'            Note: Game data is ALWAYS being sent and received. If there is no_
'                  input happening, 0's are being sent. Shoot for 60fps._
'                  60 fps / 3 connection type = 20 messages/second
'
'     0x13 = Game Cache Send(00, Cache Position)/Notification
'            1B : 00
'            1B : Cache Position
'            *256 Slots [0 to 255]. Oldest to Newest. When cache is full add new _
'            entry at position 255 and shift all the old entries down knocking off the oldest.
'            Search cache for matching data before you send. If found, send that cache position, _
'            otherwise issue a Game Data Send [0x12]. When server sends a game Data Notify to you, _
'            search for matching cache data, if not found add it to a new position*
'
'     0x14 = Drop Game Request(00,00)/Notification
'            nick\0
'            1B : player number (which player number dropped)
'
'     0x15 = Ready to Play Signal Request(00)/Notification
'            1B : 00
'
'     0x16 = Connection Rejected Notification
'            nick\0
'            2B : userId
'            message\0
'
'     0x17 = Server Information Message
'            server\0
'            message\0
'----------------------------------------------------------------------------------------------

Option Explicit

'Other
Public portNum As Long 'New Port Number from server HELLOD00D
Public msgCount As Long 'Total Message Count for Out Bound Data
Public serversLastMessage As Long 'Last Message we Received from server
Public inServer As Boolean 'Logged in Server Flag
Public step As Byte 'Step1: HELLOD00D Step2: All other messages
Public stopList As Boolean 'Stop Server List or Pinging
Public sortUserlist As Byte 'remember userlist sort column
Public sortPos As Byte 'remember userlist sort pattern
Public lastUserid As String 'last userid to join server
Public iQuit As Boolean
Public adminFeatures As Boolean
Public serverAlarm As Boolean
Public uClient As Long
Public timeoutCount As Long
Public rSwitch As Boolean
Public firstID As Long
Public allowUnload As Boolean
Public dbEdit As Boolean
Public wasAdmin As Boolean

'My Info
Public myUserId As Long
Public myGameId As Long
Public myGame As String
Public myConnectionType As Byte
Public myPing As Long

'Gameroom
Public imOwner As Boolean 'If you're the owner of the room
Public leftRoom As Boolean 'Flag to recognize if you were kicked from a room
Public inRoom As Boolean

'Server Info
Public serverIP As String
Public serverName As String

'Time
Public tCount As Byte
Public finishedTime As Byte
Public minutes As Byte
Public hours As Byte
Public days As Long

'Constants
Public Const entryMsg As String = "HELLO0.83" & vbNullChar
Public Const pingMsg As String = "PING" & vbNullChar
Public Const clientVersion As String = "5.32"
Public Const emulatorPass As String = "EmulinkerSF Admin Client v" & clientVersion
'Public Const emulatorPass As String = "Emulinker Suprclient v" & clientVersion


'Outbound Data
Public packet(0 To 1999) As Byte
Public myPackets(0 To 4) As Packets

Public Type Packets
    packet(0 To 1999) As Byte
End Type

'Inbound Data
Public myBuff() As Byte
Public serverMessages(0 To 13) As messages
Public Type messages
    msgData(0 To 1999) As Byte
    msgType As Byte
    msgLen As Long
    msgNum As Long
End Type

'Bot
Public myBot As AdminBot
Public Type AdminBot
    botStatus As Boolean
    'announcements
    announceChatroomCount As Long
    announceChatroomCount2 As Long
    announceChatroomCount3 As Long
    announceGamesCount As Long
    'login spam control
    loginTimeout As Long
    loginCount As Long
    loginIP As String
End Type

Public arUsers(0 To 299) As Users
Public Type Users
    name As String
    access As String
    ip As String
    ping As Long
    emulator As String
    connection As Byte
    userID As Long
    status As Byte
    loggedIn As Boolean
    dbPos As Long
    head As Long
    sect1 As Long
    
    linkSent As Long
    linkCount As Long
    spamRowCount As Long
    spamTimeout As Long
    gameTimeout As Long
    gameSpamCount As Long
    game As String
    gameID As Long
End Type

Public arGames(0 To 199) As Games
Public Type Games
    owner As String
    ownerAccess As String
    emulator As String
    game As String
    status As Byte
    Users As String
    gameID As Long
    opened As Boolean
End Type


Public Sub parseData()
    Dim i As Long
    Dim w As Long
    Dim temp As Long
    Dim numOfMessages As Byte
    Dim msgLen As Long
    Dim msgType As Byte
    Dim msgData() As Byte
    Dim sizeOfPacket As Long
    Dim msgNumCount As Long
    Dim msgNum As Long

    'On Error Resume Next
    
    If myBuff(5) < &H1 Or myBuff(5) > &H17 Then
        frmServerlist.List1.AddItem Time & "Received Unknown Message: " & myBuff(5)
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        Exit Sub
    End If
    
    'how many messages are in packet
    numOfMessages = myBuff(0)
    msgNum = myBuff(2) * CLng(256) + myBuff(1)
    msgLen = myBuff(4) * CLng(256) + myBuff(3)
    msgType = myBuff(5)
    sizeOfPacket = msgLen + 3
    
    If msgNum > serversLastMessage Then
        'increases speed to not process every message
        msgNumCount = msgNum - serversLastMessage
        If msgNumCount < 2 Then
            numOfMessages = 1
        End If
    ElseIf msgNum <= serversLastMessage Then
        If serversLastMessage >= 65515 Then
            msgNumCount = msgNum - (65536 - serversLastMessage)
        Else
            Exit Sub
        End If
    End If

    temp = 1
    For i = 0 To numOfMessages - 1
        serverMessages(i).msgNum = (myBuff(temp + 1) * CLng(256)) + myBuff(temp)
        temp = temp + 2
        serverMessages(i).msgLen = ((myBuff(temp + 1) * CLng(256)) + myBuff(temp)) - 1
        temp = temp + 2
        serverMessages(i).msgType = myBuff(temp)

        temp = temp + 1
        Call CopyMemory(serverMessages(i).msgData(0), myBuff(temp), ByVal serverMessages(i).msgLen)
        temp = temp + serverMessages(i).msgLen
    Next i


    If msgNumCount < 2 Then
        serversLastMessage = msgNum
        Call gotoMessageType(0, serverMessages(0).msgType)
        Exit Sub
    Else
        If msgNumCount > numOfMessages Then
            Form1.txtChatroom.SelColor = &H8000&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "*Client Alert: Dropped " & (msgNum - serversLastMessage) - numOfMessages & " Packet(s)!*" & vbCrLf
                    
            frmServerlist.List1.AddItem "*Dropped " & (msgNum - serversLastMessage) - numOfMessages & " Packet(s)*"
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
            Call MDIForm1.mnuReconnectToServer_Click
            Exit Sub
        End If
        
        serversLastMessage = msgNum
        For i = msgNumCount - 1 To 0 Step -1
            Call gotoMessageType(CByte(i), serverMessages(i).msgType)
        Next i
    End If
End Sub


Sub gotoMessageType(slot As Byte, msgType As Byte)
    '0x01 User Quit Notification
    If msgType = &H1 Then
        Call userQuitNotification(slot)
    '0x02 - User Joined
    ElseIf msgType = &H2 Then
        Call userJoined(slot)
    '0x04 - Server Status
    ElseIf msgType = &H4 Then
        Call serverStatus(slot)
    '0x05 - Server to Client ACK
    ElseIf myBuff(5) = &H5 And Trim$(frmServerlist.txtFakePing.Text) = vbNullString Then
        Call clientToServerAck
    '0x07 - Global Chat Notification
    ElseIf msgType = &H7 Then
        Call globalChatNotification(slot)
    '0x08 - Game Chat Notification
    ElseIf msgType = &H8 Then
        Call gameChatNotification(slot)
    '0x0A - Create Game Notification
    ElseIf msgType = &HA Then
        Call createGameNotification(slot)
    '0x0B - Quit Game Notification
    ElseIf msgType = &HB Then
        Call quitGameNotification(slot)
    '0x0C - Join Game Notification
    ElseIf msgType = &HC Then
        Call joinGameNotification(slot)
    '0x0D - Player Information
    ElseIf msgType = &HD Then
        If inRoom = False Then Call playerInformation(slot)
    '0x0E - Update Game Status
    ElseIf msgType = &HE Then
        Call updateGameStatus(slot)
    '0x10 - Close Game Notification
    ElseIf msgType = &H10 Then
        Call closeGameNotification(slot)
    '0x11 - Start Game Notification
    ElseIf msgType = &H11 Then
        Call startGameNotification(slot)
    '0x12 Game Data
    ElseIf msgType = &H12 Then
        Call gameDataNotification(slot)
    '0x13 Game Cache
    ElseIf msgType = &H13 Then
        Call gameCacheNotification(slot)
    '0x14 - Drop Game Notification
    ElseIf msgType = &H14 Then
        Call dropGameNotification(slot)
    '0x15 Ready to Play Notify
    ElseIf msgType = &H15 Then

    '0x16 - Connection Rejected Notification
    ElseIf msgType = &H16 Then
        Call connectionRejectedNotification(slot)
    '0x17 - Server Information Message
    ElseIf msgType = &H17 Then
        Call serverInformationMessage(slot)
    End If
End Sub


'0x13
Public Sub gameCacheNotification(msgSlot As Byte)
    Form1.txtGameChatroom.SelColor = vbRed
    Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "All players are ready!" & vbCrLf
End Sub


'0x12
Sub gameDataNotification(msgSlot As Byte)

End Sub

'0x11
Public Sub startGameNotification(msgSlot As Byte)
    Dim i As Long
    
    'message(0)=0
    'myMultiplier = serverMessages(msgSlot).msgData(2) * CLng(256) + serverMessages(msgSlot).msgData(1)
    'myPlayerNumber = serverMessages(msgSlot).msgData(3)
    'totalPlayers = serverMessages(msgSlot).msgData(4)
    
    'display
    Form1.txtGameChatroom.SelColor = &H800000
    Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:"
    If frmPreferences.chkTimeStamps.Value = vbChecked Then
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & Time & ": "
    End If
    Form1.txtGameChatroom.SelColor = &HC000C0
    Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "The host has started game!" & vbCrLf
End Sub


'0x14
Sub dropGameRequest()
    packet(0) = &H0
    packet(1) = &H0
                    
    Call constructPacket(2, &H14)
End Sub

'0x14
Sub dropGameNotification(msgSlot As Byte)
    Dim i As Long
    Dim playerNum As Byte
    Dim nick As String
    Dim temp(0 To 511) As Byte
    Dim strSize As Long

    'nick
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'for playernumber
    playerNum = serverMessages(msgSlot).msgData(i)
        
    Close #5
    Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
        Print #5, "-:" & Time & "<server> Player#: " & playerNum & ": " & nick & " dropped from the game!"
    Close #5
    
    'display
    Form1.txtGameChatroom.SelColor = &H800000
    Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:"
    If frmPreferences.chkTimeStamps.Value = vbChecked Then
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & Time & ": "
    End If
    Form1.txtGameChatroom.SelColor = &HC000C0
    Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "Player#: " & playerNum & ": " & nick & " dropped from the game!" & vbCrLf
    Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
End Sub

'0x06
Sub clientToServerAck()
    'ACK
    packet(0) = &H0
    
    packet(1) = &H0
    packet(2) = &H0
    packet(3) = &H0
    packet(4) = &H0
    
    packet(5) = &H1
    packet(6) = &H0
    packet(7) = &H0
    packet(8) = &H0
    
    packet(9) = &H2
    packet(10) = &H0
    packet(11) = &H0
    packet(12) = &H0
    
    packet(13) = &H3
    packet(14) = &H0
    packet(15) = &H0
    packet(16) = &H0
    'Call Sleep(680)
    Call constructPacket(17, &H6)
End Sub

'0x16
Sub connectionRejectedNotification(msgSlot As Byte)
    Dim i As Long
    Dim nick As String
    Dim userID As Long
    Dim msg As String
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    'nick
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'userId
    userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
    i = i + 2
    
    'message
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    msg = ByteArrayToString(temp)
    i = i + strSize + 1
    
    Form1.Timer2.Enabled = False
    frmServerlist.List1.AddItem "*" & msg & "*"
    frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
End Sub


'0x01
Sub userQuitRequest(msg As String)
    Dim temp() As Byte
    Dim dataToBeSent As String
    
    dataToBeSent = vbNullChar & "dd" & Trim$(msg) & vbNullChar
    temp = StrConv(dataToBeSent, vbFromUnicode)
                    
    temp(1) = &HFF
    temp(2) = &HFF
                    
    Call CopyMemory(packet(0), temp(0), ByVal UBound(temp) + 1)
                  
    Call constructPacket(UBound(temp) + 1, &H1)
End Sub

'0x01
Sub userQuitNotification(msgSlot As Byte)
    Dim i As Long
    Dim w As ListItem
    Dim b As Long
    Dim j As Long
    Dim nick As String
    Dim userID As Long
    Dim quit As String
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    On Error Resume Next
    
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    'nick
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'userid
    userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
    i = i + 2
    
    'quit message
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    quit = ByteArrayToString(temp)
        
    Close #3
    Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
        Print #3, "-:" & Time & ": " & nick & " [Quit]"
    Close #3
    
    If userID = myUserId And userID <> myUserId Then
        If StrComp(LCase$(quit), "ping timeout", vbBinaryCompare) = 0 Then
            iQuit = True
            'initial
            frmServerlist.List1.AddItem "*Ping Timeout (" & Time & ")*"
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
                    For b = LBound(arUsers) To UBound(arUsers)
                        arUsers(b).loggedIn = False
                    Next b
                    For b = LBound(arGames) To UBound(arGames)
                        arGames(b).opened = False
                    Next b
            Call reconnectToServer
        End If
    Else
        If frmPreferences.chkShowJoin.Value = vbChecked Then
            'display
            Form1.txtChatroom.SelColor = &H800000
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
            If frmPreferences.chkTimeStamps.Value = vbChecked Then
                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
            End If
            Form1.txtChatroom.SelColor = &H808080        '&HFF&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & nick
            Form1.txtChatroom.SelColor = &H800000
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & " [Quit] "
              
            Form1.txtChatroom.SelColor = &H808080
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & quit & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        End If
        
        'remove from userlist
        Dim here As Boolean
        For i = 1 To frmUserlist.lstUserlist.ListItems.count
            If StrComp(frmUserlist.lstUserlist.ListItems(i).SubItems(3), userID, vbBinaryCompare) = 0 Then
                 frmUserlist.lstUserlist.ListItems.remove (i)
                 Exit For
            ElseIf i = frmUserlist.lstUserlist.ListItems.count And here = False Then
                frmUserlist.lstUserlist.Refresh
                i = 0
                here = True
            End If
        Next i

        Dim k As Long
        For k = LBound(arUsers) To UBound(arUsers)
            If arUsers(k).loggedIn = True Then
                If arUsers(k).userID = userID Then
                    arUsers(k).linkSent = 0
                    arUsers(k).linkCount = 0
                    arUsers(k).spamRowCount = 0
                    arUsers(k).spamTimeout = 0
                    arUsers(k).gameTimeout = 0
                    arUsers(k).gameSpamCount = 0
                    arUsers(k).loggedIn = False
                    Exit For
                End If
            End If
        Next k

        MDIForm1.StatusBar1.Panels(1).Text = "Users: " & frmUserlist.lstUserlist.ListItems.count
        
        If StrComp(nick, Form1.txtFindUser.Text, vbTextCompare) = 0 Then
            Form1.txtFindUser.Text = vbNullString
        End If
        
        lastUserid = vbNullString
    End If
End Sub


'0x0E
Sub updateGameStatus(msgSlot As Byte)
    Dim gameID As Double
    Dim status As Byte
    Dim strStatus As String
    Dim Users As Byte
    Dim maxUsers As Byte
    Dim strMaxUsers As String
    Dim i As Long
    
    'message(0)= 0
    gameID = BytesToNumEx(serverMessages(msgSlot).msgData, 1, 4, True)
    
    status = serverMessages(msgSlot).msgData(5)

    Users = serverMessages(msgSlot).msgData(6)
    maxUsers = serverMessages(msgSlot).msgData(7)
    strMaxUsers = CStr(Users) & "/" & CStr(maxUsers)
    
        Dim here As Boolean
        For i = 1 To Form1.lstGamelist.ListItems.count
            If StrComp(Form1.lstGamelist.ListItems(i).SubItems(5), gameID, vbBinaryCompare) = 0 Then
                Form1.lstGamelist.ListItems(i).SubItems(3) = statusCheck(1, status)
                Form1.lstGamelist.ListItems(i).SubItems(4) = strMaxUsers
                If LCase$(strStatus) <> "netsync" Then Form1.lstGamelist.Sorted = True
                Exit For
            ElseIf i = Form1.lstGameUserlist.ListItems.count And here = False Then
                Form1.lstGameUserlist.Refresh
                i = 0
                here = True
            End If
        Next i
End Sub




'0x0B
Sub quitGameRequest()
    packet(0) = &H0
    packet(1) = &HFF
    packet(2) = &HFF
        
    Call constructPacket(3, &HB)
    leftRoom = True
End Sub


'0x0B
Sub quitGameNotification(msgSlot As Byte)
    Dim i As Long
    Dim w As ListItem
    Dim nick As String
    Dim userID As Long
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
        
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'userid
    userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
      
    If userID <> myUserId Then
        'display
        If frmPreferences.chkShowJoin.Value = vbChecked Then
            Form1.txtGameChatroom.SelColor = &H800000
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:"
            If frmPreferences.chkTimeStamps.Value = vbChecked Then
                Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & Time & ": "
            End If
            Form1.txtGameChatroom.SelColor = &H808080
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & nick
            Form1.txtGameChatroom.SelColor = &H800000
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & " [Quit]" & vbCrLf
            Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
        End If
        Close #5
        Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
            Print #5, "-:" & Time & ": " & nick & " [Quit]"
        Close #5
        
        'remove from userlist
        Dim here As Boolean
        For i = 1 To Form1.lstGameUserlist.ListItems.count
            If StrComp(Form1.lstGameUserlist.ListItems(i).SubItems(3), userID, vbBinaryCompare) = 0 Then
                Form1.lstGameUserlist.ListItems.remove (i)
                Exit For
            ElseIf i = Form1.lstGameUserlist.ListItems.count And here = False Then
                Form1.lstGameUserlist.Refresh
                i = 0
                here = True
            End If
        Next i
    Else
        If leftRoom = False Then
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
            Form1.txtChatroom.SelColor = &H8000&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "*Client Alert: YOU HAVE BEEN KICKED FROM GAME!*" & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        End If

        myGameId = -1
        Form1.fRoomList.Caption = "Not in a game!"
        Form1.fGameroom.Caption = "Not in a game!"
        If rSwitch = True Then Call Form1.btnToggle_Click
        leftRoom = False
        inRoom = False
        imOwner = False
        Form1.lstGameUserlist.ListItems.Clear
        Form1.txtGameChat.Text = vbNullString
        Form1.txtGameChatroom.Text = vbNullString
    End If
End Sub


'0x0D
Sub playerInformation(msgSlot As Byte)
    Dim i As Long
    Dim w As Long
    Dim userlist As ListItem
    Dim nick As String
    Dim ping As Double 'dword
    Dim userID As Long 'word
    Dim connection As Byte
    Dim numOfUsers As Double 'dword
    Dim strConnection As String
    Dim jumpOut As Boolean
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    'Get Number of users
    'message(0) = 0
    numOfUsers = BytesToNumEx(serverMessages(msgSlot).msgData, 1, 4, True)

    i = 5
    w = 0
    jumpOut = False
    'for each user
    While w < numOfUsers
        While i < UBound(serverMessages(msgSlot).msgData) And jumpOut = False
            'nick
            strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
            If strSize > 511 Then
                strSize = 511
            End If
            Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
            temp(strSize) = &H0
            nick = ByteArrayToString(temp)
            i = i + strSize + 1
            
            'ping
            ping = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
            i = i + 4
            
            'userid
            userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
            
            'connection
            connection = serverMessages(msgSlot).msgData(i + 2) '+2
                        
            'display
            Set userlist = Form1.lstGameUserlist.ListItems.Add(, , nick)
            userlist.SubItems(1) = CStr(ping)
            If Len(userlist.SubItems(1)) = 2 Then
                userlist.SubItems(1) = "0" & userlist.SubItems(1)
            ElseIf Len(userlist.SubItems(1)) = 1 Then
                userlist.SubItems(1) = "00" & userlist.SubItems(1)
            End If
            userlist.SubItems(2) = statusCheck(0, connection)
            userlist.SubItems(3) = CStr(userID)
            
            jumpOut = True
            i = i + 3 'must keep track of what byte we are at after connection: +2 then increase +1, so +3
        Wend
        w = w + 1
        jumpOut = False
    Wend
End Sub

'0x04
Sub serverStatus(msgSlot As Byte)
    Dim numOfUsers As Double 'dword
    Dim numOfGames As Double 'dword
    Dim i As Long
    Dim w As Long
    Dim k As Long
    Dim lst As ListItem
    Dim nick As String
    Dim ping As Double 'dword
    Dim status As Byte
    Dim userID As Long 'word
    Dim connection As Byte
    Dim game As String
    Dim gameID As Double
    Dim owner As String
    Dim emulator As String
    Dim str() As String
    Dim q As Long
    Dim r As Long
    Dim maxUsers As String
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    'Get Number of users
    'message(0) = 0
    numOfUsers = BytesToNumEx(serverMessages(msgSlot).msgData, 1, 4, True)
    'Get Number of games
    numOfGames = BytesToNumEx(serverMessages(msgSlot).msgData, 5, 8, True)
    i = 9

'----USERLIST----
    'for each user
    For w = 0 To numOfUsers - 1
        'nick
        strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
        If strSize > 511 Then
            strSize = 511
        End If
        'Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
        'temp(strSize) = &H0
        'nick = ByteArrayToString(temp)
        i = i + strSize + 1
            
        'ping
        'ping = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
        i = i + 4
            
        'status
        'status = serverMessages(msgSlot).msgData(i)  'is this correct?
        i = i + 1
        'userid
        'byte swap, second_byte * 256 + first_byte
        'userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
        i = i + 2
            
        'connection
        'connection = serverMessages(msgSlot).msgData(i)
        i = i + 1
        
        'userCount = userCount + 1
        'arUsers(userID).name = nick
        'arUsers(userID).ping = ping
        'arUsers(userID).connection = connection
        'arUsers(userID).status = status
        'arUsers(userID).userID = userID
        'arUsers(userID).loggedIn = True
        'Call addUsers(userID, vbNullString)
                
        'If userCount = 1 Then
        '    userBottom = userID
        '    firstID = userBottom
        'Else
        '    If userID < userBottom Then
        '        userBottom = userID
        '    ElseIf userID > userTop Then
        '        userTop = userID
        '    End If
        'End If
        
        'For k = LBound(arUsers) To UBound(arUsers)
        '    If arUsers(k).loggedIn = False Then
        '        arUsers(k).name = nick
        '        arUsers(k).ping = ping
        '        arUsers(k).connection = connection
        '        arUsers(k).status = status
        '        arUsers(k).userID = userID
        '        arUsers(k).loggedIn = True
        '        Exit For
        '    End If
        'Next k
    Next w

'----GAMELIST----
Form1.lstGamelist.Visible = False

    'for each game
    For w = 0 To numOfGames - 1
        'game
        strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
        If strSize > 511 Then
            strSize = 511
        End If
        Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
        temp(strSize) = &H0
        game = ByteArrayToString(temp)
        i = i + strSize + 1
            
        'gameid
        gameID = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
        i = i + 4
            
        'emulator
        strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
        If strSize > 511 Then
            strSize = 511
        End If
        Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
        temp(strSize) = &H0
        emulator = ByteArrayToString(temp)
        i = i + strSize + 1
            
        'for owner
        strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
        If strSize > 511 Then
            strSize = 511
        End If
        Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
        temp(strSize) = &H0
        owner = ByteArrayToString(temp)
        i = i + strSize + 1
            
        'for users/maxusers
        strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
        Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
        temp(strSize) = &H0
        maxUsers = ByteArrayToString(temp)
        i = i + strSize + 1

        'status
        status = serverMessages(msgSlot).msgData(i)
        i = i + 1
                   
        For k = LBound(arGames) To UBound(arGames)
            If arGames(k).opened = False Then
                arGames(k).game = game
                arGames(k).emulator = emulator
                arGames(k).gameID = gameID
                arGames(k).Users = maxUsers
                arGames(k).status = status
                arGames(k).owner = owner
                arGames(k).opened = True
                Exit For
            End If
        Next k
        
        game = Replace$(game, ",", ";", 1, -1, vbBinaryCompare)
               
        For q = LBound(arGames) To UBound(arGames)
            If arUsers(q).loggedIn = True Then
                If StrComp(owner, arUsers(q).name, vbTextCompare) = 0 Then
                    userID = arUsers(q).userID
                    str = Split(ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).gamesPlayed, ",")
                    For r = 0 To UBound(str)
                        If StrComp(str(r), game, vbTextCompare) = 0 Then
                            Exit For
                        ElseIf r = UBound(str) Then
                            ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).gamesPlayed = ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).gamesPlayed & "," & game
                        End If
                    Next r
                    If UBound(str) = -1 Then
                        ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).gamesPlayed = game
                    End If
                    ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).numOfGames = ipHead(arUsers(q).head).ipSect(arUsers(q).sect1).myEntries(arUsers(q).dbPos).numOfGames + 1
                    Exit For
                End If
            End If
        Next q

    Next w
    
    Form1.lstGamelist.Visible = True
    Form1.Caption = "Connected to: " & frmServerlist.txtServerIp.Text
    MDIForm1.StatusBar1.Panels(1).Text = "Users: " & frmUserlist.lstUserlist.ListItems.count
    MDIForm1.StatusBar1.Panels(2).Text = "Games: " & Form1.lstGamelist.ListItems.count
End Sub

'0x0C
Sub joinGameNotification(msgSlot As Byte)
    Dim i As Long
    Dim userlist As ListItem
    Dim username As String
    Dim pGame As Double 'dword
    Dim ping As Double 'dword
    Dim userID As Double 'word
    Dim connection As Byte
    Dim temp(0 To 511) As Byte
    Dim strSize As Long

    'message(0)=0
    'Pointer to Game on Server...has no real use
    pGame = BytesToNumEx(serverMessages(msgSlot).msgData, 1, 4, True)
    
    'for username
    i = 5
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    username = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'ping
    ping = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
    i = i + 4
    
    'for userid
    userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
    
    If inRoom = True And userID = myUserId Then Exit Sub    'connection
    
    connection = serverMessages(msgSlot).msgData(i + 1 + 1)
    'display
    Set userlist = Form1.lstGameUserlist.ListItems.Add(, , username)
    userlist.SubItems(1) = CStr(ping)
    If Len(userlist.SubItems(1)) = 2 Then
        userlist.SubItems(1) = "0" & userlist.SubItems(1)
    ElseIf Len(userlist.SubItems(1)) = 1 Then
        userlist.SubItems(1) = "00" & userlist.SubItems(1)
    End If
    userlist.SubItems(2) = statusCheck(0, connection)
    userlist.SubItems(3) = CStr(userID)
            
    If Form1.lstGameUserlist.ListItems.count < 2 And userID = myUserId Then
        Close #5
        Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
            Print #5, "***********CREATED Session - " & Time & " " & Date & "*************"
        Close #5
        Form1.txtGameChatroom.SelColor = &H800000
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "******Created Session - " & Time & " " & Date & "******" & vbCrLf
        imOwner = True
        Form1.btnKick.Visible = True
        Form1.txtKickUsers.Visible = True
        leftRoom = False
        Form1.Label24.Visible = True
        Call gameChatRequest("/maxusers " & Form1.txtKickUsers.Text)
        If rSwitch = False Then Call Form1.btnToggle_Click
    ElseIf userID = myUserId Then
        Close #5
        Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
            Print #5, "***********JOINED Session - " & Time & " " & Date & "*************"
        Close #5
        Form1.txtGameChatroom.SelColor = &H800000
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "******Joined Session - " & Time & " " & Date & "*" & vbCrLf
        Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
        imOwner = False
        Form1.btnKick.Visible = True
        Form1.txtKickUsers.Visible = True
        Form1.Label24.Visible = True
        leftRoom = False
        If rSwitch = False Then Call Form1.btnToggle_Click
    Else
        If imOwner = True Then
            If Trim$(frmPreferences.txtGameWelcomeMessage.Text) <> vbNullString Then
                Call splitRegGame(frmPreferences.txtGameWelcomeMessage.Text)
            End If
                        
            If frmPreferences.chkBeep.Value = vbChecked Then Call Beep
        End If
    End If
    
    If inRoom = False Then
        'Gameroom
        Form1.fRoomList.Caption = "Currently in: " & myGame
        Form1.fGameroom.Caption = myGame
        inRoom = True
    End If
        
    If frmPreferences.chkShowJoin.Value = vbChecked Then
        Form1.txtGameChatroom.SelColor = &H800000
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:"
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & Time & ": "
        End If
        Form1.txtGameChatroom.SelColor = &H808080
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & username
        Form1.txtGameChatroom.SelColor = &H800000
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & " [Joined]" & vbCrLf
        Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
    End If
    
    Close #5
    Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
        Print #5, "-:" & Time & ": " & username & " [Joined]"
    Close #5
End Sub

'0x0C
Sub joinGameRequest(gameID As String)
    Dim connection As Byte
    Dim temp() As Byte
    
    packet(0) = &H0
    myGameId = CLng(gameID)
    temp = LongToByteArray(myGameId)
    packet(1) = temp(0)
    packet(2) = temp(1)
    packet(3) = temp(2)
    packet(4) = temp(3)
    
    packet(5) = &H0
    
    packet(6) = &H0
    packet(7) = &H0
    packet(8) = &H0
    packet(9) = &H0
    
    packet(10) = &HFF
    packet(11) = &HFF
    
    If frmServerlist.cmbConnectionType.Text = "LAN" Then
        connection = 1
    ElseIf frmServerlist.cmbConnectionType.Text = "Excellent" Then
        connection = 2
    ElseIf frmServerlist.cmbConnectionType.Text = "Good" Then
        connection = 3
    ElseIf frmServerlist.cmbConnectionType.Text = "Average" Then
        connection = 4
    ElseIf frmServerlist.cmbConnectionType.Text = "Low" Then
        connection = 5
    ElseIf frmServerlist.cmbConnectionType.Text = "Bad" Then
        connection = 6
    End If
    
    packet(12) = connection
    
    Call constructPacket(13, &HC)
End Sub

'0x8
Sub gameChatNotification(msgSlot As Byte)
    Dim i As Long
    Dim user As ListItem
    Dim nick As String
    Dim id As Long
    Dim msg As String
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
    
    id = -1
    
    'username
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1
    
    nick = Replace$(nick, ",", ";", 1, -1, vbBinaryCompare)
    
    'user's message
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    msg = ByteArrayToString(temp)
        
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If LCase$(nick) = LCase$(arUsers(i).name) Then
                id = i
                Exit For
            End If
        End If
    Next i
    
    
    If StrComp(LCase$(nick), "server", vbBinaryCompare) = 0 Then
        If InStr(1, msg, "TO:", vbTextCompare) = 0 And InStr(1, msg, "> (", vbBinaryCompare) > 0 And InStr(1, msg, "):", vbBinaryCompare) > 0 Then
            Form1.txtGameChatroom.SelColor = &H800000
            Form1.txtGameChatroom.SelText = Form1.txtChatroom.SelText & "-:"
            If frmPreferences.chkTimeStamps.Value = vbChecked Then
                Form1.txtGameChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
            End If
            Form1.txtGameChatroom.SelColor = &H800000
            Form1.txtGameChatroom.SelText = Form1.txtChatroom.SelText & "[PM] "
            Form1.txtGameChatroom.SelColor = &HFF8000
            Form1.txtGameChatroom.SelText = Form1.txtChatroom.SelText & msg & vbCrLf
            Form1.txtGameChatroom.SelStart = Len(Form1.txtChatroom.Text)
            Exit Sub
        End If
        Form1.txtGameChatroom.SelColor = &H800000
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:"
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & Time & ": "
        End If
        Form1.txtGameChatroom.SelColor = &HC000C0
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & msg & vbCrLf
    Else
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtGameChatroom.SelColor = &H800000
            Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "-:" & Time & ": "
        End If
        
        If id <> -1 Then
            If StrComp(LCase$(arUsers(id).access), "admin", vbBinaryCompare) = 0 Or StrComp(LCase$(arUsers(id).access), "superadmin", vbBinaryCompare) = 0 Then
                Form1.txtGameChatroom.SelColor = &H8000&
            ElseIf StrComp(LCase$(arUsers(id).access), "elevated", vbBinaryCompare) = 0 Then
                Form1.txtGameChatroom.SelColor = &H4080&
            ElseIf StrComp(LCase$(arUsers(id).access), "moderator", vbBinaryCompare) = 0 Then
                Form1.txtGameChatroom.SelColor = &H8080&
            Else
                Form1.txtGameChatroom.SelColor = vbBlack
            End If
        Else
            Form1.txtGameChatroom.SelColor = vbBlack
        End If
        
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & "<" & nick & "> "
        Form1.txtGameChatroom.SelColor = vbBlack
        Form1.txtGameChatroom.SelText = Form1.txtGameChatroom.SelText & msg & vbCrLf
    End If
            
            
    Close #5
    Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
        Print #5, "-:" & Time & ": " & "<" & nick & "> " & msg
    Close #5
        

    Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
End Sub


'0x08
Sub gameChatRequest(msg As String)
    Dim temp() As Byte
    Dim dataToBeSent As String
       
    dataToBeSent = vbNullChar & Trim$(msg) & vbNullChar
    temp = StrConv(dataToBeSent, vbFromUnicode)
    
    Call CopyMemory(packet(0), temp(0), ByVal UBound(temp) + 1)
    
    Call constructPacket(UBound(temp) + 1, &H8)
End Sub


'0x11
Sub startGameRequest()
    packet(0) = &H0
    packet(1) = &HFF
    packet(2) = &HFF
    packet(3) = &HFF
    packet(4) = &HFF
    
    Call constructPacket(5, &H11)
End Sub

'0x7
Sub globalChatNotification(msgSlot As Byte)
    Dim i As Long
    Dim j As Long
    Dim nick As String
    Dim msg As String
    Dim id As Long
    Dim w As Long
    Dim temp(0 To 511) As Byte
    Dim str1 As String
    Dim strSize As Long
    On Error Resume Next
    
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    id = -1
    
    'username
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = i + strSize + 1

    'user's message
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    msg = ByteArrayToString(temp)
        
    Dim str() As String
    Dim user As ListItem
    Dim lNick As String
    Dim lMsg As String
    Static spamTime As Long
    
    lNick = LCase$(nick)
    lMsg = LCase$(msg)
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If StrComp(lNick, LCase$(arUsers(i).name), vbBinaryCompare) = 0 Then
                id = i
                Exit For
            End If
        End If
    Next i

    If frmPreferences.chkTimeStamps.Value = vbChecked Then
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:" & Time & ": "
    End If
            
    If id = -1 Then
        Form1.txtChatroom.SelColor = vbBlack
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "<" & nick & "> "

        Close #3
        Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
            Print #3, "-:" & Time & ": " & "<" & nick & "> " & msg
        Close #3
    
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & msg & vbCrLf
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        Exit Sub
    End If
    
    Dim u As Long
            If StrComp(LCase$(arUsers(id).access), "admin", vbBinaryCompare) = 0 Or StrComp(LCase$(arUsers(id).access), "superadmin", vbBinaryCompare) = 0 Then
                Form1.txtChatroom.SelColor = &H8000&

                If StrComp(lMsg, ";bot off " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(msg, ";bot off *", vbBinaryCompare) = 0 Then
                    If myBot.botStatus = True Then Call frmAdminBot.btnONOFF_Click
                    Call globalChatRequest(frmAdminBot.txtBotName.Text & " is OFF!")
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";bot on " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(lMsg, ";bot on *", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(frmAdminBot.txtBotName.Text & " Bot is ON!")
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                    If myBot.botStatus = False Then Call frmAdminBot.btnONOFF_Click
                ElseIf StrComp(lMsg, ";alerts on " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(lMsg, ";alerts on *", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(frmAdminBot.txtBotName.Text & " ALERTS are ON!")
                    frmPreferences.chkAlertOthers.Value = vbChecked
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";alerts off " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(lMsg, ";alerts off *", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(frmAdminBot.txtBotName.Text & " ALERTS are OFF!")
                    frmPreferences.chkAlertOthers.Value = vbUnchecked
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";quit " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(lMsg, ";quit *", vbBinaryCompare) = 0 Then
                    Call globalChatRequest("Logging Out!")
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                    Call MDIForm1.mnuLogOffServer_Click
                ElseIf StrComp(lMsg, ";reconnect " & LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Or StrComp(lMsg, ";reconnect *", vbBinaryCompare) = 0 Then
                    Call globalChatRequest("Reconnecting!")
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                    Call userQuitRequest("Ping Timeout")
                ElseIf StrComp(lMsg, ";date", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(CStr(Date))
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";time", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(CStr(Time))
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";ip", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(frmServerlist.txtServerIp.Text)
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";whoisalive", vbBinaryCompare) = 0 Then
                    Call globalChatRequest(";iamalive")
                    frmServerlist.List1.AddItem "<" & nick & ">" & msg
                ElseIf StrComp(lMsg, ";iamalive", vbBinaryCompare) = 0 And StrComp(lNick, LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) <> 0 Then
                    Dim vv As ListItem
                    Set vv = frmRemote.lstBots.ListItems.Add(, , nick)
                End If
            ElseIf StrComp(LCase$(arUsers(id).access), "elevated", vbBinaryCompare) = 0 Then
                Form1.txtChatroom.SelColor = &H4080&
            ElseIf StrComp(LCase$(arUsers(id).access), "moderator", vbBinaryCompare) = 0 Then
                Form1.txtChatroom.SelColor = &H8080&
            Else
                Form1.txtChatroom.SelColor = vbBlack
            End If

    Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "<" & nick & "> "
    
    Form1.txtChatroom.SelColor = vbBlack
    
    'Debug.Print LenB(Form1.txtChatroom.Text)
    'If LenB(Form1.txtChatroom.Text) > 1000 Then
    '    Form1.txtChatroom.Text = Mid$(Form1.txtChatroom.Text, 500, Len(Form1.txtChatroom.Text) - 700)
    'End If
    
    Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & msg & vbCrLf
        
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    msg = Trim$(msg)
    
    Close #3
    Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
        Print #3, "-:" & Time & ": " & "<" & nick & "> " & msg
    Close #3
    
    If StrComp(LCase$(arUsers(id).access), "normal", vbBinaryCompare) <> 0 Then
        Exit Sub
    End If
    
    Dim dl As Boolean
    If InStr(1, lMsg, "http://", vbBinaryCompare) > 0 Then
        dl = True
    ElseIf InStr(1, lMsg, "https://", vbBinaryCompare) > 0 Then
        dl = True
    ElseIf InStr(1, lMsg, "www.", vbBinaryCompare) > 0 Then
        dl = True
    ElseIf InStr(1, lMsg, "ftp://", vbBinaryCompare) > 0 Then
        dl = True
    End If
    
    If myBot.botStatus = True Then
        'Link Control
         If dl = True And frmAdminBot.chkLinkControl.Value = vbChecked Then
            arUsers(id).linkSent = arUsers(id).linkSent + 1
            If arUsers(id).linkCount <= CLng(frmAdminBot.txtLinksInterval.Text) And arUsers(id).linkSent >= frmAdminBot.txtLinkSend.Text Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(nick & " [" & frmAdminBot.txtLinkMessage.Text & "] Limit " & frmAdminBot.txtLinkSend.Text & " links/" & frmAdminBot.txtLinksInterval.Text & "s")
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtLinkMessage.Text & "] Limit " & frmAdminBot.txtLinkSend.Text & " links for every " & frmAdminBot.txtLinksInterval.Text & "s")
                End If
                If frmAdminBot.chkLinkBan.Value = vbChecked Then
                    Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtLinkMin.Text)
                ElseIf frmAdminBot.chkLinkSilenceKick = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLinkMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkLinkKick = vbChecked Then
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkLinkSilence = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLinkMin.Text)
                Else
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLinkMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                    frmAdminBot.chkSpamSilenceKick.Value = vbChecked
                End If
                Call addDamage(nick, arUsers(id).ip, "Link Spamming")
                arUsers(id).linkCount = 0
                arUsers(id).linkSent = 0
            End If
        End If
        
        'Spam Control
        If frmAdminBot.chkSpamControl.Value = vbChecked Then
            'Spam in a Row
            arUsers(id).spamRowCount = arUsers(id).spamRowCount + 1
            If Len(msg) < CLng(frmAdminBot.txtSpamChars.Text) Then
                arUsers(id).spamRowCount = 0
                arUsers(id).spamTimeout = 0
            End If
            If arUsers(id).spamRowCount >= CLng(frmAdminBot.txtSpamRow.Text) Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(nick & " [" & frmAdminBot.txtSpamMessage.Text & "]")
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtSpamMessage.Text & "]")
                End If
                If frmAdminBot.chkSpamBan.Value = vbChecked Then
                    Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtSpamMin.Text)
                ElseIf frmAdminBot.chkSpamSilenceKick = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtSpamMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkSpamKick = vbChecked Then
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkSpamSilence = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtSpamMin.Text)
                Else
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtSpamMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                    frmAdminBot.chkSpamSilenceKick.Value = vbChecked
                End If
                Call addDamage(nick, arUsers(id).ip, "Spamming Chatroom")
                arUsers(id).spamRowCount = 0
                arUsers(id).spamTimeout = 0
            End If
        End If
    
        'line wrapper
        Dim lastSpace As Long
        Dim spaceCount As Long
        Dim posBracket1, posBracket2 As Long
        Dim strNick As String
        Dim found As Boolean
        
        If frmAdminBot.chkLineWrapper.Value = vbChecked Then
            posBracket1 = InStr(1, msg, "<", vbBinaryCompare)
            posBracket2 = InStr(1, msg, ">", vbBinaryCompare)
            
            If posBracket1 > 0 And posBracket2 > posBracket1 Then
                strNick = Mid$(msg, posBracket1 + 1, posBracket2 - posBracket1 - 1)
                
                If posBracket1 >= CLng(frmAdminBot.txtMaxSpace.Text) Then
                '    For i = LBound(arUsers) To UBound(arUsers)
                '        If lcase$(strNick) = lcase$(arUsers(i).name) And arUsers(i).loggedIn = True Then
                            found = True
                '            Exit For
                '        End If
                '    Next i
                End If
                
                If found = True Then
                    If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                        Call splitReg(nick & " [" & frmAdminBot.txtLineWrapperMessage.Text & "]")
                    Else
                        Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtLineWrapperMessage.Text & "]")
                    End If
                    If frmAdminBot.chkLineWrapperBan.Value = vbChecked Then
                        Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtLineWrapperMin.Text)
                    ElseIf frmAdminBot.chkLineWrapperSilenceKick = vbChecked Then
                        Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLineWrapperMin.Text)
                        Call globalChatRequest("/kick " & arUsers(id).userID)
                    ElseIf frmAdminBot.chkLineWrapperKick = vbChecked Then
                        Call globalChatRequest("/kick " & arUsers(id).userID)
                    ElseIf frmAdminBot.chkLineWrapperSilence = vbChecked Then
                        Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLineWrapperMin.Text)
                    Else
                        Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtLineWrapperMin.Text)
                        Call globalChatRequest("/kick " & arUsers(id).userID)
                        frmAdminBot.chkLineWrapperSilenceKick.Value = vbChecked
                    End If
                    Call addDamage(nick, arUsers(id).ip, "LINE WRAPPER DETECTED: " & msg)
                End If
            End If
        End If
        
    
        'ALL CAPS
        If frmAdminBot.chkAllCaps.Value = vbChecked Then
            Dim up As String
            Dim allCaps As Boolean
            Dim h As Long
            Dim c As Long
            
            msg = removeNonLetters(msg)
            
            'Count Caps
            For h = 1 To Len(msg)
                If StrComp(Mid$(msg, h, 1), UCase$(Mid$(msg, h, 1)), vbBinaryCompare) = 0 Then
                    c = c + 1
                End If
            Next h
            
            up = UCase$(msg)
            If StrComp(up, msg, vbBinaryCompare) = 0 Then
                allCaps = True
            End If
                   
            If c <= CLng(frmAdminBot.txtTotalCaps.Text - 1) Then
                allCaps = False
            End If
    
            
    
            If allCaps = True Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(nick & " [" & frmAdminBot.txtAllCapsMessage.Text & "]")
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtAllCapsMessage.Text & "]")
                End If
                If frmAdminBot.chkAllCapsBan.Value = vbChecked Then
                    Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtAllCapsMin.Text)
                ElseIf frmAdminBot.chkAllCapsSilenceKick = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtAllCapsMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkAllCapsKick = vbChecked Then
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                ElseIf frmAdminBot.chkAllCapsSilence = vbChecked Then
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtAllCapsMin.Text)
                Else
                    Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtAllCapsMin.Text)
                    Call globalChatRequest("/kick " & arUsers(id).userID)
                    frmAdminBot.chkAllCapsSilenceKick.Value = vbChecked
                End If
                Call addDamage(nick, arUsers(id).ip, "ALL CAPS: " & msg)
            End If
        End If
    
        'Word Filter
        Dim newMsg As String
        Dim b As Long
        If frmAdminBot.chkWordFilter.Value = vbChecked Then
            If frmAdminBot.lstWord.ListCount < 0 Then Exit Sub
            
            newMsg = Trim$(lMsg)
            str = Split(newMsg, " ")
            newMsg = Replace$(newMsg, " ", vbNullString)
            For i = 0 To frmAdminBot.lstWord.ListCount - 1
                For w = 0 To UBound(str)
                    If frmAdminBot.lstWord.List(i) = str(w) Then
                        If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                            Call splitReg(nick & " [" & frmAdminBot.txtFilterMessage.Text & " Keyword: " & frmAdminBot.lstWord.List(i) & "]")
                        Else
                            Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtFilterMessage.Text & " Keyword: " & frmAdminBot.lstWord.List(i) & "]")
                        End If
                        If frmAdminBot.chkWordBan.Value = vbChecked Then
                            Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                        ElseIf frmAdminBot.chkWordSilenceKick = vbChecked Then
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                        ElseIf frmAdminBot.chkWordKick = vbChecked Then
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                        ElseIf frmAdminBot.chkWordSilence = vbChecked Then
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                        Else
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                            frmAdminBot.chkWordSilenceKick.Value = vbChecked
                        End If
                        Call addDamage(nick, arUsers(id).ip, "Word Filter: " & frmAdminBot.lstWord.List(i) & ": " & msg)
                        Exit Sub
                    End If
                Next w
            Next i
        
            Dim yes As Boolean
            
            yes = True
            For b = 0 To UBound(str)
                If Len(str(b)) <> 1 Then
                    yes = False
                    Exit For
                End If
            Next b
            
            For b = 0 To frmAdminBot.lstWord.ListCount - 1
                If Len(frmAdminBot.lstWord.List(b)) = UBound(str) + 1 Or yes Then
                    If InStr(1, newMsg, frmAdminBot.lstWord.List(b), vbBinaryCompare) > 0 Then
                        If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                            Call splitReg(nick & " [" & frmAdminBot.txtFilterMessage.Text & ": Keyword: " & frmAdminBot.lstWord.List(i) & "]")
                        Else
                            Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & nick & " [" & frmAdminBot.txtFilterMessage.Text & ": Keyword: " & frmAdminBot.lstWord.List(i) & "]")
                        End If
                        If frmAdminBot.chkWordBan.Value = vbChecked Then
                            Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                        ElseIf frmAdminBot.chkWordSilenceKick = vbChecked Then
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                        ElseIf frmAdminBot.chkWordKick = vbChecked Then
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                        ElseIf frmAdminBot.chkWordSilence = vbChecked Then
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                        Else
                            Call globalChatRequest("/silence " & arUsers(id).userID & " " & frmAdminBot.txtWordMin.Text)
                            Call globalChatRequest("/kick " & arUsers(id).userID)
                            frmAdminBot.chkWordSilenceKick.Value = vbChecked
                        End If
                        Call addDamage(nick, arUsers(id).ip, "Word Filter: " & frmAdminBot.lstWord.List(i) & ": " & msg)
                        Exit Sub
                    End If
                End If
            Next b
        End If
    End If

End Sub

'0x07
Sub globalChatRequest(message As String)
    Dim temp() As Byte
    Dim dataToBeSent As String
    
    dataToBeSent = vbNullChar & message & vbNullChar
    temp = StrConv(dataToBeSent, vbFromUnicode)
    
    Call CopyMemory(packet(0), temp(0), ByVal UBound(temp) + 1)
    
    Call constructPacket(UBound(temp) + 1, &H7)
End Sub


'0x02
Sub userJoined(msgSlot As Byte)
    Dim i As Long
    Dim k As Long
    Dim userlist As ListItem
    Dim nick As String
    Dim connection As Byte
    Dim strConnection As String
    Dim ping As Double 'dword
    Dim userID As Long 'word
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    'nick
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(0)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(0), ByVal strSize)
    temp(strSize) = &H0
    nick = ByteArrayToString(temp)
    i = strSize + 1
    
    userID = serverMessages(msgSlot).msgData(i + 1) * CLng(256) + serverMessages(msgSlot).msgData(i)
    'this is how we store our userid
    If inServer = False Then 'And unrecoveredMessage = False Then
        
        inServer = True
        
        Dim otherNames As String
        Dim t As Long
        Dim r As Long
        
        t = GetTickCount
        Form1.Timer2.Enabled = False
        
        frmServerlist.List1.AddItem "Scanning Initial Users..."
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        frmServerlist.ProgressBar1.Max = UBound(arUsers) + 1
        frmServerlist.ProgressBar1.Value = 0
        frmUserlist.lstUserlist.ListItems.Clear
        MDIForm1.Enabled = False
        For k = LBound(arUsers) To UBound(arUsers)
            If arUsers(k).loggedIn = True Then
                otherNames = checkUsers(k, arUsers(k).head, arUsers(k).sect1)
                Call addUsers(k, otherNames)
            End If
            
            frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Value + 1
            r = r + 1
            If r = 10 Then
                r = 0
                frmServerlist.lblPer.Caption = CLng((k / UBound(arUsers)) * 100) & "% Completed"
                DoEvents
            End If
        Next k
        
        For k = LBound(arGames) To UBound(arGames)
            If arGames(k).opened = True Then
                Call addGames(k)
            End If
        Next k
        
        Call sortRoomlist
        MDIForm1.Enabled = True
        frmServerlist.ProgressBar1.Value = 0
        frmServerlist.lblPer.Caption = "100% Complete in " & Abs(CSng((GetTickCount - t)) / 1000) & "s"
        frmServerlist.List1.AddItem "Finished"
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        MDIForm1.StatusBar1.Panels(1).Text = "Users: " & frmUserlist.lstUserlist.ListItems.count
        myUserId = userID 'set global
        Call weLoggedIn
    End If
    
    Form1.Caption = "Connected to: " & frmServerlist.txtServerIp.Text & "; Last UserID: " & userID
        
    'ping
    'i = i + 2
    'ping = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
    'i = i + 4
    'connection = serverMessages(msgSlot).msgData(i)

    
    If userID = myUserId Then
        myConnectionType = connection
        myPing = ping
    End If
                            
    If frmPreferences.chkShowJoin.Value = vbChecked Then
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
        End If
        
        Form1.txtChatroom.SelColor = &H808080
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & nick
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & " [Joined]" & vbCrLf
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    End If

    Close #3
    Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
        Print #3, "-:" & Time & ": " & nick & " [Joined]"
    Close #3
        
    Form1.txtFindUser.Text = nick
    lastUserid = userID
End Sub

'0x0A
Sub createGameNotification(msgSlot As Byte)
    Dim i As Long
    Dim w As Long
    Dim gamelist As ListItem
    Dim owner As String
    Dim id As Long
    Dim game As String
    Dim vv As ListItem
    Dim emulator As String
    Dim str() As String
    Dim gameID As Double 'word
    Dim temp(0 To 511) As Byte
    Dim strSize As Long
    On Error Resume Next
    
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    id = -1
    
    'for owner
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    owner = ByteArrayToString(temp)
    i = i + strSize + 1
    
    owner = Replace$(owner, ",", ";", 1, -1, vbBinaryCompare)
    
    'for game
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    game = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'for emulator
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    emulator = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'for gameid
    gameID = BytesToNumEx(serverMessages(msgSlot).msgData, i, i + 3, True)
            
    If frmPreferences.chkShowOpen.Value = vbChecked Then
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
        End If
        Form1.txtChatroom.SelColor = vbRed
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & owner
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & " [Created Game] "
    
        Form1.txtChatroom.SelColor = vbRed
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & game & vbCrLf
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    End If
    
    Close #3
    Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
        Print #3, "-:" & Time & ": " & owner & " [Created Game] " & game
    Close #3
    
    Dim s() As String
    Dim head As Long
    Dim lOwner As String
    Dim lGame As String
    Dim lEmulator As String
    
    lEmulator = LCase$(emulator)
    lGame = LCase$(game)
    lOwner = LCase$(owner)

    If StrComp(lOwner, LCase$(frmServerlist.txtUsername.Text), vbBinaryCompare) = 0 Then
        myGame = game
        myGameId = gameID
    End If
            
            
    Dim k As Long
    Dim p As Long
    For k = LBound(arGames) To UBound(arGames)
        If arGames(k).opened = False Then
            arGames(k).game = game
            arGames(k).emulator = emulator
            arGames(k).gameID = gameID
            arGames(k).Users = "1/2"
            arGames(k).status = 0
            For p = LBound(arUsers) To UBound(arUsers)
                If StrComp(LCase$(arUsers(p).name), lOwner, vbBinaryCompare) = 0 Then
                    arGames(k).ownerAccess = arUsers(p).access
                    Exit For
                End If
            Next p
            arGames(k).owner = owner
            arGames(k).opened = True
            Call addGames(k)
            Exit For
        End If
    Next k
    
    MDIForm1.StatusBar1.Panels(2).Text = "Games: " & Form1.lstGamelist.ListItems.count
    Call sortRoomlist
    game = Replace$(game, ",", ";", 1, -1, vbBinaryCompare)
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If StrComp(lOwner, LCase(arUsers(i).name), vbBinaryCompare) = 0 Then
                id = i
                head = CLng(Left$(arUsers(i).ip, 3))
                str = Split(ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).gamesPlayed, ",")
                For w = 0 To UBound(str)
                    If StrComp(LCase$(str(w)), lGame, vbBinaryCompare) = 0 Then
                        Exit For
                    ElseIf w = UBound(str) Then
                        ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).gamesPlayed = ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).gamesPlayed & "," & game
                    End If
                Next w
                If UBound(str) = -1 Then
                    ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).gamesPlayed = game
                End If
                ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).numOfGames = ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).numOfGames + 1
                Exit For
            End If
        End If
    Next i
 
    If myBot.botStatus = True And frmAdminBot.chkCreateGame.Value = vbChecked And Trim$(frmAdminBot.txtCreateGame.Text) > vbNullString Then
        Call globalChatRequest("/announcegame " & gameID & " " & frmAdminBot.txtCreateGame.Text)
    End If
    
    If id = -1 Then Exit Sub
    
    If LCase$(arUsers(id).access) <> "normal" Then
        Exit Sub
    End If
    
    'game disable control
    If frmAdminBot.chkGameDisable.Value = vbChecked And myBot.botStatus = True Then
        For i = 1 To frmAdminBot.lstDisableHosting.ListItems.count
            If StrComp(frmAdminBot.lstDisableHosting.ListItems(i).Text, arUsers(id).ip, vbBinaryCompare) = 0 Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(owner & " [Disabled from hosting a game]")
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & owner & " [Disabled from hosting a game]")
                End If
                
                If StrComp(LCase$(frmAdminBot.lstDisableHosting.ListItems(i).SubItems(2)), "close game", vbBinaryCompare) = 0 Then
                    Call globalChatRequest("/closegame " & gameID)
                Else
                    Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.lstDisableHosting.ListItems(i).SubItems(3))
                End If
                Call addDamage(owner, arUsers(id).ip, "Disabled from Hosting")
                Exit For
            End If
        Next i
    End If
        
    'game control
    If frmAdminBot.chkGameControl.Value = vbChecked And myBot.botStatus = True Then
    
        For i = 1 To frmAdminBot.lstGame.ListItems.count
            If StrComp(LCase$(frmAdminBot.lstGame.ListItems(i).Text), lGame, vbBinaryCompare) = 0 Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg("Blocked [" & game & "] from being hosted!")
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & " has blocked [" & game & "] from being hosted!")
                End If
                'Call globalChatRequest("/closegame " & CStr(gameID))
                Call addDamage(owner, arUsers(id).ip, "Game Block: " & game)
                Call globalChatRequest("/kick " & arUsers(id).userID)
                Exit Sub
            ElseIf StrComp(LCase$(frmAdminBot.lstGame.ListItems(i).Text), lEmulator, vbBinaryCompare) = 0 Then
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(owner & ", you're not allowed to use that emulator [" & emulator & "]!")
                    Call splitReg(frmAdminBot.lstGame.ListItems(i).SubItems(1))
                Else
                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & owner & ", you're not allowed to use that emulator [" & emulator & "]!")
                    Call splitAnnounce(frmAdminBot.lstGame.ListItems(i).SubItems(1))
                End If
                'Call globalChatRequest("/closegame " & CStr(gameID))
                Call globalChatRequest("/kick " & arUsers(id).userID)
                Call addDamage(owner, arUsers(id).ip, "Emulator Block: " & emulator)
                Exit Sub
            End If
        Next i
    End If


    'Game Spamming
    If frmAdminBot.chkGameSpamControl.Value = vbChecked And myBot.botStatus = True Then
        arUsers(id).gameSpamCount = arUsers(id).gameSpamCount + 1
        If arUsers(id).gameSpamCount >= CLng(frmAdminBot.txtGameroomNum.Text) Then
            Call addDamage(owner, arUsers(id).ip, "Game Spamming")
            If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                Call splitReg(owner & " " & frmAdminBot.txtGameroomMessage.Text)
            Else
                Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & owner & " " & frmAdminBot.txtGameroomMessage.Text)
            End If
            If frmAdminBot.chkGameSpamBan.Value = vbChecked Then
                Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtGameroomBan.Text)
            ElseIf frmAdminBot.chkGameSpamKick = vbChecked Then
                Call globalChatRequest("/kick " & arUsers(id).userID)
            Else
                Call globalChatRequest("/ban " & arUsers(id).userID & " " & frmAdminBot.txtGameroomBan.Text)
                frmAdminBot.chkGameSpamBan.Value = vbChecked
            End If
            arUsers(id).gameSpamCount = 0
            arUsers(id).gameTimeout = 0
        End If
    End If
End Sub

'0x0A
Sub createGameRequest(gameName As String)
    Dim temp() As Byte
    Dim dataToBeSent
    
    dataToBeSent = vbNullChar & Trim$(gameName) & vbNullChar & "ddddd"
    temp = StrConv(dataToBeSent, vbFromUnicode)
    
    temp(UBound(temp) - 4) = &H0
    temp(UBound(temp) - 3) = &HFF
    temp(UBound(temp) - 2) = &HFF
    temp(UBound(temp) - 1) = &HFF
    temp(UBound(temp)) = &HFF
    
    Call CopyMemory(packet(0), temp(0), ByVal UBound(temp) + 1)

    Call constructPacket(UBound(temp) + 1, &HA)
    
    If rSwitch = False Then Call Form1.btnToggle_Click
End Sub

'0x09
Sub clientKeepAlive()
    packet(0) = &H0
    Call constructPacket(1, &H9)
End Sub
    
Sub constructPacket(pLen As Long, pType As Byte)
    Dim i As Long
    Dim j As Long
    Dim num As Byte
    Dim sizeP As Long
    Dim mySize As Long
    Dim temp() As Byte
    Dim globalPacket() As Byte
    
    If inServer = False And pType <> &H6 And pType <> &H3 And pType <> &H1 Then
        Exit Sub
    End If
    
    'increase message count each send
    msgCount = msgCount + 1
    
    'overflow watch
    If msgCount = 65536 Then msgCount = 0
'*****************************************************************************
    'num of messages in packet
    If msgCount >= 2 Then
        num = &H3
    ElseIf msgCount = 1 Then
        num = &H2
    Else
        num = &H1
    End If
    
'***************************************************************************
    For i = (num - 1) To 0 Step -1
        If i = 0 Then
            'message number
            temp = LongToByteArray(msgCount)
            myPackets(0).packet(0) = temp(0)
            myPackets(0).packet(1) = temp(1)
            'size of packet [includes message type]
            temp = LongToByteArray(pLen + 1)
            myPackets(0).packet(2) = temp(0)
            myPackets(0).packet(3) = temp(1)
            'message type
            myPackets(0).packet(4) = pType
        
            Call CopyMemory(myPackets(0).packet(5), packet(0), ByVal pLen)

            sizeP = sizeP + pLen + 5
        Else
            'store data
            mySize = myPackets(i - 1).packet(3) * CLng(256) + myPackets(i - 1).packet(2)
            Call CopyMemory(myPackets(i).packet(0), myPackets(i - 1).packet(0), ByVal (mySize + 4))
            sizeP = sizeP + mySize + 4
        End If
    Next i
    
    sizeP = sizeP + 1
    ReDim globalPacket(0 To sizeP)
    
    If msgCount >= 2 Then
        globalPacket(0) = &H3
    Else
        globalPacket(0) = num
    End If
             
    j = 1
    For i = 0 To globalPacket(0) - 1
        mySize = myPackets(i).packet(3) * CLng(256) + myPackets(i).packet(2)
        Call CopyMemory(globalPacket(j), myPackets(i).packet(0), ByVal (mySize + 4))
        j = j + mySize + 4
    Next i

    frmServerlist.Winsock1.SendData globalPacket
End Sub


'0x03
Sub userLoginInformation(nick As String)
    Dim temp() As Byte
    Dim connection As Byte
    Dim dataToBeSent As String
    
    dataToBeSent = Trim$(nick) & vbNullChar & Trim$(emulatorPass) & vbNullChar & "d" 'extra byte so we can insert connection type
    'convert to byte array
    temp = StrConv(dataToBeSent, vbFromUnicode)
    
    'connection type
    If frmServerlist.cmbConnectionType.Text = "LAN" Then
        connection = &H1
    ElseIf frmServerlist.cmbConnectionType.Text = "Excellent" Then
        connection = &H2
    ElseIf frmServerlist.cmbConnectionType.Text = "Good" Then
        connection = &H3
    ElseIf frmServerlist.cmbConnectionType.Text = "Average" Then
        connection = &H4
    ElseIf frmServerlist.cmbConnectionType.Text = "Low" Then
        connection = &H5
    ElseIf frmServerlist.cmbConnectionType.Text = "Bad" Then
        connection = &H6
    End If

    temp(UBound(temp)) = connection
    
    Call CopyMemory(packet(0), temp(0), ByVal UBound(temp) + 1)

    Call constructPacket(UBound(temp) + 1, &H3)

End Sub


'0x10
Sub closeGameNotification(msgSlot As Byte)
    Dim w As ListItem
    Dim gameID As Double 'dword
    Dim i As Long
    Dim pos As Long
    
    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    'message(0)=0
    gameID = BytesToNumEx(serverMessages(msgSlot).msgData, 1, 4, True)
    'remove from gamelist
        
        For i = 1 To Form1.lstGamelist.ListItems.count
            If StrComp(Form1.lstGamelist.ListItems(i).SubItems(5), gameID, vbBinaryCompare) = 0 Then
                 pos = i
                 Exit For
            ElseIf i = Form1.lstGamelist.ListItems.count Then
                Exit Sub
            End If
        Next i
        
            If frmPreferences.chkShowOpen.Value = vbChecked Then
                Form1.txtChatroom.SelColor = &H800000
                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
                If frmPreferences.chkTimeStamps.Value = vbChecked Then
                    Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
                End If
                Form1.txtChatroom.SelColor = &HFF&
                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Form1.lstGamelist.ListItems(pos).SubItems(2)
                Form1.txtChatroom.SelColor = &H800000
                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & " [Closed Game]" & vbCrLf
                Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
            End If

        Close #3
        Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
            Print #3, "-:" & Time & ": " & Form1.lstGamelist.ListItems(pos).SubItems(2) & " [Closed Game]"
        Close #3
        
        'display
        Form1.lstGamelist.ListItems.remove (pos)

        Dim k As Long
        For k = LBound(arGames) To UBound(arGames)
            If arGames(k).opened = True And arGames(k).gameID = gameID Then
                arGames(k).opened = False
                Exit For
            End If
        Next k
    
        'gameCount = gameCount - 1
        MDIForm1.StatusBar1.Panels(2).Text = "Games: " & Form1.lstGamelist.ListItems.count
        
        If myGameId = gameID Then
            myGameId = -1
            Form1.fRoomList.Caption = "Not in a game!"
            Form1.fGameroom.Caption = "Not in a game!"
            If rSwitch = True Then Call Form1.btnToggle_Click
            leftRoom = True
            inRoom = False
            imOwner = False
            Form1.fGameroom.Caption = "Not in a room!"
            Form1.lstGameUserlist.ListItems.Clear
            Form1.txtGameChat.Text = vbNullString
            'Form1.txtGameChatroom.Text = vbNullString
        End If
End Sub






'0x0F
Sub kickRequest(userID As String)
    Dim temp() As Byte
    
        
    packet(0) = &H0

    temp = LongToByteArray(CLng(userID))
    packet(1) = temp(0)
    packet(2) = temp(1)
   
    Call constructPacket(3, &HF)
End Sub


Public Sub addUsers(id As Long, aliases As String)
    Dim q As ListItem
    Dim str() As String
    Dim newStr As String
    Dim i As Long
    Dim w As Long
    On Error Resume Next


        If aliases = vbNullString Then
            Set q = frmUserlist.lstUserlist.ListItems.Add(, , arUsers(id).name)
            q.ForeColor = vbBlack
            q.ListSubItems(1).ForeColor = vbBlack
            q.ListSubItems(2).ForeColor = vbBlack
            q.ListSubItems(3).ForeColor = vbBlack
            q.ListSubItems(4).ForeColor = vbBlack
            q.ListSubItems(5).ForeColor = vbBlack
            q.ListSubItems(6).ForeColor = vbBlack
        Else
            str = Split(aliases, ",")
            For w = 0 To UBound(str)
                If LCase$(arUsers(id).name) <> LCase$(str(w)) Then
                    newStr = newStr & ", " & str(w)
                End If
            Next w
            newStr = Mid$(newStr, 3, Len(newStr) - 2)
            'aliases = Replace$(aliases, ",", ", ", 1, -1, vbTextCompare)
            Set q = frmUserlist.lstUserlist.ListItems.Add(, , arUsers(id).name & " [is] " & newStr)
        End If


        q.SubItems(1) = arUsers(id).access
    
        q.SubItems(2) = arUsers(id).ip
        q.SubItems(3) = arUsers(id).userID
        q.ListSubItems(3).Tag = Format(arUsers(id).userID, "00000")
        q.SubItems(4) = arUsers(id).ping
        If Len(q.SubItems(4)) = 2 Then
            q.SubItems(4) = "0" & q.SubItems(4)
        ElseIf Len(q.SubItems(4)) = 1 Then
            q.SubItems(4) = "00" & q.SubItems(4)
        End If
        q.ListSubItems(4).Tag = Format(q.SubItems(4), "000")
        q.SubItems(5) = statusCheck(0, arUsers(id).connection)
        q.SubItems(6) = statusCheck(2, arUsers(id).status)
        q.SubItems(7) = "null"
        
        'set appropriate colors
        If LCase$(arUsers(id).access) = "admin" Then
            q.ForeColor = &H8000&
            q.ListSubItems(1).ForeColor = &H8000&
            q.ListSubItems(2).ForeColor = &H8000&
            q.ListSubItems(3).ForeColor = &H8000&
            q.ListSubItems(4).ForeColor = &H8000&
            q.ListSubItems(5).ForeColor = &H8000&
            q.ListSubItems(6).ForeColor = &H8000&
            Form1.txtChatroom.SelColor = &H8000&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "LOCAL ALERT: Admin is connected: " & arUsers(id).name & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        ElseIf LCase$(arUsers(id).access) = "superadmin" Then
            q.ForeColor = &H8000&
            Form1.txtChatroom.SelColor = &H8000&
            q.ListSubItems(1).ForeColor = &H8000&
            q.ListSubItems(2).ForeColor = &H8000&
            q.ListSubItems(3).ForeColor = &H8000&
            q.ListSubItems(4).ForeColor = &H8000&
            q.ListSubItems(5).ForeColor = &H8000&
            q.ListSubItems(6).ForeColor = &H8000&
            Form1.txtChatroom.SelColor = &H8000&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "LOCAL ALERT: Super Admin is connected: " & arUsers(id).name & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        ElseIf LCase$(arUsers(id).access) = "elevated" Then
            q.ForeColor = &H4080&
            q.ListSubItems(1).ForeColor = &H4080&
            q.ListSubItems(2).ForeColor = &H4080&
            q.ListSubItems(3).ForeColor = &H4080&
            q.ListSubItems(4).ForeColor = &H4080&
            q.ListSubItems(5).ForeColor = &H4080&
            q.ListSubItems(6).ForeColor = &H4080&
            Form1.txtChatroom.SelColor = &H4080&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "LOCAL ALERT: Elevated User is connected: " & arUsers(id).name & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        ElseIf LCase$(arUsers(id).access) = "moderator" Then
            q.ForeColor = &H8080&
            q.ListSubItems(1).ForeColor = &H8080&
            q.ListSubItems(2).ForeColor = &H8080&
            q.ListSubItems(3).ForeColor = &H8080&
            q.ListSubItems(4).ForeColor = &H8080&
            q.ListSubItems(5).ForeColor = &H8080&
            q.ListSubItems(6).ForeColor = &H8080&
            Form1.txtChatroom.SelColor = &H8080&
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "LOCAL ALERT: Moderator is connected: " & arUsers(id).name & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        ElseIf aliases <> vbNullString Then
            q.ForeColor = vbRed
            q.ListSubItems(1).ForeColor = vbRed
            q.ListSubItems(2).ForeColor = vbRed
            q.ListSubItems(3).ForeColor = vbRed
            q.ListSubItems(4).ForeColor = vbRed
            q.ListSubItems(5).ForeColor = vbRed
            q.ListSubItems(6).ForeColor = vbRed
        End If
End Sub

Public Sub addGames(ByVal id As Long)
    Dim q As ListItem
    
    Set q = Form1.lstGamelist.ListItems.Add(, , arGames(id).game)
        
    q.SubItems(1) = arGames(id).emulator
    q.SubItems(2) = arGames(id).owner
    q.SubItems(3) = statusCheck(1, arGames(id).status)
    q.SubItems(4) = arGames(id).Users
    q.SubItems(5) = arGames(id).gameID
    q.ListSubItems(5).Tag = Format(id, "00000")
    q.SubItems(6) = "null"
End Sub


'0x17
Sub serverInformationMessage(msgSlot As Byte)
    Dim i As Long
    Dim num As Long
    Dim count As Long
    Dim q As ListItem
    Dim server As String
    Dim msg As String
    Dim str() As String
    Dim str1 As String
    Dim userInfo() As String
    Dim gameInfo() As String
    Dim strAlert As String
    Dim strTemp As String
    Dim sip() As String
    Dim temp(0 To 511) As Byte
    Dim lMsg As String
    Dim strSize As Long
    On Error Resume Next

    Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    
    'for server
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    server = ByteArrayToString(temp)
    i = i + strSize + 1
    
    'for message
    strSize = lstrlen(VarPtr(serverMessages(msgSlot).msgData(i)))
    If strSize > 511 Then
        strSize = 511
    End If
    Call CopyMemory(temp(0), serverMessages(msgSlot).msgData(i), ByVal strSize)
    temp(strSize) = &H0
    msg = ByteArrayToString(temp)
    i = i + strSize + 1

    Dim k As Long
    Dim t As Long
    Dim otherNames As String
    
    lMsg = LCase$(msg)
    ':USERINFO=userID1,userIP1,userAccessLevel1;
    'Debug.Print msg
    If StrComp(Left$(lMsg, Len(":userinfo=")), ":userinfo=", vbBinaryCompare) = 0 Then
        str = Split(Mid$(msg, Len(":userinfo=") + 1, Len(msg) - Len(":userinfo=")), ChrW$(&H3))
        t = LBound(arUsers)
        For i = 0 To UBound(str)
            If LenB(str(i)) < 1 Then Exit For
            'userInfo(0) = User ID
            'userInfo(1) = IP Address
            'userInfo(2) = Access Level
            'userInfo(3) = Nick
            'userInfo(4) = Ping
            'userInfo(5) = Status
            'userInfo(6) = Connection Type
            userInfo = Split(str(i), ChrW$(&H2))
            sip = Split(userInfo(1), ".")
            If Len(sip(0)) = 2 Then
                userInfo(1) = "0" & userInfo(1)
            ElseIf Len(sip(0)) = 1 Then
                userInfo(1) = "00" & userInfo(1)
            End If
            
            For k = t To UBound(arUsers)
                If arUsers(k).loggedIn = False Then
                    arUsers(k).userID = Trim$(userInfo(0))
                    arUsers(k).ip = Trim$(userInfo(1))
                    arUsers(k).head = sip(0)
                    arUsers(k).sect1 = sip(1)
                    arUsers(k).access = Trim$(userInfo(2))
                    arUsers(k).name = userInfo(3)
                    arUsers(k).ping = Trim$(userInfo(4))
                    arUsers(k).status = Trim$(userInfo(5))
                    arUsers(k).connection = Trim$(userInfo(6))
                    arUsers(k).loggedIn = True
                    If myUserId <> -1 Then
                        otherNames = checkUsers(k, arUsers(k).head, arUsers(k).sect1)
                        Call addUsers(k, otherNames)
                        If LenB(otherNames) <> 0 Then
                            If frmPreferences.chkAlertOthers.Value = vbChecked And StrComp(LCase$(userInfo(2)), "normal", vbBinaryCompare) = 0 Then
                                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                    Call splitReg("ALERT: " & userInfo(3) & " [is] " & otherNames)
                                Else
                                    Call splitAnnounce("ALERT: " & userInfo(3) & " [is] " & otherNames)
                                End If
                            Else
                                Form1.txtChatroom.SelColor = &HC000C0
                                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "LOCAL ALERT: " & userInfo(3) & " [is] " & otherNames & vbCrLf
                                Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
                            End If
                        End If
                    End If
                    t = k
                    Exit For
                End If
            Next k
            
        Next i
                
        MDIForm1.StatusBar1.Panels(1).Text = "Users: " & frmUserlist.lstUserlist.ListItems.count
                
        If UBound(str) = 0 Then
                Dim aPos As Long
                'Hit List
                If frmAdminBot.chkBanIP.Value = vbChecked And myBot.botStatus = True Then
                        Set q = frmAdminBot.lstBanIP.FindItem(userInfo(1), lvwText, 1, lvwWhole)
                        If q Is Nothing Then
                            For i = 1 To frmAdminBot.lstBanIP.ListItems.count
                                aPos = InStr(1, frmAdminBot.lstBanIP.ListItems(i).Text, "*", vbTextCompare)
                                If aPos > 0 Then
                                    If Left$(userInfo(1), aPos - 1) = Left$(frmAdminBot.lstBanIP.ListItems(i).Text, aPos - 1) Then
                                        If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                            Call splitReg(userInfo(3) & "->" & frmAdminBot.lstBanIP.ListItems(i).SubItems(2) & " [" & frmAdminBot.lstBanIP.ListItems(i).SubItems(4) & "]")
                                        Else
                                            Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & userInfo(3) & "->" & frmAdminBot.lstBanIP.ListItems(i).SubItems(2) & " [" & frmAdminBot.lstBanIP.ListItems(i).SubItems(4) & "]")
                                        End If
                                    
                                        'If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                        '    Call splitReg(userInfo(3) & "->" & q.SubItems(2) & " [" & q.SubItems(4) & "]")
                                        'Else
                                        '    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & userInfo(3) & "->" & q.SubItems(2) & " [" & q.SubItems(4) & "]")
                                        'End If
    
                                        If frmAdminBot.lstBanIP.ListItems(i).SubItems(2) = "Silence" Then
                                            Call globalChatRequest("/silence " & userInfo(0) & " " & frmAdminBot.lstBanIP.ListItems(i).SubItems(3))
                                        Else
                                            Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.lstBanIP.ListItems(i).SubItems(3))
                                        End If
                                
                                        Call addDamage(userInfo(3), userInfo(1), "Hit List")
                                        Exit Sub
                                    End If
                                End If
                            Next i
                        Else
                            If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                Call splitReg(userInfo(3) & "->" & q.SubItems(2) & " [" & q.SubItems(4) & "]")
                            Else
                                Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & userInfo(3) & "->" & q.SubItems(2) & " [" & q.SubItems(4) & "]")
                            End If
                                
                            If q.SubItems(2) = "Silence" Then
                                Call globalChatRequest("/silence " & userInfo(0) & " " & q.SubItems(3))
                            Else
                                Call globalChatRequest("/ban " & userInfo(0) & " " & q.SubItems(3))
                            End If
                                
                            Call addDamage(userInfo(3), userInfo(1), "Hit List")
                            Exit Sub
                        End If
                End If
               
                'Login Spamming
                If frmAdminBot.chkLoginSpamControl.Value = vbChecked And myBot.botStatus = True Then
                    If StrComp(LCase$(userInfo(2)), "normal", vbBinaryCompare) = 0 Then
                        If StrComp(myBot.loginIP, userInfo(1), vbBinaryCompare) <> 0 Then
                            myBot.loginCount = 0
                            myBot.loginTimeout = 0
                        End If
                        myBot.loginIP = userInfo(1)
                        myBot.loginCount = myBot.loginCount + 1
                        If myBot.loginCount >= CLng(frmAdminBot.txtLoginNum.Text) Then
                            If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                Call splitReg(userInfo(3) & " [" & frmAdminBot.txtLoginMessage.Text & "]")
                            Else
                                Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & userInfo(3) & " [" & frmAdminBot.txtLoginMessage.Text & "]")
                            End If
                            
                            If frmAdminBot.chkLoginBan.Value = vbChecked Then
                                Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtLoginMin.Text)
                            ElseIf frmAdminBot.chkLoginKick = vbChecked Then
                                Call globalChatRequest("/kick " & userInfo(0))
                            Else
                                Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtLoginMin.Text)
                                frmAdminBot.chkLoginBan.Value = vbChecked
                            End If
                            
                            Call addDamage(userInfo(3), userInfo(1), "Login Spamming")
                            myBot.loginCount = 0
                        End If
                                               
                        'check for sameip
                        For i = LBound(arUsers) To UBound(arUsers)
                            If arUsers(i).loggedIn = True Then
                                If StrComp(arUsers(i).ip, userInfo(1), vbBinaryCompare) = 0 Then
                                    count = count + 1
                                End If
                                
                                If count > CLng(frmAdminBot.txtLoginSameIP.Text) Then
                                
                                    If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                        Call splitReg("Only " & CLng(frmAdminBot.txtLoginSameIP.Text) & " connections from the same address is allowed !!!->" & userInfo(3))
                                    Else
                                        Call splitAnnounce(frmAdminBot.txtBotName.Text & ": Only " & CLng(frmAdminBot.txtLoginSameIP.Text) & " connections from the same address is allowed !!!->" & userInfo(3))
                                    End If
                                        
                                    If frmAdminBot.chkLoginBan.Value = vbChecked Then
                                        Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtLoginMin.Text)
                                    ElseIf frmAdminBot.chkLoginKick = vbChecked Then
                                        Call globalChatRequest("/kick " & userInfo(0))
                                    Else
                                        Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtLoginMin.Text)
                                        frmAdminBot.chkLoginBan.Value = vbChecked
                                    End If
                    
                                    Call addDamage(userInfo(3), userInfo(1), "SAME IP LOGIN MAX")
                                    Exit Sub
                                End If
                            End If
                        Next i
                    End If
                End If
        
                Dim choice As Byte
                Dim bfound As Boolean
                'Username Filter
                If frmAdminBot.chkUserNameFilter.Value = vbChecked And myBot.botStatus = True Then
                    If frmAdminBot.lstUsername.ListCount < 0 Then Exit Sub
                    If InStr(1, userInfo(3), ChrW$(160), vbBinaryCompare) > 0 Then
                        choice = 1
                        bfound = True
                    End If
                    
                    For i = 0 To frmAdminBot.lstUsername.ListCount - 1
                        If Right$(frmAdminBot.lstUsername.List(i), 1) = "*" Then
                            If InStr(1, userInfo(3), Left$(frmAdminBot.lstUsername.List(i), Len(frmAdminBot.lstUsername.List(i)) - 1), vbTextCompare) > 0 Then
                                choice = 1
                            End If
                        Else
                            If StrComp(userInfo(3), frmAdminBot.lstUsername.List(i), vbTextCompare) = 0 Then
                                choice = 1
                            End If
                        End If
                        
                        If choice = 1 Then
                            If StrComp(LCase$(userInfo(2)), "normal", vbBinaryCompare) <> 0 Then
                                Exit For
                            Else
                                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                                    Call splitReg(userInfo(3) & " [" & frmAdminBot.txtUsernameMessage.Text & "]")
                                Else
                                    Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & userInfo(3) & " [" & frmAdminBot.txtUsernameMessage.Text & "]")
                                End If
                                
                                If frmAdminBot.chkUsernameBan.Value = vbChecked Then
                                    Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtUsernameMin.Text)
                                ElseIf frmAdminBot.chkUsernameKick = vbChecked Then
                                    Call globalChatRequest("/kick " & userInfo(0))
                                Else
                                    Call globalChatRequest("/ban " & userInfo(0) & " " & frmAdminBot.txtUsernameMin.Text)
                                    frmAdminBot.chkUsernameBan.Value = vbChecked
                                End If
    
                                If bfound = True Then
                                    Call addDamage(userInfo(3), userInfo(1), "Username Filter: Ascii Character 160")
                                Else
                                    Call addDamage(userInfo(3), userInfo(1), "Username Filter: " & userInfo(3))
                                End If
                                Exit Sub
                            End If
                        End If
                    Next i
                End If
                
                
                'Welcome Messages
                If frmAdminBot.chkWelcomeMessages.Value = vbChecked And myBot.botStatus = True Then
                    For i = 1 To frmAdminBot.lstWelcomeMessages.ListItems.count
                        If StrComp(userInfo(1), frmAdminBot.lstWelcomeMessages.ListItems(i).Text, vbBinaryCompare) = 0 Then
                            Call splitAnnounce(frmAdminBot.lstWelcomeMessages.ListItems(i).SubItems(2))
                            Exit For
                        End If
                    Next i
                End If
                
                
                'Silence Users who enter
                If frmMassive.chkSilenceAll.Value = vbChecked Then
                    If StrComp(LCase$(userInfo(2)), "normal", vbBinaryCompare) = 0 Then
                        If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                            Call splitReg(userInfo(3) & ": The server is currently silencing every user who joins.")
                        Else
                            Call splitAnnounce(userInfo(3) & ": The server is currently silencing every user who joins.")
                        End If
                        Call globalChatRequest("/silence " & userInfo(0) & " " & frmMassive.txtMin.Text)
                    End If
                End If
        Else
            If frmPreferences.chkRoomOnConnect.Value = vbChecked And frmPreferences.txtRoomOnConnect.Text <> vbNullString Then
                Call createGameRequest(frmPreferences.txtRoomOnConnect.Text)
                imOwner = True
                myGame = frmPreferences.txtRoomOnConnect.Text
                Form1.fRoomList.Caption = "Currently in: " & frmPreferences.txtRoomOnConnect.Text
                Form1.fGameroom.Caption = frmPreferences.txtRoomOnConnect.Text
                If rSwitch = False Then Call Form1.btnToggle_Click
            End If
        End If
    ElseIf StrComp(Left$(lMsg, Len(":alivecheck")), ":alivecheck", vbBinaryCompare) = 0 Then
        'just to make sure we are still alive.
    ':ACCESS=userLevel
    ElseIf StrComp(Left$(lMsg, Len(":access=")), ":access=", vbBinaryCompare) = 0 Then
        If Mid$(lMsg, Len(":access=") + 1, Len(msg) - Len(":access=")) = "superadmin" Then 'Or Mid$(lMsg, Len(":access=") + 1, Len(msg) - Len(":access=")) = "admin" Then
            MDIForm1.StatusBar1.Panels(5).Text = "Access: SuperAdmin"
            If frmPreferences.chkStartBot.Value = vbChecked Then
                'Start/Stop Bot
                If myBot.botStatus = False Then Call frmAdminBot.btnONOFF_Click
            End If
            adminFeatures = True
            wasAdmin = True
        ElseIf Mid$(lMsg, Len(":access=") + 1, Len(msg) - Len(":access=")) = "admin" Then
            MDIForm1.StatusBar1.Panels(5).Text = "Access: Admin"
            adminFeatures = True
            wasAdmin = True
        Else
            MDIForm1.StatusBar1.Panels(5).Text = "Access Level: < Admin"
            adminFeatures = False
            wasAdmin = False
        End If
    Else
        
        Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
            Print #3, "[CHAT - " & Time & "]: <" & server & "> " & msg
        Close #3
               
        'display
        If InStr(1, lMsg, "wants to play: ", vbBinaryCompare) > 0 Then
            Exit Sub
        ElseIf InStr(1, lMsg, "created game: ", vbBinaryCompare) > 0 Then
            Exit Sub
        ElseIf InStr(1, lMsg, "loaded the game: ", vbBinaryCompare) > 0 Then
            Exit Sub
        ElseIf InStr(1, lMsg, "has loaded ", vbBinaryCompare) > 0 Then
            Exit Sub
        ElseIf InStr(1, lMsg, "TO:", vbTextCompare) = 0 And InStr(1, lMsg, "> (", vbBinaryCompare) > 0 And InStr(1, lMsg, "):", vbBinaryCompare) > 0 Then
            Form1.txtChatroom.SelColor = &H800000
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
            If frmPreferences.chkTimeStamps.Value = vbChecked Then
                Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
            End If
            Form1.txtChatroom.SelColor = &H800000
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "[PM] "
            Form1.txtChatroom.SelColor = &HFF8000
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & msg & vbCrLf
            Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
            Exit Sub
        End If
        
        Form1.txtChatroom.SelColor = &H800000
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "-:"
        If frmPreferences.chkTimeStamps.Value = vbChecked Then
            Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & Time & ": "
        End If
        Form1.txtChatroom.SelColor = &HC000C0
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & msg & vbCrLf
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    End If
End Sub














