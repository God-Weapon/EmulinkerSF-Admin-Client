Attribute VB_Name = "Other"
Option Explicit


Private Type CtrlProportions
    HeightProportions As Single
    WidthProportions As Single
    TopProportions As Single
    LeftProportions As Single
End Type

Public form1Array() As CtrlProportions
Public frmUserlistArray() As CtrlProportions
Public frmDBVArray() As CtrlProportions

Public dbPos As Long


Public Sub weLoggedIn()
        'Chatroom
        If wasAdmin = False Then
            Call frmServerlist.btnExit_Click
            Call MsgBox("You are NOT an Admin. Application is shutting down!", vbOKOnly, "Admin Alert!")
            Unload MDIForm1
        End If
        Form1.Caption = "Connected to: " & frmServerlist.txtServerIp.Text & " " & serverName
        frmServerlist.List1.AddItem ":-" & Time & ": Connected!"
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        Form1.Visible = True
        frmUserlist.Visible = True
        frmUserlist.Left = Form1.Width + 200
        frmUserlist.Top = 0
        Form1.Top = 0
        Form1.Left = 0
        frmServerlist.WindowState = vbMinimized
        Call fixFramesButtons(2)
        MDIForm1.mnuAsciiArt.Enabled = True
        MDIForm1.mnuTriviaCommands.Enabled = True
        MDIForm1.mnuMassiveCommands.Enabled = True
        MDIForm1.mnuChatroom.Enabled = True
        MDIForm1.mnuUserList.Enabled = True
        MDIForm1.mnuReconnectToServer.Enabled = True
        MDIForm1.mnuLogOffServer.Enabled = True
        MDIForm1.mnuRemoteControl.Enabled = True
        MDIForm1.Toolbar1.Buttons(2).Enabled = True
        MDIForm1.Toolbar1.Buttons(3).Enabled = True
        MDIForm1.Toolbar1.Buttons(6).Enabled = True
        MDIForm1.Toolbar1.Buttons(7).Enabled = True
        If frmPreferences.chkRoomOnConnect.Value = vbChecked And frmPreferences.txtRoomOnConnect.Text <> vbNullString Then
            Call createGameRequest(frmPreferences.txtRoomOnConnect.Text)
            imOwner = True
            myGame = frmPreferences.txtRoomOnConnect.Text
            Form1.fRoomList.Caption = "Currently in: " & frmPreferences.txtRoomOnConnect.Text
            Form1.fGameroom.Caption = frmPreferences.txtRoomOnConnect.Text
            If rSwitch = True Then Call Form1.btnToggle_Click
        End If
End Sub


Public Sub sortRoomlist()
    Dim temp1, temp2, temp3, temp4, temp5 As String
    Dim i, w As Long
    
    For i = 1 To Form1.lstGamelist.ListItems.count
        If Form1.lstGamelist.ListItems(i).SubItems(3) <> "Waiting" Then Exit For
        For w = 1 To Form1.lstGamelist.ListItems.count
            If Form1.lstGamelist.ListItems(w).SubItems(3) <> "Waiting" Then Exit For
            If CLng(Form1.lstGamelist.ListItems(w).SubItems(5)) < CLng(Form1.lstGamelist.ListItems(i).SubItems(5)) Then
                temp1 = Form1.lstGamelist.ListItems(i).Text
                temp2 = Form1.lstGamelist.ListItems(i).SubItems(1)
                temp3 = Form1.lstGamelist.ListItems(i).SubItems(2)
                temp4 = Form1.lstGamelist.ListItems(i).SubItems(4)
                temp5 = Form1.lstGamelist.ListItems(i).SubItems(5)
                
                Form1.lstGamelist.ListItems(i).Text = Form1.lstGamelist.ListItems(w).Text
                Form1.lstGamelist.ListItems(i).SubItems(1) = Form1.lstGamelist.ListItems(w).SubItems(1)
                Form1.lstGamelist.ListItems(i).SubItems(2) = Form1.lstGamelist.ListItems(w).SubItems(2)
                Form1.lstGamelist.ListItems(i).SubItems(5) = Form1.lstGamelist.ListItems(w).SubItems(5)
                Form1.lstGamelist.ListItems(i).SubItems(4) = Form1.lstGamelist.ListItems(w).SubItems(4)
                
                Form1.lstGamelist.ListItems(w).Text = temp1
                Form1.lstGamelist.ListItems(w).SubItems(1) = temp2
                Form1.lstGamelist.ListItems(w).SubItems(2) = temp3
                Form1.lstGamelist.ListItems(w).SubItems(4) = temp4
                Form1.lstGamelist.ListItems(w).SubItems(5) = temp5
            End If
        Next w
    Next i
End Sub


Public Function LongToByteArray(ByVal lng As Long) As Byte()
    Dim ByteArray(0 To 3) As Byte
    CopyMemory ByteArray(0), ByVal VarPtr(lng), Len(lng)
    LongToByteArray = ByteArray
End Function

Public Function statusCheck(statusType As Byte, b As Byte) As String
    Dim str As String
    
    Select Case statusType
        Case 0
            If b = 1 Then
                str = "LAN"
            ElseIf b = 2 Then
                str = "Excellent"
            ElseIf b = 3 Then
                str = "Good"
            ElseIf b = 4 Then
                str = "Average"
            ElseIf b = 5 Then
                str = "Low"
            ElseIf b = 6 Then
                str = "Bad"
            End If
            statusCheck = str
            Exit Function
        Case 1
            If b = 0 Then
                str = "Waiting"
            ElseIf b = 2 Then
                str = "Playing"
            ElseIf b = 1 Then
                str = "Netsync"
            End If
            statusCheck = str
            Exit Function
        Case 2
            If b = 0 Then
                str = "Playing"
            ElseIf b = 1 Then
                str = "Idle"
            ElseIf b = 2 Then
                str = "Netsync"
            End If
            statusCheck = str
            Exit Function
    End Select
    
End Function



Public Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
 End Function
Sub LV_ColumnSort(ByRef oListView As MSComctlLib.ListView, ByRef oColumnHeader As MSComctlLib.ColumnHeader)
'-- Sorts all list items correctly according to data type.
'-- Requirements:
'--     Any items without tag data will be sorted alphabetically.
'--     When creating the list, add a dummy column to the end, width = 0.
'--     Must be the last column in the list.
'--     Create the dummy column subitems as you fill the loop.
'--     Set .Sorted property = True.

    Dim oListItem           As MSComctlLib.ListItem
    Dim i                   As Long
    Dim iTempColIndex       As Long
    Dim bNoTagInColumn      As Boolean
    
    With oListView
    
        '-- If 0 or 1 items or -1(uninitialized), then don't try to sort.
        If .ListItems.count < 2 Then GoTo Exit_Point
        
        iTempColIndex = .ColumnHeaders.count - 1


        '-- Add the tag data from the clicked-on column to the dummy column.
        If oColumnHeader.Index = 1 Then
            '-- First column gets special treatment.
            For i = 1 To .ListItems.count
                Set oListItem = .ListItems(i)
                oListItem.ListSubItems(iTempColIndex) = oListItem.Tag
            Next
            If Len(Trim(oListItem.Tag)) = 0 Then bNoTagInColumn = True
        Else
            '-- Subcolumns.
            For i = 1 To .ListItems.count
                Set oListItem = .ListItems(i)
                oListItem.ListSubItems(iTempColIndex) = oListItem.ListSubItems(oColumnHeader.Index - 1).Tag
            Next
            If Len(Trim(oListItem.ListSubItems(iTempColIndex))) = 0 Then bNoTagInColumn = True
        End If
        
        
        If bNoTagInColumn Then
            '-- If the tag is blank, sort by default - alphabetically.
            .SortKey = oColumnHeader.Index - 1
        Else
            '-- Otherwise sort by the dummy column.
            .SortKey = iTempColIndex
        End If
        
        '-- Sort.
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        
        
        '-- Remove the data so no peeking.
        For i = 1 To .ListItems.count
            Set oListItem = .ListItems(i)
            oListItem.ListSubItems(iTempColIndex) = ""
        Next

    End With
    
Exit_Point:
    Set oListItem = Nothing

End Sub


Sub splitAnnounce(str As String)
    Dim i, w, t, count As Long
    Dim s() As String
    Dim newStr As String
    On Error Resume Next

    s = Split(str, " ")
    If s(0) = "" Then Exit Sub
    For i = 0 To UBound(s)
        If w >= 102 Then
            t = GetTickCount
            Do Until GetTickCount - t >= 20
                DoEvents
            Loop
            Call globalChatRequest("/announce " & Left$(newStr, Len(newStr) - 1))
            w = 0
            newStr = vbNullString
        End If
        w = w + Len(s(i))
        count = count + Len(s(i))
        If count >= CLng(frmPreferences.txtMaxChars.Text) Then
            Call globalChatRequest("/announce " & Left$(newStr, Len(newStr) - 1))
            Exit Sub
        End If
        newStr = newStr & s(i) & " "
    Next i
    Call globalChatRequest("/announce " & Left$(newStr, Len(newStr) - 1))
End Sub

Sub splitReg(str As String)
    Dim i, w, t, count  As Long
    Dim s() As String
    Dim newStr As String
    On Error Resume Next

    s = Split(str, " ")
    For i = 0 To UBound(s)
        If w >= 95 Then
            t = GetTickCount
            Do Until GetTickCount - t >= 20
                DoEvents
            Loop
            Call globalChatRequest(Left$(newStr, Len(newStr) - 1))
            w = 0
            newStr = vbNullString
        End If
        w = w + Len(s(i))
        count = count + Len(s(i))
        If count >= CLng(frmPreferences.txtMaxChars.Text) Then
            Call globalChatRequest(Left$(newStr, Len(newStr) - 1))
            Exit Sub
        End If
        newStr = newStr & s(i) & " "
    Next i
    Call globalChatRequest(Left$(newStr, Len(newStr) - 1))
End Sub

Sub splitRegGame(str As String)
    Dim i, w, t, count  As Long
    Dim s() As String
    Dim newStr As String
    On Error Resume Next

    s = Split(str, " ")
    For i = 0 To UBound(s)
        If w >= 102 Then
            t = GetTickCount
            Do Until GetTickCount - t >= 20
                DoEvents
            Loop
            Call gameChatRequest(Left$(newStr, Len(newStr) - 1))
            w = 0
            newStr = vbNullString
        End If
        w = w + Len(s(i))
        count = count + Len(s(i))
        If count >= CLng(frmPreferences.txtMaxChars.Text) Then
            Call gameChatRequest(Left$(newStr, Len(newStr) - 1))
            Exit Sub
        End If
        newStr = newStr & s(i) & " "
    Next i
    Call gameChatRequest(Left$(newStr, Len(newStr) - 1))
End Sub

Sub reconnectToServer()
    frmServerlist.List1.AddItem "Reconnecting..... (" & Time & ")*"
    frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
    Call frmServerlist.btnLogin_Click
End Sub

Sub fixFramesButtons(itemNum As Byte)
    Dim i As Long
    On Error Resume Next

    With Form1
        Select Case itemNum
            'Initial & Log Out
            Case 0
                frmServerlist.Winsock1.Close
                MDIForm1.mnuAsciiArt.Enabled = False
                MDIForm1.mnuTriviaCommands.Enabled = False
                MDIForm1.mnuMassiveCommands.Enabled = False
                MDIForm1.mnuChatroom.Enabled = False
                MDIForm1.mnuUserList.Enabled = False
                MDIForm1.mnuReconnectToServer.Enabled = False
                MDIForm1.mnuLogOffServer.Enabled = False
                MDIForm1.mnuRemoteControl.Enabled = False
                MDIForm1.Toolbar1.Buttons(2).Enabled = False
                MDIForm1.Toolbar1.Buttons(3).Enabled = False
                MDIForm1.Toolbar1.Buttons(6).Enabled = False
                MDIForm1.Toolbar1.Buttons(7).Enabled = False
                frmServerlist.WindowState = vbNormal
                frmServerlist.Top = 0
                frmServerlist.Left = 0
                frmAbout.Hide
                frmAdminBot.Hide
                .Hide
                frmTrivia.Hide
                frmMassive.Hide
                frmRemote.Hide
                frmAscii.Hide
                frmUserlist.Hide
                'frames
                If rSwitch = True Then Call Form1.btnToggle_Click
                If inServer = True Then
                    Call saveUsers(True)
                    Call saveUsers
                Else
                    For i = LBound(arUsers) To UBound(arUsers)
                        arUsers(i).loggedIn = False
                    Next i
                    For i = LBound(arGames) To UBound(arGames)
                        arGames(i).opened = False
                    Next i
                End If
                MDIForm1.Enabled = True
                
                Call resetBotValues
                'Other
                .Caption = "Server List"
                'close
                inRoom = False
                dbEdit = False
                adminFeatures = False
                .Timer4.Enabled = False
                serversLastMessage = -1
                leftRoom = False
                frmServerlist.btnLogin.Enabled = True
                frmServerlist.txtUsername.Locked = False
                frmServerlist.cmbConnectionType.Enabled = True
                frmServerlist.txtServerIp.Locked = False
                frmServerlist.WindowState = vbNormal
                frmUserlist.lstUserlist.ListItems.Clear
                .lstGamelist.ListItems.Clear
                .lstGameUserlist.ListItems.Clear
                inServer = False
                .btnStealthMode.Caption = "Turn Stealth Mode ON"
                If myBot.botStatus = True Then Call frmAdminBot.btnONOFF_Click
                days = 0
                hours = 0
                myUserId = -1
                myGameId = -1
                timeoutCount = 0
                minutes = 0
                hours = 0
                uClient = 0

                iQuit = False
                frmUserlist.lstUserlist.ListItems.Clear
                .lstGamelist.ListItems.Clear
                .lstGameUserlist.ListItems.Clear
                minutes = 0
                finishedTime = 0
                MDIForm1.StatusBar1.Panels(1).Text = "Users: 0"
                MDIForm1.StatusBar1.Panels(2).Text = "Games: 0"
                MDIForm1.StatusBar1.Panels(3).Text = "Session: " & days & "d:" & hours & "h:" & minutes & "m:" & finishedTime & "s"
                MDIForm1.StatusBar1.Panels(5).Text = "Access Level: Normal "
                MDIForm1.StatusBar1.Panels(6).Text = "Admin Bot: OFF"
                .txtChatroom.Text = vbNullString
                .txtGameChatroom = vbNullString
            'Main Chatroom
            Case 2
                'frames
                If inRoom = True Then
                    .fGameroom.Visible = True
                    .fRoomList.Visible = False
                Else
                    .fGameroom.Visible = False
                    .fRoomList.Visible = True
                End If
            'Gameroom
            Case 3
                'frames
                .fGameroom.Visible = True
                .fRoomList.Visible = False
            'Login Attempt
            Case 7
                myUserId = -1
                myGameId = -1
                step = 1
                msgCount = -1
                portNum = 0
                finishedTime = 0
                minutes = 0
                hours = 0
                days = 0
                frmServerlist.btnLogin.Enabled = False
                frmServerlist.txtUsername.Locked = True
                frmServerlist.txtServerIp.Locked = True
                frmServerlist.btnExit.Enabled = True
                frmServerlist.cmbConnectionType.Enabled = False
            'toggle
            Case 9
                .fGameroom.Visible = False
                .fRoomList.Visible = True
        End Select
    End With
End Sub

Public Function removeNonLetters(str As String) As String
        Dim r(0 To 14) As String
        Dim i As Long
        
        r(0) = "0"
        r(1) = "1"
        r(2) = "2"
        r(3) = "3"
        r(4) = "4"
        r(5) = "5"
        r(6) = "6"
        r(7) = "7"
        r(8) = "8"
        r(9) = "9"
        r(10) = "?"
        r(11) = "!"
        r(12) = "'"
        r(13) = ","
        r(14) = "."
            
        removeNonLetters = str
        For i = 0 To UBound(r)
            removeNonLetters = Replace$(removeNonLetters, r(i), vbNullString)
        Next i
        
End Function

Sub fixBotFrames(num As Byte)
    On Error Resume Next
    Select Case num
        'admin bot
        Case 0
            'visiblity
            frmAdminBot.fAdminMain.Visible = True
            frmAdminBot.fControlChatroom.Visible = False
            frmAdminBot.fControlGames.Visible = False
            frmAdminBot.fAnnouncements.Visible = False
        'announcements
        Case 1
            'visiblity
            frmAdminBot.fAdminMain.Visible = False
            frmAdminBot.fControlChatroom.Visible = False
            frmAdminBot.fControlGames.Visible = False
            frmAdminBot.fAnnouncements.Visible = True
        'control chatroom
        Case 2
            'visiblity
            frmAdminBot.fAdminMain.Visible = False
            frmAdminBot.fControlChatroom.Visible = True
            frmAdminBot.fControlGames.Visible = False
            frmAdminBot.fAnnouncements.Visible = False
        'control games
        Case 3
            'visiblity
            frmAdminBot.fAdminMain.Visible = False
            frmAdminBot.fControlChatroom.Visible = False
            frmAdminBot.fControlGames.Visible = True
            frmAdminBot.fAnnouncements.Visible = False
        End Select
End Sub




Public Sub saveConfig()
    Dim i As Long
    Dim str As String
    Dim strSplit() As String
    
    Close #1
    Open App.Path & "\config.txt" For Output As #1
        Print #1, "ip=" & frmServerlist.txtServerIp.Text
        Print #1, "username=" & frmServerlist.txtUsername.Text
        Print #1, "connection=" & Left$(frmServerlist.cmbConnectionType.Text, 1)
        Print #1, "quit=" & frmPreferences.txtQuit.Text
        Print #1, "maxusers=" & Form1.txtKickUsers.Text
        Print #1, "gameWelcomeMessage=" & frmPreferences.txtGameWelcomeMessage.Text
        Print #1, "beep=" & frmPreferences.chkBeep.Value
        Print #1, "usernameStoreage=" & Replace$(frmServerlist.txtUsernameStoreage.Text, vbCrLf, ";&|")
        Print #1, "connectLoad=" & frmServerlist.chkLoading.Value
        'Print #1, "X=" & MDIForm1.Left
        'Print #1, "Y=" & MDIForm1.Top
        
        Print #1, "timeStamps=" & frmPreferences.chkTimeStamps.Value
        Print #1, "showJoin=" & frmPreferences.chkShowJoin.Value
        Print #1, "showOpen=" & frmPreferences.chkShowOpen.Value
        
        Print #1, "roomConnect=" & frmPreferences.chkRoomOnConnect.Value
        Print #1, "roomConnectName=" & frmPreferences.txtRoomOnConnect.Text
        Print #1, "startbot=" & frmPreferences.chkStartBot.Value
        Print #1, "quickban=" & Form1.txtQuickBan.Text
        Print #1, "alertothers=" & frmPreferences.chkAlertOthers.Value
        
        Print #1, vbNullString
        
        'Bot*********************
        Print #1, "botName=" & frmAdminBot.txtBotName.Text
        Print #1, "spamRow=" & frmAdminBot.txtSpamRow.Text
        Print #1, "allcaps=" & frmAdminBot.chkAllCaps.Value
        Print #1, "spamMin=" & frmAdminBot.txtSpamMin.Text
        Print #1, "spamChars=" & frmAdminBot.txtSpamChars.Text
        Print #1, "spamMessage=" & frmAdminBot.txtSpamMessage.Text
        Print #1, "activateCheck=" & frmAdminBot.chkSpamControl.Value
        Print #1, "announceInterval=" & frmAdminBot.txtAnnounceInterval.Text
        Print #1, "announceInterval2=" & frmAdminBot.txtAnnounceInterval2.Text
        Print #1, "announceInterval3=" & frmAdminBot.txtAnnounceInterval3.Text
        Print #1, "announceMessage=" & Replace$(frmAdminBot.txtAnnounceMessage.Text, vbCrLf, ";&|")
        Print #1, "announceMessage2=" & Replace$(frmAdminBot.txtAnnounceMessage2.Text, vbCrLf, ";&|")
        Print #1, "announceMessage3=" & Replace$(frmAdminBot.txtAnnounceMessage3.Text, vbCrLf, ";&|")

        Print #1, "announceReg=" & frmAdminBot.chkAnnounceReg.Value
        Print #1, "annouonceChatroomCheck=" & frmAdminBot.chkAnnounceChatroom.Value
        Print #1, "gameAnnounceInterval=" & frmAdminBot.txtGameInterval.Text
        Print #1, "gameAnnounceMessage=" & frmAdminBot.txtGameMessage.Text
        Print #1, "announceGameCheck=" & frmAdminBot.chkAnnounceGames.Value
        Print #1, "ALLCAPS=" & frmAdminBot.chkAllCaps.Value
        
        str = vbNullString
        For i = 0 To frmAdminBot.lstWord.ListCount - 1
            str = str & frmAdminBot.lstWord.List(i) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        
        Print #1, "filterWords=" & str
        Print #1, "filterMessage=" & frmAdminBot.txtFilterMessage.Text
        Print #1, "filterCheck=" & frmAdminBot.chkWordFilter.Value
        Print #1, "filterMin=" & frmAdminBot.txtWordMin.Text

        str = vbNullString
        For i = 1 To frmAdminBot.lstGame.ListItems.count
            str = str & frmAdminBot.lstGame.ListItems(i).Text & ";&," & frmAdminBot.lstGame.ListItems(i).SubItems(1) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #1, "controlledgames=" & str
        Print #1, "gameControlCheck=" & frmAdminBot.chkGameControl.Value
        Print #1, "gamespammessage=" & frmAdminBot.txtGameroomMessage.Text
        Print #1, "gameSpamControlCheck=" & frmAdminBot.chkGameSpamControl.Value
        Print #1, "gamespamlimit=" & frmAdminBot.txtGameroomBan.Text
        Print #1, "gamespamconsecutive=" & frmAdminBot.txtGameroomNum.Text
        
        str = vbNullString
        For i = 1 To frmAdminBot.lstWelcomeMessages.ListItems.count
            str = str & frmAdminBot.lstWelcomeMessages.ListItems(i) & ";&," & frmAdminBot.lstWelcomeMessages.ListItems(i).SubItems(1) & ";&," & frmAdminBot.lstWelcomeMessages.ListItems(i).SubItems(2) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #1, "welcomeMessages=" & str
        Print #1, "activateWelcome=" & frmAdminBot.chkWelcomeMessages.Value
        Print #1, "silenceAll=" & frmMassive.chkSilenceAll.Value
        Print #1, "loginSpamControl=" & frmAdminBot.chkLoginSpamControl.Value
        Print #1, "loginNum=" & frmAdminBot.txtLoginNum.Text
        Print #1, "loginMin=" & frmAdminBot.txtLoginMin.Text
        Print #1, "loginMessage=" & frmAdminBot.txtLoginMessage.Text
        
        str = vbNullString
        For i = 0 To frmAdminBot.lstUsername.ListCount - 1
            str = str & frmAdminBot.lstUsername.List(i) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        
        Print #1, "filterUsername=" & str
        Print #1, "usernameCheck=" & frmAdminBot.chkUserNameFilter.Value
        Print #1, "disableHosting=" & frmAdminBot.chkGameDisable.Value
        Print #1, "chkBanIP=" & frmAdminBot.chkBanIP.Value
        
        str = vbNullString
        For i = 1 To frmAdminBot.lstDisableHosting.ListItems.count
            str = str & frmAdminBot.lstDisableHosting.ListItems(i) & ";&," & frmAdminBot.lstDisableHosting.ListItems(i).SubItems(1) & ";&," & frmAdminBot.lstDisableHosting.ListItems(i).SubItems(2) & ";&," & frmAdminBot.lstDisableHosting.ListItems(i).SubItems(3) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #1, "disabledList=" & str

        str = vbNullString
        For i = 1 To frmAdminBot.lstBanIP.ListItems.count
            str = str & frmAdminBot.lstBanIP.ListItems(i).Text & ";&," & frmAdminBot.lstBanIP.ListItems(i).SubItems(1) & ";&," & frmAdminBot.lstBanIP.ListItems(i).SubItems(2) & ";&," & frmAdminBot.lstBanIP.ListItems(i).SubItems(3) & ";&," & frmAdminBot.lstBanIP.ListItems(i).SubItems(4) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #1, "bannedList=" & str
        Print #1, "spamExpire=" & frmAdminBot.txtSpamExpire.Text
        Print #1, "gameExpire=" & frmAdminBot.txtGameExpire.Text
        Print #1, "loginExpire=" & frmAdminBot.txtLoginExpire.Text
        Print #1, "totalCaps=" & frmAdminBot.txtTotalCaps.Text
        
        Print #1, "maximumChars=" & frmPreferences.txtMaxChars.Text
        
        Print #1, "chkWordBan=" & frmAdminBot.chkWordBan.Value
        Print #1, "chkWordKick=" & frmAdminBot.chkWordKick.Value
        Print #1, "chkWordSilence=" & frmAdminBot.chkWordSilence.Value
        Print #1, "chkWordSilenceKick=" & frmAdminBot.chkWordSilenceKick.Value

        Print #1, "chkSpamBan=" & frmAdminBot.chkSpamBan.Value
        Print #1, "chkSpamKick=" & frmAdminBot.chkSpamKick.Value
        Print #1, "chkSpamSilence=" & frmAdminBot.chkSpamSilence.Value
        Print #1, "chkSpamSilenceKick=" & frmAdminBot.chkSpamSilenceKick.Value

        Print #1, "chkLineWrapperBan=" & frmAdminBot.chkLineWrapperBan.Value
        Print #1, "chkLineWrapperKick=" & frmAdminBot.chkLineWrapperKick.Value
        Print #1, "chkLineWrapperSilence=" & frmAdminBot.chkLineWrapperSilence.Value
        Print #1, "chkLineWrapperSilenceKick=" & frmAdminBot.chkLineWrapperSilenceKick.Value
        Print #1, "lineWrapperMessage=" & frmAdminBot.txtLineWrapperMessage.Text
        Print #1, "lineWrapperMin=" & frmAdminBot.txtLineWrapperMin.Text
        Print #1, "maxSpace=" & frmAdminBot.txtMaxSpace.Text
        
        Print #1, "chkAllCapsBan=" & frmAdminBot.chkAllCapsBan.Value
        Print #1, "chkAllCapsKick=" & frmAdminBot.chkAllCapsKick.Value
        Print #1, "chkAllCapsSilence=" & frmAdminBot.chkAllCapsSilence.Value
        Print #1, "chkAllCapsSilenceKick=" & frmAdminBot.chkAllCapsSilenceKick.Value
        Print #1, "allCapsMessage=" & frmAdminBot.txtAllCapsMessage.Text
        Print #1, "allCapsMin=" & frmAdminBot.txtAllCapsMin.Text

        Print #1, "chkLoginBan=" & frmAdminBot.chkLoginBan.Value
        Print #1, "chkLoginKick=" & frmAdminBot.chkLoginKick.Value

        Print #1, "chkUsernameBan=" & frmAdminBot.chkUsernameBan.Value
        Print #1, "chkUsernameKick=" & frmAdminBot.chkUsernameKick.Value
        Print #1, "usernameMessage=" & frmAdminBot.txtUsernameMessage.Text
        Print #1, "usernameMin=" & frmAdminBot.txtUsernameMin.Text


        Print #1, "chkGameSpamBan=" & frmAdminBot.chkGameSpamBan.Value
        Print #1, "chkGameSpamKick=" & frmAdminBot.chkGameSpamKick.Value

        Print #1, "loginSameIP=" & frmAdminBot.txtLoginSameIP.Text
        
        Print #1, "reconnectHours=" & frmPreferences.txtReconnect.Text
        
        Print #1, "chkLineWrapper=" & frmAdminBot.chkLineWrapper.Value
        
        
        
        Print #1, "linkMessage=" & frmAdminBot.txtLinkMessage.Text
        Print #1, "linkMin=" & frmAdminBot.txtLinkMin.Text
        Print #1, "linksInterval=" & frmAdminBot.txtLinksInterval.Text
        Print #1, "chkLinkBan=" & frmAdminBot.chkLinkBan.Value
        Print #1, "chkLinkSilence=" & frmAdminBot.chkLinkSilence.Value
        Print #1, "chkLinkSilenceKick=" & frmAdminBot.chkLinkSilenceKick.Value
        Print #1, "chkLinkKick=" & frmAdminBot.chkLinkKick.Value
        Print #1, "linkSend=" & frmAdminBot.txtLinkSend.Text
        Print #1, "chkLink=" & frmAdminBot.chkLinkControl.Value
        
        Print #1, "announceCreateMessage=" & frmAdminBot.txtCreateGame.Text
        Print #1, "chkCreateGameAnnounce=" & frmAdminBot.chkCreateGame.Value
    Close #1
End Sub





Public Sub resetBotValues()
    Dim i As Long
        myBot.announceChatroomCount = 0
        myBot.announceChatroomCount2 = 0
        myBot.announceChatroomCount3 = 0
        myBot.announceGamesCount = 0
        myBot.loginCount = 0
        myBot.loginIP = vbNullString
        
        For i = LBound(arUsers) To UBound(arUsers)
            arUsers(i).linkSent = 0
            arUsers(i).linkCount = 0
            arUsers(i).spamRowCount = 0
            arUsers(i).spamTimeout = 0
            arUsers(i).gameTimeout = 0
            arUsers(i).gameSpamCount = 0
        Next i
End Sub

Sub textboxStuff(txt As TextBox, ascNum As Integer)
    If ascNum = 1 Then
        ascNum = 0
        txt.SelStart = 0
        txt.SelLength = Len(txt.Text)
    ElseIf ascNum = 13 Or ascNum = 10 Then
        ascNum = 0
    End If
End Sub

Function BytesToNumEx(ByteArray() As Byte, ByVal StartRec As Long, ByVal EndRec As Long, UnSigned As Boolean) As Double
' ###################################################
' Author                : Imran Zaheer
' Contact               : imraanz@mail.com
' Date                  : January 2000
' Function BytesToNumEx : Convertes the specified byte array
'                         into the corresponding Integer or Long
'                         or any signed/unsigned
'                        ;(non-float) data type.
'
' * BYTES : LIKE NUMBERS(Integer/Long etc.) STORED IN A
' * BINARY FILE

' Parameters :
'  (All parameters are reuuired: No Optional)
'     ByteArray() : byte array containg a number in byte format
'  StartRec    : specify the starting array record within the
                 ' array
'     EndRec      : specify the end array record within the array
'     UnSigned    : when False process bytes for both -ve and
'                   +ve values.
'                   when true only process the bytes for +ve
'                   values.
'
' Note: If both "StartRec" and "EndRec" Parameters are zero,
'       then the complete array will be processed.
'
' Example Calls :
'      dim myArray(1 To 4) as byte
'      dim myVar1 as Long
'      dim myVar2 as Long
'
'      myArray(1) = 255
'      myArray(2) = 127
'      myVar1 = BytesToNumEx(myArray(), 1, 2, False)
'  after execution of above statement myVar1 will be 32767
'
'      myArray(1) = 0
'      myArray(2) = 0
'      myArray(3) = 0
'      myArray(4) = 128
'      myVar2 = BytesToNumEx(myArray(), 1, 4, False)
'  after execution of above statement myVar2 will be -2147483648
'
'
'####################################################
On Error GoTo ErrorHandler
Dim i As Long
Dim lng256 As Double
Dim lngReturn As Double
    
    lng256 = 1
    lngReturn = 0
    
    If EndRec < 1 Then
        EndRec = UBound(ByteArray)
    End If
    
    If StartRec > EndRec Or StartRec < 0 Then
        MsgBox _
         "Start record can not be greater then End record...!", _
          vbInformation
        BytesToNumEx = -1
        Exit Function
    End If
    
    lngReturn = lngReturn + (ByteArray(StartRec))
    For i = (StartRec + 1) To EndRec
        lng256 = lng256 * 256
        If i < EndRec Then
            lngReturn = lngReturn + (ByteArray(i) * lng256)
        Else
           ' if -ve

            If ByteArray(i) > 127 And UnSigned = False Then
             lngReturn = (lngReturn + ((ByteArray(i) - 256) _
                  * lng256))
            Else
                lngReturn = lngReturn + (ByteArray(i) * lng256)
            End If
        End If
    Next i
    
    BytesToNumEx = lngReturn
ErrorHandler:
End Function


Public Sub addDamage(nick As String, ip As String, rReason As String)
    Dim lst As ListItem
    Dim i As Long
    
    Set lst = frmAdminBot.lstDamage.ListItems.Add(, , nick)
    lst.SubItems(1) = ip
    lst.SubItems(2) = Time
    lst.SubItems(3) = rReason
    lst.SubItems(4) = "#"
    frmAdminBot.lstDamage.ListItems.Item(frmAdminBot.lstDamage.ListItems.count).EnsureVisible
    frmAdminBot.lstDamage.ListItems.Item(frmAdminBot.lstDamage.ListItems.count).Selected = True
    Close #2
    Open App.Path & "\EmulinkerSF_Logs\bot.txt" For Append As #2
        Print #2, Time & ": " & nick & "; " & ip & "; " & rReason
    Close #2
    
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If arUsers(i).ip = ip Then
                ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).numOfBotHits = ipHead(arUsers(i).head).ipSect(arUsers(i).sect1).myEntries(arUsers(i).dbPos).numOfBotHits + 1
                Exit Sub
            End If
        End If
    Next i

End Sub




Sub InitResizeArray()

    Dim i As Long
    
    On Error Resume Next
    
    ReDim form1Array(0 To Form1.Controls.count - 1)
    
    For i = 0 To Form1.Controls.count - 1
        With form1Array(i)
            .HeightProportions = Form1.Controls(i).Height / Form1.ScaleHeight
            .WidthProportions = Form1.Controls(i).Width / Form1.ScaleWidth
            .TopProportions = Form1.Controls(i).Top / Form1.ScaleHeight
            .LeftProportions = Form1.Controls(i).Left / Form1.ScaleWidth
        End With
    Next i

    ReDim frmUserlistArray(0 To frmUserlist.Controls.count - 1)
    
    For i = 0 To frmUserlist.Controls.count - 1
        With frmUserlistArray(i)
            .HeightProportions = frmUserlist.Controls(i).Height / frmUserlist.ScaleHeight
            .WidthProportions = frmUserlist.Controls(i).Width / frmUserlist.ScaleWidth
            .TopProportions = frmUserlist.Controls(i).Top / frmUserlist.ScaleHeight
            .LeftProportions = frmUserlist.Controls(i).Left / frmUserlist.ScaleWidth
        End With
    Next i

    ReDim frmDBVArray(0 To frmDBV.Controls.count - 1)
    For i = 0 To frmDBV.Controls.count - 1
        With frmDBVArray(i)
            .HeightProportions = frmDBV.Controls(i).Height / frmDBV.ScaleHeight
            .WidthProportions = frmDBV.Controls(i).Width / frmDBV.ScaleWidth
            .TopProportions = frmDBV.Controls(i).Top / frmDBV.ScaleHeight
            .LeftProportions = frmDBV.Controls(i).Left / frmDBV.ScaleWidth
        End With
    Next i
End Sub

Sub ResizeControls(str As String)
    On Error Resume Next
    Dim i As Long
    
    If str = "form1" Then
        For i = 0 To Form1.Controls.count - 1
            With form1Array(i)
                ' move and resize controls
                Form1.Controls(i).Move .LeftProportions * Form1.ScaleWidth, .TopProportions * Form1.ScaleHeight, .WidthProportions * Form1.ScaleWidth, .HeightProportions * Form1.ScaleHeight
            End With
        Next i
    ElseIf str = "frmUserlist" Then
        For i = 0 To frmUserlist.Controls.count - 1
            With frmUserlistArray(i)
                ' move and resize controls
                frmUserlist.Controls(i).Move .LeftProportions * frmUserlist.ScaleWidth, .TopProportions * frmUserlist.ScaleHeight, .WidthProportions * frmUserlist.ScaleWidth, .HeightProportions * frmUserlist.ScaleHeight
            End With
        Next i
    ElseIf str = "frmDBV" Then
        For i = 0 To frmDBV.Controls.count - 1
            With frmDBVArray(i)
                ' move and resize controls
                frmDBV.Controls(i).Move .LeftProportions * frmDBV.ScaleWidth, .TopProportions * frmDBV.ScaleHeight, .WidthProportions * frmDBV.ScaleWidth, .HeightProportions * frmDBV.ScaleHeight
            End With
        Next i
    End If
End Sub



