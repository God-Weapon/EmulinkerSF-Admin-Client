VERSION 5.00
Begin VB.Form frmAdminBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Bot"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   10230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnGames 
      Caption         =   "Control Games"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton btnChatroom 
      Caption         =   "Control Chatroom"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton btnAnnouncements 
      Caption         =   "Announcements"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton btnAdminBot 
      Caption         =   "Admin Bot"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mnuWord 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveWord 
         Caption         =   "Remove Word"
      End
      Begin VB.Menu bar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveWordAll 
         Caption         =   "Remove All"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveGame 
         Caption         =   "Remove Game"
      End
      Begin VB.Menu bar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveGameAll 
         Caption         =   "Remove All"
      End
   End
   Begin VB.Menu mnuDamage 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDamageClearIP 
         Caption         =   "Clear IP"
      End
      Begin VB.Menu bar333 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDamageClear 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu bar44 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveDamageAll 
         Caption         =   "Clear All"
      End
   End
   Begin VB.Menu mnuWelcome 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuWelcomeRemove 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu bar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWelcomeRemoveAll 
         Caption         =   "Remove All"
      End
   End
End
Attribute VB_Name = "frmAdminBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnAddGame_Click()
    Dim lst As ListItem
    
    If txtNumOfGames.Text = vbNullString Then
        txtNumOfGames.Text = "Supply #of games here!"
        Call textboxStuff(txtNumOfGames, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If
    
    If txtGameName.Text = vbNullString Then
        txtGameName.Text = "Supply #of games here!"
        Call textboxStuff(txtGameName, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If

    Set lst = lstGame.ListItems.Add(, , Trim$(txtGameName.Text))
    lst.SubItems(1) = Trim$(txtNumOfGames.Text)

    txtNumOfGames.Text = vbNullString
    txtGameName.Text = vbNullString
End Sub

Private Sub btnAddWord_Click()
    If txtWordFilter.Text <> vbNullString Then
        lstWord.AddItem Trim$(txtWordFilter.Text)
    Else
        txtWordFilter.Text = "Supply word here!"
        Call textboxStuff(txtWordFilter, 1)
        txtWordFilter.SetFocus
    End If
    txtWordFilter.Text = vbNullString
End Sub

Private Sub btnAdminBot_Click()
    Call fixBotFrames(0)
End Sub

Private Sub btnAnnouncements_Click()
    Call fixBotFrames(1)
End Sub

Private Sub btnChatroom_Click()
    Call fixBotFrames(2)
End Sub

Private Sub btnGames_Click()
    Call fixBotFrames(3)
End Sub

Public Sub btnONOFF_Click()
    Static choice As Boolean
    
    If inServer = True And adminFeatures = True Then
        If choice = False Then
            Call btnSet_Click
            btnONOFF.Caption = "Turn Bot OFF"
            fAdminBot.Caption = "Admin Bot: ON"
            Form1.StatusBar1.Panels(7).Text = "Admin Bot: ON"
            Me.Hide
            Form1.lstUserlist.SetFocus
            choice = True
            myBot.status = True
            txtName.Enabled = False
            myBot.name = txtName.Text
            msgCount = msgCount + 1
            Call globalChatRequest("/announce " & myBot.name & ": Activated!")
        Else
            btnONOFF.Caption = "Turn Bot ON"
            choice = False
            myBot.status = False
            myBot.announceStatus = False
            myBot.filterStatus = False
            myBot.gameControlStatus = False
            myBot.announceCount = 0
            myBot.kickStatus = False
            fAutoAnnounce.Caption = "Chat Announce: OFF"
            fGameAnnounce.Caption = "Game Announce: OFF"
            fWelcome.Caption = "Welcome Messages: OFF"
            fGameSpam.Caption = "Game Spam Control: OFF"
            fGameControl.Caption = "Game Control: OFF"
            Form1.StatusBar1.Panels(7).Text = "Admin Bot: OFF"
            fAutoKick.Caption = "Spam Control: OFF"
            fAdminBot.Caption = "Admin Bot: OFF"
            fWordFilter.Caption = "Word Filter Control: OFF"
            myBot.kickCount = 0
            txtName.Enabled = True
            myBot.announceCount = myBot.announceInterval - 1 'initial announcement
            If inServer = True And myBot.status = True And adminFeatures = True Then
                msgCount = msgCount + 1
                Call globalChatRequest("/announce " & myBot.name & ": Deactivated!")
            End If
        End If
    Else
        Call MsgBox("You need to be logged into a Server with Admin Status to use this feature!", vbOKOnly, "Admin Bot Alert!")
        btnONOFF.Caption = "Turn Bot ON"
    End If
End Sub


Private Sub saveChanges()
    Dim str As String
    Dim strSplit() As String
    Dim i As Long
    
     Open App.Path & "\bot.txt" For Output As #5
        Print #5, "#Comments are done with #"
        Print #5, "#Reads from top to bottom"
        
        Print #5, vbNullString
        Print #5, "#Your bot's name"
        Print #5, "botName=" & txtName.Text
        
        Print #5, vbNullString
        Print #5, "#How many messages in a row with out other users typing before kick"
        Print #5, "kickConsecutive=" & txtKickConsecutive.Text
        Print #5, vbNullString
        Print #5, "#silence users who type in all caps"
        Print #5, "allcaps=" & chkAllCaps.Value
        Print #5, vbNullString
        Print #5, "#How many seconds in respect to kick number of messages before kick"
        Print #5, "kickInterval=" & txtKickInterval.Text
        Print #5, vbNullString
        Print #5, "#How many message in respect to kick interval before kick"
        Print #5, "kickMessageNum=" & txtKickMessageNum.Text
        Print #5, vbNullString
        Print #5, "#Your kick message"
        Print #5, "kickMessage=" & txtKickMessage.Text
        Print #5, vbNullString
        Print #5, "kickCheck=" & chkActivateAutoSpamKick.Value
        Print #5, vbNullString
        Print #5, "kicksilence=" & txtKickSilence.Text
               
        Print #5, vbNullString
        Print #5, "#Interval to display announcement"
        Print #5, "announceInterval=" & txtAnnounceInterval.Text
        Print #5, vbNullString
        Print #5, "#Your announcement"
        strSplit = Split(txtAnnounceMessage.Text, vbCrLf)
        If UBound(strSplit) > -1 Then
            For i = 0 To UBound(strSplit)
                If i = UBound(strSplit) Then
                    str = str & strSplit(i)
                Else
                    str = str & strSplit(i) & "\0\0"
                End If
            Next i
        Else
            str = txtAnnounceMessage.Text
        End If
        Print #5, "announceMessage=" & str
        Print #5, vbNullString
        Print #5, "annouonceChatroomCheck=" & chkActivateAutoAnnounceChatroom.Value
        
        Print #5, vbNullString
        Print #5, "#Interval to display game announcement"
        Print #5, "gameAnnounceInterval=" & txtGameInterval.Text
        Print #5, vbNullString
        Print #5, "#Your game announcement"
        Print #5, "gameAnnounceMessage=" & txtGameMessage.Text
        Print #5, vbNullString
        Print #5, "announceGameCheck=" & chkActivateAnnounceGame.Value
        Print #5, vbNullString
        Print #5, "ALLCAPS=" & chkAllCaps.Value
        
        Print #5, vbNullString
        Print #5, "#Words to be filtered"
        str = vbNullString
        For i = 0 To lstWord.ListCount - 1
            str = str & lstWord.List(i) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        
        Print #5, "filterWords=" & str
        Print #5, vbNullString
        Print #5, "#message to display after kick from filter"
        Print #5, "filterMessage=" & txtFilterMessage.Text
        Print #5, vbNullString
        Print #5, "filterCheck=" & chkActivateAutoWordFilter.Value
        Print #5, vbNullString
        Print #5, "kicksBeforeBan=" & txtKicksBeforeBan.Text

        Print #5, vbNullString
        str = vbNullString
        For i = 1 To lstGame.ListItems.Count
            str = str & lstGame.ListItems(i) & ";&," & lstGame.ListItems(i).SubItems(1) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #5, "controlledgames=" & str
        
        
        
        Print #5, vbNullString
        Print #5, "activategamecontrol=" & chkActivateGameControl.Value
        
        Print #5, vbNullString
        Print #5, "gamespammessage=" & txtGameroomMessage.Text
        Print #5, vbNullString
        Print #5, "activategamespam=" & chkActivateGameroomSpamControl.Value
        Print #5, vbNullString
        Print #5, "gamespamlimit=" & txtGameroomBan.Text
        Print #5, vbNullString
        Print #5, "gamespamconsecutive=" & txtGameroomNum.Text
        
        Print #5, vbNullString
        str = vbNullString
        For i = 1 To lstWelcome.ListItems.Count
            str = str & lstWelcome.ListItems(i) & ";&," & lstWelcome.ListItems(i).SubItems(1) & ";&," & lstWelcome.ListItems(i).SubItems(2) & ";&,"
        Next i
        If str <> vbNullString Then str = Left$(str, Len(str) - 3)
        Print #5, "welcomeMessages=" & str
        Print #5, vbNullString
        Print #5, "activateWelcome=" & chkWelcome.Value
        
        Print #5, "gameBan=" & chkGameBan.Value
        Print #5, "gameKick=" & chkGameKick.Value
        Print #5, "silenceAll=" & chkSilenceAll.Value
        Print #5, "loginSpamControl=" & chkLoginSpamControl.Value
        Print #5, "dupeNick=" & chkDupeNick.Value
        Print #5, "wordSilence=" & chkWordSilence.Value
        Print #5, "wordBan=" & chkWordBan.Value
        Print #5, "wordKick=" & chkWordKick.Value
        Print #5, "loginNum=" & txtLoginNum.Text
        Print #5, "loginMin=" & txtLoginMin.Text
        Print #5, "loginMessage=" & txtLoginMessage.Text
        Print #5, "spamSilence=" & chkSpamSilence.Value
        Print #5, "spamBan=" & chkSpamBan.Value
        Print #5, "spamKick=" & chkSpamKick.Value
        Print #5, "silenceAllInterval=" & txtSilenceAll.Text
    Close #5

End Sub


Private Sub btnSet_Click()
    Dim str As String
    Dim strSplit() As String
    Dim i As Long
           
    With myBot
        'bot name
        If .name <> vbNullString Then .name = txtName.Text
        
        'announce
        If txtAnnounceMessage.Text <> vbNullString And txtAnnounceInterval.Text <> vbNullString And chkActivateAutoAnnounceChatroom.Value = vbChecked Then
            .announceMessage = txtAnnounceMessage.Text
            .announceInterval = txtAnnounceInterval.Text
            .announceCount = .announceInterval - 1
            .announceStatus = True
            chkActivateAutoAnnounceChatroom.ForeColor = vbBlack
            fAutoAnnounce.Caption = "Announce: ON"
        Else
            fAutoAnnounce.Caption = "Announce: OFF"
            .announceStatus = False
            chkActivateAutoAnnounceChatroom.Value = vbUnchecked
            If txtAnnounceMessage.Text = vbNullString Or txtAnnounceInterval.Text = vbNullString Then
                chkActivateAutoAnnounceChatroom.ForeColor = vbBlue
            End If
        End If
        
        'announce games
        If txtGameMessage.Text <> vbNullString And txtGameInterval.Text <> vbNullString And chkActivateAnnounceGame.Value = vbChecked Then
            .gameMessage = txtGameMessage.Text
            .gameInterval = txtGameInterval.Text
            .gameCount = .gameInterval - 1
            .gameStatus = True
            chkActivateAnnounceGame.ForeColor = vbBlack
            fGameAnnounce.Caption = "Announce Games: ON"
            If CLng(txtGameInterval.Text) < 10 Then
                txtGameInterval.Text = "10"
                Call txtGameInterval_KeyPress(1)
            End If
        Else
            fGameAnnounce.Caption = "Announce Games: OFF"
            .gameStatus = False
            chkActivateAnnounceGame.Value = vbUnchecked
            If txtGameMessage.Text = vbNullString Or txtGameInterval.Text = vbNullString Then
                chkActivateAnnounceGame.ForeColor = vbBlue
            End If
        End If

         'login spam control
        If txtLoginMin.Text <> vbNullString And txtLoginNum.Text <> vbNullString And txtLoginMessage.Text <> vbNullString And chkLoginSpamControl.Value = vbChecked Then
            .loginSpamConsecutive = CInt(txtLoginNum.Text)
            .loginSpamMessage = txtLoginMessage.Text
            .loginSpamLimit = txtLoginMin.Text
            .loginSpamStatus = True
            chkLoginSpamControl.ForeColor = vbBlack
            fLoginControl.Caption = "Login Spam Control: ON"
        Else
            fLoginControl.Caption = "Login Spam Control: OFF"
            .loginSpamStatus = False
            chkLoginSpamControl.Value = vbUnchecked
            If txtLoginNum.Text <> vbNullString Or txtLoginMin.Text = vbNullString Or txtLoginMessage.Text = vbNullString Then
                chkActivateGameroomSpamControl.ForeColor = vbBlue
            End If
        End If
        
        
        'gameroom spam control
        If txtGameroomNum.Text <> vbNullString And txtGameroomBan.Text <> vbNullString And txtGameroomMessage.Text <> vbNullString And chkActivateGameroomSpamControl.Value = vbChecked Then
        If chkGameBan.Value = vbChecked Or chkGameKick.Value = vbChecked Then
            .gameSpamConsecutive = CInt(txtGameroomNum.Text)
            .gameSpamMessage = txtGameroomMessage.Text
            .gameSpamLimit = txtGameroomBan.Text
            .gameSpamStatus = True
            chkActivateGameroomSpamControl.ForeColor = vbBlack
            chkGameBan.ForeColor = vbBlack
            chkGameKick.ForeColor = vbBlack
            If chkGameBan.Value = vbUnchecked And chkGameKick.Value = vbUnchecked Then chkGameKick.Value = vbUnchecked
            fGameSpam.Caption = "Game Spam Control: ON"
        Else
            fGameSpam.Caption = "Game Spam Control: OFF"
            .gameSpamStatus = False
            chkActivateGameroomSpamControl.Value = vbUnchecked
            If txtGameroomNum.Text <> vbNullString Or txtGameroomBan.Text = vbNullString Or txtGameroomMessage.Text = vbNullString Then
                chkActivateGameroomSpamControl.ForeColor = vbBlue
            End If
            If chkGameBan.Value = vbUnchecked And chkGameKick.Value = vbUnchecked Then
                If chkGameBan.Value = vbUnchecked Then
                    chkGameBan.ForeColor = vbBlue
                ElseIf chkGameKick.Value = vbUnchecked Then
                    chkGameKick.ForeColor = vbBlue
                End If
            End If
        End If
        End If
        'auto kick
        If txtKickInterval.Text <> vbNullString And txtKickMessage.Text <> vbNullString And txtKickMessageNum.Text <> vbNullString And txtKickConsecutive.Text <> vbNullString And txtKickSilence.Text <> vbNullString And chkActivateAutoSpamKick.Value = vbChecked And chkSpamBan.Value = vbUnchecked Then
        If chkSpamKick.Value = vbUnchecked Or chkSpamSilence.Value = vbUnchecked Then
            .kickInterval = txtKickInterval.Text
            .kickMessage = txtKickMessage.Text
            .kickMessageNum = txtKickMessageNum.Text
            .kickConsecutive = txtKickConsecutive.Text
            .kickSilence = txtKickSilence.Text
            .kickStatus = True
            chkActivateAutoSpamKick.ForeColor = vbBlack
            chkSpamBan.ForeColor = vbBlack
            chkSpamSilence.ForeColor = vbBlack
            chkSpamKick.ForeColor = vbBlack
            fAutoKick.Caption = "Spam Control: ON"
            If chkSpamSilence.Value = vbChecked And chkSpamBan.Value = vbChecked _
            Or chkSpamSilence.Value = vbChecked And chkSpamKick.Value _
            Or chkSpamKick.Value = vbChecked And chkSpamBan.Value = vbChecked Then
                chkSpamSilence.Value = vbChecked
                chkSpamBan.Value = vbUnchecked
                chkSpamKick.Value = vbUnchecked
            End If
        Else
            fAutoKick.Caption = "Spam Control: OFF"
            .kickStatus = False
            chkActivateAutoSpamKick.Value = vbUnchecked
            If txtKickInterval.Text = vbNullString Or txtKickMessage.Text = vbNullString Or txtKickSilence.Text = vbNullString Or txtKickMessageNum.Text = vbNullString Then
                chkActivateAutoSpamKick.ForeColor = vbBlue
            End If
            If chkSpamBan.Value = vbUnchecked And chkSpamKick.Value = vbUnchecked And chkSpamSilence.Value = vbUnchecked Then
                If chkSpamBan.Value = vbUnchecked Then
                    chkSpamBan.ForeColor = vbBlue
                ElseIf chkGameKick.Value = vbUnchecked Then
                    chkSpamKick.ForeColor = vbBlue
                ElseIf chkSpamSilence.Value = vbUnchecked Then
                    chkSpamSilence.ForeColor = vbBlue
                End If
            End If
        End If
        End If
        'word filter
        str = vbNullString
        If lstWord.ListCount > 0 And txtFilterMessage.Text <> vbNullString And txtKicksBeforeBan <> vbNullString And chkActivateAutoWordFilter.Value = vbChecked And chkWordBan.Value = vbUnchecked Then
         If chkWordKick.Value = vbUnchecked Or chkWordSilence.Value = vbUnchecked Then
            'fill array with words
            For i = 0 To lstWord.ListCount - 1
                str = str & lstWord.List(i) & ", "
            Next i
            str = Left$(str, Len(str) - 2)
            strSplit = Split(str, ", ")
            ReDim .filterWords(0 To UBound(strSplit))
            For i = 0 To UBound(.filterWords)
                .filterWords(i) = strSplit(i)
            Next i
            .filterSilence = txtKicksBeforeBan.Text
            .filterMessage = txtFilterMessage.Text
            .filterStatus = True
            chkActivateAutoWordFilter.ForeColor = vbBlack
            chkWordBan.ForeColor = vbBlack
            chkWordSilence.ForeColor = vbBlack
            chkWordKick.ForeColor = vbBlack
            If chkWordSilence.Value = vbChecked And chkWordBan.Value = vbChecked _
            Or chkWordSilence.Value = vbChecked And chkWordKick.Value _
            Or chkWordKick.Value = vbChecked And chkWordBan.Value = vbChecked Then
                chkWordSilence.Value = vbChecked
                chkWordBan.Value = vbUnchecked
                chkWordKick.Value = vbUnchecked
            End If
            fWordFilter.Caption = "Word Filter Control: ON"
        Else
            fWordFilter.Caption = "Word Filter Control: OFF"
            .filterStatus = False
            chkActivateAutoWordFilter.Value = vbUnchecked
            If lstWord.ListCount < 0 Or txtFilterMessage.Text = vbNullString Or txtKicksBeforeBan.Text = vbNullString Then
                chkActivateAutoWordFilter.ForeColor = vbBlue
            End If
            If chkWordBan.Value = vbUnchecked And chkWordKick.Value = vbUnchecked And chkWordSilence.Value = vbUnchecked Then
                chkActivateAutoWordFilter.ForeColor = vbBlue
            End If
        End If
       End If
        'game control
        str = vbNullString
        If lstGame.ListItems.Count > 0 And chkActivateGameControl.Value = vbChecked Then
            For i = 0 To lstGame.ListItems.Count
                str = str & lstGame.ListItems(i) & ", " & lstGame.ListItems(i).SubItems(1) & ", "
            Next i
            str = Left$(str, Len(str) - 2)
            'fill array with words
            strSplit = Split(str, ", ")
            ReDim .gameControlGames(0 To ((UBound(strSplit) + 1) / 2) - 1)
            ReDim .gameControlMaxNum(0 To ((UBound(strSplit) + 1) / 2) - 1)
            'fill games and max num
            Dim w, s As Long
            For i = 0 To UBound(strSplit)
                If i Mod 2 = 0 Then
                    .gameControlGames(w) = strSplit(i)
                    w = w + 1
                Else
                    .gameControlMaxNum(s) = strSplit(i)
                    s = s + 1
                End If
            Next i
            
            .gameControlStatus = True
            chkActivateGameControl.ForeColor = vbBlack
            fGameControl.Caption = "Game Control: ON"
        Else
            fGameControl.Caption = "Game Control: OFF"
            .gameControlStatus = False
            chkActivateGameControl.Value = vbUnchecked
            If lstGame.ListItems.Count < 1 Then '4 because the minimum allowed is (game, maxNum): w, 5 = 4
                chkActivateAutoWordFilter.ForeColor = vbBlue
            End If
        End If
        
        'welcome messages
        str = vbNullString
        If lstWelcome.ListItems.Count > 0 And chkWelcome.Value = vbChecked Then
            For i = 1 To lstWelcome.ListItems.Count
                str = str & lstWelcome.ListItems(i) & ", " & lstWelcome.ListItems(i).SubItems(1) & ", " & lstWelcome.ListItems(i).SubItems(2) & ", "
            Next i
            str = Left$(str, Len(str) - 2)
            strSplit = Split(str, ", ")
            w = 0
            For i = 0 To UBound(strSplit)
                ReDim Preserve .welcomeIP(0 To w)
                .welcomeIP(w) = strSplit(i)
                i = i + 1
                ReDim Preserve .welcomeNicks(0 To w)
                .welcomeNicks(w) = strSplit(i)
                i = i + 1
                ReDim Preserve .welcomeMessages(0 To w)
                .welcomeMessages(w) = strSplit(i)
                w = w + 1
            Next i
            .welcomeStatus = True
            chkWelcome.ForeColor = vbBlack
            fWelcome.Caption = "Welcome Messages: ON"
        Else
            fWelcome.Caption = "Welcome Messages: OFF"
            chkWelcome.Value = vbUnchecked
            .welcomeStatus = False
            If lstWelcome.ListItems.Count < 1 Then '4 because the minimum allowed is (ip(x.x.x.x), nick, message): w, 5 = 4
                chkWelcome.ForeColor = vbBlue
            End If
        End If
        
    End With
    
    myBot.kickCount = 0
    myBot.kickmessageNumCount = 0
    myBot.kickConsecutiveCount = 0
        
    Call saveChanges
        
End Sub





Private Sub btnAddUser_Click()
    Dim lst As ListItem
    
    If txtWelcomeMessage.Text = vbNullString Then
        txtWelcomeMessage.Text = "Supply message here!"
        Call textboxStuff(txtWelcomeMessage, 1)
        txtWelcomeMessage.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeNick.Text = vbNullString Then
        txtWelcomeNick.Text = "Supply nick here!"
        Call textboxStuff(txtWelcomeNick, 1)
        txtWelcomeNick.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeIP.Text = vbNullString Then
        txtWelcomeIP.Text = "Supply IP here!"
        Call textboxStuff(txtWelcomeIP, 1)
        txtWelcomeIP.SetFocus
        Exit Sub
    End If
    
    Set lst = lstWelcome.ListItems.Add(, , Trim$(txtWelcomeIP.Text))
    lst.SubItems(1) = Trim$(txtWelcomeNick.Text)
    lst.SubItems(2) = Trim$(txtWelcomeMessage.Text)
    
    txtWelcomeIP.Text = vbNullString
    txtWelcomeNick.Text = vbNullString
    txtWelcomeMessage.Text = vbNullString
End Sub

Private Sub Form_Activate()
    Call fixBotFrames(fPos)
End Sub

Private Sub Form_Load()
    If frmAdminBot.WindowState <> vbMinimized Then
        Me.Height = 7725
        Me.Width = 6915
    End If
            
    If txtName.Text = vbNullString Then
        btnONOFF.Enabled = False
        btnSet.Enabled = False
    Else
        btnONOFF.Enabled = True
        btnSet.Enabled = True
    End If
    'incase file doesn't exist, make it
    Open App.Path & "\bot.txt" For Append As #5
        'need something here to check for empty file and fill it
    Close #5
    'read from it
    Dim splitwords() As String
    Dim str() As String
    Dim wordlist As ListItem
    Dim strBuff As String
    Dim i As Long
    
    Open App.Path & "\bot.txt" For Input As #5
    Do Until EOF(5)
        Line Input #5, strBuff
        If Left$(strBuff, Len("botName=")) = "botName=" Then
            txtName.Text = Right$(strBuff, Len(strBuff) - Len("botName="))
        ElseIf Left$(strBuff, Len("ALLCAPS=")) = "ALLCAPS=" Then
            chkAllCaps.Value = Right$(strBuff, Len(strBuff) - Len("ALLCAPS="))
        ElseIf Left$(strBuff, Len("silenceAllInterval=")) = "silenceAllInterval=" Then
            txtSilenceAll.Text = Right$(strBuff, Len(strBuff) - Len("silenceAllInterval="))
        ElseIf Left$(strBuff, Len("kickConsecutive=")) = "kickConsecutive=" Then
            txtKickConsecutive.Text = Right$(strBuff, Len(strBuff) - Len("kickConsecutive="))
        ElseIf Left$(strBuff, Len("kickInterval=")) = "kickInterval=" Then
            txtKickInterval.Text = Right$(strBuff, Len(strBuff) - Len("kickInterval="))
        ElseIf Left$(strBuff, Len("filterCheck=")) = "filterCheck=" Then
            chkActivateAutoWordFilter.Value = Right$(strBuff, Len(strBuff) - Len("filterCheck="))
        ElseIf Left$(strBuff, Len("kickCheck=")) = "kickCheck=" Then
            chkActivateAutoSpamKick.Value = Right$(strBuff, Len(strBuff) - Len("kickCheck="))
        ElseIf Left$(strBuff, Len("annouonceChatroomCheck=")) = "annouonceChatroomCheck=" Then
            chkActivateAutoAnnounceChatroom.Value = Right$(strBuff, Len(strBuff) - Len("annouonceChatroomCheck="))
        ElseIf Left$(strBuff, Len("kicksBeforeBan=")) = "kicksBeforeBan=" Then
            txtKicksBeforeBan.Text = Right$(strBuff, Len(strBuff) - Len("kicksBeforeBan="))
        ElseIf Left$(strBuff, Len("kickMessageNum=")) = "kickMessageNum=" Then
            txtKickMessageNum.Text = Right$(strBuff, Len(strBuff) - Len("kickMessageNum="))
        ElseIf Left$(strBuff, Len("allcaps=")) = "callcaps=" Then
            chkAllCaps.Value = Right$(strBuff, Len(strBuff) - Len("callcaps="))
        ElseIf Left$(strBuff, Len("kickMessage=")) = "kickMessage=" Then
            txtKickMessage.Text = Right$(strBuff, Len(strBuff) - Len("kickMessage="))
        ElseIf Left$(strBuff, Len("announceInterval=")) = "announceInterval=" Then
            txtAnnounceInterval.Text = Right$(strBuff, Len(strBuff) - Len("announceInterval="))
        ElseIf Left$(strBuff, Len("announceMessage=")) = "announceMessage=" Then
            ReDim str(0)
            str = Split(Right$(strBuff, Len(strBuff) - Len("announceMessage=")), "\0\0")
            txtAnnounceMessage.Text = vbNullString
            If UBound(str) > -1 Then
                For i = 0 To UBound(str)
                    If i = UBound(str) Then
                        txtAnnounceMessage.Text = txtAnnounceMessage.Text & str(i)
                    Else
                        txtAnnounceMessage.Text = txtAnnounceMessage.Text & str(i) & vbCrLf
                    End If
                Next i
            Else
                txtAnnounceMessage.Text = Right$(strBuff, Len(strBuff) - Len("announceMessage="))
            End If
            
            
        ElseIf Left$(strBuff, Len("gameAnnounceInterval=")) = "gameAnnounceInterval=" Then
                txtGameInterval.Text = Right$(strBuff, Len(strBuff) - Len("gameAnnounceInterval="))
        ElseIf Left$(strBuff, Len("gameAnnounceMessage=")) = "gameAnnounceMessage=" Then
                txtGameMessage.Text = Right$(strBuff, Len(strBuff) - Len("gameAnnounceMessage="))
        ElseIf Left$(strBuff, Len("announceGameCheck=")) = "announceGameCheck=" Then
                chkActivateAnnounceGame.Value = Right$(strBuff, Len(strBuff) - Len("announceGameCheck="))
                                     
        Dim temp As String
        ElseIf Left$(strBuff, Len("filterWords=")) = "filterWords=" Then
                temp = Right$(strBuff, Len(strBuff) - Len("filterWords="))
        ReDim str(0)
        str = Split(temp, ";&,")
        For i = 0 To UBound(str)
            lstWord.AddItem str(i)
        Next i

       
        ElseIf Left$(strBuff, Len("filterMessage=")) = "filterMessage=" Then
            txtFilterMessage.Text = Right$(strBuff, Len(strBuff) - Len("filterMessage="))
        ElseIf Left$(strBuff, Len("activategamecontrol=")) = "activategamecontrol=" Then
            chkActivateGameControl.Value = Right$(strBuff, Len(strBuff) - Len("activategamecontrol="))
        ElseIf Left$(strBuff, Len("controlledgames=")) = "controlledgames=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("controlledgames="))
        
        ReDim str(0)
        str = Split(temp, ";&,")
        
            For i = 0 To UBound(str)
                Set wordlist = lstGame.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                i = i + 1
            Next i
       
        
        ElseIf Left$(strBuff, Len("activateWelcome=")) = "activateWelcome=" Then
            chkWelcome.Value = Right$(strBuff, Len(strBuff) - Len("activateWelcome="))
        ElseIf Left$(strBuff, Len("welcomeMessages=")) = "welcomeMessages=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("welcomeMessages="))
        
        ReDim str(0)
        str = Split(temp, ";&,")
        
            For i = 0 To UBound(str)
                Set wordlist = lstWelcome.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                i = i + 1
                wordlist.SubItems(2) = str(i + 1)
                i = i + 1
            Next i
        
        
        ElseIf Left$(strBuff, Len("gamespammessage=")) = "gamespammessage=" Then
            txtGameroomMessage.Text = Right$(strBuff, Len(strBuff) - Len("gamespammessage="))
        ElseIf Left$(strBuff, Len("activategamespam=")) = "activategamespam=" Then
            chkActivateGameroomSpamControl.Value = Right$(strBuff, Len(strBuff) - Len("activategamespam="))
        ElseIf Left$(strBuff, Len("gamespamlimit=")) = "gamespamlimit=" Then
            txtGameroomBan.Text = Right$(strBuff, Len(strBuff) - Len("gamespamlimit="))
        ElseIf Left$(strBuff, Len("gamespamconsecutive=")) = "gamespamconsecutive=" Then
            txtGameroomNum.Text = Right$(strBuff, Len(strBuff) - Len("gamespamconsecutive="))
        ElseIf Left$(strBuff, Len("gameBan=")) = "gameBan=" Then
            chkGameBan.Value = Right$(strBuff, Len(strBuff) - Len("gameBan="))
        ElseIf Left$(strBuff, Len("gameKick=")) = "gameKick=" Then
            chkGameKick.Value = Right$(strBuff, Len(strBuff) - Len("gameKick="))
        ElseIf Left$(strBuff, Len("silenceAll=")) = "silenceAll=" Then
            chkSilenceAll.Value = Right$(strBuff, Len(strBuff) - Len("silenceAll="))
        ElseIf Left$(strBuff, Len("loginSpamControl=")) = "loginSpamControl=" Then
            chkLoginSpamControl.Value = Right$(strBuff, Len(strBuff) - Len("loginSpamControl="))
        ElseIf Left$(strBuff, Len("dupeNick=")) = "dupeNick=" Then
            chkDupeNick.Value = Right$(strBuff, Len(strBuff) - Len("dupeNick="))
        ElseIf Left$(strBuff, Len("wordSilence=")) = "wordSilence=" Then
            chkWordSilence.Value = Right$(strBuff, Len(strBuff) - Len("wordSilence="))
        ElseIf Left$(strBuff, Len("wordBan=")) = "wordBan=" Then
            chkWordBan.Value = Right$(strBuff, Len(strBuff) - Len("wordBan="))
        ElseIf Left$(strBuff, Len("wordKick=")) = "wordKick=" Then
            chkWordKick.Value = Right$(strBuff, Len(strBuff) - Len("wordKick="))
        ElseIf Left$(strBuff, Len("loginNum=")) = "loginNum=" Then
            txtLoginNum.Text = Right$(strBuff, Len(strBuff) - Len("loginNum="))
        ElseIf Left$(strBuff, Len("loginMin=")) = "loginMin=" Then
            txtLoginMin.Text = Right$(strBuff, Len(strBuff) - Len("loginMin="))
        ElseIf Left$(strBuff, Len("loginMessage=")) = "loginMessage=" Then
            txtLoginMessage.Text = Right$(strBuff, Len(strBuff) - Len("loginMessage="))
        ElseIf Left$(strBuff, Len("spamSilence=")) = "spamSilence=" Then
            chkSpamSilence.Value = Right$(strBuff, Len(strBuff) - Len("spamSilence="))
        ElseIf Left$(strBuff, Len("spamBan=")) = "spamBan=" Then
            chkSpamBan.Value = Right$(strBuff, Len(strBuff) - Len("spamBan="))
        ElseIf Left$(strBuff, Len("spamKick=")) = "spamKick=" Then
            chkSpamKick.Value = Right$(strBuff, Len(strBuff) - Len("spamKick="))
        End If
    Loop
    Close #5
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call saveChanges
    If inServer = True Then
        Cancel = 1
        Me.Hide
        Form1.txtDummy.SetFocus
    End If
End Sub



Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub txtAnnounceAllInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8))) Then
        KeyAscii = 0
    End If
End Sub


Private Sub lstDamage_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstDamage, ColumnHeader)
End Sub

Private Sub lstDamage_DblClick()
    If lstDamage.ListItems.Count > 0 Then
        lstDamage.ListItems.Remove (lstDamage.SelectedItem.Index)
    End If
End Sub

Private Sub lstDamage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstDamage.ListItems.Count > 0 Then
        If Button = 2 Then PopupMenu mnuDamage, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub lstGame_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstGame, ColumnHeader)
End Sub

Private Sub lstGame_DblClick()
    If lstGame.ListItems.Count > 0 Then
        lstGame.ListItems.Remove (lstGame.SelectedItem.Index)
    End If
End Sub

Private Sub lstGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstGame.ListItems.Count > 0 Then
        If Button = 2 Then PopupMenu mnuGame, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub lstWelcome_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstWelcome, ColumnHeader)
End Sub

Private Sub lstWelcome_DblClick()
    If lstWelcome.ListItems.Count > 0 Then
        lstWelcome.ListItems.Remove (lstWelcome.SelectedItem.Index)
    End If
End Sub

Private Sub lstWelcome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstWelcome.ListItems.Count > 0 Then
        If Button = 2 Then PopupMenu mnuWelcome, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub lstWord_DblClick()
    If lstWord.ListCount > -1 Then
        lstWord.RemoveItem (lstWord.ListIndex)
    End If
End Sub

Private Sub lstWord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstWord.ListCount > -1 Then
        If Button = 2 Then PopupMenu mnuWord, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub mnuDamageClear_Click()
    lstDamage.ListItems.Remove (lstDamage.SelectedItem.Index)
End Sub

Private Sub mnuDamageClearIP_Click()
    msgCount = msgCount + 1
    Call globalChatRequest("/clear " & lstDamage.SelectedItem.SubItems(1))
    msgCount = msgCount + 1
    Call globalChatRequest("/announce " & Form1.txtUsername.Text & " has cleared <" & lstDamage.SelectedItem & "> of all damage!")
End Sub

Private Sub mnuRemoveDamageAll_Click()
    lstDamage.ListItems.Clear
End Sub

Private Sub mnuRemoveGame_Click()
    lstGame.ListItems.Remove (lstGame.SelectedItem.Index)
End Sub

Private Sub mnuRemoveGameAll_Click()
    lstGame.ListItems.Clear
End Sub

Private Sub mnuRemoveWord_Click()
    If lstWord.ListIndex = -1 Then lstWord.ListIndex = 0
    lstWord.RemoveItem (lstWord.ListIndex)
End Sub

Private Sub mnuRemoveWordAll_Click()
    lstWord.Clear
End Sub

Private Sub mnuWelcomeRemove_Click()
    lstWelcome.ListItems.Remove (lstWelcome.SelectedItem.Index)
End Sub

Private Sub mnuWelcomeRemoveAll_Click()
    lstWelcome.ListItems.Clear
End Sub

Private Sub txtAnnounceInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8))) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtGameSpam_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8))) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAnnounceMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then
        KeyAscii = 0
        txtAnnounceMessage.SelStart = 0
        txtAnnounceMessage.SelLength = Len(txtAnnounceMessage.Text)
    End If
End Sub


Private Sub txtFilterMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtFilterMessage, KeyAscii)
End Sub

Private Sub txtGameControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txtGameInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtGameroomBan, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtGameMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtGameMessage, KeyAscii)
End Sub

Private Sub txtGameName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    Dim lst As ListItem
    
    If txtNumOfGames.Text = vbNullString Then
        txtNumOfGames.Text = "Supply #of games here!"
        Call textboxStuff(txtNumOfGames, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If
    
    If txtGameName.Text = vbNullString Then
        txtGameName.Text = "Supply #of games here!"
        Call textboxStuff(txtGameName, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If

    Set lst = lstGame.ListItems.Add(, , Trim$(txtGameName.Text))
    lst.SubItems(1) = Trim$(txtNumOfGames.Text)

    txtNumOfGames.Text = vbNullString
    txtGameName.Text = vbNullString
End If
End Sub

Private Sub txtGameName_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtGameName, KeyAscii)
End Sub

Private Sub txtGameroomBan_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtGameroomBan, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtGameroomMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtGameroomMessage, KeyAscii)
End Sub

Private Sub txtGameroomNum_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtGameroomNum, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKickConsecutive_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtKickConsecutive, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKickInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtKickInterval, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtKickMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtKickMessage, KeyAscii)
End Sub

Private Sub txtKickMessageNum_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtKickMessageNum, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtKicksBeforeBan_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtKicksBeforeBan, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtKickSilence_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtKickSilence, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLoginMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLoginMessage, KeyAscii)
End Sub

Private Sub txtLoginMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtLoginMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLoginNum_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtLoginNum, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Change()
    If txtName.Text = vbNullString Then
        btnONOFF.Enabled = False
        btnSet.Enabled = False
    Else
        btnONOFF.Enabled = True
        btnSet.Enabled = True
    End If
End Sub


Private Sub txtSpamKickSilence_KeyPress(KeyAscii As Integer)
    Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8))) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtName, KeyAscii)
End Sub


Private Sub txtNumOfGames_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    Dim lst As ListItem
    
    If txtNumOfGames.Text = vbNullString Then
        txtNumOfGames.Text = "Supply #of games here!"
        Call textboxStuff(txtNumOfGames, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If
    
    If txtGameName.Text = vbNullString Then
        txtGameName.Text = "Supply #of games here!"
        Call textboxStuff(txtGameName, 1)
        txtNumOfGames.SetFocus
        Exit Sub
    End If

    Set lst = lstGame.ListItems.Add(, , Trim$(txtGameName.Text))
    lst.SubItems(1) = Trim$(txtNumOfGames.Text)

    txtNumOfGames.Text = vbNullString
    txtGameName.Text = vbNullString
End If
End Sub

Private Sub txtNumOfGames_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtNumOfGames, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSilenceAll_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtSilenceAll, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWelcomeIP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    Dim lst As ListItem
    
    If txtWelcomeMessage.Text = vbNullString Then
        txtWelcomeMessage.Text = "Supply message here!"
        Call textboxStuff(txtWelcomeMessage, 1)
        txtWelcomeMessage.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeNick.Text = vbNullString Then
        txtWelcomeNick.Text = "Supply nick here!"
        Call textboxStuff(txtWelcomeNick, 1)
        txtWelcomeNick.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeIP.Text = vbNullString Then
        txtWelcomeIP.Text = "Supply IP here!"
        Call textboxStuff(txtWelcomeIP, 1)
        txtWelcomeIP.SetFocus
        Exit Sub
    End If
    
    Set lst = lstWelcome.ListItems.Add(, , Trim$(txtWelcomeIP.Text))
    lst.SubItems(1) = Trim$(txtWelcomeNick.Text)
    lst.SubItems(2) = Trim$(txtWelcomeMessage.Text)
    
    txtWelcomeIP.Text = vbNullString
    txtWelcomeNick.Text = vbNullString
    txtWelcomeMessage.Text = vbNullString
End If
End Sub

Private Sub txtWelcomeIP_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtWelcomeIP, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = "." Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWelcomeMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    Dim lst As ListItem
    
    If txtWelcomeMessage.Text = vbNullString Then
        txtWelcomeMessage.Text = "Supply message here!"
        Call textboxStuff(txtWelcomeMessage, 1)
        txtWelcomeMessage.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeNick.Text = vbNullString Then
        txtWelcomeNick.Text = "Supply nick here!"
        Call textboxStuff(txtWelcomeNick, 1)
        txtWelcomeNick.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeIP.Text = vbNullString Then
        txtWelcomeIP.Text = "Supply IP here!"
        Call textboxStuff(txtWelcomeIP, 1)
        txtWelcomeIP.SetFocus
        Exit Sub
    End If
    
    Set lst = lstWelcome.ListItems.Add(, , Trim$(txtWelcomeIP.Text))
    lst.SubItems(1) = Trim$(txtWelcomeNick.Text)
    lst.SubItems(2) = Trim$(txtWelcomeMessage.Text)
    
    txtWelcomeIP.Text = vbNullString
    txtWelcomeNick.Text = vbNullString
    txtWelcomeMessage.Text = vbNullString
End If
End Sub

Private Sub txtWelcomeMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtWelcomeMessage, KeyAscii)
End Sub

Private Sub txtWelcomeNick_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    Dim lst As ListItem
    
    If txtWelcomeMessage.Text = vbNullString Then
        txtWelcomeMessage.Text = "Supply message here!"
        Call textboxStuff(txtWelcomeMessage, 1)
        txtWelcomeMessage.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeNick.Text = vbNullString Then
        txtWelcomeNick.Text = "Supply nick here!"
        Call textboxStuff(txtWelcomeNick, 1)
        txtWelcomeNick.SetFocus
        Exit Sub
    End If
    
    If txtWelcomeIP.Text = vbNullString Then
        txtWelcomeIP.Text = "Supply IP here!"
        Call textboxStuff(txtWelcomeIP, 1)
        txtWelcomeIP.SetFocus
        Exit Sub
    End If
    
    Set lst = lstWelcome.ListItems.Add(, , Trim$(txtWelcomeIP.Text))
    lst.SubItems(1) = Trim$(txtWelcomeNick.Text)
    lst.SubItems(2) = Trim$(txtWelcomeMessage.Text)
    
    txtWelcomeIP.Text = vbNullString
    txtWelcomeNick.Text = vbNullString
    txtWelcomeMessage.Text = vbNullString
End If
End Sub

Private Sub txtWelcomeNick_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtWelcomeNick, KeyAscii)
End Sub

Private Sub txtWordFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
            If txtWordFilter.Text <> vbNullString Then
        lstWord.AddItem Trim$(txtWordFilter.Text)
    Else
        txtWordFilter.Text = "Supply word here!"
        Call textboxStuff(txtWordFilter, 1)
        txtWordFilter.SetFocus
    End If
    txtWordFilter.Text = vbNullString
    End If
End Sub

Private Sub txtWordFilter_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtWordFilter, KeyAscii)
End Sub
