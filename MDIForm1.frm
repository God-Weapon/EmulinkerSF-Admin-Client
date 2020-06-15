VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "EmulinkerSF Admin Client"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1799
      ButtonWidth     =   2514
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Server List"
            Description     =   "View Admin Bot"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Chatroom"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "User List"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Admin Bot"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Massive"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DB Search"
            Description     =   "Search DB by IP"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DB View All"
            Description     =   "Search All DB Entries"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preferences"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6855
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Users: 0"
            TextSave        =   "Users: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Games: 0"
            TextSave        =   "Games: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Session: 0"
            TextSave        =   "Session: 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Text            =   "Version: "
            TextSave        =   "Version: "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Access Level: Normal"
            TextSave        =   "Access Level: Normal"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "Admin Bot: OFF"
            TextSave        =   "Admin Bot: OFF"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Total Users in DB: 0"
            TextSave        =   "Total Users in DB: 0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuServers 
         Caption         =   "Server List"
      End
      Begin VB.Menu mnuBar89 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChatroom 
         Caption         =   "Chatroom"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUserList 
         Caption         =   "User List"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar98 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReconnectToServer 
         Caption         =   "Reconnect"
      End
      Begin VB.Menu mnuLogOffServer 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuBaradsfewr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAdministrativeTools 
      Caption         =   "&Administrative Tools"
      Begin VB.Menu mnuAdminBot 
         Caption         =   "Admin Bot"
      End
      Begin VB.Menu mnuMassiveCommands 
         Caption         =   "Massive Commands"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRemoteControl 
         Caption         =   "Remote Control"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "View All"
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "&Other"
      Begin VB.Menu mnuAsciiArt 
         Caption         =   "ASCII Art"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutEmulinkerSFAC 
         Caption         =   "About EmulinkerSF Admin Client"
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFindUser 
         Caption         =   "Find User"
      End
      Begin VB.Menu mnuBarG 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDBEntry 
         Caption         =   "View Database Entry"
      End
      Begin VB.Menu mnuShowDB 
         Caption         =   "Show Database to Chatroom"
      End
      Begin VB.Menu mnuBar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoDamage 
         Caption         =   "Clear IP Address"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSilence 
         Caption         =   "Silence"
         Begin VB.Menu mnuSilence5 
            Caption         =   "5 minutes"
         End
         Begin VB.Menu mnuSilence15 
            Caption         =   "15 minutes"
         End
         Begin VB.Menu mnuSilence30 
            Caption         =   "30 minutes"
         End
         Begin VB.Menu mnuSilence45 
            Caption         =   "45 minutes"
         End
         Begin VB.Menu mnuBar6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSilence60 
            Caption         =   "1 hour"
         End
         Begin VB.Menu mnuSilence180 
            Caption         =   "3 hours"
         End
         Begin VB.Menu mnuSilence300 
            Caption         =   "5 hours"
         End
         Begin VB.Menu mnuBar7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSilence1440 
            Caption         =   "1 day"
         End
         Begin VB.Menu mnuSilence30000 
            Caption         =   "For a very long time..."
         End
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick from SERVER"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
         Begin VB.Menu mnuBan15 
            Caption         =   "15 minutes"
         End
         Begin VB.Menu mnuBan30 
            Caption         =   "30 minutes"
         End
         Begin VB.Menu mnuBan45 
            Caption         =   "45 minutes"
         End
         Begin VB.Menu mnuBar3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBan60 
            Caption         =   "1 hour"
         End
         Begin VB.Menu mnuBan180 
            Caption         =   "3 hours"
         End
         Begin VB.Menu mnuBan300 
            Caption         =   "5 hours"
         End
         Begin VB.Menu mnuBar4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBan1440 
            Caption         =   "1 day"
         End
         Begin VB.Menu mnuBan30000 
            Caption         =   "For a very long time..."
         End
      End
      Begin VB.Menu mnuBar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTempAdmin 
         Caption         =   "Give Temporary Admin Status"
         Begin VB.Menu mnuTempAdmin15 
            Caption         =   "15 minutes"
         End
         Begin VB.Menu mnuTempAdmin30 
            Caption         =   "30 minutes"
         End
         Begin VB.Menu mnuTempAdmin45 
            Caption         =   "45 minutes"
         End
         Begin VB.Menu mnuBar8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTempAdmin60 
            Caption         =   "1 hour"
         End
         Begin VB.Menu mnuTempAdmin180 
            Caption         =   "3 hours"
         End
         Begin VB.Menu mnuTempAdmin300 
            Caption         =   "5 hours"
         End
         Begin VB.Menu mnuBar9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTempAdmin1440 
            Caption         =   "1 day"
         End
         Begin VB.Menu mnuTempAdmin30000 
            Caption         =   "For a very long time..."
         End
      End
      Begin VB.Menu mnuTempElevated 
         Caption         =   "Give Temporary Elevated Status"
         Begin VB.Menu mnuTempElevated15 
            Caption         =   "15 minutes"
         End
         Begin VB.Menu mnuTempElevated30 
            Caption         =   "30 minutes"
         End
         Begin VB.Menu mnuTempElevated45 
            Caption         =   "45 minutes"
         End
         Begin VB.Menu mnuBar10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTempElevated60 
            Caption         =   "1 hour"
         End
         Begin VB.Menu mnuTempElevated180 
            Caption         =   "3 hours"
         End
         Begin VB.Menu mnuTempElevated300 
            Caption         =   "5 hours"
         End
         Begin VB.Menu mnuBar11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTempElevated1440 
            Caption         =   "1 day"
         End
         Begin VB.Menu mnuTempElevated30000 
            Caption         =   "For a very long time..."
         End
      End
      Begin VB.Menu mnuBar14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommandCopy 
         Caption         =   "Copy"
         Begin VB.Menu mnuCopyUsername 
            Caption         =   "Username"
         End
         Begin VB.Menu mnuCopyIPAddress 
            Caption         =   "IP Address"
         End
         Begin VB.Menu mnuCopyUserID 
            Caption         =   "User ID"
         End
      End
   End
   Begin VB.Menu mnuGameCommands 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFindGames 
         Caption         =   "Get Game Details"
      End
      Begin VB.Menu mnuBar16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJoin 
         Caption         =   "Join this game"
      End
      Begin VB.Menu mnuCloseGame 
         Caption         =   "Close this game"
      End
      Begin VB.Menu mnuBar18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyGameCommands 
         Caption         =   "Copy"
         Begin VB.Menu mnuGCCGame 
            Caption         =   "Game"
         End
         Begin VB.Menu mnuGCCEmulator 
            Caption         =   "Emulator"
         End
         Begin VB.Menu mnuGCCOwner 
            Caption         =   "Owner"
         End
         Begin VB.Menu mnuGCCGameID 
            Caption         =   "Game ID"
         End
      End
   End
   Begin VB.Menu mnuDamage 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuUndoD 
         Caption         =   "Clear IP Address"
      End
      Begin VB.Menu mnuBar22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDamageClS 
         Caption         =   "Clear Screen"
      End
      Begin VB.Menu mnuBar21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDamageCopy 
         Caption         =   "Copy"
         Begin VB.Menu mnuDamageIP 
            Caption         =   "IP Address"
         End
         Begin VB.Menu mnuDamageU 
            Caption         =   "Username"
         End
         Begin VB.Menu mnuDamageR 
            Caption         =   "Reason"
         End
      End
   End
   Begin VB.Menu mnuGameRoomCommands 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuGRCKick 
         Caption         =   "Kick from game"
      End
      Begin VB.Menu mnuBar26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGRCCopy 
         Caption         =   "Copy"
         Begin VB.Menu mnuGRCNick 
            Caption         =   "Nick"
         End
         Begin VB.Menu mnuGRCUserID 
            Caption         =   "UserID"
         End
      End
   End
   Begin VB.Menu mnuDBVEdit 
      Caption         =   ""
      Begin VB.Menu mnuDBEdit 
         Caption         =   "Edit Database"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Activate()
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    'this is an option to show icon without minimizing the form
End Sub

Private Sub MDIForm_Initialize()
    dbPos = 1
    Call Load(frmMassive)
    Call Load(frmAdminBot)
    Call Load(frmPreferences)
    Call Load(frmAscii)
    Call Load(frmDB)
    Call Load(Form1)
    Call Load(frmUserlist)
    Call Load(frmDBV)
    Call frmServerlist.Show
    'frmServerlist.SetFocus
    Call InitResizeArray
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    
    If App.PrevInstance Then
        ActivatePrevInstance
    End If
            
    'emulatorPass = kInfo.appName
    StatusBar1.Panels(4).Text = "Version: " & emulatorPass

    'if the folder is not here than make it
    If GetAttr("EmulinkerSF_Logs") And vbDirectory = False Then
        Call MkDir(App.Path & "\EmulinkerSF_Logs")
    End If

    'incase file doesn't exist, make it
    Close #1
    Open App.Path & "\config.txt" For Append As #1
        'need something here to check for empty file and fill it
    Close #1

    'incase file doesn't exist, make it
    Close #2
    Open App.Path & "\EmulinkerSF_Logs\bot.txt" For Append As #2
        Print #2, "*****************************************"
        Print #2, "BOT LOG"
        Print #2, "*****************************************"
        Print #2, "Today is: " & Date & ", " & Time
        Print #2, "*****************************************"
        Print #2, "*****************************************"
    Close #2

    'incase file doesn't exist, make it
    Close #3
    Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
        Print #3, "*****************************************"
        Print #3, "CHAT LOG"
        Print #3, "*****************************************"
        Print #3, "Today is: " & Date & ", " & Time
        Print #3, "*****************************************"
        Print #3, "*****************************************"
    Close #3

    'incase file doesn't exist, make it
    Close #5
    Open App.Path & "\EmulinkerSF_Logs\gamechat.txt" For Append As #5
        Print #5, "*****************************************"
        Print #5, "GAME CHAT LOG"
        Print #5, "*****************************************"
        Print #5, "Today is: " & Date & ", " & Time
        Print #5, "*****************************************"
        Print #5, "*****************************************"
    Close #5
    
    
    
    Close #4
    Open App.Path & "\EmulinkerSF_Logs\user.txt" For Append As #4
    Close #4
    Open App.Path & "\EmulinkerSF_Logs\user.bak" For Append As #4
    Close #4
    
    
    
    
    

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim q As Long
        
    If inServer = True Then
        Call userQuitRequest(Trim$(frmPreferences.txtQuit.Text))
    End If
        
    If iQuit = False Then
        frmServerlist.WindowState = vbNormal
        frmServerlist.Top = MDIForm1.Top + 100
        frmServerlist.Left = 0
        frmServerlist.List1.AddItem "Saving Database..."
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        frmServerlist.Winsock1.Close
        If dbEdit = True Then
            Call saveUsers(True)
            Call saveUsers
        End If
        frmServerlist.List1.AddItem "Finished"
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        'call fixframesbuttons(0)
    End If
    
    Call saveConfig
    allowUnload = True
    Close #1
    Close #2
    Close #3
    Close #4
    Close #5
    Close #6
End Sub

Private Sub MDIForm_Resize()
    'this is the option to show icon on minimizing the form
    If Me.WindowState = 1 Then
    'The user has minimized his window
    
    

    ' Add the form's icon to the tray
    
    Me.Hide
    
    ' Hide the button at the taskbar
    
    End If
End Sub

Private Sub MDIForm_Terminate()
    allowUnload = True
    Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, IconData
    Unload frmAbout
    Unload frmAdminBot
    Unload frmDB
    Unload frmDBV
    Unload frmRemote
    Unload frmAscii
    Unload frmPreferences
    Unload frmMassive
    Unload Form1
    Unload Me
    End
End Sub

Private Sub mnuAboutEmulinkerSFAC_Click()
    frmAbout.Show
    frmAbout.WindowState = vbNormal
    'frmAbout.SetFocus
End Sub

Private Sub mnuAdminBot_Click()
    frmAdminBot.Show
    If frmAdminBot.WindowState = vbNormal Then
        frmAdminBot.Left = Form1.Left
    End If
    frmAdminBot.WindowState = vbNormal
    Call frmAdminBot.ZOrder(vbBringToFront)
    'frmAdminBot.SetFocus
End Sub

Private Sub mnuAsciiArt_Click()
    If inServer = False Then Exit Sub
    
    frmAscii.Show
    frmAscii.WindowState = vbNormal
    Call frmAscii.ZOrder(vbBringToFront)
    'frmAscii.SetFocus
End Sub

Private Sub mnuChatroom_Click()
    If inServer = False Then Exit Sub
        
    Form1.Show
    Form1.WindowState = vbNormal
    Call Form1.ZOrder(vbBringToFront)
    'Form1.SetFocus
End Sub

Private Sub mnuDBEdit_Click()
    Dim str() As String
    Dim entryID As Long
    Dim head As Long
    Dim sect1 As Long
    
    str = Split(frmDBV.lstDBV.SelectedItem.SubItems(1), ".")
    head = str(0)
    sect1 = str(1)
    entryID = frmDBV.lstDBV.SelectedItem.SubItems(3)
    Call frmDB.Show
    Call showDBEntry(entryID, frmDBV.lstDBV.SelectedItem.SubItems(2), head, sect1)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFindUser_Click()
    Dim str() As String
    
    str = Split(frmUserlist.lstUserlist.SelectedItem, " [is] ")
    
    Call globalChatRequest("/finduser " & str(0))
End Sub

Public Sub mnuLogOffServer_Click()
    Dim i As Long
    
    If inServer = True Then
        If inRoom = True Then Call Form1.btnGameExit_Click
        Form1.Timer2.Enabled = False
        Call userQuitRequest(Trim$(frmPreferences.txtQuit.Text))
        iQuit = True
        Call fixFramesButtons(0)
    End If
End Sub

Private Sub mnuMassiveCommands_Click()
    If inServer = False Then Exit Sub
       
    frmMassive.Show
    frmMassive.WindowState = vbNormal
    Call frmMassive.ZOrder(vbBringToFront)
    'frmMassive.SetFocus
End Sub

Private Sub mnuPreferences_Click()
    frmPreferences.Show
    frmPreferences.WindowState = vbNormal
    Call frmPreferences.ZOrder(vbBringToFront)
    'frmPreferences.SetFocus
End Sub

Public Sub mnuReconnectToServer_Click()
    iQuit = True
    'initial
    Call userQuitRequest("Reconnecting...")
    frmServerlist.List1.AddItem "*Ping Timeout (" & Time & ")*"
    frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
    Call frmServerlist.btnExit_Click
    DoEvents
    Sleep (1000)
    Call frmServerlist.btnLogin_Click 'reconnectToServer
End Sub

Private Sub mnuRemoteControl_Click()
    If inServer = False Then Exit Sub
        
    frmRemote.Show
    frmRemote.WindowState = vbNormal
    Call frmRemote.ZOrder(vbBringToFront)
    'frmRemote.SetFocus
End Sub

Private Sub mnuSearch_Click()
    frmDB.Show
    frmDB.WindowState = vbNormal
    Call frmDB.ZOrder(vbBringToFront)
    'frmDB.SetFocus
End Sub




Private Sub mnuServers_Click()
    frmServerlist.Show
    frmServerlist.WindowState = vbNormal
    Call frmServerlist.ZOrder(vbBringToFront)

    'If frmServerlist.Enabled = True Then frmServerlist.SetFocus
End Sub

Private Sub mnuUserList_Click()
    If inServer = False Then Exit Sub
    frmUserlist.Show
    frmUserlist.WindowState = vbNormal
    Call frmUserlist.ZOrder(vbBringToFront)
    'frmUserlist.SetFocus
End Sub

Private Sub mnuViewAll_Click()
    frmDBV.Show
    frmDBV.WindowState = vbNormal
    Call frmDBV.ZOrder(vbBringToFront)
    'frmDBV.SetFocus
End Sub

Private Sub mnuBan1440_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 1440")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 1440min")
End Sub

Private Sub mnuBan15_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 15")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 15min")
End Sub

Private Sub mnuBan180_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 180")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 180min")
End Sub

Private Sub mnuBan30_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 30min")
End Sub


Private Sub mnuBan300_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 300")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban [approx. 45.5 days]")
End Sub

Private Sub mnuBan45_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 45")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 45min")
End Sub

Private Sub mnuBan60_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 60")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban 60min")
End Sub

Private Sub mnuBan30000_Click()
    Call globalChatRequest("/ban " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30000")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Ban [approx. 20.83 days]")
End Sub


Private Sub mnuCloseGame_Click()
    Call Form1.btnCloseGame_Click
End Sub

Private Sub mnuCopyIpAddress_Click()
    Clipboard.Clear
    Clipboard.SetText frmUserlist.lstUserlist.SelectedItem.SubItems(2)
End Sub

Private Sub mnuCopyUserId_Click()
    Clipboard.Clear
    Clipboard.SetText frmUserlist.lstUserlist.SelectedItem.SubItems(3)
End Sub

Private Sub mnuCopyUsername_Click()
    Clipboard.Clear
    Clipboard.SetText frmUserlist.lstUserlist.SelectedItem
End Sub

Private Sub mnuDamageCLS_Click()
    frmAdminBot.lstDamage.ListItems.Clear
End Sub

Private Sub mnuDamageIP_Click()
    Clipboard.Clear
    Clipboard.SetText frmAdminBot.lstDamage.SelectedItem.SubItems(1)
End Sub

Private Sub mnuDamageR_Click()
    Clipboard.Clear
    Clipboard.SetText frmAdminBot.lstDamage.SelectedItem.SubItems(3)
End Sub

Private Sub mnuDamageU_Click()
    Clipboard.Clear
    Clipboard.SetText frmAdminBot.lstDamage.SelectedItem.Text
End Sub



Private Sub mnuFindGames_Click()
    Call globalChatRequest("/findgame " & Form1.lstGamelist.SelectedItem)
End Sub

Private Sub mnuGCCEmulator_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGamelist.SelectedItem.SubItems(1)
End Sub

Private Sub mnuGCCGame_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGamelist.SelectedItem
End Sub

Private Sub mnuGCCGameId_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGamelist.SelectedItem.SubItems(5)
End Sub

Private Sub mnuGCCOwner_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGamelist.SelectedItem.SubItems(2)
End Sub

Private Sub mnuGRCKick_Click()
    Call kickRequest(Form1.lstGameUserlist.SelectedItem.SubItems(3))
End Sub

Private Sub mnuGRCNick_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGameUserlist.SelectedItem.Text
End Sub

Private Sub mnuGRCUserID_Click()
    Clipboard.Clear
    Clipboard.SetText Form1.lstGameUserlist.SelectedItem.SubItems(3)
End Sub

Private Sub mnuJoin_Click()
    Call Form1.btnJoin_Click
End Sub



Private Sub mnuShowDB_Click()
    Dim str() As String
    Dim userID As Long
    Dim i As Long
    
    str = Split(frmUserlist.lstUserlist.SelectedItem.Text, " [is] ")
    userID = CLng(frmUserlist.lstUserlist.SelectedItem.SubItems(3))
    
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If arUsers(i).userID = userID Then
                Call showDBEntry(arUsers(i).dbPos, str(0), arUsers(i).head, arUsers(i).sect1)
            End If
        End If
    Next i
    Call frmDB.btnDBShow_Click
End Sub


Public Sub mnuKick_Click()
    Call globalChatRequest("/kick " & frmUserlist.lstUserlist.SelectedItem.SubItems(3))
End Sub





Private Sub mnuSilence1440_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 1440")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 1440min")
End Sub

Private Sub mnuSilence15_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 15")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 15min")
End Sub

Private Sub mnuSilence180_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 180")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 180min")
End Sub

Private Sub mnuSilence30_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 30min")
End Sub

Private Sub mnuSilence300_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 300")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 5 hours")
End Sub

Private Sub mnuSilence45_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 45")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 45min")
End Sub

Private Sub mnuSilence5_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 5")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 5min")
End Sub

Private Sub mnuSilence60_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 60")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence 60min")
End Sub


Private Sub mnuTempAdmin1440_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 1440")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 1440min")
End Sub

Private Sub mnuTempAdmin15_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 15")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 15min")
End Sub

Private Sub mnuTempAdmin180_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 180")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 180min")
End Sub

Private Sub mnuTempAdmin30_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 30min")
End Sub


Private Sub mnuTempAdmin300_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 300")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 5 hours")
End Sub

Private Sub mnuTempAdmin45_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 45")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 45min")
End Sub

Private Sub mnuTempAdmin60_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 60")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin 1 hour")
End Sub

Private Sub mnuUndoD_Click()
    Dim str() As String
    
    str = Split(frmAdminBot.lstDamage.SelectedItem.Text, " [is] ")
    Call globalChatRequest("/clear " & frmAdminBot.lstDamage.SelectedItem.SubItems(1))
    Call globalChatRequest("/announce <" & str(0) & "'s> address has been cleared!")
End Sub

Private Sub mnuUndoDamage_Click()
    Dim str() As String
    
    str = Split(frmUserlist.lstUserlist.SelectedItem.Text, " [is] ")
    Call globalChatRequest("/clear " & frmUserlist.lstUserlist.SelectedItem.SubItems(2))
    Call globalChatRequest("/announce <" & str(0) & "'s> address has been cleared!")
End Sub

Private Sub mnuSilence30000_Click()
    Call globalChatRequest("/silence " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30000")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Silence [approx. 20.83 days]")
End Sub

Private Sub mnuTempAdmin30000_Click()
    Call globalChatRequest("/tempadmin " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30000")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Admin [approx. 20.83 days]")
End Sub

Private Sub mnuTempElevated1440_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 1440")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 1 day")
End Sub

Private Sub mnuTempElevated15_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 15")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 15min")
End Sub

Private Sub mnuTempElevated180_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 180")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 180min")
End Sub

Private Sub mnuTempElevated30_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 30min")
End Sub

Private Sub mnuTempElevated300_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 300")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 5 hours")
End Sub

Private Sub mnuTempElevated45_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 45")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 45min")
End Sub

Private Sub mnuTempElevated60_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 60")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated 1 hour")
End Sub

Private Sub mnuTempElevated30000_Click()
    Call globalChatRequest("/tempelevated " & frmUserlist.lstUserlist.SelectedItem.SubItems(3) & " 30000")
    Call addDamage(frmUserlist.lstUserlist.SelectedItem.Text, frmUserlist.lstUserlist.SelectedItem.SubItems(2), "Manual Temp Elevated [approx. 20.83 days]")
End Sub


Private Sub mnuViewDBEntry_Click()
    Dim userID As Long
    Dim str() As String
    Dim i As Long
    
    userID = CLng(frmUserlist.lstUserlist.SelectedItem.SubItems(3))
    str = Split(frmUserlist.lstUserlist.SelectedItem.Text, " [is] ")
    Call frmDB.Show
    
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            If arUsers(i).userID = userID Then
                Call showDBEntry(arUsers(i).dbPos, str(0), arUsers(i).head, arUsers(i).sect1)
            End If
        End If
    Next i
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button = "Server List" Then
        Call mnuServers_Click
    ElseIf Button = "Chatroom" Then
        Call mnuChatroom_Click
    ElseIf Button = "User List" Then
        Call mnuUserList_Click
    ElseIf Button = "Admin Bot" Then
        Call mnuAdminBot_Click
    ElseIf Button = "DB Search" Then
        Call mnuSearch_Click
    ElseIf Button = "DB View All" Then
        Call mnuViewAll_Click
    ElseIf Button = "Preferences" Then
        Call mnuPreferences_Click
    ElseIf Button = "Massive" Then
        Call mnuMassiveCommands_Click
    End If
End Sub






























Public Sub mnuExit1_Click()

Unload Me
' Unload the form

End
' Just to be sure the program has ended

End Sub

Public Sub mnuShow_Click()

Me.WindowState = vbMaximized
'Shell_NotifyIcon NIM_DELETE, IconData
Me.Show

End Sub




