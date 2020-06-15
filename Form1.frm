VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Chatroom"
   ClientHeight    =   8415
   ClientLeft      =   1065
   ClientTop       =   2445
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "Form1.frx":11F6
   ScaleHeight     =   8415
   ScaleWidth      =   9015
   Begin VB.CommandButton btnWipeOut 
      Caption         =   "Wipe Out!"
      Height          =   315
      Left            =   6720
      TabIndex        =   30
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton btnMe 
      Caption         =   "Me"
      Height          =   315
      Left            =   7200
      TabIndex        =   29
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton btnAnnounceAll 
      Caption         =   "Announce All"
      Height          =   315
      Left            =   5400
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnAnnounce 
      Caption         =   "Announce"
      Height          =   315
      Left            =   7920
      TabIndex        =   25
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton btnChat 
      Caption         =   "Chat"
      Height          =   315
      Left            =   6480
      TabIndex        =   24
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtChat 
      Height          =   315
      Left            =   600
      TabIndex        =   23
      Top             =   4800
      Width           =   5775
   End
   Begin VB.CommandButton btnToggle 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ä"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Toggle roomlist & gameroom"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton btnHistory 
      Caption         =   "History View"
      Height          =   315
      Left            =   7800
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnQuickBan 
      Caption         =   "Quick Ban"
      Height          =   315
      Left            =   3480
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtQuickBan 
      Height          =   315
      Left            =   2880
      TabIndex        =   19
      Text            =   "900"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFindUser 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fGameroom 
      Caption         =   "Gameroom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton btnGameMe 
         Caption         =   "Me"
         Height          =   315
         Left            =   3840
         TabIndex        =   31
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton btnHistory1 
         Caption         =   "History View"
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton btnGameExit 
         Caption         =   "Leave"
         Height          =   315
         Left            =   8040
         TabIndex        =   13
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton btnKick 
         Caption         =   "Kick"
         Height          =   315
         Left            =   7320
         TabIndex        =   8
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtKickUsers 
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         Text            =   "8"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtGameChat 
         Height          =   315
         Left            =   120
         MaxLength       =   109
         TabIndex        =   6
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton btnGameChat 
         Caption         =   "Chat"
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   2760
         Width           =   615
      End
      Begin MSComctlLib.ListView lstGameUserlist 
         Height          =   2415
         Left            =   5760
         TabIndex        =   9
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nick"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ping"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Connection"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UserID"
            Object.Width           =   1411
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGameChatroomD 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4260
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":1780
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtGameChatroom 
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4260
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         TextRTF         =   $"Form1.frx":1804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Users:"
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Frame fRoomList 
      Caption         =   "Room List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8775
      Begin VB.CommandButton btnStealthMode 
         Caption         =   "Turn Stealth Mode ON"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCreateGame 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Text            =   "*Chat (not game)"
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton btnCloseGame 
         Caption         =   "Close"
         Height          =   315
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnJoin 
         Caption         =   "Join"
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnCreate 
         Caption         =   "Create"
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin MSComctlLib.ListView lstGamelist 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4260
         SortKey         =   3
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Game"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emulator"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Owner"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Users"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "dummy"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox txtChatroomD 
      Height          =   4215
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":1888
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChatroom 
      Height          =   4215
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":190C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents FormSubClass1 As clsSubClass
Attribute FormSubClass1.VB_VarHelpID = -1
Private WithEvents FormSubClass2 As clsSubClass
Attribute FormSubClass2.VB_VarHelpID = -1

Private Sub EnableAutoURLDetection()
    SendMessage txtChatroomD.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    SendMessage txtChatroom.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    Set FormSubClass1 = New clsSubClass
    FormSubClass1.Enable Me.hwnd
    SendMessage txtChatroom.hwnd, EM_SETEVENTMASK, 0&, ByVal ENM_LINK
    SendMessage txtChatroomD.hwnd, EM_SETEVENTMASK, 0&, ByVal ENM_LINK
    

    SendMessage txtGameChatroomD.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    SendMessage txtGameChatroom.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    Set FormSubClass2 = New clsSubClass
    FormSubClass2.Enable fGameroom.hwnd
    SendMessage txtGameChatroom.hwnd, EM_SETEVENTMASK, 0&, ByVal ENM_LINK
    SendMessage txtGameChatroomD.hwnd, EM_SETEVENTMASK, 0&, ByVal ENM_LINK
End Sub


Private Sub btnGameMe_Click()
    If Trim$(txtGameChat.Text) = vbNullString Then Exit Sub 'Or Trim$(txtChat.Text) = "You don't have admin status!" Then Exit Sub
    Call splitRegGame("/me " & Trim$(txtGameChat.Text))
    txtGameChat.Text = vbNullString
End Sub

Private Sub btnMe_Click()
    If Trim$(txtChat.Text) = vbNullString Then Exit Sub 'Or Trim$(txtChat.Text) = "You don't have admin status!" Then Exit Sub
    Call splitReg("/me " & Trim$(txtChat.Text))
    txtChat.Text = vbNullString
End Sub

Private Sub FormSubClass1_WMArrival(hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, lRetVal As Long)
Dim notifyCode As NMHDR
Dim LinkData As ENLINK
Dim URL As String


    Select Case uMsg
    Case WM_NOTIFY

        CopyMemory notifyCode, ByVal lParam, LenB(notifyCode)
        If notifyCode.code = EN_LINK Then
        'A RTB sends EN_LINK notifications when it receives certain mouse messages
        'while the mouse pointer is over text that has the CFE_LINK effect:
        
        'To receive EN_LINK notifications, specify the ENM_LINK flag in the mask
        'sent with the EM_SETEVENTMASK message.
        
        'If you send the EM_AUTOURLDETECT message to enable automatic URL detection,
        'the RTB automatically sets the CFE_LINK effect for modified text that it
        'identifies as a URL.
        
            CopyMemory LinkData, ByVal lParam, Len(LinkData)
            If LinkData.msg = WM_LBUTTONUP Then
                'user clicked on a hyperlink
                'get text with CFE_LINK effect that caused message to be sent
                If notifyCode.hWndFrom = txtChatroom.hwnd Then
                    URL = Mid(txtChatroom.Text, LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
                Else
                    URL = Mid(txtChatroomD.Text, LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
                End If
                'launch the browser here
                ShellExecute 0&, "OPEN", URL, vbNullString, "C:\", SW_SHOWNORMAL
            End If

        End If
        lRetVal = FormSubClass1.callWindProc(hwnd, uMsg, wParam, lParam)
        
    Case Else
        lRetVal = FormSubClass1.callWindProc(hwnd, uMsg, wParam, lParam)
    End Select


End Sub

Private Sub FormSubClass2_WMArrival(hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, lRetVal As Long)
Dim notifyCode As NMHDR
Dim LinkData As ENLINK
Dim URL As String


    Select Case uMsg
    Case WM_NOTIFY

        CopyMemory notifyCode, ByVal lParam, LenB(notifyCode)
        If notifyCode.code = EN_LINK Then
        'A RTB sends EN_LINK notifications when it receives certain mouse messages
        'while the mouse pointer is over text that has the CFE_LINK effect:
        
        'To receive EN_LINK notifications, specify the ENM_LINK flag in the mask
        'sent with the EM_SETEVENTMASK message.
        
        'If you send the EM_AUTOURLDETECT message to enable automatic URL detection,
        'the RTB automatically sets the CFE_LINK effect for modified text that it
        'identifies as a URL.
        
            CopyMemory LinkData, ByVal lParam, Len(LinkData)
            If LinkData.msg = WM_LBUTTONUP Then
                'user clicked on a hyperlink
                'get text with CFE_LINK effect that caused message to be sent
                If notifyCode.hWndFrom = txtGameChatroom.hwnd Then
                    URL = Mid(txtGameChatroom.Text, LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
                Else
                    URL = Mid(txtGameChatroomD.Text, LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
                End If                'launch the browser here
                ShellExecute 0&, "OPEN", URL, vbNullString, "C:\", SW_SHOWNORMAL
            End If

        End If
        lRetVal = FormSubClass2.callWindProc(hwnd, uMsg, wParam, lParam)
        
    Case Else
        lRetVal = FormSubClass2.callWindProc(hwnd, uMsg, wParam, lParam)
    End Select


End Sub







Private Sub btnAnnounceAll_Click()
    If Trim$(txtChat.Text) = vbNullString Then Exit Sub 'Or Trim$(txtChat.Text) = "You don't have admin status!" Then Exit Sub
    Call globalChatRequest("/announceall " & Trim$(txtChat.Text))
    txtChat.Text = vbNullString
End Sub






Private Sub btnCreate_Click()
    Dim str As String
    
    str = Trim$(txtCreateGame.Text)
    If str = vbNullString Then
        txtCreateGame.Text = "*Chat (not game)"
        Exit Sub
    End If
    
    Call createGameRequest(str)
    imOwner = True
    myGame = str
    Form1.fRoomList.Caption = "Currently in: " & myGame
    Form1.fGameroom.Caption = myGame
End Sub







Private Sub btnGameChat_Click()
    If Trim$(txtGameChat.Text) = vbNullString Then Exit Sub
    Call splitRegGame(Trim$(Form1.txtGameChat.Text))
    txtGameChat.Text = vbNullString
End Sub

Public Sub btnGameExit_Click()
    leftRoom = True
    If rSwitch = True Then Call Form1.btnToggle_Click
    Call quitGameRequest
End Sub

Private Sub btnHistory_Click()
    Static here As Boolean
    
    If here = False Then
        txtChatroomD.Visible = True
        btnHistory.Caption = "Scroll View"
        Form1.txtChatroomD = vbNullString
        'Set insert point (can be at ANY point i
        '     n rtb1)
        Form1.txtChatroomD.SelStart = Len(Form1.txtChatroomD.Text)
        'Select rich text to add
        Form1.txtChatroom.SelStart = 0
        Form1.txtChatroom.SelLength = Len(Form1.txtChatroom.Text)
        'Add the selected rich text
        Form1.txtChatroomD.SelRTF = Form1.txtChatroom.SelRTF
        Form1.txtChatroomD.SelStart = Len(Form1.txtChatroomD.Text)
        
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
        here = True
    Else
        txtChatroomD.Visible = False
        btnHistory.Caption = "History View"
        here = False
    End If
End Sub

Private Sub btnHistory1_Click()
    Static here As Boolean
    
    If here = False Then
        txtGameChatroomD.Visible = True
        btnHistory1.Caption = "Scroll View"
        Form1.txtGameChatroomD = vbNullString
        'Set insert point (can be at ANY point i
        '     n rtb1)
        Form1.txtGameChatroomD.SelStart = Len(Form1.txtGameChatroomD.Text)
        'Select rich text to add
        Form1.txtGameChatroom.SelStart = 0
        Form1.txtGameChatroom.SelLength = Len(Form1.txtGameChatroom.Text)
        'Add the selected rich text
        Form1.txtGameChatroomD.SelRTF = Form1.txtGameChatroom.SelRTF
        Form1.txtGameChatroomD.SelStart = Len(Form1.txtGameChatroomD.Text)
        
        Form1.txtGameChatroom.SelStart = Len(Form1.txtGameChatroom.Text)
        here = True
    Else
        txtGameChatroomD.Visible = False
        btnHistory1.Caption = "History View"
        here = False
    End If
End Sub

Public Sub btnJoin_Click()
    Dim w As Long
    
    If lstGamelist.ListItems.count > 0 Then
        If myGameId = Form1.lstGamelist.SelectedItem.SubItems(5) Then Exit Sub
        
        If inRoom = True Then
            leftRoom = True
            Call quitGameRequest
        End If
    
        w = GetTickCount
        Do Until inRoom = False Or GetTickCount - w >= 3000
            DoEvents
        Loop
        If inRoom = True Then Exit Sub
        imOwner = False
        'Form1.btnKick.Visible = False
        myGame = Form1.lstGamelist.SelectedItem
        Call joinGameRequest(Form1.lstGamelist.SelectedItem.SubItems(5))
        If rSwitch = False Then Call Form1.btnToggle_Click
        Form1.fRoomList.Caption = "Currently in: " & myGame
        Form1.fGameroom.Caption = myGame
    Else
        txtChatroom.SelColor = &HFF&
        txtChatroom.SelText = txtChatroom.SelText & "*No games are available*" & vbCrLf
    End If
End Sub



Public Sub btnKick_Click()
    If myGameId = -1 Then
        Exit Sub
    ElseIf inRoom = False And imOwner = False Then
        Exit Sub
    End If
    'If lstGameUserlist.SelectedItem.SubItems(3) = myUserId Then
        'txtGameChatroom.SelColor = &HFF&
        'txtGameChatroom.SelText = txtGameChatroom.SelText & "*You can't kick yourself*" & vbCrLf
        'txtGameChatroom.SelStart = Len(txtGameChatroom.Text)
    'Else
    If lstGameUserlist.ListItems.count > 0 Then
        Call kickRequest(lstGameUserlist.SelectedItem.SubItems(3))
    End If
End Sub

Private Sub btnQuickBan_Click()
    If lastUserid <> vbNullString Then
        Call globalChatRequest("/ban " & lastUserid & " " & txtQuickBan.Text)
    End If
End Sub

Private Sub btnStealthMode_Click()
    If btnStealthMode.Caption = "Turn Stealth Mode ON" Then
        Call globalChatRequest("/stealthon")
        btnStealthMode.Caption = "Turn Stealth Mode OFF"
    Else
        Call globalChatRequest("/stealthoff")
        btnStealthMode.Caption = "Turn Stealth Mode ON"
    End If
End Sub


Public Sub btnWipeOut_Click()
    If Trim$(txtChat.Text) = vbNullString Then Exit Sub
    Call globalChatRequest("/announce " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & txtChat.Text)
    txtChat.Text = vbNullString
End Sub







Public Sub btnCloseGame_Click()
    If lstGamelist.ListItems.count > 0 Then
        Call globalChatRequest("/closegame " & Form1.lstGamelist.SelectedItem.SubItems(5))
        Form1.txtChatroom.SelColor = &H8000&
        Form1.txtChatroom.SelText = Form1.txtChatroom.SelText & "*Client Alert: Closed->" & Form1.lstGamelist.SelectedItem.SubItems(2) & " (" & Form1.lstGamelist.SelectedItem.SubItems(5) & "): " & Form1.lstGamelist.SelectedItem.Text & vbCrLf
        Form1.txtChatroom.SelStart = Len(Form1.txtChatroom.Text)
    Else
        Call MsgBox("You can't close a game because there are no games \=", vbOKOnly, "Client Alert!")
    End If
End Sub

Private Sub btnAnnounce_Click()
    If Trim$(txtChat.Text) = vbNullString Then Exit Sub 'Or Trim$(txtChat.Text) = "You don't have admin status!" Then Exit Sub
    Call splitAnnounce(Trim$(txtChat.Text))
    txtChat.Text = vbNullString
End Sub

Private Sub btnChat_Click()
    If Trim$(txtChat.Text) = vbNullString Then Exit Sub 'Or Trim$(txtChat.Text) = "You don't have admin status!" Then Exit Sub
    Call splitReg(Trim$(txtChat.Text))
    txtChat.Text = vbNullString
End Sub





Public Sub btnToggle_Click()
                
    If rSwitch = False Then
        Call fixFramesButtons(3)
        btnToggle.Caption = "Å"
        rSwitch = True
    Else
        Call fixFramesButtons(9)
        btnToggle.Caption = "Ä"
        rSwitch = False
    End If

End Sub


Private Sub Form_Initialize()
    Me.Width = 9240
    Me.Height = 8985
End Sub

Private Sub Form_Resize()
    Call ResizeControls("form1")
End Sub



Private Sub lstGamelist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstGamelist, ColumnHeader)
End Sub

Private Sub lstGamelist_DblClick()
    Call btnJoin_Click
End Sub


Private Sub lstGamelist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstGamelist.ListItems.count > 0 Then
        'creates popup menu when you click on the left mouse button
        If Button = 2 Then PopupMenu MDIForm1.mnuGameCommands, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub lstGameUserlist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstGameUserlist, ColumnHeader)
End Sub

Private Sub lstGameUserlist_DblClick()
    'If lstGameUserlist.ListItems.count > 0 Then
        Call btnKick_Click
    'End If
End Sub

Private Sub lstGameUserlist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstGameUserlist.ListItems.count > 0 Then
        'creates popup menu when you click on the left mouse button
        If Button = 2 Then PopupMenu MDIForm1.mnuGameRoomCommands, vbPopupMenuCenterAlign
    End If
End Sub


Private Sub Form_Load()
    Dim strBuff As String
    
    On Error Resume Next
        
    Call EnableAutoURLDetection

    fGameroom.Top = 5160
    fGameroom.Left = 120
    fRoomList.Top = 5160
    fRoomList.Left = 120
    
    'read from it
    Open App.Path & "\config.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, strBuff

        If Left$(strBuff, 9) = "maxusers=" Then
            txtKickUsers.Text = Right$(strBuff, Len(strBuff) - 9)
        ElseIf Left$(strBuff, 9) = "quickban=" Then
            txtQuickBan.Text = Right$(strBuff, Len(strBuff) - 9)
        End If
        'ElseIf Left$(strBuff, 2) = "X=" Then
        '    'If Form1.WindowState = vbNormal Then Form1.Left = Right$(strBuff, Len(strBuff) - 2)
        'ElseIf Left$(strBuff, 2) = "Y=" Then
        '    'If Form1.WindowState = vbNormal Then Form1.Top = Right$(strBuff, Len(strBuff) - 2)
        Loop
    Close #1
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub




Public Sub sessionTime()
    Dim i, w As Long
    On Error Resume Next
    
    finishedTime = finishedTime + 1
    
    If myBot.botStatus = True Then
        For i = LBound(arUsers) To UBound(arUsers)
            If arUsers(i).loggedIn = True Then
            
                If frmAdminBot.chkSpamControl.Value = vbChecked Then
                    arUsers(i).spamTimeout = arUsers(i).spamTimeout + 1
                    If arUsers(i).spamTimeout >= CLng(frmAdminBot.txtSpamExpire.Text) Then
                        arUsers(i).spamTimeout = 0
                        arUsers(i).spamRowCount = 0
                    End If
                End If
                
                If frmAdminBot.chkLinkControl.Value = vbChecked Then
                    arUsers(i).linkCount = arUsers(i).linkCount + 1
                    If arUsers(i).linkCount > CLng(frmAdminBot.txtLinksInterval.Text) Then
                        arUsers(i).linkSent = 0
                        arUsers(i).linkCount = 0
                    End If
                End If
                
                
                If frmAdminBot.chkGameSpamControl.Value = vbChecked Then
                    arUsers(i).gameTimeout = arUsers(i).gameTimeout + 1
                    If arUsers(i).gameTimeout >= CLng(frmAdminBot.txtSpamExpire.Text) Then
                        arUsers(i).gameTimeout = 0
                        arUsers(i).gameSpamCount = 0
                    End If
                End If
    
            End If
        Next i

        If frmAdminBot.chkLoginSpamControl.Value = vbChecked Then
            myBot.loginTimeout = myBot.loginTimeout + 1
            If myBot.loginTimeout >= CLng(frmAdminBot.txtLoginExpire.Text) Then
                myBot.loginTimeout = 0
                myBot.loginCount = 0
            End If
        End If
        
        If frmAdminBot.chkAnnounceChatroom.Value = vbChecked Then
            myBot.announceChatroomCount = myBot.announceChatroomCount + 1
            If myBot.announceChatroomCount >= CLng(frmAdminBot.txtAnnounceInterval.Text) Then
                myBot.announceChatroomCount = 0
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(frmAdminBot.txtAnnounceMessage.Text)
                Else
                    If Trim$(frmAdminBot.txtAnnounceMessage.Text) > vbNullString Then
                        Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & frmAdminBot.txtAnnounceMessage.Text)
                    End If
                End If
            End If
            myBot.announceChatroomCount2 = myBot.announceChatroomCount2 + 1
            If myBot.announceChatroomCount2 >= CLng(frmAdminBot.txtAnnounceInterval2.Text) Then
                myBot.announceChatroomCount2 = 0
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(frmAdminBot.txtAnnounceMessage2.Text)
                Else
                    If Trim$(frmAdminBot.txtAnnounceMessage2.Text) > vbNullString Then
                        Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & frmAdminBot.txtAnnounceMessage2.Text)
                    End If
                End If
            End If
            myBot.announceChatroomCount3 = myBot.announceChatroomCount3 + 1
            If myBot.announceChatroomCount3 >= CLng(frmAdminBot.txtAnnounceInterval3.Text) Then
                myBot.announceChatroomCount3 = 0
                If frmAdminBot.chkAnnounceReg.Value = vbChecked Then
                    Call splitReg(frmAdminBot.txtAnnounceMessage3.Text)
                Else
                    If Trim$(frmAdminBot.txtAnnounceMessage3.Text) > vbNullString Then
                        Call splitAnnounce(frmAdminBot.txtBotName.Text & ": " & frmAdminBot.txtAnnounceMessage3.Text)
                    End If
                End If
            End If
        End If
        
        If frmAdminBot.chkAnnounceGames.Value = vbChecked Then
            myBot.announceGamesCount = myBot.announceGamesCount + 1
            If myBot.announceGamesCount >= CLng(frmAdminBot.txtGameInterval) Then
                myBot.announceGamesCount = 0
                
                For i = LBound(arGames) To UBound(arGames)
                    If arGames(i).opened = True Then
                        If Trim$(frmAdminBot.txtGameMessage.Text) > vbNullString Then
                            Call globalChatRequest("/announcegame " & arGames(i).gameID & " " & frmAdminBot.txtBotName.Text & ": " & frmAdminBot.txtGameMessage.Text)
                            w = GetTickCount
                            Do Until GetTickCount - w >= 50
                                DoEvents
                            Loop
                        End If
                    End If
                Next i
            End If
        End If
    End If
       
        
    'time
    If finishedTime > 59 Then
        Call clientKeepAlive
        finishedTime = 0
        minutes = minutes + 1
    ElseIf finishedTime = 50 Then
        Call clientKeepAlive
    ElseIf finishedTime = 40 Then
        Call clientKeepAlive
    ElseIf finishedTime = 30 Then
        Call clientKeepAlive
    ElseIf finishedTime = 20 Then
        Call clientKeepAlive
    ElseIf finishedTime = 10 Then
        Call clientKeepAlive
    End If
    
    'clear the text in chatroom every hour
    If minutes = 59 And finishedTime = 59 Then
        'clear chatroom
        txtChatroom.Text = vbNullString
        'txtGameChatroom.Text = vbNullString
        txtChatroom.SelColor = &H808000
        'txtGameChatroom.SelColor = &H808000
        'frmServerlist.List1.Clear
        'frmServerlist.List1.AddItem "*Listbox is Automatically Cleared Every Hour.*"
        'txtGameChatroom.SelText = txtGameChatroom.SelText & "*Client Alert: Gameroom Text is Automatically Cleared Every Hour.*" & vbCrLf
        txtChatroom.SelText = txtChatroom.SelText & "*Client Alert: Chatroom Text is Automatically Cleared Every Hour.*" & vbCrLf
        txtChatroom.SelText = txtChatroom.SelText & "*Client Alert: Saving Database and Config...*" & vbCrLf
        'txtGameChatroom.SelStart = Len(txtGameChatroom.Text)
        Call saveUsers(True)
        Call saveUsers
        Call saveConfig
        txtChatroom.SelText = txtChatroom.SelText & "*Client Alert: Saved!*" & vbCrLf
        txtChatroom.SelStart = Len(txtChatroom.Text)
    ElseIf minutes > 59 Then
        minutes = 0
        hours = hours + 1
        If hours Mod CLng(frmPreferences.txtReconnect.Text) = 0 Then
            Call userQuitRequest("Ping Timeout")
        End If
    End If
    
    If hours > 23 Then
        hours = 0
        days = days + 1
    End If
    MDIForm1.StatusBar1.Panels(3).Text = "Session: " & days & "d:" & hours & "h:" & minutes & "m:" & finishedTime & "s"
    Exit Sub
End Sub











Private Sub Timer2_Timer()
    Static i As Long
    
    If inServer = False Then
        i = i + 1
        If i = 100 Then
            i = 0
            frmServerlist.List1.Clear
        End If
        frmServerlist.List1.AddItem ":-" & Time & ": Timed Out! Server may be down. Retrying..."
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        Call MDIForm1.mnuReconnectToServer_Click
    Else
        Timer2.Enabled = False
    End If
End Sub



Public Sub Timer4_Timer()
    Call sessionTime
    timeoutCount = timeoutCount + 1
    If timeoutCount = 30 Then
        Call globalChatRequest("/alivecheck")
    ElseIf timeoutCount = 35 Then
        Call reconnectToServer
    End If
End Sub







Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And vbCtrlMask = vbCtrlMask And KeyCode = 13 Then
        Call btnAnnounce_Click
    ElseIf KeyCode = vbKeyReturn Then
        Call btnChat_Click
    'ElseIf Shift And vbCtrlMask = vbCtrlMask And KeyCode = vbKeyW Then
    '    Call btnWipeOut_Click
    End If
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    Call textboxStuff(Form1.txtChat, KeyAscii)
End Sub

Private Sub txtFindUser_KeyPress(KeyAscii As Integer)
    Call textboxStuff(Form1.txtFindUser, KeyAscii)
End Sub

'Private Sub txtGameChat_Change()
    'If txtGameChat.Text = vbNullString Then
        'btnGameChat.Enabled = False
    'Else
        'btnGameChat.Enabled = True
    'End If
'End Sub

Private Sub txtgameChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim$(txtGameChat.Text) = vbNullString Then Exit Sub
        Call splitRegGame(Trim$(Form1.txtGameChat.Text))
        txtGameChat.Text = vbNullString
    End If
End Sub

Private Sub txtGameChat_KeyPress(KeyAscii As Integer)
    Call textboxStuff(Form1.txtGameChat, KeyAscii)
End Sub






Private Sub txtKickUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call gameChatRequest("/maxusers " & Form1.txtKickUsers.Text)
    End If
End Sub

Private Sub txtKickUsers_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(Form1.txtKickUsers, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub





Private Sub txtQuickBan_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtQuickBan, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


