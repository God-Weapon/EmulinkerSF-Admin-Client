VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      Caption         =   "Gameroom"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   9015
      Begin VB.CheckBox chkBeep 
         Caption         =   "Beep on Join"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtGameWelcomeMessage 
         Height          =   315
         Left            =   120
         MaxLength       =   2999
         TabIndex        =   16
         Text            =   "https://github.com/God-Weapon"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "MOTD:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Admin"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   9015
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear IP Address"
         Height          =   315
         Left            =   7440
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtClear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5640
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "127.0.0.1"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkAlertOthers 
         Caption         =   "Announce User Alerts to All"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkStartBot 
         Caption         =   "Start Bot when entering server"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   9015
      Begin VB.TextBox txtQuit 
         Height          =   315
         Left            =   120
         MaxLength       =   128
         TabIndex        =   13
         Text            =   "https://github.com/God-Weapon"
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkRoomOnConnect 
         Caption         =   "Create Room on Connect"
         Height          =   255
         Left            =   6720
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtRoomOnConnect 
         Height          =   315
         Left            =   3840
         MaxLength       =   128
         TabIndex        =   7
         Text            =   "*Chat (not game)"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Quit Message:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Name:"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chatroom"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9015
      Begin VB.TextBox txtReconnect 
         Height          =   315
         Left            =   7560
         TabIndex        =   19
         Text            =   "24"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtMaxChars 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Text            =   "256"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkTimeStamps 
         Caption         =   "Use Time Stamps"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkShowJoin 
         Caption         =   "Show Join/Exit Messages"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkShowOpen 
         Caption         =   "Show Create/Close Messages"
         Height          =   255
         Left            =   4680
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reconnect to server in how many hours?"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Characters to send to chatroom?"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub
Private Sub btnClear_Click()
    Call globalChatRequest("/clear " & txtClear.Text)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub txtClear_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtClear, KeyAscii)
End Sub

Private Sub txtGameWelcomeMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtGameWelcomeMessage, KeyAscii)
End Sub



Private Sub Form_Load()
    Dim strBuff As String
    On Error Resume Next
    
    chkBeep.ToolTipText = "Default System Beep will occur when users join the room."
    
    'read from it
    Open App.Path & "\config.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, strBuff
        If Left$(strBuff, 5) = "beep=" Then
            chkBeep.Value = Right$(strBuff, Len(strBuff) - 5)
        ElseIf Left$(strBuff, 19) = "gameWelcomeMessage=" Then
            txtGameWelcomeMessage.Text = Right$(strBuff, Len(strBuff) - 19)
        ElseIf Left$(strBuff, 5) = "quit=" Then
            txtQuit.Text = Right$(strBuff, Len(strBuff) - 5)
        ElseIf Left$(strBuff, 9) = "showOpen=" Then
            chkShowOpen.Value = Right$(strBuff, Len(strBuff) - 9)
        ElseIf Left$(strBuff, 9) = "showJoin=" Then
            chkShowJoin.Value = Right$(strBuff, Len(strBuff) - 9)
        ElseIf Left$(strBuff, 11) = "timeStamps=" Then
            chkTimeStamps.Value = Right$(strBuff, Len(strBuff) - 11)
        ElseIf Left$(strBuff, Len("alertothers=")) = "alertothers=" Then
            chkAlertOthers.Value = Right$(strBuff, Len(strBuff) - Len("alertothers="))
        ElseIf Left$(strBuff, 9) = "startbot=" Then
            chkStartBot.Value = Right$(strBuff, Len(strBuff) - 9)
        ElseIf Left$(strBuff, 12) = "roomConnect=" Then
            chkRoomOnConnect.Value = Right$(strBuff, Len(strBuff) - 12)
        ElseIf Left$(strBuff, 16) = "roomConnectName=" Then
            txtRoomOnConnect.Text = Right$(strBuff, Len(strBuff) - 16)
        ElseIf Left$(strBuff, Len("maximumChars=")) = "maximumChars=" Then
            txtMaxChars.Text = Right$(strBuff, Len(strBuff) - Len("maximumChars="))
        ElseIf Left$(strBuff, Len("reconnectHours=")) = "reconnectHours=" Then
            txtReconnect.Text = Right$(strBuff, Len(strBuff) - Len("reconnectHours="))
        End If
    Loop
    Close #1
End Sub

Private Sub txtQuit_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtQuit, KeyAscii)
End Sub

Private Sub txtMaxChars_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtMaxChars, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRoomOnConnect_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtRoomOnConnect, KeyAscii)
End Sub


