VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Bot Commands"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   5865
   Begin VB.CommandButton btnGetBot 
      Caption         =   "Get Current Bots"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton btnAlertsOn 
      Caption         =   "Alerts ON"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton btnAlertsOff 
      Caption         =   "Alerts OFF"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnBotOff 
      Caption         =   "Bot OFF"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnBotOn 
      Caption         =   "Bot ON"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin MSComctlLib.ListView lstBots 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton btnDate 
      Caption         =   "Date"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton btnTime 
      Caption         =   "Time"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton btnIP 
      Caption         =   "Server IP"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton btnReconnect 
      Caption         =   "Reconnect"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   5640
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   2760
   End
End
Attribute VB_Name = "frmRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnAlertsOff_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";alerts off " & lstBots.SelectedItem.Text)
    End If
End Sub

Private Sub btnBotOff_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";bot off " & lstBots.SelectedItem.Text)
    End If
End Sub

Private Sub btnBotOn_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";bot on " & lstBots.SelectedItem.Text)
    End If
End Sub

Private Sub btnDate_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";date")
    End If
End Sub

Private Sub btnIP_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";ip")
    End If
End Sub

Private Sub btnQuit_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";quit " & lstBots.SelectedItem.Text)
    End If
End Sub

Private Sub btnReconnect_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";reconnect " & lstBots.SelectedItem.Text)
    End If
End Sub


Private Sub btnAlertsOn_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";alerts on " & lstBots.SelectedItem.Text)
    End If
End Sub

Private Sub btnTime_Click()
    If lstBots.ListItems.count > 0 Then
        Call globalChatRequest(";time")
    End If
End Sub

Private Sub btnGetBot_Click()
    Call globalChatRequest(";whoisalive")
End Sub

Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Load()
    Dim vv As ListItem

    lstBots.ListItems.Clear
    Set vv = lstBots.ListItems.Add(, , "*")
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        lstBots.ListItems.Clear
        Dim vv As ListItem
    
        lstBots.ListItems.Clear
        Set vv = lstBots.ListItems.Add(, , "*")
        Me.Hide
    Else
        Unload Me
    End If
End Sub

'*************************************************************














