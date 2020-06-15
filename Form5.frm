VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Bot"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12600
   Begin VB.Frame fControlGames 
      Caption         =   "Specifics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   99
      Top             =   600
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Frame Frame2 
         Caption         =   "Hit List"
         Height          =   3615
         Left            =   120
         TabIndex        =   127
         Top             =   4440
         Width           =   8415
         Begin MSComctlLib.ListView lstBanIP 
            Height          =   2775
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP Address"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Punishment"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Minutes"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Message"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "dummy"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.CommandButton btnBanIpAdd 
            Caption         =   "Add IP"
            Height          =   315
            Left            =   5760
            TabIndex        =   129
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton btnBanIpRemove 
            Caption         =   "Remove IP"
            Height          =   315
            Left            =   7080
            TabIndex        =   128
            Top             =   3120
            Width           =   1215
         End
      End
      Begin VB.Frame fGameControl 
         Caption         =   "Game/Emulator Type Block"
         Height          =   4095
         Left            =   6480
         TabIndex        =   104
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton btnAddGame 
            Caption         =   "Add "
            Height          =   315
            Left            =   2880
            TabIndex        =   106
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CommandButton btnRemoveGame 
            Caption         =   "Remove "
            Height          =   315
            Left            =   4320
            TabIndex        =   105
            Top             =   3720
            Width           =   1335
         End
         Begin MSComctlLib.ListView lstGame 
            Height          =   3375
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   5953
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name of Game or Emulator"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Message"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "dummy"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Disable Game Hosting"
         Height          =   4095
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   6255
         Begin VB.CommandButton btnRemoveIp 
            Caption         =   "Remove IP"
            Height          =   315
            Left            =   4920
            TabIndex        =   102
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton btnAddIp 
            Caption         =   "Add IP"
            Height          =   315
            Left            =   3600
            TabIndex        =   101
            Top             =   3720
            Width           =   1215
         End
         Begin MSComctlLib.ListView lstDisableHosting 
            Height          =   3375
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   5953
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP Address"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Punishment"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Minutes"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "dummy"
               Object.Width           =   0
            EndProperty
         End
      End
   End
   Begin VB.Frame fAnnouncements 
      Caption         =   "Announcements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   84
      Top             =   600
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Frame Frame4 
         Caption         =   "Announce to Gameroom as soon as the User Creates it"
         Height          =   1095
         Left            =   6240
         TabIndex        =   152
         Top             =   6960
         Width           =   6015
         Begin VB.TextBox txtCreateGame 
            Height          =   315
            Left            =   120
            TabIndex        =   153
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fWelcome 
         Caption         =   "Welcome Messages"
         Height          =   4215
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   12135
         Begin VB.CommandButton btnAddUser 
            Caption         =   "Add User"
            Height          =   315
            Left            =   8520
            TabIndex        =   97
            Top             =   3840
            Width           =   1695
         End
         Begin VB.CommandButton btnRemoveUser 
            Caption         =   "Remove User"
            Height          =   315
            Left            =   10320
            TabIndex        =   96
            Top             =   3840
            Width           =   1695
         End
         Begin MSComctlLib.ListView lstWelcomeMessages 
            Height          =   3495
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP Address"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nick"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Message"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "dummy"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame fGameAnnounce 
         Caption         =   "Gameroom Announce"
         Height          =   1095
         Left            =   120
         TabIndex        =   90
         Top             =   6960
         Width           =   6015
         Begin VB.TextBox txtGameInterval 
            Height          =   315
            Left            =   5160
            MaxLength       =   5
            TabIndex        =   92
            Text            =   "600"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtGameMessage 
            Height          =   315
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (s):"
            Height          =   255
            Left            =   5160
            TabIndex        =   94
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fAutoAnnounce 
         Caption         =   "Chatroom Announce"
         Height          =   2295
         Left            =   120
         TabIndex        =   85
         Top             =   4560
         Width           =   12135
         Begin VB.TextBox txtAnnounceInterval3 
            Height          =   315
            Left            =   11400
            MaxLength       =   5
            TabIndex        =   146
            Text            =   "300"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtAnnounceInterval2 
            Height          =   315
            Left            =   7680
            MaxLength       =   5
            TabIndex        =   145
            Text            =   "300"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtAnnounceMessage3 
            Height          =   1695
            Left            =   8400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   144
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtAnnounceMessage2 
            Height          =   1695
            Left            =   4200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   143
            Top             =   480
            Width           =   4095
         End
         Begin VB.TextBox txtAnnounceMessage 
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtAnnounceInterval 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   86
            Text            =   "300"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (s):"
            Height          =   255
            Index           =   2
            Left            =   10440
            TabIndex        =   148
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (s):"
            Height          =   255
            Index           =   0
            Left            =   6840
            TabIndex        =   147
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (s):"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   88
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.Frame fControlChatroom 
      Caption         =   "Chatroom Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Frame Frame3 
         Caption         =   "Link Control"
         Height          =   2655
         Left            =   8280
         TabIndex        =   132
         Top             =   3000
         Width           =   3975
         Begin VB.TextBox txtLinkSend 
            Height          =   315
            Left            =   3360
            TabIndex        =   149
            Text            =   "2"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtLinkMin 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   141
            Text            =   "5"
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox chkLinkSilenceKick 
            Caption         =   "Silence/Kick"
            Height          =   195
            Left            =   120
            TabIndex        =   139
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkLinkBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   120
            TabIndex        =   138
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox chkLinkKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   120
            TabIndex        =   137
            Top             =   1680
            Width           =   735
         End
         Begin VB.CheckBox chkLinkSilence 
            Caption         =   "Silence"
            Height          =   195
            Left            =   120
            TabIndex        =   136
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtLinkMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   135
            Text            =   "Punished for Link Spamming!"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtLinksInterval 
            Height          =   315
            Left            =   3360
            TabIndex        =   133
            Text            =   "300"
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "How many Link sends?"
            Height          =   255
            Left            =   1560
            TabIndex        =   150
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   142
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3360
            TabIndex        =   140
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label27 
            Caption         =   "How many Seconds before next link send?"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1080
            Width           =   3135
         End
      End
      Begin VB.Frame fGameSpam 
         Caption         =   "Game Spam Control"
         Height          =   2175
         Left            =   8280
         TabIndex        =   116
         Top             =   5880
         Width           =   3975
         Begin VB.TextBox txtGameExpire 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   122
            Text            =   "20"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtGameroomBan 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   121
            Text            =   "5"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtGameroomNum 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   120
            Text            =   "4"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtGameroomMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   119
            Text            =   "Punished for Spamming Games!"
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox chkGameSpamKick 
            Caption         =   "Kick"
            Height          =   255
            Left            =   2880
            TabIndex        =   118
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox chkGameSpamBan 
            Caption         =   "Ban"
            Height          =   255
            Left            =   2880
            TabIndex        =   117
            Top             =   1560
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3360
            TabIndex        =   126
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "In how many seconds?"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "# of Games in a row?"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Frame fWordFilter 
         Caption         =   "Word Filter Control"
         Height          =   5505
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtFilterMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   81
            Text            =   "Punished for Profanity!"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtWordMin 
            Height          =   315
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   80
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin VB.ListBox lstWord 
            Height          =   4350
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   79
            Top             =   960
            Width           =   2415
         End
         Begin VB.CommandButton btnAddWord 
            Caption         =   "Add"
            Height          =   435
            Left            =   2640
            TabIndex        =   78
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CommandButton btnRemoveWord 
            Caption         =   "Remove"
            Height          =   435
            Left            =   2640
            TabIndex        =   77
            Top             =   4920
            Width           =   1215
         End
         Begin VB.CheckBox chkWordKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   76
            Top             =   1440
            Width           =   735
         End
         Begin VB.CheckBox chkWordSilence 
            Caption         =   "Silence"
            Height          =   195
            Left            =   2640
            TabIndex        =   75
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkWordBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   2640
            TabIndex        =   74
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox chkWordSilenceKick 
            Caption         =   "Silence/Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   73
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3240
            TabIndex        =   83
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fLoginControl 
         Caption         =   "Login Control"
         Height          =   2535
         Left            =   8280
         TabIndex        =   61
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtLoginSameIP 
            Height          =   315
            Left            =   3000
            MaxLength       =   5
            TabIndex        =   114
            Text            =   "20"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtLoginNum 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   67
            Text            =   "4"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtLoginMin 
            Height          =   315
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   66
            Text            =   "5"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtLoginMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   65
            Text            =   "Punished for Login Spamming!"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtLoginExpire 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   64
            Text            =   "20"
            Top             =   1440
            Width           =   495
         End
         Begin VB.CheckBox chkLoginKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2880
            TabIndex        =   63
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkLoginBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   2880
            TabIndex        =   62
            Top             =   1440
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum connections from same IP?"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "# of Logins in a row?"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   70
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "In how many seconds?"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   1695
         End
      End
      Begin VB.Frame fAutoKick 
         Caption         =   "Spam Control"
         Height          =   2175
         Left            =   4200
         TabIndex        =   46
         Top             =   3600
         Width           =   3975
         Begin VB.TextBox txtSpamChars 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   55
            Text            =   "32"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtSpamRow 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   54
            Text            =   "3"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtSpamMin 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   53
            Text            =   "10"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtSpamMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   52
            Text            =   "Punished for Spamming!"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtSpamExpire 
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   51
            Text            =   "12"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CheckBox chkSpamSilence 
            Caption         =   "Silence"
            Height          =   195
            Left            =   2640
            TabIndex        =   50
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkSpamKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   49
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkSpamBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   2640
            TabIndex        =   48
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox chkSpamSilenceKick 
            Caption         =   "Silence/Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   47
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "# of Messages in a row?"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3360
            TabIndex        =   58
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "# of Chars in message?"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "In how many seconds?"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Username Filter"
         Height          =   3255
         Left            =   4200
         TabIndex        =   36
         Top             =   240
         Width           =   3975
         Begin VB.ListBox lstUsername 
            Height          =   2205
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   435
            Left            =   2760
            TabIndex        =   42
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Remove"
            Height          =   435
            Left            =   2760
            TabIndex        =   41
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CheckBox chkUsernameKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2760
            TabIndex        =   40
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkUsernameBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   2760
            TabIndex        =   39
            Top             =   1440
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.TextBox txtUsernameMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   38
            Text            =   "Punished for using that Username!"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtUsernameMin 
            Height          =   315
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   37
            Text            =   "5"
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3240
            TabIndex        =   44
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "ALL CAPS"
         Height          =   2175
         Left            =   120
         TabIndex        =   25
         Top             =   5880
         Width           =   3975
         Begin VB.TextBox txtTotalCaps 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Text            =   "8"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtAllCapsMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   31
            Text            =   "Punished for SHOUTING!"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtAllCapsMin 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   30
            Text            =   "10"
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox chkAllCapsSilence 
            Caption         =   "Silence"
            Height          =   195
            Left            =   2640
            TabIndex        =   29
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkAllCapsKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   28
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkAllCapsBan 
            Caption         =   "Ban"
            Height          =   195
            Left            =   2640
            TabIndex        =   27
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox chkAllCapsSilenceKick 
            Caption         =   "Silence/Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   26
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "# of Chars in message?"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3360
            TabIndex        =   33
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Line Wrapper Control"
         Height          =   2175
         Left            =   4200
         TabIndex        =   14
         Top             =   5880
         Width           =   3975
         Begin VB.TextBox txtMaxSpace 
            Height          =   315
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   21
            Text            =   "57"
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox chkLineWrapperBan 
            Caption         =   "Ban"
            Height          =   255
            Left            =   2640
            TabIndex        =   20
            Top             =   1560
            Width           =   735
         End
         Begin VB.CheckBox chkLineWrapperSilence 
            Caption         =   "Silence"
            Height          =   195
            Left            =   2640
            TabIndex        =   19
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkLineWrapperKick 
            Caption         =   "Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   18
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkLineWrapperSilenceKick 
            Caption         =   "Silence/Kick"
            Height          =   195
            Left            =   2640
            TabIndex        =   17
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtLineWrapperMessage 
            Height          =   315
            Left            =   120
            MaxLength       =   56
            TabIndex        =   16
            Text            =   "Punished for Line Wrapping!"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtLineWrapperMin 
            Height          =   315
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   15
            Text            =   "10"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "How many characters until New Line?"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   255
            Left            =   3360
            TabIndex        =   22
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.Frame fAdminMain 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12375
      Begin VB.CheckBox chkCreateGame 
         Caption         =   "Announce on Create Game"
         Height          =   255
         Left            =   4200
         TabIndex        =   155
         Top             =   7800
         Width           =   2295
      End
      Begin VB.CheckBox chkLinkControl 
         Caption         =   "Link Control"
         Height          =   255
         Left            =   2160
         TabIndex        =   151
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CheckBox chkBanIP 
         Caption         =   "Hit List"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6600
         TabIndex        =   131
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CheckBox chkGameDisable 
         Caption         =   "Game Disable Control"
         Height          =   255
         Left            =   6600
         TabIndex        =   113
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CheckBox chkGameSpamControl 
         Caption         =   "Game Spam Control"
         Height          =   255
         Left            =   2160
         TabIndex        =   112
         Top             =   7440
         Width           =   1935
      End
      Begin VB.CheckBox chkGameControl 
         Caption         =   "Game Type Control"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6600
         TabIndex        =   111
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CheckBox chkWelcomeMessages 
         Caption         =   "Welcome Messages"
         Height          =   255
         Left            =   4200
         TabIndex        =   110
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CheckBox chkAnnounceChatroom 
         Caption         =   "Announce Chatroom"
         Height          =   255
         Left            =   4200
         TabIndex        =   109
         Top             =   7080
         Width           =   1935
      End
      Begin VB.CheckBox chkAnnounceGames 
         Caption         =   "Announce Games"
         Height          =   255
         Left            =   4200
         TabIndex        =   108
         Top             =   7440
         Width           =   1935
      End
      Begin VB.TextBox txtBotName 
         Height          =   315
         Left            =   120
         MaxLength       =   32
         TabIndex        =   10
         Text            =   "Admin Bot"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton btnONOFF 
         Caption         =   "Click to Turn Bot ON"
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkAnnounceReg 
         Caption         =   "Use Normal Chat Method "
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkSpamControl 
         Caption         =   "Spam Control"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CheckBox chkUserNameFilter 
         Caption         =   "Username Filter"
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CheckBox chkLoginSpamControl 
         Caption         =   "Login Spam Control"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CheckBox chkAllCaps 
         Caption         =   "ALL CAPS Control"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CheckBox chkWordFilter 
         Caption         =   "Word Filter"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CheckBox chkLineWrapper 
         Caption         =   "Line Wrapper Control"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   7080
         Width           =   2055
      End
      Begin MSComctlLib.ListView lstDamage 
         Height          =   5775
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nick"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Reason"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "dummy"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bot's Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip TabStrip3 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Announcements"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chatroom Control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Specifics"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdminBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddGame_Click()
    Dim str, str1 As String
    Dim lst As ListItem
    
    str = InputBox("Please enter the name of the game or emulator.", "Game/Emulator Control")
    If Trim$(str) = vbNullString Then Exit Sub
    
    str1 = InputBox("Please enter a message to be sent.", "Game/Emulator Control")
    If Trim$(str1) = vbNullString Then Exit Sub
    
    Set lst = lstGame.ListItems.Add(, , Trim$(str))
    lst.SubItems(1) = Trim$(str1)
    lst.SubItems(2) = "#"
End Sub

Private Sub btnAddIp_Click()
    Dim str, str1, str2, str3 As String
    Dim lst As ListItem
    Dim sip() As String
    Dim defa As String
    Dim i As Long
    On Error Resume Next
    
    defa = "127.0.0.1"
    Do
        str = InputBox("Please enter the IP Address you want to ban. Use * for a range.", "Disable Hosting", defa)
        If Trim$(str) = vbNullString Then Exit Sub
        
        sip = Split(str, ".")
        If UBound(sip) = 3 Then
            If Len(sip(0)) = 2 Then
                str = "0" & str
            ElseIf Len(sip(0)) = 1 Then
                str = "00" & str
            End If
            Exit Do
        Else
            defa = "Invalid IP Format! Please try again."
        End If
    Loop


    
    str1 = InputBox("Please enter his Username. NOTE: This is for your use only.", "Disable Hosting")
    If Trim$(str1) = vbNullString Then Exit Sub
    
    
    Do
        str3 = InputBox("Would you like to Close his game or Ban him?" & vbCrLf & vbCrLf & "1=Close Game  2=Ban", "Disable Hosting", "2")
        str3 = Trim$(str3)
        If str3 = vbNullString Then Exit Sub
        If str3 = "1" Or str3 = "2" Then Exit Do
    Loop
    
    
    defa = "10"
    Do Until str3 = "1"
        str2 = InputBox("Please enter for how long (min)." & vbCrLf & vbCrLf & "1 to 30,000", "Disable Hosting", defa)
        If Trim$(str2) = vbNullString Then Exit Sub
        
        If CLng(str2) > 0 And CLng(str2) <= 30000 Then
            For i = 1 To Len(str2)
                If Asc(Mid$(str2, i, 1)) > 57 Or Asc(Mid$(str2, i, 1)) < 47 Then
                    defa = "Invalid choice! Please try again."
                    Exit For
                End If
            Next i
            
            If i = Len(str2) + 1 Then
                Exit Do
            End If
        Else
            defa = "Invalid choice! Please try again."
        End If
        
    Loop
    
    Set lst = lstDisableHosting.ListItems.Add(, , Trim$(str))
    lst.SubItems(1) = Trim$(str1)
    If str3 = "1" Then
        lst.SubItems(2) = "Close Game"
        lst.SubItems(3) = "Null"
    Else
        lst.SubItems(2) = "Ban"
        lst.SubItems(3) = Trim$(str2)
    End If
    lst.SubItems(4) = "#"
End Sub

Private Sub btnAddUser_Click()
    Dim str, str1, str2 As String
    Dim lst As ListItem
    
    str = InputBox("Please enter the IP address of the user.", "Welcome Message")
    If Trim$(str) = vbNullString Then Exit Sub
    
    str1 = InputBox("Please enter the name of the user. NOTE: This is for your use only.", "Welcome Message")
    If Trim$(str1) = vbNullString Then Exit Sub
    
    str2 = InputBox("Please enter the Welcome Message for the user.", "Welcome Message")
    If Trim$(str2) = vbNullString Then Exit Sub
    
    Set lst = lstWelcomeMessages.ListItems.Add(, , Trim$(str))
    lst.SubItems(1) = Trim$(str1)
    lst.SubItems(2) = Trim$(str2)
    lst.SubItems(3) = "#"
End Sub

Private Sub btnAddWord_Click()
    Dim str As String
    Dim temp As String
    Dim s As String
    Dim i As Long
    Dim w As Long
    Dim iRep(0 To 4) As String
    
    iRep(0) = "!"
    iRep(1) = "1"
    iRep(2) = "|"
    iRep(3) = "j"
    iRep(4) = "l"
    
    str = InputBox("Please enter a word you would like to add to the filter. Do not enter the plural form of the word. This program will take care of that.", "Word Filter")
    
    str = LCase$(Trim$(str))
    If str = vbNullString Then Exit Sub
    
    lstWord.AddItem str
    lstWord.AddItem str & "s"
    lstWord.AddItem str & "z"
    
    
    If InStr(1, str, "i", vbBinaryCompare) > 0 Then
        For i = 0 To UBound(iRep)
            s = Replace$(str, "i", iRep(i))
            lstWord.AddItem s
            lstWord.AddItem s & "s"
            lstWord.AddItem s & "z"
        Next i
    End If
End Sub

Private Sub btnBanIpAdd_Click()
    Dim str, str1, str2, str3, str4 As String
    Dim lst As ListItem
    Dim sip() As String
    Dim defa As String
    Dim i As Long
    On Error Resume Next
    
    defa = "127.0.0.1 OR 127.0.*.*"
    Do
        str = InputBox("Please enter the IP Address you want to ban. Use * for a range.", "Hit List", defa)
        If Trim$(str) = vbNullString Then Exit Sub
        
        sip = Split(str, ".")
        If UBound(sip) = 3 Then
            If Len(sip(0)) = 2 Then
                str = "0" & str
            ElseIf Len(sip(0)) = 1 Then
                str = "00" & str
            End If
            Exit Do
        Else
            defa = "Invalid IP Format! Please try again."
        End If
    Loop
    
    str1 = InputBox("Please enter his Username. NOTE: This is for your use only.", "Hit List")
    If Trim$(str1) = vbNullString Then Exit Sub


    Do
        str4 = InputBox("Would you like to Silence or Ban him?" & vbCrLf & vbCrLf & "1=Silence  2=Ban", "Hit List", "2")
        str4 = Trim$(str4)
        If str4 = vbNullString Then Exit Sub
        If str4 = "1" Or str4 = "2" Then Exit Do
    Loop

    defa = "10"
    Do
        str2 = InputBox("Please enter for how long (min)." & vbCrLf & vbCrLf & "1 to 30,000", "Hit List", defa)
        If Trim$(str2) = vbNullString Then Exit Sub
        
        If CLng(str2) > 0 And CLng(str2) <= 30000 Then
            For i = 1 To Len(str2)
                If Asc(Mid$(str2, i, 1)) > 57 Or Asc(Mid$(str2, i, 1)) < 47 Then
                    defa = "Invalid choice! Please try again."
                    Exit For
                End If
            Next i
            
            If i = Len(str2) + 1 Then
                Exit Do
            End If
        Else
            defa = "Invalid choice! Please try again."
        End If
    Loop

    str3 = InputBox("Please enter a message.", "Hit List")
    If Trim$(str3) = vbNullString Then Exit Sub
    
    Set lst = lstBanIP.ListItems.Add(, , Trim$(str))
    lst.SubItems(1) = Trim$(str1)
    If str4 = "1" Then
        lst.SubItems(2) = "Silence"
    Else
        lst.SubItems(2) = "Ban"
    End If
    lst.SubItems(3) = Trim$(str2)
    lst.SubItems(4) = Trim$(str3)
    lst.SubItems(5) = "#"
End Sub

Private Sub btnBanIpRemove_Click()
    If lstBanIP.ListItems.count > 0 Then
        lstBanIP.ListItems.remove (lstBanIP.SelectedItem.Index)
    End If
End Sub

Private Sub btnRemoveIp_Click()
    If lstDisableHosting.ListItems.count > 0 Then
        lstDisableHosting.ListItems.remove (lstDisableHosting.SelectedItem.Index)
    End If
End Sub

Private Sub btnRemoveUser_Click()
    If lstWelcomeMessages.ListItems.count > 0 Then
        lstWelcomeMessages.ListItems.remove (lstWelcomeMessages.SelectedItem.Index)
    End If
End Sub

Private Sub btnRemoveWord_Click()
    If lstWord.ListCount > 0 Then
        If lstWord.ListIndex = -1 Then
            lstWord.ListIndex = 0
        End If
        lstWord.RemoveItem (lstWord.ListIndex)
    End If
End Sub

Private Sub btnRemoveGame_Click()
    If lstGame.ListItems.count > 0 Then
        lstGame.ListItems.remove (lstGame.SelectedItem.Index)
    End If
End Sub


Private Sub chkAllCapsBan_Click()
    If chkAllCapsBan.Value = vbUnchecked _
    And chkAllCapsKick.Value = vbUnchecked _
    And chkAllCapsSilence.Value = vbUnchecked _
    And chkAllCapsSilenceKick.Value = vbUnchecked Then
        chkAllCapsBan.Value = vbChecked
    ElseIf chkAllCapsBan.Value = vbChecked Then
        chkAllCapsSilence.Value = vbUnchecked
        chkAllCapsSilenceKick.Value = vbUnchecked
        chkAllCapsKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkAllCapsKick_Click()
    If chkAllCapsKick.Value = vbUnchecked _
    And chkAllCapsSilenceKick.Value = vbUnchecked _
    And chkAllCapsSilence.Value = vbUnchecked _
    And chkAllCapsBan.Value = vbUnchecked Then
        chkAllCapsKick.Value = vbChecked
    ElseIf chkAllCapsKick.Value = vbChecked Then
        chkAllCapsSilence.Value = vbUnchecked
        chkAllCapsBan.Value = vbUnchecked
        chkAllCapsSilenceKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkAllCapsSilence_Click()
    If chkAllCapsSilence.Value = vbUnchecked _
    And chkAllCapsKick.Value = vbUnchecked _
    And chkAllCapsSilenceKick.Value = vbUnchecked _
    And chkAllCapsBan.Value = vbUnchecked Then
        chkAllCapsSilence.Value = vbChecked
    ElseIf chkAllCapsSilence.Value = vbChecked Then
        chkAllCapsSilenceKick.Value = vbUnchecked
        chkAllCapsBan.Value = vbUnchecked
        chkAllCapsKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkAllCapsSilenceKick_Click()
    If chkAllCapsSilenceKick.Value = vbUnchecked _
    And chkAllCapsKick.Value = vbUnchecked _
    And chkAllCapsSilence.Value = vbUnchecked _
    And chkAllCapsBan.Value = vbUnchecked Then
        chkAllCapsSilenceKick.Value = vbChecked
    ElseIf chkAllCapsSilenceKick.Value = vbChecked Then
        chkAllCapsSilence.Value = vbUnchecked
        chkAllCapsBan.Value = vbUnchecked
        chkAllCapsKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkGameSpamBan_Click()
    If chkGameSpamBan.Value = vbUnchecked _
    And chkGameSpamKick.Value = vbUnchecked Then
        chkGameSpamBan.Value = vbChecked
    ElseIf chkGameSpamBan.Value = vbChecked Then
        chkGameSpamKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkGameSpamKick_Click()
    If chkGameSpamKick.Value = vbUnchecked _
    And chkGameSpamBan.Value = vbUnchecked Then
        chkGameSpamKick.Value = vbChecked
    ElseIf chkGameSpamKick.Value = vbChecked Then
        chkGameSpamBan.Value = vbUnchecked
    End If
End Sub

Private Sub chkLineWrapperBan_Click()
    If chkLineWrapperBan.Value = vbUnchecked _
    And chkLineWrapperKick.Value = vbUnchecked _
    And chkLineWrapperSilence.Value = vbUnchecked _
    And chkLineWrapperSilenceKick.Value = vbUnchecked Then
        chkLineWrapperBan.Value = vbChecked
    ElseIf chkLineWrapperBan.Value = vbChecked Then
        chkLineWrapperSilence.Value = vbUnchecked
        chkLineWrapperSilenceKick.Value = vbUnchecked
        chkLineWrapperKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLineWrapperKick_Click()
    If chkLineWrapperKick.Value = vbUnchecked _
    And chkLineWrapperSilenceKick.Value = vbUnchecked _
    And chkLineWrapperSilence.Value = vbUnchecked _
    And chkLineWrapperBan.Value = vbUnchecked Then
        chkLineWrapperSilenceKick.Value = vbChecked
    ElseIf chkLineWrapperKick.Value = vbChecked Then
        chkLineWrapperSilence.Value = vbUnchecked
        chkLineWrapperBan.Value = vbUnchecked
        chkLineWrapperSilenceKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLineWrapperSilence_Click()
    If chkLineWrapperSilence.Value = vbUnchecked _
    And chkLineWrapperSilenceKick.Value = vbUnchecked _
    And chkLineWrapperKick.Value = vbUnchecked _
    And chkLineWrapperBan.Value = vbUnchecked Then
        chkLineWrapperSilence.Value = vbChecked
    ElseIf chkLineWrapperSilence.Value = vbChecked Then
        chkLineWrapperSilenceKick.Value = vbUnchecked
        chkLineWrapperBan.Value = vbUnchecked
        chkLineWrapperKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLineWrapperSilenceKick_Click()
    If chkLineWrapperSilenceKick.Value = vbUnchecked _
    And chkLineWrapperKick.Value = vbUnchecked _
    And chkLineWrapperSilence.Value = vbUnchecked _
    And chkLineWrapperBan.Value = vbUnchecked Then
        chkLineWrapperSilenceKick.Value = vbChecked
    ElseIf chkLineWrapperSilenceKick.Value = vbChecked Then
        chkLineWrapperSilence.Value = vbUnchecked
        chkLineWrapperBan.Value = vbUnchecked
        chkLineWrapperKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLinkBan_Click()
    If chkLinkBan.Value = vbUnchecked _
    And chkLinkKick.Value = vbUnchecked _
    And chkLinkSilence.Value = vbUnchecked _
    And chkLinkSilenceKick.Value = vbUnchecked Then
        chkLinkSilenceKick.Value = vbChecked
    ElseIf chkLinkBan.Value = vbChecked Then
        chkLinkSilence.Value = vbUnchecked
        chkLinkSilenceKick.Value = vbUnchecked
        chkLinkKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLinkKick_Click()
    If chkLinkKick.Value = vbUnchecked _
    And chkLinkBan.Value = vbUnchecked _
    And chkLinkSilence.Value = vbUnchecked _
    And chkLinkSilenceKick.Value = vbUnchecked Then
        chkLinkSilenceKick.Value = vbChecked
    ElseIf chkLinkKick.Value = vbChecked Then
        chkLinkSilence.Value = vbUnchecked
        chkLinkSilenceKick.Value = vbUnchecked
        chkLinkBan.Value = vbUnchecked
    End If
End Sub

Private Sub chkLinkSilence_Click()
    If chkLinkSilence.Value = vbUnchecked _
    And chkLinkBan.Value = vbUnchecked _
    And chkLinkKick.Value = vbUnchecked _
    And chkLinkSilenceKick.Value = vbUnchecked Then
        chkLinkSilenceKick.Value = vbChecked
    ElseIf chkLinkSilence.Value = vbChecked Then
        chkLinkBan.Value = vbUnchecked
        chkLinkSilenceKick.Value = vbUnchecked
        chkLinkKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLinkSilenceKick_Click()
    If chkLinkSilenceKick.Value = vbUnchecked _
    And chkLinkKick.Value = vbUnchecked _
    And chkLinkSilence.Value = vbUnchecked _
    And chkLinkBan.Value = vbUnchecked Then
        chkLinkSilenceKick.Value = vbChecked
    ElseIf chkLinkSilenceKick.Value = vbChecked Then
        chkLinkSilence.Value = vbUnchecked
        chkLinkBan.Value = vbUnchecked
        chkLinkKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLoginBan_Click()
    If chkLoginBan.Value = vbUnchecked _
    And chkLoginKick.Value = vbUnchecked Then
        chkLoginBan.Value = vbChecked
    ElseIf chkLoginBan.Value = vbChecked Then
        chkLoginKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkLoginKick_Click()
    If chkLoginKick.Value = vbUnchecked _
    And chkLoginBan.Value = vbUnchecked Then
        chkLoginKick.Value = vbChecked
    ElseIf chkLoginKick.Value = vbChecked Then
        chkLoginBan.Value = vbUnchecked
    End If
End Sub

Private Sub chkSpamBan_Click()
    If chkSpamBan.Value = vbUnchecked _
    And chkSpamKick.Value = vbUnchecked _
    And chkSpamSilence.Value = vbUnchecked _
    And chkSpamSilenceKick.Value = vbUnchecked Then
        chkSpamSilenceKick.Value = vbChecked
    ElseIf chkSpamBan.Value = vbChecked Then
        chkSpamSilence.Value = vbUnchecked
        chkSpamSilenceKick.Value = vbUnchecked
        chkSpamKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkSpamKick_Click()
    If chkSpamKick.Value = vbUnchecked _
    And chkSpamBan.Value = vbUnchecked _
    And chkSpamSilence.Value = vbUnchecked _
    And chkSpamSilenceKick.Value = vbUnchecked Then
        chkSpamSilenceKick.Value = vbChecked
    ElseIf chkSpamKick.Value = vbChecked Then
        chkSpamSilence.Value = vbUnchecked
        chkSpamSilenceKick.Value = vbUnchecked
        chkSpamBan.Value = vbUnchecked
    End If
End Sub

Private Sub chkSpamSilence_Click()
    If chkSpamSilence.Value = vbUnchecked _
    And chkSpamBan.Value = vbUnchecked _
    And chkSpamKick.Value = vbUnchecked _
    And chkSpamSilenceKick.Value = vbUnchecked Then
        chkSpamSilenceKick.Value = vbChecked
    ElseIf chkSpamSilence.Value = vbChecked Then
        chkSpamBan.Value = vbUnchecked
        chkSpamSilenceKick.Value = vbUnchecked
        chkSpamKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkSpamSilenceKick_Click()
    If chkSpamSilenceKick.Value = vbUnchecked _
    And chkSpamKick.Value = vbUnchecked _
    And chkSpamSilence.Value = vbUnchecked _
    And chkSpamBan.Value = vbUnchecked Then
        chkSpamSilenceKick.Value = vbChecked
    ElseIf chkSpamSilenceKick.Value = vbChecked Then
        chkSpamSilence.Value = vbUnchecked
        chkSpamBan.Value = vbUnchecked
        chkSpamKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkUsernameBan_Click()
    If chkUsernameBan.Value = vbUnchecked _
    And chkUsernameKick.Value = vbUnchecked Then
        chkUsernameBan.Value = vbChecked
    ElseIf chkUsernameBan.Value = vbChecked Then
        chkUsernameKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkUsernameKick_Click()
    If chkUsernameKick.Value = vbUnchecked _
    And chkUsernameBan.Value = vbUnchecked Then
        chkUsernameKick.Value = vbChecked
    ElseIf chkUsernameKick.Value = vbChecked Then
        chkUsernameBan.Value = vbUnchecked
    End If
End Sub

Private Sub chkWordBan_Click()
    If chkWordBan.Value = vbUnchecked _
    And chkWordKick.Value = vbUnchecked _
    And chkWordSilence.Value = vbUnchecked _
    And chkWordSilenceKick.Value = vbUnchecked Then
        chkWordBan.Value = vbChecked
    ElseIf chkWordBan.Value = vbChecked Then
        chkWordSilence.Value = vbUnchecked
        chkWordSilenceKick.Value = vbUnchecked
        chkWordKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkWordKick_Click()
    If chkWordKick.Value = vbUnchecked _
    And chkWordSilenceKick.Value = vbUnchecked _
    And chkWordSilence.Value = vbUnchecked _
    And chkWordBan.Value = vbUnchecked Then
        chkWordKick.Value = vbChecked
    ElseIf chkWordKick.Value = vbChecked Then
        chkWordSilence.Value = vbUnchecked
        chkWordBan.Value = vbUnchecked
        chkWordSilenceKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkWordSilence_Click()
    If chkWordSilence.Value = vbUnchecked _
    And chkWordKick.Value = vbUnchecked _
    And chkWordSilenceKick.Value = vbUnchecked _
    And chkWordBan.Value = vbUnchecked Then
        chkWordSilence.Value = vbChecked
    ElseIf chkWordSilence.Value = vbChecked Then
        chkWordSilenceKick.Value = vbUnchecked
        chkWordBan.Value = vbUnchecked
        chkWordKick.Value = vbUnchecked
    End If
End Sub

Private Sub chkWordSilenceKick_Click()
    If chkWordSilenceKick.Value = vbUnchecked _
    And chkWordKick.Value = vbUnchecked _
    And chkWordSilence.Value = vbUnchecked _
    And chkWordBan.Value = vbUnchecked Then
        chkWordSilenceKick.Value = vbChecked
    ElseIf chkWordSilenceKick.Value = vbChecked Then
        chkWordSilence.Value = vbUnchecked
        chkWordBan.Value = vbUnchecked
        chkWordKick.Value = vbUnchecked
    End If
End Sub

Private Sub Command1_Click()
    Dim str As String
    
    str = InputBox("Please enter a username you would like to block." & vbCrLf & vbCrLf & " If you put a * after the name, you will block any name that contains that.  For example: putting supr*  will block anyone from having SupraFast, Suprbot, and etc.", "Username Filter")
    If Trim$(str) = vbNullString Then Exit Sub
    lstUsername.AddItem Trim$(str)
End Sub



Public Sub btnONOFF_Click()
    Static c As Byte
    
    If c = 0 Then
        myBot.botStatus = True
        btnONOFF.Caption = "Click to Turn Bot OFF"
        MDIForm1.StatusBar1.Panels(6).Text = "Admin Bot: ON"
        Call resetBotValues
        c = 1
    Else
        myBot.botStatus = False
        btnONOFF.Caption = "Click to Turn Bot ON"
        MDIForm1.StatusBar1.Panels(6).Text = "Admin Bot: OFF"
        c = 0
    End If
End Sub

Private Sub Command3_Click()
    If lstUsername.ListCount > 0 Then
        If lstUsername.ListIndex = -1 Then
            lstUsername.ListIndex = 0
        End If
        lstUsername.RemoveItem (lstUsername.ListIndex)
    End If
End Sub

Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strBuff As String
    Dim splitwords() As String
    Dim str() As String
    Dim temp As String
    Dim wordlist As ListItem
    
    On Error Resume Next
    
    fAnnouncements.Top = 480
    fAnnouncements.Left = 120
    
    fAdminMain.Top = 480
    fAdminMain.Left = 120
    
    fControlChatroom.Top = 480
    fControlChatroom.Left = 120
    
    fControlGames.Top = 480
    fControlGames.Left = 120
    
    

    'read from it
    Open App.Path & "\config.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, strBuff
        
        If Left$(strBuff, Len("botName=")) = "botName=" Then
            txtBotName.Text = Right$(strBuff, Len(strBuff) - Len("botName="))
        ElseIf Left$(strBuff, Len("ALLCAPS=")) = "ALLCAPS=" Then
            chkAllCaps.Value = Right$(strBuff, Len(strBuff) - Len("ALLCAPS="))
            'chkAllCaps.Value = vbUnchecked
        ElseIf Left$(strBuff, Len("spamRow=")) = "spamRow=" Then
            txtSpamRow.Text = Right$(strBuff, Len(strBuff) - Len("spamRow="))
        ElseIf Left$(strBuff, Len("spamMin=")) = "spamMin=" Then
            txtSpamMin.Text = Right$(strBuff, Len(strBuff) - Len("spamMin="))
        ElseIf Left$(strBuff, Len("filterCheck=")) = "filterCheck=" Then
            chkWordFilter.Value = Right$(strBuff, Len(strBuff) - Len("filterCheck="))
        ElseIf Left$(strBuff, Len("activateCheck=")) = "activateCheck=" Then
            chkSpamControl.Value = Right$(strBuff, Len(strBuff) - Len("activateCheck="))
        ElseIf Left$(strBuff, Len("annouonceChatroomCheck=")) = "annouonceChatroomCheck=" Then
            chkAnnounceChatroom.Value = Right$(strBuff, Len(strBuff) - Len("annouonceChatroomCheck="))
        ElseIf Left$(strBuff, Len("spamMin=")) = "spamMin=" Then
            txtSpamMin.Text = Right$(strBuff, Len(strBuff) - Len("spamMin="))
        ElseIf Left$(strBuff, Len("spamChars=")) = "spamChars=" Then
            txtSpamChars.Text = Right$(strBuff, Len(strBuff) - Len("spamChars="))
        ElseIf Left$(strBuff, Len("allcaps=")) = "callcaps=" Then
            chkAllCaps.Value = Right$(strBuff, Len(strBuff) - Len("callcaps="))
        ElseIf Left$(strBuff, Len("spamMessage=")) = "spamMessage=" Then
            txtSpamMessage.Text = Right$(strBuff, Len(strBuff) - Len("spamMessage="))
        ElseIf Left$(strBuff, Len("announceInterval=")) = "announceInterval=" Then
            txtAnnounceInterval.Text = Right$(strBuff, Len(strBuff) - Len("announceInterval="))
        ElseIf Left$(strBuff, Len("announceInterval2=")) = "announceInterval2=" Then
            txtAnnounceInterval2.Text = Right$(strBuff, Len(strBuff) - Len("announceInterval2="))
        ElseIf Left$(strBuff, Len("announceInterval3=")) = "announceInterval3=" Then
            txtAnnounceInterval3.Text = Right$(strBuff, Len(strBuff) - Len("announceInterval3="))
        
        ElseIf Left$(strBuff, Len("announceMessage=")) = "announceMessage=" Then
            txtAnnounceMessage.Text = Right$(strBuff, Len(strBuff) - Len("announceMessage="))
            txtAnnounceMessage.Text = Replace$(txtAnnounceMessage.Text, ";&|", vbCrLf)
        ElseIf Left$(strBuff, Len("announceMessage2=")) = "announceMessage2=" Then
            txtAnnounceMessage2.Text = Right$(strBuff, Len(strBuff) - Len("announceMessage2="))
            txtAnnounceMessage2.Text = Replace$(txtAnnounceMessage2.Text, ";&|", vbCrLf)
        ElseIf Left$(strBuff, Len("announceMessage3=")) = "announceMessage3=" Then
            txtAnnounceMessage3.Text = Right$(strBuff, Len(strBuff) - Len("announceMessage3="))
            txtAnnounceMessage3.Text = Replace$(txtAnnounceMessage3.Text, ";&|", vbCrLf)
        
        ElseIf Left$(strBuff, Len("gameAnnounceInterval=")) = "gameAnnounceInterval=" Then
            txtGameInterval.Text = Right$(strBuff, Len(strBuff) - Len("gameAnnounceInterval="))
        ElseIf Left$(strBuff, Len("gameAnnounceMessage=")) = "gameAnnounceMessage=" Then
            txtGameMessage.Text = Right$(strBuff, Len(strBuff) - Len("gameAnnounceMessage="))
        ElseIf Left$(strBuff, Len("announceGameCheck=")) = "announceGameCheck=" Then
            chkAnnounceGames.Value = Right$(strBuff, Len(strBuff) - Len("announceGameCheck="))
        ElseIf Left$(strBuff, Len("filterWords=")) = "filterWords=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("filterWords="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                lstWord.AddItem str(i)
            Next i
        ElseIf Left$(strBuff, Len("filterUsername=")) = "filterUsername=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("filterUsername="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                lstUsername.AddItem str(i)
            Next i
        ElseIf Left$(strBuff, Len("usernameCheck=")) = "usernameCheck=" Then
            chkUserNameFilter.Value = Right$(strBuff, Len(strBuff) - Len("usernameCheck="))
        ElseIf Left$(strBuff, Len("filterMessage=")) = "filterMessage=" Then
            txtFilterMessage.Text = Right$(strBuff, Len(strBuff) - Len("filterMessage="))
        ElseIf Left$(strBuff, Len("gameControlCheck=")) = "gameControlCheck=" Then
            chkGameControl.Value = Right$(strBuff, Len(strBuff) - Len("gameControlCheck="))
        ElseIf Left$(strBuff, Len("controlledgames=")) = "controlledgames=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("controlledgames="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                Set wordlist = lstGame.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                i = i + 1
            Next i
        ElseIf Left$(strBuff, Len("disabledList=")) = "disabledList=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("disabledList="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                Set wordlist = lstDisableHosting.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                wordlist.SubItems(2) = str(i + 2)
                wordlist.SubItems(3) = str(i + 3)
                i = i + 3
            Next i
        ElseIf Left$(strBuff, Len("bannedList=")) = "bannedList=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("bannedList="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                Set wordlist = lstBanIP.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                wordlist.SubItems(2) = str(i + 2)
                wordlist.SubItems(3) = str(i + 3)
                wordlist.SubItems(4) = str(i + 4)
                i = i + 4
            Next i
        ElseIf Left$(strBuff, Len("activateWelcome=")) = "activateWelcome=" Then
            chkWelcomeMessages.Value = Right$(strBuff, Len(strBuff) - Len("activateWelcome="))
        ElseIf Left$(strBuff, Len("welcomeMessages=")) = "welcomeMessages=" Then
            temp = Right$(strBuff, Len(strBuff) - Len("welcomeMessages="))
            ReDim str(0)
            str = Split(temp, ";&,")
            For i = 0 To UBound(str)
                Set wordlist = lstWelcomeMessages.ListItems.Add(, , Trim$(str(i)))
                wordlist.SubItems(1) = str(i + 1)
                i = i + 1
                wordlist.SubItems(2) = str(i + 1)
                i = i + 1
            Next i
        ElseIf Left$(strBuff, Len("gamespammessage=")) = "gamespammessage=" Then
            txtGameroomMessage.Text = Right$(strBuff, Len(strBuff) - Len("gamespammessage="))
        ElseIf Left$(strBuff, Len("gameSpamControlCheck=")) = "gameSpamControlCheck=" Then
            chkGameSpamControl.Value = Right$(strBuff, Len(strBuff) - Len("gameSpamControlCheck="))
        ElseIf Left$(strBuff, Len("gamespamlimit=")) = "gamespamlimit=" Then
            txtGameroomBan.Text = Right$(strBuff, Len(strBuff) - Len("gamespamlimit="))
        ElseIf Left$(strBuff, Len("gamespamconsecutive=")) = "gamespamconsecutive=" Then
            txtGameroomNum.Text = Right$(strBuff, Len(strBuff) - Len("gamespamconsecutive="))
        ElseIf Left$(strBuff, Len("silenceAll=")) = "silenceAll=" Then
            frmMassive.chkSilenceAll.Value = Right$(strBuff, Len(strBuff) - Len("silenceAll="))
        ElseIf Left$(strBuff, Len("loginSpamControl=")) = "loginSpamControl=" Then
            chkLoginSpamControl.Value = Right$(strBuff, Len(strBuff) - Len("loginSpamControl="))
        ElseIf Left$(strBuff, Len("loginNum=")) = "loginNum=" Then
            txtLoginNum.Text = Right$(strBuff, Len(strBuff) - Len("loginNum="))
        ElseIf Left$(strBuff, Len("loginMin=")) = "loginMin=" Then
            txtLoginMin.Text = Right$(strBuff, Len(strBuff) - Len("loginMin="))
        ElseIf Left$(strBuff, Len("loginMessage=")) = "loginMessage=" Then
            txtLoginMessage.Text = Right$(strBuff, Len(strBuff) - Len("loginMessage="))
        ElseIf Left$(strBuff, Len("announceReg=")) = "announceReg=" Then
            chkAnnounceReg.Value = Right$(strBuff, Len(strBuff) - Len("announceReg="))
        ElseIf Left$(strBuff, Len("disableHosting=")) = "disableHosting=" Then
            chkGameDisable.Value = Right$(strBuff, Len(strBuff) - Len("disableHosting="))
        ElseIf Left$(strBuff, Len("spamExpire=")) = "spamExpire=" Then
            txtSpamExpire.Text = Right$(strBuff, Len(strBuff) - Len("spamExpire="))
        ElseIf Left$(strBuff, Len("gameExpire=")) = "gameExpire=" Then
            txtGameExpire.Text = Right$(strBuff, Len(strBuff) - Len("gameExpire="))
        ElseIf Left$(strBuff, Len("loginExpire=")) = "loginExpire=" Then
            txtLoginExpire.Text = Right$(strBuff, Len(strBuff) - Len("loginExpire="))
        ElseIf Left$(strBuff, Len("totalCaps=")) = "totalCaps=" Then
            txtTotalCaps.Text = Right$(strBuff, Len(strBuff) - Len("totalCaps="))
        ElseIf Left$(strBuff, Len("filterMin=")) = "filterMin=" Then
            txtWordMin.Text = Right$(strBuff, Len(strBuff) - Len("filterMin="))
        ElseIf Left$(strBuff, Len("maxSpace=")) = "maxSpace=" Then
            txtMaxSpace.Text = Right$(strBuff, Len(strBuff) - Len("maxSpace="))
            
        ElseIf Left$(strBuff, Len("chkWordBan=")) = "chkWordBan=" Then
            chkWordBan.Value = Right$(strBuff, Len(strBuff) - Len("chkWordBan="))
        ElseIf Left$(strBuff, Len("chkWordKick=")) = "chkWordKick=" Then
            chkWordKick.Value = Right$(strBuff, Len(strBuff) - Len("chkWordKick="))
        ElseIf Left$(strBuff, Len("chkWordSilence=")) = "chkWordSilence=" Then
            chkWordSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkWordSilence="))
        ElseIf Left$(strBuff, Len("chkWordSilenceKick=")) = "chkWordSilenceKick=" Then
            chkWordSilenceKick.Value = Right$(strBuff, Len(strBuff) - Len("chkWordSilenceKick="))
        ElseIf Left$(strBuff, Len("chkSpamBan=")) = "chkSpamBan=" Then
            chkSpamBan.Value = Right$(strBuff, Len(strBuff) - Len("chkSpamBan="))
        ElseIf Left$(strBuff, Len("chkSpamKick=")) = "chkSpamKick=" Then
            chkSpamKick.Value = Right$(strBuff, Len(strBuff) - Len("chkSpamKick="))
        ElseIf Left$(strBuff, Len("chkSpamSilence=")) = "chkSpamSilence=" Then
            chkSpamSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkSpamSilence="))
        ElseIf Left$(strBuff, Len("chkSpamSilenceKick=")) = "chkSpamSilenceKick=" Then
            chkSpamSilenceKick.Value = Right$(strBuff, Len(strBuff) - Len("chkSpamSilenceKick="))
        ElseIf Left$(strBuff, Len("chkLineWrapperBan=")) = "chkLineWrapperBan=" Then
            chkLineWrapperBan.Value = Right$(strBuff, Len(strBuff) - Len("chkLineWrapperBan="))
        ElseIf Left$(strBuff, Len("chkLineWrapperKick=")) = "chkLineWrapperKick=" Then
            chkLineWrapperKick.Value = Right$(strBuff, Len(strBuff) - Len("chkLineWrapperKick="))
        ElseIf Left$(strBuff, Len("chkLineWrapperSilence=")) = "chkLineWrapperSilence=" Then
            chkLineWrapperSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkLineWrapperSilence="))
        ElseIf Left$(strBuff, Len("chkLineWrapperSilenceKick=")) = "chkLineWrapperSilenceKick=" Then
            chkLineWrapperSilenceKick.Value = Right$(strBuff, Len(strBuff) - Len("chkLineWrapperSilenceKick="))
        ElseIf Left$(strBuff, Len("lineWrapperMessage=")) = "lineWrapperMessage=" Then
            txtLineWrapperMessage.Text = Right$(strBuff, Len(strBuff) - Len("lineWrapperMessage="))
        ElseIf Left$(strBuff, Len("lineWrapperMin=")) = "lineWrapperMin=" Then
            txtLineWrapperMin.Text = Right$(strBuff, Len(strBuff) - Len("lineWrapperMin="))
        ElseIf Left$(strBuff, Len("chkAllCapsBan=")) = "chkAllCapsBan=" Then
            chkAllCapsBan.Value = Right$(strBuff, Len(strBuff) - Len("chkAllCapsBan="))
        ElseIf Left$(strBuff, Len("chkAllCapsKick=")) = "chkAllCapsKick=" Then
            chkAllCapsKick.Value = Right$(strBuff, Len(strBuff) - Len("chkAllCapsKick="))
        ElseIf Left$(strBuff, Len("chkAllCapsSilence=")) = "chkAllCapsSilence=" Then
            chkAllCapsSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkAllCapsSilence="))
        ElseIf Left$(strBuff, Len("chkAllCapsSilence=")) = "chkAllCapsSilence=" Then
            chkAllCapsSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkAllCapsSilence="))
        ElseIf Left$(strBuff, Len("chkAllCapsSilenceKick=")) = "chkAllCapsSilenceKick=" Then
            chkAllCapsSilenceKick.Value = Right$(strBuff, Len(strBuff) - Len("chkAllCapsSilenceKick="))
        ElseIf Left$(strBuff, Len("allCapsMessage=")) = "allCapsMessage=" Then
            txtAllCapsMessage.Text = Right$(strBuff, Len(strBuff) - Len("allCapsMessage="))
        ElseIf Left$(strBuff, Len("allCapsMin=")) = "allCapsMin=" Then
            txtAllCapsMin.Text = Right$(strBuff, Len(strBuff) - Len("allCapsMin="))
        ElseIf Left$(strBuff, Len("chkLoginBan=")) = "chkLoginBan=" Then
            chkLoginBan.Value = Right$(strBuff, Len(strBuff) - Len("chkLoginBan="))
        ElseIf Left$(strBuff, Len("chkLoginKick=")) = "chkLoginKick=" Then
            chkLoginKick.Value = Right$(strBuff, Len(strBuff) - Len("chkLoginKick="))
        ElseIf Left$(strBuff, Len("chkUsernameBan=")) = "chkUsernameBan=" Then
            chkUsernameBan.Value = Right$(strBuff, Len(strBuff) - Len("chkUsernameBan="))
        ElseIf Left$(strBuff, Len("chkUsernameKick=")) = "chkUsernameKick=" Then
            chkUsernameKick.Value = Right$(strBuff, Len(strBuff) - Len("chkUsernameKick="))
        ElseIf Left$(strBuff, Len("usernameMessage=")) = "usernameMessage=" Then
            txtUsernameMessage.Text = Right$(strBuff, Len(strBuff) - Len("usernameMessage="))
        ElseIf Left$(strBuff, Len("usernameMin=")) = "usernameMin=" Then
            txtUsernameMin.Text = Right$(strBuff, Len(strBuff) - Len("usernameMin="))
        ElseIf Left$(strBuff, Len("chkGameSpamBan=")) = "chkGameSpamBan=" Then
            chkGameSpamBan.Value = Right$(strBuff, Len(strBuff) - Len("chkGameSpamBan="))
        ElseIf Left$(strBuff, Len("chkGameSpamKick=")) = "chkGameSpamKick=" Then
            chkGameSpamKick.Value = Right$(strBuff, Len(strBuff) - Len("chkGameSpamKick="))
        ElseIf Left$(strBuff, Len("loginSameIP=")) = "loginSameIP=" Then
            txtLoginSameIP.Text = Right$(strBuff, Len(strBuff) - Len("loginSameIP="))
        ElseIf Left$(strBuff, Len("chkLineWrapper=")) = "chkLineWrapper=" Then
            chkLineWrapper.Value = Right$(strBuff, Len(strBuff) - Len("chkLineWrapper="))
        ElseIf Left$(strBuff, Len("chkBanIP=")) = "chkBanIP=" Then
            chkBanIP.Value = Right$(strBuff, Len(strBuff) - Len("chkBanIP="))
        ElseIf Left$(strBuff, Len("linkSend=")) = "linkSend=" Then
            txtLinkSend.Text = Right$(strBuff, Len(strBuff) - Len("linkSend="))
        ElseIf Left$(strBuff, Len("linkMessage=")) = "linkMessage=" Then
            txtLinkMessage.Text = Right$(strBuff, Len(strBuff) - Len("linkMessage="))
        ElseIf Left$(strBuff, Len("linkMin=")) = "linkMin=" Then
            txtLinkMin.Text = Right$(strBuff, Len(strBuff) - Len("linkMin="))
        ElseIf Left$(strBuff, Len("linksInterval=")) = "linksInterval=" Then
            txtLinksInterval.Text = Right$(strBuff, Len(strBuff) - Len("linksInterval="))
        ElseIf Left$(strBuff, Len("chkLinkBan=")) = "chkLinkBan=" Then
            chkLinkBan.Value = Right$(strBuff, Len(strBuff) - Len("chkLinkBan="))
        ElseIf Left$(strBuff, Len("chkLinkSilence=")) = "chkLinkSilence=" Then
            chkLinkSilence.Value = Right$(strBuff, Len(strBuff) - Len("chkLinkSilence="))
        ElseIf Left$(strBuff, Len("chkLinkKick=")) = "chkLinkKick=" Then
            chkLinkKick.Value = Right$(strBuff, Len(strBuff) - Len("chkLinkKick="))
        ElseIf Left$(strBuff, Len("chkLinkSilenceKick=")) = "chkLinkSilenceKick=" Then
            chkLinkSilenceKick.Value = Right$(strBuff, Len(strBuff) - Len("chkLinkSilenceKick="))
        ElseIf Left$(strBuff, Len("chkLink=")) = "chkLink=" Then
            chkLinkControl.Value = Right$(strBuff, Len(strBuff) - Len("chkLink="))
        ElseIf Left$(strBuff, Len("chkCreateGameAnnounce=")) = "chkCreateGameAnnounce=" Then
            chkCreateGame.Value = Right$(strBuff, Len(strBuff) - Len("chkCreateGameAnnounce="))
        ElseIf Left$(strBuff, Len("announceCreateMessage=")) = "announceCreateMessage=" Then
            txtCreateGame.Text = Right$(strBuff, Len(strBuff) - Len("announceCreateMessage="))
        
  
        End If
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


Private Sub txtAllCapsMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtAllCapsMessage, KeyAscii)
End Sub

Private Sub txtAllCapsMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtAllCapsMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtCreateGame_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtCreateGame, KeyAscii)
End Sub

Private Sub txtLineWrapperMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLineWrapperMessage, KeyAscii)
End Sub

Private Sub txtLineWrapperMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtLineWrapperMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLinkMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLinkMessage, KeyAscii)
End Sub

Private Sub txtLinkMin_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLinkMin, KeyAscii)
End Sub

Private Sub txtLinkSend_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLinkSend, KeyAscii)
End Sub

Private Sub txtLinksInterval_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtLinksInterval, KeyAscii)
End Sub

Private Sub txtLoginSameIP_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtLoginSameIP, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMaxSpace_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtMaxSpace, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUsernameMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtUsernameMessage, KeyAscii)
End Sub

Private Sub txtUsernameMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtUsernameMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWordMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtWordMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpamChars_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtSpamChars, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtSpamExpire_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtSpamExpire, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpamMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtSpamMessage, KeyAscii)
End Sub


Private Sub txtSpamMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtSpamMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpamRow_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtSpamRow, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTotalCaps_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtTotalCaps, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLoginExpire_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtLoginExpire, KeyAscii)
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




Private Sub txtGameInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtGameInterval, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtGameMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtGameMessage, KeyAscii)
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




Private Sub txtFilterMessage_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtFilterMessage, KeyAscii)
End Sub

Private Sub txtAnnounceInterval_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtAnnounceInterval, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
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


Private Sub txtGameExpire_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtGameExpire, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub lstWelcomeMessages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstWelcomeMessages, ColumnHeader)
End Sub





Private Sub TabStrip3_Click()
    If TabStrip3.SelectedItem.Index = 1 Then
        'main
        Call fixBotFrames(0)
    ElseIf TabStrip3.SelectedItem.Index = 2 Then
        'announcements
        Call fixBotFrames(1)
    ElseIf TabStrip3.SelectedItem.Index = 3 Then
        'chatroom control
        Call fixBotFrames(2)
    ElseIf TabStrip3.SelectedItem.Index = 4 Then
        'gameroom control
        Call fixBotFrames(3)
    End If
End Sub


Private Sub lstDamage_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstDamage, ColumnHeader)
End Sub

Private Sub lstDamage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstDamage.ListItems.count > 0 Then
        If Button = 2 Then PopupMenu MDIForm1.mnuDamage, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub lstGame_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstGame, ColumnHeader)
End Sub
