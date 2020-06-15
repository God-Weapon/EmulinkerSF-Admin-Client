VERSION 5.00
Begin VB.Form frmDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   8415
   Begin VB.Frame fUserDb 
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   6975
      End
      Begin VB.Frame fUserinfo 
         Caption         =   "User Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7935
         Begin VB.TextBox txtDBAliases 
            Height          =   2655
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   2760
            Width           =   7695
         End
         Begin VB.CommandButton btnDBEntryDelete 
            Caption         =   "Mark for Delete"
            Height          =   315
            Left            =   6240
            TabIndex        =   10
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton btnDBEntrySave 
            Caption         =   "Save Changes"
            Height          =   315
            Left            =   6240
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton btnDBEntryRefresh 
            Caption         =   "Refresh"
            Height          =   315
            Left            =   6240
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtDBGames 
            Height          =   1395
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   5760
            Width           =   7695
         End
         Begin VB.CommandButton btnDBEntryNoDelete 
            Caption         =   "Unmark for Delete"
            Height          =   315
            Left            =   6240
            TabIndex        =   6
            Top             =   2280
            Width           =   1575
         End
         Begin VB.ListBox lstDB 
            Height          =   2400
            Left            =   3840
            TabIndex        =   5
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton btnDBEntryReset 
            Caption         =   "Reset"
            Height          =   315
            Left            =   6240
            TabIndex        =   4
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton btnDBShow 
            Caption         =   "Show Chatroom"
            Height          =   315
            Left            =   6240
            TabIndex        =   3
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Aliases:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lblDBIPAddress 
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblDBTotalLogins 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Logins:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblDBTotalGames 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Games:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label lblDBTotalAliases 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Aliases:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblDBDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Date of Login:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2160
            Width           =   3615
         End
         Begin VB.Label lblDBAvgPing 
            BackStyle       =   0  'Transparent
            Caption         =   "Average Ping:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblDBFDate 
            BackStyle       =   0  'Transparent
            Caption         =   "First Date of Login:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Created Games:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   5520
            Width           =   1215
         End
         Begin VB.Label lblDBBotHits 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bot Hits:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.CommandButton btnFindNext 
         Caption         =   "Find Next"
         Height          =   315
         Left            =   7200
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Users In Database: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnDBEntryDelete_Click()
    dbEdit = True
    ipHead(tempHeadPos).ipSect(tempSect1Pos).myEntries(tempDBPos).remove = True
    lstDB.AddItem "Marked entry for delete."
    lstDB.TopIndex = lstDB.ListCount - 1
End Sub

Private Sub btnDBEntryRefresh_Click()
    Dim str() As String
    
    str = Split(fUserinfo.Caption, " - ")
    If UBound(str) = -1 Then
        Call showDBEntry(tempDBPos, ipHead(tempHeadPos).ipSect(tempSect1Pos).myEntries(tempDBPos).ip, tempHeadPos, tempSect1Pos)
    Else
        Call showDBEntry(tempDBPos, str(0), tempHeadPos, tempSect1Pos)
    End If
    
    lstDB.AddItem "Refreshed Entry."
    lstDB.TopIndex = lstDB.ListCount - 1
End Sub

Public Sub btnDBShow_Click()
    Dim str() As String
    
    str = Split(frmUserlist.lstUserlist.SelectedItem.Text, " [is] ")
    
    Call splitAnnounce(str(0))
    Call splitAnnounce(lblDBFDate.Caption)
    Call splitAnnounce(lblDBDate.Caption)
    Call splitAnnounce(lblDBAvgPing.Caption)
    Call splitAnnounce(lblDBTotalAliases.Caption)
    Call splitAnnounce(lblDBTotalGames.Caption)
    Call splitAnnounce(lblDBTotalLogins.Caption)
    Call splitAnnounce(lblDBBotHits.Caption)
End Sub

Private Sub btnFindNext_Click()
    Dim i As Long
    Dim found As Boolean
    Dim sip() As String
    Dim ip As String
    Dim sect1 As Long
    Dim head As Long
    Dim r As Long
    On Error Resume Next
    
    txtFind.Enabled = False
    btnFindNext.Enabled = False
    txtFind.Text = Trim$(txtFind.Text)
    'split ip sectors
    sip = Split(txtFind.Text, ".")
    If Len(sip(0)) = 2 Then
        txtFind.Text = "0" & txtFind.Text
    ElseIf Len(sip(0)) = 1 Then
        txtFind.Text = "00" & txtFind.Text
    End If
    
    head = sip(0)
    sect1 = sip(1)
    Do While i <= UBound(ipHead(head).ipSect(sect1).myEntries)
        'check for ip match
        If StrComp(ipHead(head).ipSect(sect1).myEntries(i).ip, txtFind.Text, vbBinaryCompare) = 0 Then
            found = True
            tempDBPos = i
            tempHeadPos = head
            tempSect1Pos = sect1
            Call showDBEntry(tempDBPos, "Results for search", tempHeadPos, tempSect1Pos)
            txtFind.Enabled = True
            btnFindNext.Enabled = True
            Exit Sub
        End If
        i = i + 1
        
        r = r + 1
        If r = 1000 Then
            r = 0
            DoEvents
        End If
    Loop

    txtFind.Enabled = True
    btnFindNext.Enabled = True
    Call MsgBox("No IP Address match is found!", vbOKOnly, "Database Alert!")
End Sub

Private Sub btnDBEntrySave_Click()
    dbEdit = True
    lstDB.AddItem "Saved Entry."
    lstDB.TopIndex = lstDB.ListCount - 1
    ipHead(tempHeadPos).ipSect(tempSect1Pos).myEntries(tempDBPos).aliases = Replace$(txtDBAliases.Text, ", ", ",", 1, -1, vbTextCompare)
    ipHead(tempHeadPos).ipSect(tempSect1Pos).myEntries(tempDBPos).gamesPlayed = Replace$(txtDBGames.Text, ", ", ",", 1, -1, vbTextCompare)
End Sub

Private Sub btnDBEntryNoDelete_Click()
    dbEdit = True
    ipHead(tempHeadPos).ipSect(tempSect1Pos).myEntries(tempDBPos).remove = False
    lstDB.AddItem "Unmarked entry for delete."
    lstDB.TopIndex = lstDB.ListCount - 1
End Sub

Private Sub btnDBEntryReset_Click()
    dbEdit = True
    Call showDBEntryReset(tempDBPos, tempHeadPos, tempSect1Pos)
    lstDB.AddItem "Entry has been reset."
    lstDB.TopIndex = lstDB.ListCount - 1
End Sub

Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
        
    If Shift And vbCtrlMask = vbCtrlMask Then
        Exit Sub
    ElseIf KeyCode >= 1 And KeyCode <= 9 Then
        Exit Sub
    ElseIf KeyCode = &HC Or KeyCode = &HD Then
        Exit Sub
    ElseIf KeyCode >= &H10 And KeyCode <= &H14 Then
        Exit Sub
    ElseIf KeyCode = &H1B Then
        Exit Sub
    ElseIf KeyCode >= &H20 And KeyCode <= &H2F Then
        Exit Sub
    ElseIf KeyCode = &H90 Then
        Exit Sub
    ElseIf KeyCode >= &H70 And KeyCode <= &H7F Then
        Exit Sub
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtFind, KeyAscii)
End Sub
