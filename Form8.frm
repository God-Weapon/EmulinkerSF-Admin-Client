VERSION 5.00
Begin VB.Form frmMassive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Massive Commands"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   3900
   Begin VB.CommandButton btnClearServer 
      Caption         =   "Clear All Users in Server"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CheckBox chkSilenceAll 
      Caption         =   "Silence All Users who Enter Server"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtMin 
      Height          =   435
      Left            =   3240
      TabIndex        =   4
      Text            =   "5"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton btnSilenceAll 
      Caption         =   "Silence All Users"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton btnKickAll 
      Caption         =   "Kick All Users"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnCloseAll 
      Caption         =   "Close All Games"
      Height          =   435
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnBanAll 
      Caption         =   "Ban All Users"
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmMassive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim g As Long


Private Sub btnBanAll_Click()
    Dim i, w As Long
    
    On Error Resume Next
    
    If MsgBox("Are you sure you want to Ban the entire server?", vbYesNo, "Massive Server Ban?") = vbNo Then Exit Sub
    
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            Call globalChatRequest("/ban " & CStr(arUsers(i).userID) & " " & txtMin.Text)
            w = GetTickCount
            Do Until GetTickCount - w >= 50
                DoEvents
            Loop
        End If
    Next i
    Call globalChatRequest("/announce The server has banned everyone!!!")
End Sub



Private Sub btnClearServer_Click()
    Dim i, w As Long
    
    On Error Resume Next
    
    If MsgBox("Are you sure you want to Clear the entire server?", vbYesNo, "Massive Server Clear?") = vbNo Then Exit Sub
    
    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            Call globalChatRequest("/clear " & CStr(arUsers(i).ip))
            w = GetTickCount
            Do Until GetTickCount - w >= 50
                DoEvents
            Loop
        End If
    Next i
    Call globalChatRequest("/announce The server has cleared everyone!!!")
End Sub

Private Sub btnCloseAll_Click()
    Dim i, w As Long
    
    On Error Resume Next
    
    If MsgBox("Are you sure you want to Close every game in server?", vbYesNo, "Massive Game Close?") = vbNo Then Exit Sub

    For i = LBound(arGames) To UBound(arGames)
        If arGames(i).opened = True Then
            Call globalChatRequest("/closegame " & arGames(i).gameID)
            w = GetTickCount
            Do Until GetTickCount - w >= 20
                DoEvents
            Loop
        End If
    Next i
    Call globalChatRequest("/announce The server has closed all games!!!")
End Sub

Private Sub btnKickAll_Click()
    Dim i, w As Long
    
    On Error Resume Next
    
    If MsgBox("Are you sure you want to Kick the entire server?", vbYesNo, "Massive Server Kick?") = vbNo Then Exit Sub

    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            Call globalChatRequest("/kick " & CStr(arUsers(i).userID))
            w = GetTickCount
            Do Until GetTickCount - w >= 20
                DoEvents
            Loop
        End If
    Next i
    Call globalChatRequest("/announce The server has kicked everyone!")
End Sub

Private Sub btnSilenceAll_Click()
    Dim i, w As Long
    
    On Error Resume Next
    
    If MsgBox("Are you sure you want to Silence the entire server?", vbYesNo, "Massive Server Silence?") = vbNo Then Exit Sub

    For i = LBound(arUsers) To UBound(arUsers)
        If arUsers(i).loggedIn = True Then
            Call globalChatRequest("/silence " & CStr(arUsers(i).userID) & " " & txtMin.Text)
            w = GetTickCount
            Do Until GetTickCount - w >= 20
                DoEvents
            Loop
        End If
    Next i
    Call globalChatRequest("/announce The server has silenced everyone!")
End Sub


Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Load()
    g = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtMin, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub
