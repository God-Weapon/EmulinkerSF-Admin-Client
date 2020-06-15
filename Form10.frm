VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServerlist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Screen"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9630
   Begin VB.Frame frame2 
      Caption         =   "Login Info:"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6960
         Top             =   5400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.CheckBox chkLoading 
         Caption         =   "Connect to Server after loading?"
         Height          =   255
         Left            =   5880
         TabIndex        =   14
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtUsernameStoreage 
         Height          =   2835
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Form10.frx":11F6
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "Connect"
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   5160
         MaxLength       =   31
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtServerIp 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Text            =   "127.0.0.1:27888"
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cmbConnectionType 
         Height          =   315
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   5520
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   5880
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblPer 
         BackStyle       =   0  'Transparent
         Caption         =   "0% Complete"
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   5520
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection:"
         Height          =   255
         Index           =   0
         Left            =   7920
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP Address:"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmServerlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub btnExit_Click()
    Dim i As Long
    btnExit.Enabled = False
    
    If inServer = True Then
        Form1.Timer2.Enabled = False
        Call userQuitRequest(Trim$(frmPreferences.txtQuit.Text))
        iQuit = True
        Call fixFramesButtons(0)
    Else
        Form1.Timer2.Enabled = False
        Call fixFramesButtons(0)
        List1.AddItem "*Disconnected (" & Time & ")*"
        List1.TopIndex = List1.ListCount - 1
    End If
End Sub


Public Sub btnLogin_Click()
    Dim here As Boolean
    Dim i As Long
    Dim serverList As ListItem
    
    Form1.Timer2.Enabled = True
    'List1.Clear
    
    'initial
    Call fixFramesButtons(0)
    Form1.txtChatroom.Text = vbNullString
    'Login
    Call fixFramesButtons(7)
    List1.AddItem ":-" & Time & ": Connecting..."
    List1.TopIndex = List1.ListCount - 1
    
    If Trim$(txtServerIp.Text) = vbNullString Then
        List1.AddItem "Please choose an IP Address [server]!"
        List1.TopIndex = List1.ListCount - 1
        Call btnExit_Click
        Exit Sub
    ElseIf Trim$(txtUsername.Text) = vbNullString Then
        List1.AddItem "Please choose a Username!"
        List1.TopIndex = List1.ListCount - 1
        Call btnExit_Click
        Exit Sub
    End If
    
    txtUsername.Text = Replace$(txtUsername.Text, ",", ";", 1, -1, vbTextCompare)
        
    'Call saveConfig
    
    Dim str() As String
    If InStr(1, txtServerIp.Text, ":", vbTextCompare) < 1 Then
        txtServerIp.Text = txtServerIp.Text & ":27888"
    End If
    str = Split(Trim$(txtServerIp.Text), ":")

    serverIP = str(0)
        
    Form1.Timer4.Enabled = True
    Winsock1.RemotePort = str(1)
    Winsock1.RemoteHost = str(0)
    Winsock1.SendData entryMsg
End Sub







Private Sub Form_Activate()
    Static c As Byte
    Dim i As Long
    

    
    If c = 0 Then
        Call loadUsers
        If frmServerlist.chkLoading.Value = vbChecked Then
            Call frmServerlist.btnLogin_Click
        End If
        c = 1
    End If
End Sub


Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim serverList As ListItem
    Dim str() As String
    Dim strBuff As String
    
    On Error Resume Next
    
    cmbConnectionType.AddItem "LAN"
    cmbConnectionType.AddItem "Excellent"
    cmbConnectionType.AddItem "Good"
    cmbConnectionType.AddItem "Average"
    cmbConnectionType.AddItem "Low"
    cmbConnectionType.AddItem "Bad"
    cmbConnectionType.Text = "Good"

    'read from it
    Open App.Path & "\config.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, strBuff
        If Left$(strBuff, 3) = "ip=" Then
            txtServerIp.Text = Right$(strBuff, Len(strBuff) - Len("ip="))
        ElseIf Left$(strBuff, 9) = "username=" Then
            txtUsername.Text = Right$(strBuff, Len(strBuff) - Len("username="))
        ElseIf Left$(strBuff, 11) = "connection=" Then
            If Right$(strBuff, Len(strBuff) - Len("connection=")) = "1" Then
                cmbConnectionType.Text = "LAN"
            ElseIf Right$(strBuff, Len(strBuff) - Len("connection=")) = "2" Then
                cmbConnectionType.Text = "Excellent"
            ElseIf Right$(strBuff, Len(strBuff) - Len("connection=")) = "3" Then
                cmbConnectionType.Text = "Good"
            ElseIf Right$(strBuff, Len(strBuff) - Len("connection=")) = "4" Then
                cmbConnectionType.Text = "Average"
            ElseIf Right$(strBuff, Len(strBuff) - Len("connection=")) = "5" Then
                cmbConnectionType.Text = "Low"
            ElseIf Right$(strBuff, Len(strBuff) - Len("connection=")) = "6" Then
                cmbConnectionType.Text = "Bad"
            Else
                cmbConnectionType.Text = "Good"
            End If
        ElseIf Left$(strBuff, Len("connectLoad=")) = "connectLoad=" Then
            chkLoading.Value = Right$(strBuff, Len(strBuff) - Len("connectLoad="))
        ElseIf Left$(strBuff, Len("usernameStoreage=")) = "usernameStoreage=" Then
            txtUsernameStoreage.Text = Right$(strBuff, Len(strBuff) - Len("usernameStoreage="))
            txtUsernameStoreage.Text = Replace$(txtUsernameStoreage.Text, ";&|", vbCrLf)
        End If
    Loop
    Close #1
    
    
    
    
    
    
    
        With IconData

        .cbSize = Len(IconData)
        ' The length of the NOTIFYICONDATA type
        
        .hIcon = Me.Icon
        ' A reference to the form's icon
        
        .hwnd = Me.hwnd
        ' hWnd of the form
        
        .szTip = emulatorPass & Chr(0)
        ' Tooltip string delimited with a null character
        
        .uCallbackMessage = WM_MOUSEMOVE
        ' The icon we're placing will send messages to the MouseMove event
        
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ' It will have message handling and a tooltip
        
        .uID = vbNull
        ' uID is not used by VB, so it's set to a Null value

    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim msg As Long

    msg = x / Screen.TwipsPerPixelX

    ' The message is passed to the X value
    
    ' You must set your form's ScaleMode property to pixels in order to get the correct message
    
    If msg = WM_LBUTTONDBLCLK Then
        ' The user has double-clicked your icon
        Call MDIForm1.mnuShow_Click
        ' Show the window
    
    ElseIf msg = WM_RBUTTONDOWN Then
        ' Right-click
        
        PopupMenu MDIForm1.mnuPopup
        ' Popup the menu
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.WindowState = vbMinimized
    Else
        Unload Me
    End If
End Sub

Private Sub txtServerIp_KeyPress(KeyAscii As Integer)
    Dim ch As String
    
    Call textboxStuff(txtServerIp, KeyAscii)
    ch = Chr$(KeyAscii)
    If Not ((ch >= "0" And ch <= "9" Or ch = "." Or ch = ":" Or ch = Chr(8) Or KeyAscii = 1 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)) Then
        KeyAscii = 0
    End If
End Sub




Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtUsername, KeyAscii)
End Sub




Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim str As String
    
    On Error Resume Next
    
    'Check the socket
    If bytesTotal < 1 Then Exit Sub
    
    If step = 1 Then
        timeoutCount = 0
        Winsock1.GetData str, vbString
        If Left$(str, Len("HELLOD00D")) = "HELLOD00D" Then
            step = 2
            portNum = CLng(Trim$(Mid$(str, Len("HELLOD00D") + 1, Len(str) - Len("HELLOD00D"))))
            frmServerlist.List1.AddItem ":-" & Time & ": Server Response switch to Port#: " & CStr(portNum)
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
            frmServerlist.Winsock1.RemotePort = portNum
            frmServerlist.Winsock1.RemoteHost = serverIP
            Close #3
            Open App.Path & "\EmulinkerSF_Logs\chat.txt" For Append As #3
                Print #3, "***********New Session - " & Time & " " & Date & "*************"
            Close #3
            Call userLoginInformation(frmServerlist.txtUsername.Text)
        ElseIf Left$(str, 3) = "TOO" Then
            frmServerlist.List1.AddItem Time & "Server is Full!"
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
            Form1.Timer2.Enabled = True
        ElseIf Left$(str, 4) = "PONG" Then
            frmServerlist.List1.AddItem Time & "PONG"
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        Else
            frmServerlist.List1.AddItem Time & "Unknown Message: " & str
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        End If
    Else
        Winsock1.GetData myBuff, vbArray + vbByte
        timeoutCount = 0
        Call parseData
    End If
End Sub


