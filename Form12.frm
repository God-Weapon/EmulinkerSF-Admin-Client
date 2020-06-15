VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBV 
   Caption         =   "View All Database Entries"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   8520
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load All Users"
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkSearchIP 
      Caption         =   "Search by IP Address"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Next"
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin MSComctlLib.ListView lstDBV 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13361
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aliases"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Login"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Entry"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "dummy"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblPer 
      BackStyle       =   0  'Transparent
      Caption         =   "0% Completed"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmDBV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnLoad_Click()
    Dim i As Long
    Dim t As Long
    Dim w As Long
    Dim q As Long
    Dim r As Long
   
    lstDBV.ListItems.Clear
    
    ProgressBar1.Max = entryCount
    ProgressBar1.Value = 0
    lstDBV.Visible = False
    t = GetTickCount
    For i = LBound(ipHead) To UBound(ipHead)
        For q = LBound(ipHead(0).ipSect) To UBound(ipHead(0).ipSect)
            If UBound(ipHead(i).ipSect(q).myEntries) > 0 Then
                For w = LBound(ipHead(i).ipSect(q).myEntries) To UBound(ipHead(i).ipSect(q).myEntries) - 1
                    Call loadEntry(w, i, False, q)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    r = r + 1
                    If r = 1000 Then
                        r = 0
                        lblPer.Caption = CLng((ProgressBar1.Value / ProgressBar1.Max) * 100) & "% Completed"
                        DoEvents
                    End If
                Next w
            End If
        Next q
    Next i

    Me.Caption = lstDBV.ListItems.count & " Database Entries Loaded Successfully!"
    
    ProgressBar1.Value = 0
    lstDBV.Visible = True
    lblPer.Caption = "100% Completed in " & Abs(CSng((GetTickCount - t)) / 1000) & "s"
End Sub

Private Sub Command2_Click()
    Dim str() As String
    Dim i As Long
    Dim w As Long
    Dim q As ListItem
    
    On Error Resume Next
    
    If lstDBV.ListItems.count < 1 Then Exit Sub
    
    If txtFind.Text = vbNullString Then dbPos = 1
    
    If chkSearchIP.Value = vbChecked Then
        Set q = lstDBV.FindItem(txtFind.Text, lvwSubItem, 1, lvwWhole)
        If q Is Nothing Then
            Call MsgBox("Could not find a match!", vbOKOnly, "User Database Search")
        Else
            lstDBV.ListItems.Item(q.Index).EnsureVisible
            lstDBV.ListItems.Item(q.Index).Selected = True
        End If
    Else
        For i = dbPos + 1 To lstDBV.ListItems.count
            str = Split(lstDBV.ListItems(i).Text, ", ")
            For w = 0 To UBound(str)
                If InStr(1, str(w), txtFind.Text, vbTextCompare) > 0 Then
                    lstDBV.ListItems.Item(i).EnsureVisible
                    lstDBV.ListItems.Item(i).Selected = True
                    dbPos = i
                    Exit Sub
                End If
            Next w
            'If i Mod 1000 = 0 Then DoEvents
        Next i
        dbPos = 1
        Call MsgBox("There are no more occurences!", vbOKOnly, "User Database Search")
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = 8640
    Me.Height = 9105
    Me.Top = 0
    Me.Left = 3000
End Sub


Private Sub Form_Resize()
    Call ResizeControls("frmDBV")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub lstDBV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call LV_ColumnSort(lstDBV, ColumnHeader)
End Sub

Private Sub lstDBV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstDBV.ListItems.count > 0 Then
        'creates popup menu when you click on the left mouse button
        If Button = 2 Then PopupMenu MDIForm1.mnuDBVEdit, vbPopupMenuCenterAlign
    End If
End Sub

Private Sub txtFind_Change()
    If lstDBV.ListItems.count < 1 Then Exit Sub
    If txtFind.Text = vbNullString Or Len(txtFind.Text) < 2 Then dbPos = 1
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim w As ListItem
    Dim str() As String
    Dim i As Long
    Dim p As Long
    
    On Error Resume Next
    
    If chkSearchIP.Value = vbChecked And KeyCode = vbKeyReturn Then
        Set w = lstDBV.FindItem(txtFind.Text, lvwSubItem, 1, lvwWhole)
        If w Is Nothing Then
                Call MsgBox("Could not find a match!", vbOKOnly, "User Database Search")
            Exit Sub
        Else
            lstDBV.ListItems.Item(w.Index).EnsureVisible
            lstDBV.ListItems.Item(w.Index).Selected = True
        End If
        Exit Sub
    ElseIf chkSearchIP.Value = vbChecked Then
        Exit Sub
    End If
    
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
    
    If txtFind.Text = vbNullString Then dbPos = 1
    
    For i = dbPos To lstDBV.ListItems.count
        str = Split(lstDBV.ListItems(i).Text, ", ")
        For p = 0 To UBound(str)
            If InStr(1, str(p), txtFind.Text, vbTextCompare) > 0 Then
                lstDBV.ListItems.Item(i).EnsureVisible
                lstDBV.ListItems.Item(i).Selected = True
                dbPos = i
                Exit Sub
            End If
        Next p
        'If i Mod 1000 = 0 Then DoEvents
    Next i
    dbPos = 1
    Call MsgBox("Could not find a match!", vbOKOnly, "User Database Search")

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Call textboxStuff(txtFind, KeyAscii)
End Sub
