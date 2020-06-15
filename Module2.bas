Attribute VB_Name = "UserDb"
Option Explicit

Public tempDBPos As Long
Public tempHeadPos As Long
Public tempSect1Pos As Long
Public entryCount As Long

Public Type Entries
    aliases As String
    ip As String
    
    avgPing As Long
    lastLogin As String
    firstLogin As String
    gamesPlayed As String
    numOfLogins As Long
    numOfGames As Long
    numOfNames As Long
    numOfBotHits As Long
    
    remove As Boolean
End Type

Public ipHead(0 To 255) As IP_HEAD

Public Type IP_SECTOR
    myEntries() As Entries
End Type

Public Type IP_HEAD
    ipSect(0 To 255) As IP_SECTOR
End Type


Sub saveUsers(Optional bak As Boolean = False)
    Dim i As Long
    Dim r As Long
    Dim w As Long
    Dim q As Long
    Dim t As Long
    On Error Resume Next
              
        Close #4
        If bak = True Then
            Call FileCopy(App.Path & "\EmulinkerSF_Logs\user.txt", App.Path & "\EmulinkerSF_Logs\user.bak")
            frmServerlist.List1.AddItem "Backup DB Saved!:"
            frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
            Exit Sub
        End If
        
        MDIForm1.Enabled = False
        'save to file
        t = GetTickCount
        frmServerlist.List1.AddItem Time & ": Saving Database..."
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        
        
        
            Open App.Path & "\EmulinkerSF_Logs\user.txt" For Output As #4
        
            frmServerlist.ProgressBar1.Max = entryCount - 1
            frmServerlist.ProgressBar1.Value = 0
            For i = LBound(ipHead) To UBound(ipHead)
                For q = LBound(ipHead(0).ipSect) To UBound(ipHead(0).ipSect)
                    If UBound(ipHead(i).ipSect(q).myEntries) > 0 Then
                        For w = LBound(ipHead(i).ipSect(q).myEntries) To UBound(ipHead(i).ipSect(q).myEntries) - 1
                            If ipHead(i).ipSect(q).myEntries(w).remove = False Then
                                Print #4, ipHead(i).ipSect(q).myEntries(w).ip & ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).aliases & _
                                ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).lastLogin & ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).avgPing & _
                                ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).numOfLogins & ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).numOfGames & _
                                ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).firstLogin & ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).gamesPlayed & _
                                ChrW$(&H2) & ipHead(i).ipSect(q).myEntries(w).numOfBotHits
                            End If
                            frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Value + 1
                            r = r + 1
                            If r = 1000 Then
                                r = 0
                                frmServerlist.lblPer.Caption = CLng((frmServerlist.ProgressBar1.Value / frmServerlist.ProgressBar1.Max) * 100) & "% Complete"
                                DoEvents
                                'Sleep (1)
                            End If
                        Next w
                    End If
                Next q
            Next i
        Close #4
        
        frmServerlist.lblPer.Caption = "100% Complete in " & Abs(CSng((GetTickCount - t)) / 1000) & "s"
        frmServerlist.ProgressBar1.Value = 0
        frmServerlist.List1.AddItem "Finished"
        frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
        MDIForm1.Enabled = True
End Sub

Function checkUsers(id As Long, head As Long, sect1 As Long) As String
    Dim i As Long
    Dim w As Long
    Dim str() As String
    Dim strT() As String
    Dim found As Boolean
    Dim r As Long
    
    dbEdit = True
    
    
    Do While i <= UBound(ipHead(head).ipSect(sect1).myEntries)
        'check for ip match
        If StrComp(ipHead(head).ipSect(sect1).myEntries(i).ip, arUsers(id).ip, vbBinaryCompare) = 0 Then
            found = True
            arUsers(id).dbPos = i
            ipHead(head).ipSect(sect1).myEntries(i).avgPing = (ipHead(head).ipSect(sect1).myEntries(i).avgPing + arUsers(id).ping) / 2
            ipHead(head).ipSect(sect1).myEntries(i).lastLogin = Date & ": " & Time
            If myUserId <> -1 Then ipHead(head).ipSect(sect1).myEntries(i).numOfLogins = ipHead(head).ipSect(sect1).myEntries(i).numOfLogins + 1
            'check for name match
            str = Split(ipHead(head).ipSect(sect1).myEntries(i).aliases, ",")
            
            For w = 0 To UBound(str)
                If StrComp(str(w), arUsers(id).name, vbTextCompare) = 0 Then
                    If UBound(str) = 0 Then
                        checkUsers = vbNullString
                        Exit Function
                    Else
                        Exit For
                    End If
                ElseIf w = UBound(str) Then
                    ipHead(head).ipSect(sect1).myEntries(i).aliases = ipHead(head).ipSect(sect1).myEntries(i).aliases & "," & arUsers(id).name
                    ipHead(head).ipSect(sect1).myEntries(i).numOfNames = ipHead(head).ipSect(sect1).myEntries(i).numOfNames + 1
                End If
            Next w
            checkUsers = Replace$(ipHead(head).ipSect(sect1).myEntries(i).aliases, ",", ", ", 1, -1, vbBinaryCompare)
            Exit Function
        End If
        i = i + 1
        
        r = r + 1
        If r = 100 Then
            r = 0
            DoEvents
        End If
    Loop
    
    'if user is not in the database, then add him
    If found = False Then
        arUsers(id).dbPos = UBound(ipHead(head).ipSect(sect1).myEntries)
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).aliases = arUsers(id).name
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).ip = arUsers(id).ip
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).avgPing = arUsers(id).ping
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).numOfLogins = 1
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).numOfNames = 1
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).firstLogin = Date & ": " & Time
        ipHead(head).ipSect(sect1).myEntries(arUsers(id).dbPos).lastLogin = Date & ": " & Time
        entryCount = entryCount + 1
        ReDim Preserve ipHead(head).ipSect(sect1).myEntries(0 To UBound(ipHead(head).ipSect(sect1).myEntries) + 1)
        frmDBV.Caption = "View All Database Entries: Total Entries: " & entryCount
        frmDB.lblTotal.Caption = "Total Users In Database: " & entryCount
        MDIForm1.StatusBar1.Panels(7).Text = "Total Users in DB: " & entryCount
        checkUsers = vbNullString
    End If
End Function

Sub loadUsers()
    Dim str As String
    Dim s() As String
    Dim sip() As String
    Dim buff() As String
    Dim i As Long
    Dim w As Long
    Dim t As Long
    Dim head As Long
    Dim sect1 As Long
    Dim r As Long
    
    On Error GoTo here
    
    'incase file doesn't exist, make it
    Close #4
    Open App.Path & "\EmulinkerSF_Logs\user.txt" For Append As #4
        'do nothing
    Close #4
    
    'Call saveUsers(True)
    
    entryCount = 0
    frmServerlist.ProgressBar1.Value = 0
    frmServerlist.List1.AddItem "Loading Database..."
    frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
    MDIForm1.Enabled = False
    t = GetTickCount

    'Open App.Path & "\EmulinkerSF_Logs\user.txt" For Input As #4
    '    Do Until EOF(4)
    '        Line Input #4, str
    '        i = i + 1
    '    Loop
    'Close #4
    
    For i = LBound(ipHead) To UBound(ipHead)
        For w = LBound(ipHead(0).ipSect) To UBound(ipHead(0).ipSect)
            ReDim ipHead(i).ipSect(w).myEntries(0 To 0)
        Next w
    Next i
        
    
    Dim f1 As Long
    Dim f2 As Long
    Dim f As String
       
    Open App.Path & "\EmulinkerSF_Logs\user.txt" For Input As #4
    f1 = LOF(4)
    Close #4
    Open App.Path & "\EmulinkerSF_Logs\user.bak" For Input As #4
    f2 = LOF(4)
    Close #4
    
    f = "user.txt"
    If f1 < f2 Then
        f = "user.bak"
    End If
    
    Open App.Path & "\EmulinkerSF_Logs\user.txt" For Input As #4
        
        If LOF(4) > 0 Then
            ReDim myEntries(0 To i)
            frmServerlist.ProgressBar1.Max = LOF(4)
            frmServerlist.ProgressBar1.Value = 0
            Do Until EOF(4)
                Line Input #4, str
                'remove users that haven't logged in for num days
                buff = Split(str, ChrW$(&H2))

                sip = Split(buff(0), ".")
                If Len(sip(0)) = 2 Then
                    buff(0) = "0" & buff(0)
                ElseIf Len(sip(0)) = 1 Then
                    buff(0) = "00" & buff(0)
                End If
                               
                               
                               
                'IP Address
                head = sip(0) 'Left$(buff(0), 3)
                sect1 = sip(1)
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).ip = buff(0)
                               
                'Aliases
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).aliases = Replace$(buff(1), ", ", ",", 1, -1, vbBinaryCompare)
                                                        
                'Last date of login
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).lastLogin = buff(2)
      
                'Number Aliases
                s = Split(buff(1), ",")
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).numOfNames = UBound(s) + 1
                    
                'Average Ping
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).avgPing = buff(3)
                    
                'Num of Logins
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).numOfLogins = buff(4)
    
                'Num of Games
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).numOfGames = buff(5)
                    
                'First Date of Login
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).firstLogin = buff(6)
                    
                'Num of Games Played
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).gamesPlayed = buff(7)
                
                'Num of Bot Hits
                ipHead(head).ipSect(sect1).myEntries(UBound(ipHead(head).ipSect(sect1).myEntries)).numOfBotHits = buff(8)
                
                entryCount = entryCount + 1
                ReDim Preserve ipHead(head).ipSect(sect1).myEntries(0 To UBound(ipHead(head).ipSect(sect1).myEntries) + 1)
                
                frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Value + Len(str)
                r = r + 1
                If r = 1000 Then
                    r = 0
                    frmServerlist.lblPer.Caption = CLng((frmServerlist.ProgressBar1.Value / frmServerlist.ProgressBar1.Max) * 100) & "% Complete"
                    DoEvents
                End If
            Loop
        End If
    Close #4

    frmDBV.Caption = "View All Database Entries: Total Entries: " & entryCount
    frmDB.lblTotal.Caption = "Total Users In Database: " & entryCount
    If entryCount = 0 Then
        MDIForm1.StatusBar1.Panels(7).Text = "Total Users in DB: " & "0"
    Else
        MDIForm1.StatusBar1.Panels(7).Text = "Total Users in DB: " & entryCount
    End If
    frmServerlist.lblPer.Caption = "100% Complete"
    frmServerlist.ProgressBar1.Value = 0
    frmServerlist.lblPer.Caption = "0% Complete"
    
    If entryCount > 1 Then
        frmServerlist.lblPer.Caption = "100% Complete in " & Abs(CSng((GetTickCount - t)) / 1000) & "s"
        frmServerlist.ProgressBar1.Value = 0
    End If
    
    frmServerlist.List1.AddItem "Finished"
    frmServerlist.List1.TopIndex = frmServerlist.List1.ListCount - 1
    
    MDIForm1.Enabled = True
    
    Exit Sub
here:
    MsgBox "There was an error Loading the Database!", vbExclamation
End Sub



Public Sub showDBEntry(pos As Long, name As String, head As Long, sect1 As Long)
    tempDBPos = pos
    tempHeadPos = head
    frmDB.fUserinfo.Caption = name & " - " & ipHead(head).ipSect(sect1).myEntries(pos).ip
    If ipHead(head).ipSect(sect1).myEntries(pos).remove = True Then
        frmDB.fUserinfo.Caption = name & " - " & ipHead(head).ipSect(sect1).myEntries(pos).ip & "; MARKED FOR DELETE"
    End If
    frmDB.txtFind.Text = ipHead(head).ipSect(sect1).myEntries(pos).ip
    frmDB.lblDBIPAddress.Caption = "IP Address: " & ipHead(head).ipSect(sect1).myEntries(pos).ip
    frmDB.lblDBDate.Caption = "Last Date of Login: " & ipHead(head).ipSect(sect1).myEntries(pos).lastLogin
    frmDB.lblDBFDate.Caption = "First Date of Login: " & ipHead(head).ipSect(sect1).myEntries(pos).firstLogin
    frmDB.lblDBTotalAliases.Caption = "Total Aliases: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfNames
    frmDB.lblDBTotalGames.Caption = "Total Games: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfGames
    frmDB.lblDBTotalLogins.Caption = "Total Logins: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfLogins
    frmDB.txtDBAliases.Text = Replace$(ipHead(head).ipSect(sect1).myEntries(pos).aliases, ",", ", ", 1, -1, vbBinaryCompare)
    frmDB.txtDBGames.Text = Replace$(ipHead(head).ipSect(sect1).myEntries(pos).gamesPlayed, ",", ", ", 1, -1, vbBinaryCompare)
    frmDB.lblDBAvgPing.Caption = "Average Ping: " & ipHead(head).ipSect(sect1).myEntries(pos).avgPing
    frmDB.lblDBBotHits.Caption = "Total Bot Hits: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfBotHits
    frmDB.lstDB.Clear
End Sub

Public Sub showDBEntryReset(pos As Long, head As Long, sect1 As Long)
    frmDB.fUserinfo.Caption = ipHead(head).ipSect(sect1).myEntries(pos).ip
    frmDB.lblDBIPAddress.Caption = "IP Address: " & ipHead(head).ipSect(sect1).myEntries(pos).ip
    
    ipHead(head).ipSect(sect1).myEntries(pos).lastLogin = Date & ": " & Time
    frmDB.lblDBDate.Caption = "Last Date of Login: " & ipHead(head).ipSect(sect1).myEntries(pos).lastLogin
    
    ipHead(head).ipSect(sect1).myEntries(pos).firstLogin = Date & ": " & Time
    frmDB.lblDBFDate.Caption = "First Date of Login: " & ipHead(head).ipSect(sect1).myEntries(pos).firstLogin
    
    ipHead(head).ipSect(sect1).myEntries(pos).numOfNames = 0
    frmDB.lblDBTotalAliases.Caption = "Total Aliases: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfNames
    
    ipHead(head).ipSect(sect1).myEntries(pos).numOfGames = 0
    frmDB.lblDBTotalGames.Caption = "Total Games: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfGames
    
    ipHead(head).ipSect(sect1).myEntries(pos).numOfLogins = 0
    frmDB.lblDBTotalLogins.Caption = "Total Logins: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfLogins
    
    ipHead(head).ipSect(sect1).myEntries(pos).aliases = vbNullString
    frmDB.txtDBAliases.Text = ipHead(head).ipSect(sect1).myEntries(pos).aliases
    
    ipHead(head).ipSect(sect1).myEntries(pos).gamesPlayed = vbNullString
    frmDB.txtDBGames.Text = ipHead(head).ipSect(sect1).myEntries(pos).gamesPlayed
    
    ipHead(head).ipSect(sect1).myEntries(pos).avgPing = 0
    frmDB.lblDBAvgPing.Caption = "Average Ping: " & ipHead(head).ipSect(sect1).myEntries(pos).avgPing

    ipHead(head).ipSect(sect1).myEntries(pos).numOfBotHits = 0
    frmDB.lblDBBotHits.Caption = "Total Bot Hits: " & ipHead(head).ipSect(sect1).myEntries(pos).numOfBotHits
End Sub



Public Sub loadEntry(e As Long, head As Long, updateIt As Boolean, sect1 As Long)
    Dim q As ListItem
    
    On Error Resume Next
    
    If updateIt = False Then
        Set q = frmDBV.lstDBV.ListItems.Add(, , ipHead(head).ipSect(sect1).myEntries(e).aliases)
    Else
        frmDBV.lstDBV.Refresh
        Set q = frmDBV.lstDBV.FindItem(ipHead(head).ipSect(sect1).myEntries(e).ip, lvwSubItem, 1, lvwWhole)
        If q Is Nothing Then
            Exit Sub
        Else
            q.Text = ipHead(head).ipSect(sect1).myEntries(e).aliases
        End If
    End If
    
    q.SubItems(1) = ipHead(head).ipSect(sect1).myEntries(e).ip
    q.SubItems(2) = ipHead(head).ipSect(sect1).myEntries(e).lastLogin
    q.SubItems(3) = e
    q.SubItems(4) = "dummy"
End Sub






