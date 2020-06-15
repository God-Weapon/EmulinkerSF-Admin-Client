Attribute VB_Name = "MergeSort"
Option Explicit

Public Sub Merge(ByRef lngArray() As Entries)
    Dim arrTemp() As Entries
    Dim iSegSize As Long
    Dim iLBound As Long
    Dim iUBound As Long
            
    iLBound = LBound(lngArray)
    iUBound = UBound(lngArray)
        
    ReDim arrTemp(iLBound To iUBound)


    'frmServerlist.ProgressBar1.Max = iUBound - iLBound
    'frmServerlist.ProgressBar1.Value = 1
    
    iSegSize = 1
    Do While iSegSize < iUBound - iLBound
        
        'Merge from A to B
        InnerMergePass lngArray, arrTemp, iLBound, iUBound, iSegSize
        iSegSize = iSegSize + iSegSize
        
        'If frmServerlist.ProgressBar1.Value + iSegSize < frmServerlist.ProgressBar1.Max Then
        '    frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Value + iSegSize
        '    frmServerlist.lblPer.Caption = CLng((frmServerlist.ProgressBar1.Value / frmServerlist.ProgressBar1.Max) * 100) & "% Complete"
        '    DoEvents
        'End If
        
        'Merge from B to A
        InnerMergePass arrTemp, lngArray, iLBound, iUBound, iSegSize
        iSegSize = iSegSize + iSegSize
        
        'If frmServerlist.ProgressBar1.Value + iSegSize < frmServerlist.ProgressBar1.Max Then
        '    frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Value + iSegSize
        '    frmServerlist.lblPer.Caption = CLng((frmServerlist.ProgressBar1.Value / frmServerlist.ProgressBar1.Max) * 100) & "% Complete"
        '    DoEvents
        'End If
    Loop

    'frmServerlist.ProgressBar1.Value = frmServerlist.ProgressBar1.Max
    'frmServerlist.lblPer.Caption = "100% Complete"
End Sub

Private Sub InnerMergePass(ByRef lngSrc() As Entries, ByRef lngDest() As Entries, ByVal iLBound As Long, iUBound As Long, ByVal iSegSize As Long)
    Dim iSegNext As Long
    
    iSegNext = iLBound
    
    Do While iSegNext <= iUBound - (2 * iSegSize)
        'Merge 2 segments from src to dest
        InnerMerge lngSrc, lngDest, iSegNext, iSegNext + iSegSize - 1, iSegNext + iSegSize + iSegSize - 1
        
        iSegNext = iSegNext + iSegSize + iSegSize
    Loop
    
    'Fewer than 2 full segments remain
    If iSegNext + iSegSize <= iUBound Then
        '2 segs remain
        InnerMerge lngSrc, lngDest, iSegNext, iSegNext + iSegSize - 1, iUBound
    Else
        '1 seg remains, just copy it
        For iSegNext = iSegNext To iUBound
            lngDest(iSegNext) = lngSrc(iSegNext)
        Next iSegNext
    End If

End Sub

Private Sub InnerMerge(ByRef lngSrc() As Entries, ByRef lngDest() As Entries, ByVal iStartFirst As Long, ByVal iEndFirst As Long, ByVal iEndSecond As Long)
    Dim iFirst As Long
    Dim iSecond As Long
    Dim iResult As Long
    Dim iOuter As Long
    
    iFirst = iStartFirst
    iSecond = iEndFirst + 1
    iResult = iStartFirst
    
    Do While (iFirst <= iEndFirst) And (iSecond <= iEndSecond)
    
        'Select the smaller value and place in the output
        'Since the subarrays are already sorted, only one comparison is needed
        If lngSrc(iFirst).ip <= lngSrc(iSecond).ip Then
            lngDest(iResult) = lngSrc(iFirst)
            iFirst = iFirst + 1
        Else
            lngDest(iResult) = lngSrc(iSecond)
            iSecond = iSecond + 1
        End If
        
        iResult = iResult + 1
    Loop
    
    'Take care of any leftover values
    If iFirst > iEndFirst Then
        'Got some leftover seconds
        For iOuter = iSecond To iEndSecond
            lngDest(iResult) = lngSrc(iOuter)
            iResult = iResult + 1
        Next iOuter
    Else
        'Got some leftover firsts
        For iOuter = iFirst To iEndFirst
            lngDest(iResult) = lngSrc(iOuter)
            iResult = iResult + 1
        Next iOuter
    End If
End Sub
