Attribute VB_Name = "mdlCallback"
Option Explicit

Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
        
Public Const LVIF_TEXT As Long = &H1
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETITEM As Long = (LVM_FIRST + 5)
Public Const LVM_SETITEM As Long = (LVM_FIRST + 6)
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Public Const MEM_COMMIT As Long = &H1000
Public Const PAGE_READWRITE As Long = &H4
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SYNCHRONIZE As Long = &H100000
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const MEM_RELEASE As Long = &H8000

Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long 'Notice not string but long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public Declare Function OpenProcess Lib "kernel32.dll" ( _
     ByVal dwDesiredAccess As Long, _
     ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As Long
     
Public Declare Function WriteProcessMemory Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByRef lpBaseAddress As Any, _
     ByRef lpBuffer As Any, _
     ByVal nSize As Long, _
     ByRef lpNumberOfBytesWritten As Long) As Long
     
Public Declare Function ReadProcessMemory Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByRef lpBaseAddress As Any, _
     ByRef lpBuffer As Any, _
     ByVal nSize As Long, _
     ByRef lpNumberOfBytesWritten As Long) As Long
     
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByRef lpdwProcessId As Long) As Long
     
     
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
     ByVal hwndParent As Long, _
     ByVal hwndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszWindow As Long) As Long
     

Public Declare Function VirtualAllocEx Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByVal lpAddress As Long, _
     ByVal dwSize As Long, _
     ByVal flAllocationType As Long, _
     ByVal flProtect As Long) As Long

Public Declare Function VirtualFreeEx Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByVal lpAddress As Long, _
     ByVal dwSize As Long, _
     ByVal dwFreeType As Long) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" ( _
     ByVal hObject As Long) As Long
     
Public Declare Function TerminateProcess Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     ByVal uExitCode As Long) As Long
     
Public hProcess As Long





Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal code As Long, ByVal maptype As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal c As Byte) As Integer
Public Declare Sub keybd_event Lib "user32" (ByVal vk As Byte, ByVal scan As Byte, ByVal Flags As Long, ByVal extra As Long)




Public Declare Function FlashWindowEx Lib "user32" (FWInfo As FLASHWINFO) As Boolean

Public Type FLASHWINFO
  cbSize As Long     ' size of structure
  hwnd As Long       ' hWnd of window to use
  dwFlags As Long    ' Flags, see below
  uCount As Long     ' Number of times to flash window
  dwTimeout As Long  ' Flash rate of window in milliseconds. 0 is default.
End Type

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE As Integer = 0
Public Const SW_SHOW As Integer = 5

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_APPLICATION = &H80         ' look for application specific association
Public Const SND_ALIAS = &H10000     ' name is a WIN.INI [sounds] entry
Public Const SND_ALIAS_ID = &H110000    ' name is a WIN.INI [sounds] entry identifier
Public Const SND_ASYNC = &H1         ' play asynchronously
Public Const SND_FILENAME = &H20000     ' name is a file name
Public Const SND_LOOP = &H8         ' loop the sound until next sndPlaySound
Public Const SND_MEMORY = &H4         ' lpszSoundName points to a memory file
Public Const SND_NODEFAULT = &H2         ' silence not default, if sound not found
Public Const SND_NOSTOP = &H10        ' don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000      ' don't wait if the driver is busy
Public Const SND_PURGE = &H40               ' purge non-static events for task
Public Const SND_RESOURCE = &H40004     ' name is a resource name or atom
Public Const SND_SYNC = &H0         ' play synchronously (default)

Public Const HWND_TOPMOST = -1


Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long



Public Const FLASHW_STOP = 0
Public Const FLASHW_CAPTION = 1
Public Const FLASHW_TRAY = 2
Public Const FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY
Public Const FLASHW_TIMER = 4
Public Const FLASHW_TIMERNOFG = 12
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINE = &HC4

Public Const MONITOR_DEFAULTTONULL As Long = &H0    'If the monitor is not found, return 0
Public Const MONITOR_DEFAULTTOPRIMARY As Long = &H1 'If the monitor is not found, return the primary monitor
Public Const MONITOR_DEFAULTTONEAREST As Long = &H2 'If the monitor is not found, return the nearest monitor

Public Type POINTAPI
   x  As Long
   y  As Long
End Type
'
Public Declare Sub Sleep Lib "kernel32" _
 _
        (ByVal dwMilliseconds As Long)

Public Declare Function MonitorFromPoint Lib "user32" _
  (ByVal x As Long, ByVal y As Long, _
   ByVal dwFlags As Long) As Long
   
Public Declare Function ClientToScreen Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long
Public Const SW_SHOWNORMAL As Long = 1
      Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOP = 0
      Public Const HWND_BOTTOM = 1
      Public Const HWND_NOTOPMOST = -2

      Public Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

Public Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long _
) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Const GWL_HWNDPARENT = (-8)
Public Const GWLP_HINSTANCE = (-6)

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)

Public Const GWLP_WNDPROC = (-4)
Public Const GWLP_USERDATA = (-21)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302

Public Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String _
) As Long



Public Const WM_USER = &H400
Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const WM_NOTIFY = &H4E
Public Const EN_LINK = &H70B
Public Const ENM_LINK = &H4000000
Public Const WM_LBUTTONUP = &H202


Public Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Public Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type ENLINK
    NMHDR As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CharRange
End Type

Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Long
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(LF_FACESIZE - 1) As Byte
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lcid As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type



' GetWindowsLong Constants
Public Const GWL_WNDPROC = (-4)

' Windows Message Constants
Public Const WM_DESTROY = &H2

' Column Header Notification Meassage Constants
Public Const HDN_FIRST = -300&
Public Const HDN_BEGINTRACK = (HDN_FIRST - 6)

' Column Header Item Info Message Constants
Public Const HDI_WIDTH = &H1



' Notify Message Header for Listview
Public Type NMHEADER
hdr As NMHDR
iItem As Long
iButton As Long
lPtrHDItem As Long ' HDITEM FAR* pItem
End Type

' Header Item Type
Public Type HDITEM
mask As Long
cxy As Long
pszText As Long
hbm As Long
cchTextMax As Long
fmt As Long
lParam As Long
iImage As Long
iOrder As Long
End Type

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public mlPrevWndProc As Long



Public Function RTBGetLine(RB As RichTextBox, LLineNum As Long) As String
'--get text of line.
'--get text of line. RB is RTB. LLineNum is line number.
Dim LineLen As Long, lIndex As Long
Dim sBuf As String
On Error GoTo woops
RTBGetLine = ""

'-- get line length from character, in this case index. Then
'-- put line length info. into first two bytes of GETLINE buffer.
'--
lIndex = SendMessageLong(RB.hwnd, EM_LINEINDEX, LLineNum, 0&)
LineLen = SendMessageLong(RB.hwnd, EM_LINELENGTH, lIndex, 0&) + 1
sBuf = String$((LineLen + 2), 0)
Mid$(sBuf, 1, 1) = Chr$(LineLen And &HFF)
Mid$(sBuf, 2, 1) = Chr$(LineLen \ &H100)

'--finally, call GETLINE:

LineLen = SendMessageString(RB.hwnd, EM_GETLINE, LLineNum, sBuf)
If (LineLen > 0) Then RTBGetLine = Left$(sBuf, LineLen)

Exit Function
woops:
End Function



Public Function MainCLSProc(ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
Dim clsRefToCLS As clsSubClass
Dim pUserData As Long
    
    pUserData = GetProp(hwnd, "objptr")
    
    If pUserData Then
        Set clsRefToCLS = ObjFromPtr(pUserData)
        MainCLSProc = clsRefToCLS.CLSProc(hwnd, msg, wParam, lParam)
        Set clsRefToCLS = Nothing
    End If
    
End Function


Private Function ObjFromPtr(ByVal lpObject As Long) As Object
Dim objTemp As Object
    CopyMemory objTemp, lpObject, 4&
    Set ObjFromPtr = objTemp
    CopyMemory objTemp, 0&, 4&
End Function

      Sub ActivatePrevInstance()
         Dim OldTitle As String
         Dim PrevHndl As Long
         Dim result As Long

         'Save the title of the application.
         OldTitle = App.Title

         'Rename the title of this application so FindWindow
         'will not find this application instance.
         App.Title = "unwanted instance"

         'Attempt to get window handle using VB4 class name.
         PrevHndl = FindWindow("ThunderRTMain", OldTitle)

         'Check for no success.
         If PrevHndl = 0 Then
            'Attempt to get window handle using VB5 class name.
            PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
         End If

         'Check if found
         If PrevHndl = 0 Then
         'Attempt to get window handle using VB6 class name
         PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
         End If

         'Check if found
         If PrevHndl = 0 Then
            'No previous instance found.
            Exit Sub
         End If

         'Get handle to previous window.
         PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

         'Restore the program.
         result = OpenIcon(PrevHndl)

         'Activate the application.
         result = SetForegroundWindow(PrevHndl)

         'End the application.
         Call MsgBox("You can't load 2 instances of: " & emulatorPass, vbExclamation, "Can't load 2 instances!")
         Unload MDIForm1
      End Sub


Private Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNMH As NMHDR
Dim tNMHEADER As NMHEADER
Dim tITEM As HDITEM

Select Case msg
Case WM_NOTIFY
' Copy the Notify Message Header to a Header Structure
CopyMemory tNMH, ByVal lParam, Len(tNMH)

Select Case tNMH.code
Case HDN_BEGINTRACK
' If the user is trying to Size a Column Header...

' Extract Information about the Header being Sized
CopyMemory tNMHEADER, ByVal lParam, Len(tNMHEADER)

' Get Item Info. about the header (i.e. Width)
CopyMemory tITEM, ByVal tNMHEADER.lPtrHDItem, Len(tITEM)

' Don't allow Zero Width Columns to be Sized.
If (tITEM.mask And HDI_WIDTH) = HDI_WIDTH And tITEM.cxy = 0 Then
WindowProc = 1
Exit Function
End If
End Select

Case WM_DESTROY
' Remove Subclassing when Listview is Destroyed (Form unloaded.)
WindowProc = CallWindowProc(mlPrevWndProc, hwnd, msg, wParam, lParam)
Call SetWindowLong(hwnd, GWL_WNDPROC, mlPrevWndProc)
Exit Function

End Select

' Call Default Window Handler
WindowProc = CallWindowProc(mlPrevWndProc, hwnd, msg, wParam, lParam)
End Function

Public Sub SubClassHwnd(ByVal hwnd As Long)
mlPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub


