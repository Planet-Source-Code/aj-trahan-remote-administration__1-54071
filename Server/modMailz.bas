Attribute VB_Name = "modMailz"
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128
Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Global fCaption As String
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer 'The key states
Public Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer 'the key states
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function
Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function
Public Function GetInternetIP(pboolReturnExternalIP As Boolean) As String
Dim hostname As String * 256
Dim hostent_addr As Long
Dim host As HOSTENT
Dim hostip_addr As Long
Dim temp_ipaddress() As Byte
Dim iCounter As Integer
Dim strIPaddress As String
Dim strCheckIP As String
Dim strInternIP As String
Dim strExternIP As String
GetInternetIP = ""
If gethostname(hostname, 256) = SOCKET_ERROR Then
    MsgBox "Socket-Error with Winsock.dll."
Else
    hostname = Trim$(hostname)
    hostent_addr = gethostbyname(hostname)
    If hostent_addr = 0 Then
        MsgBox "There's an error with Winsock.dll."
    Else
        RtlMoveMemory host, hostent_addr, LenB(host)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
        Do
            ReDim temp_ipaddress(1 To host.hLength)
            RtlMoveMemory temp_ipaddress(1), hostip_addr, host.hLength
            For iCounter = 1 To host.hLength
                strIPaddress = strIPaddress & temp_ipaddress(iCounter) & "."
            Next
            strIPaddress = Mid$(strIPaddress, 1, Len(strIPaddress) - 1)
            strInternIP = strCheckIP
            strExternIP = strIPaddress
            strCheckIP = strIPaddress
            host.hAddrList = host.hAddrList + LenB(host.hAddrList)
            RtlMoveMemory hostip_addr, host.hAddrList, 4
            strIPaddress = ""
        Loop While (hostip_addr <> 0)
        If Trim(strInternIP) = "" Then ' same as External
            strInternIP = strExternIP
        End If
        If Trim(strExternIP) = "" Then 'just for sure
            strExternIP = strInternIP  ' no one knows, what
        End If                         ' micrososft does next :>
        GetInternetIP = strInternIP
        If pboolReturnExternalIP = True Then
            GetInternetIP = strExternIP
        End If
    End If
End If
End Function
Public Function GetString(hKey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
If lRegResult <> ERROR_SUCCESS Then
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function WindowsDirectory() As String
Dim WinPath As String
Dim temp
WinPath = String(145, Chr(0))
temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function

Public Function SystemDirectory() As String
Dim SysPath As String
Dim temp
SysPath = String(145, Chr(0))
temp = GetSystemDirectory(SysPath, 145)
SystemDirectory = Left(SysPath, InStr(SysPath, Chr(0)) - 1)
End Function
Private Function GetDefaultAccount() As String
Dim RegKey As String
RegKey = "Software\Microsoft\Internet Account Manager"
GetDefaultAccount = GetString(HKEY_CURRENT_USER, RegKey, "Default Mail Account")
End Function
Public Function GetDefaultSmtp() As String
Dim sAccount As String
Dim RegKey As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
GetDefaultSmtp = GetString(HKEY_CURRENT_USER, hKey, "SMTP Server")
End Function
Public Function GetmailAcc() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  DEFAULT MAIL ACCOUNT :  " & GetString(HKEY_CURRENT_USER, RegKey, "Default Mail Account") & "   |"
GetmailAcc = RInfo
End Function
Public Function smtpServer() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  SMTP SERVER          :  " & GetString(HKEY_CURRENT_USER, hKey, "SMTP Server") & "|"
smtpServer = RInfo
End Function
Public Function MailAddr() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  EMAIL ADDRESS        :  " & GetString(HKEY_CURRENT_USER, hKey, "SMTP Email Address") & "   |"
MailAddr = RInfo
End Function
Public Function PopUser() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  POP3 USER NAME       :  " & GetString(HKEY_CURRENT_USER, hKey, "POP3 User Name") & "   |"
PopUser = RInfo
End Function

Public Function SmtpDisplay() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  SMTP DISPLAY NAME    :  " & GetString(HKEY_CURRENT_USER, hKey, "SMTP Display Name") & "   |"
SmtpDisplay = RInfo
End Function
Public Function CountryDisplay() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String
sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount
RInfo = " |  Country    :  " & GetString(HKEY_CURRENT_USER, "Control Panel\International", "sCountry") & "   |"
CountryDisplay = RInfo
End Function

