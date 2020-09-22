Attribute VB_Name = "modEnumDrives"
Option Explicit
Global Const SW_SHOWNORMAL = 1
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const vbAllFileSpec = "*.*"
Public Const MAX_PATH = 260
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Dim hFind As Long
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cShortFileName As String * 14
End Type
Public wfd As WIN32_FIND_DATA
Public Const INVALID_HANDLE_VALUE = -1
Public lNodeCount As Long
Public Const vbBackslash = "\"
Public Const vbAscDot = 46
Public Function Enum_Drives() As String
Dim strDrive    As String
Dim strMessage  As String
Dim intCnt      As Integer
Dim rtn As String
strMessage = "|DRVS|"
For intCnt = 65 To 86
    strDrive = Chr(intCnt)
    Select Case GetDriveType(strDrive + ":\")
           Case DRIVE_REMOVABLE
                rtn = "Floppy Drive"
           Case DRIVE_FIXED
                rtn = "Hard Drive"
           Case DRIVE_REMOTE
                rtn = "Network Drive"
           Case DRIVE_CDROM
                rtn = "CD-ROM Drive"
           Case DRIVE_RAMDISK
                rtn = "RAM Disk"
           Case Else
                rtn = ""
    End Select
    If rtn <> "" Then
        strMessage = strMessage & strDrive & "," & GetDriveType(strDrive + ":\") & "|"
    End If
Next intCnt
Enum_Drives = Mid$(strMessage, 1, Len(strMessage) - 1)
End Function
Function ParseString(ByVal sString As String, ByVal Delimiter As String) As Collection
On Error GoTo ParseString_Error
Dim CurPos       As Long
Dim NextPos      As Long
Dim DelLen       As Integer
Dim nCount       As Integer
Dim TStr         As String
Set ParseString = New Collection
' Add delimiters to start and end of string to make loop simpler:
sString = Delimiter & sString & Delimiter
' Calculate the delimiter length only once:
DelLen = Len(Delimiter)
' Initialize the count and position:
nCount = 0
CurPos = 1
NextPos = InStr(CurPos + DelLen, sString, Delimiter)
' Loop searching for delimiters:
Do Until NextPos = 0
    ' Extract a sub-string:
    ParseString.Add Mid$(sString, CurPos + DelLen, NextPos - CurPos - DelLen)
    ' Increment the sub string counter:
    nCount = nCount + 1
    ' Position to the last found delimiter:
    CurPos = NextPos
    ' Find the next delimiter:
    NextPos = InStr(CurPos + DelLen, sString, Delimiter)
Loop
ParseString_Exit:
    Exit Function
ParseString_Error:
    Err.Raise Err.Number, "ParseString"
    Exit Function
End Function

Public Function Get_File_Name(sString As String) As String
Dim lLoop As Long
For lLoop = Len(sString) To 1 Step -1
    If Mid$(sString, lLoop, 1) = "\" Then
        Get_File_Name = Mid$(sString, lLoop + 1, Len(sString))
    End If
Next lLoop
End Function
Public Function Enum_Files(sParentPath As String) As String
Dim wfd As WIN32_FIND_DATA
Dim hFind As Long
Dim strString As String
Dim sFileName As String
strString = "|FILES|"
sParentPath = NormalizePath(sParentPath)
hFind = FindFirstFile(sParentPath & "\" & vbAllFileSpec, wfd)
If (hFind <> INVALID_HANDLE_VALUE) Then
    Do
        sFileName = Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
            If sFileName <> "." And sFileName <> ".." Then
                If wfd.dwFileAttributes <> vbDirectory Then
                    strString = strString & sParentPath & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1) & "|" & FileLen(sParentPath & wfd.cFileName) & "^"
                End If
            End If
    Loop While FindNextFile(hFind, wfd)
    Call FindClose(hFind)
End If

If strString <> "|FILES|" Then
    Enum_Files = Mid$(strString, 1, Len(strString) - 1)
Else
    Enum_Files = strString
    'DoSizeShit strString
End If
End Function
' normalizing path through "\"
Public Function NormalizePath(sPath As String) As String
  If Right$(sPath, 1) <> "\" Then
    NormalizePath = sPath & "\"
  Else
    NormalizePath = sPath
  End If
End Function
Public Function Enum_Folders(sParentPath As String) As String
Dim strMessage  As String
Dim wfd As WIN32_FIND_DATA
Dim hFind As Long
strMessage = "|FOLDERS|"
sParentPath = NormalizePath(sParentPath)
hFind = FindFirstFile(sParentPath & vbAllFileSpec, wfd)
If (hFind <> INVALID_HANDLE_VALUE) Then
    Do
        If (wfd.dwFileAttributes And vbDirectory) Then
            ' If not a  "." or ".." DOS subdir...
            If (Asc(wfd.cFileName) <> vbAscDot) Then
                strMessage = strMessage & sParentPath & "^" & Mid$(wfd.cFileName, 1, InStr(wfd.cFileName, vbNullChar) - 1) & "|"
            End If
        End If
    Loop While FindNextFile(hFind, wfd)
    Call FindClose(hFind)
End If
Screen.MousePointer = vbDefault
If strMessage <> "|FOLDERS|" Then
    Enum_Folders = Mid$(strMessage, 1, Len(strMessage) - 1)
Else
    Enum_Folders = Mid$(strMessage, 1, Len(strMessage))
End If

End Function
