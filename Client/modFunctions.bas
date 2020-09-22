Attribute VB_Name = "modFunctions"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Sub Pause(HowLong As Long)
Dim u%, tick As Long
tick = GetTickCount()
Do
    u% = DoEvents
Loop Until tick + HowLong < GetTickCount
End Sub
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
Public Sub SendFile(FileName As String, WinS As Winsock)
Dim FreeF As Integer
Dim LenFile As Long
Dim nCnt As Long
Dim LocData As String
Dim LoopTimes As Long
Dim i As Long
FreeF = FreeFile
Open FileName For Binary As #99
    nCnt = 1
    LenFile = LOF(99)
    WinS.SendData "|FILESIZE|" & LenFile
    DoEvents
    Sleep (400)
    Do Until nCnt >= (LenFile)
        LocData = Space$(1024)
        Get #99, nCnt, LocData
        If nCnt + 1024 > LenFile Then
            WinS.SendData Mid$(LocData, 1, (LenFile - nCnt))
        Else
            WinS.SendData LocData
        End If
        nCnt = nCnt + 1024
    Loop
Close #99
End Sub
