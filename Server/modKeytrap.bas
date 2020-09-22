Attribute VB_Name = "modktrap"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Function fSaveGuiToFile(theFile As String) As Boolean
Dim lString As String
On Error Resume Next 'GoTo Trap
'Check if the File Exist
If Dir(theFile) <> "" Then Exit Function
'Clipboard.Clear
'To get the Entire Screen
Call keybd_event(vbKeySnapshot, 1, 0, 0)
'To get the Active Window
SavePicture Clipboard.GetData(vbCFBitmap), theFile
fSaveGuiToFile = True
Exit Function
End Function
