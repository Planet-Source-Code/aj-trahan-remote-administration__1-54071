Attribute VB_Name = "modAPIS"
'The GetWindowTextLength function retrieves the length, in characters, of the specified window's title bar text (if the window has a title bar). If the specified window is a control, the function retrieves the length of the text within the control. However, GetWindowTextLength cannot retrieve the length of the text of an edit control in another application.
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
'The GetWindowText function copies the text of the specified window's title bar (if it has one) into a buffer. If the specified window is a control, the text of the control is copied. However, GetWindowText cannot retrieve the text of a control in another application.
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Integer) As Integer
'The GetWindow function retrieves a handle to a window that has the specified relationship (Z order or owner) to the specified window.
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'The IsIconic function determines whether the specified window is minimized (iconic).
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'The IsZoomed function determines whether a window is maximized.
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
'The IsWindowVisible function retrieves the visibility state of the specified window.
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'The AppendMenu function appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. You can use this function to specify the content, appearance, and behavior of the menu item.
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'The GetSystemMenu function allows the application to access the window menu (also known as the system menu or the control menu) for copying and modifying.
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
'The ExitWindowsEx function either logs off the current user, shuts down the system, or shuts down and restarts the system. It sends the WM_QUERYENDSESSION message to all applications to determine if they can be terminated.
Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'The SystemParametersInfo function queries or sets systemwide parameters. This function can also update the user profile while setting a parameter
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'The RegisterServiceProcess function registers or unregisters a service process. A service process continues to run after the user logs off.
Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
'The GetCurrentProcessId function returns the process identifier of the calling process.
Public Declare Function GetCurrentProcessId Lib _
   "kernel32.dll" () As Long
'The ExitWindows function logs off the current user.
Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Option Explicit
Public Type ProcData
    AppHwnd As Long
    title As String
    Placement As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'The EnumWindows function enumerates all top-level windows on the screen by passing the handle to each window, in turn, to an application-defined callback function. EnumWindows continues until the last top-level window is enumerated or the callback function returns FALSE.
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As _
Any, ByVal lParam As Long) As Long
'The FindWindow function retrieves a handle to the top-level window whose class name and window name match the specified strings. This function does not search child windows. This function does not perform a case-sensitive search.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
lpClassName As String, ByVal lpWindowName As String) As Long
'The GetClassName function retrieves the name of the class to which the specified window belongs.
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal _
hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'The Most Important API
'The SendMessage function sends the specified message to a window or windows. The function calls the window procedure for the specified window and does not return until the window procedure has processed the message.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'The keybd_event function synthesizes a keystroke. The system can use such a synthesized keystroke to generate a WM_KEYUP or WM_KEYDOWN message. The keyboard driver's interrupt handler calls the keybd_event function.
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
'The FindFirstFile function searches a directory for a file whose name matches the specified filename. FindFirstFile examines subdirectory names as well as filenames.
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'The FindClose function closes the specified search handle. The FindFirstFile and FindNextFile functions use the search handle to locate files with names that match a given name.
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'The ReleaseDC function releases a device context (DC), freeing it for use by other applications. The effect of the ReleaseDC function depends on the type of device context. It frees only common and window device contexts. It has no effect on class or private device contexts.
Public Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal _
    HDC As Integer) As Integer
' to open or execute any file based upon its extension
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'The mciSendString function sends a command string to an MCI device. The device that the command is sent to is specified in the command string.
' we use it to open close cd door
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'The GetNextWindow function retrieves a handle to the next or previous window in the Z order. The next window is below the specified window; the previous window is above. If the specified window is a topmost window, the function retrieves a handle to the next (or previous) topmost window
Public Declare Function GetNextWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) _
As Long
'The GetTickCount function retrieves the number of milliseconds that have elapsed since the system was started. It is limited to the resolution of the system timer.
Declare Function GetTickCount& Lib "kernel32" ()
'The GetDesktopWindow function returns a handle to the desktop window. The desktop window covers the entire screen. The desktop window is the area on top of which all icons and other windows are painted.
Private Declare Function GetDesktopWindow Lib "User" () As Integer
'The GetDC function retrieves a handle to a display device context for the client area of a specified window or for the entire screen. You can use the returned handle in subsequent GDI functions to draw in the device context.
Private Declare Function GetDC Lib "User" (ByVal hwnd%) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'The GetDriveType function determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive.
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Global LastX As Integer
Global LastY As Integer
Global TransBuff As String
Public Type POINTAPI
        x As Long
        y As Long
End Type
'Functions for mouse inverse
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCursor Lib "user32" () As Long
'Functions for screen lock
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWnd2 As Long, ByVal x As Long, ByVal y As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal l As Long) As Long
'The GetSystemMetrics function retrieves various system metrics (widths and heights of display elements) and system configuration settings. All dimensions retrieved by GetSystemMetrics are in pixels.
Public Declare Function GetSystemMetrics Lib "user32" (ByVal n As Long) As Long
Const SHOWS = &H40
Const FLAG = 2 Or 1
