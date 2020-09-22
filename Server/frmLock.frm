VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Admin - System Locked"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Remote Administration Ver 5.0.1"
      Height          =   2775
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   120
         Picture         =   "frmLock.frx":628A
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Label lblUnlock 
      Caption         =   "FALSE"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "The Remote User Has Locked Your System For Administration."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DisableCTRLaltDEL(huh As Boolean)
'Disable CTRL+ALT+DEL'
GD = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub
Public Function SetTopMostWindow(hwnd As Long, Top As Boolean) As Long
    SetTopMostWindow = SetWindowPos(hwnd, -1, 0, 0, 0, 0, FLAG)
End Function
Private Sub Form_Load()
StayOnTop frmLock
Me.Width = Screen.Width
Me.Height = Screen.Height
Dim X As Long
Dim y As Long
X = GetSystemMetrics(0)
y = GetSystemMetrics(1)
SetWindowPos Me.hwnd, 0, 0, 0, X, y, SHOWS
SetTopMostWindow Me.hwnd, True
DisableCTRLaltDEL (1)
Me.Width = Screen.Width
Me.Height = Screen.Height
lblInfo.Left = (Me.Width - lblInfo.Width) / 2
lblInfo.Top = (Me.Height - lblInfo.Height) / 2
Frame2.Left = (Me.Width - Frame2.Width) / 2
Frame2.Top = lblInfo.Top + lblInfo.Height + 200
End Sub
Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lblUnlock.Caption = "FALSE" Then
    Cancel = 1
Else
    Unload Me
End If
End Sub
Private Sub Form_Terminate()
If lblUnlock.Caption = "FALSE" Then
    DisableCTRLaltDEL (0)
Else
    DisableCTRLaltDEL (1)
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lblUnlock.Caption = "FALSE" Then
    DisableCTRLaltDEL (0)
Else
    DisableCTRLaltDEL (1)
    Unload Me
End If
End Sub
