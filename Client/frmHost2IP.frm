VERSION 5.00
Begin VB.Form frmHost2IP 
   Caption         =   "Resolve Host to IP"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   Icon            =   "frmHost2IP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHN 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdResolve 
      Caption         =   "Resolve"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtHostName 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblHostName 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Host Domain Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "127.0.0.1"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Host Resolved IP"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "frmHost2IP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128
Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type
Private Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADESCRIPTION_LEN) As Byte
   szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub ResolveHost()
Dim lHostName As Long
lHostName = ConvertToLong(lblIP.Caption)
txtHN.Text = GetHostNamez(lHostName)
lblHostName.Caption = txtHN.Text
End Sub
Private Function ConvertToLong(IP As String) As Long
ConvertToLong = inet_addr(IP)
End Function
Private Sub cmdResolve_Click()
Dim hostent_addr As Long
Dim host As HOSTENT
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String
hostent_addr = gethostbyname(txtHostName.Text)
If hostent_addr = 0 Then
    MsgBox "Can't resolve name."
    Exit Sub
End If
RtlMoveMemory host, hostent_addr, LenB(host)
RtlMoveMemory hostip_addr, host.hAddrList, 4
ReDim temp_ip_address(1 To host.hLength)
RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
For i = 1 To host.hLength
    ip_address = ip_address & temp_ip_address(i) & "."
Next
ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
lblIP.Caption = ip_address
ResolveHost
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SocketsCleanup
End Sub
Sub SocketsCleanup()
   Dim lReturn As Long
   lReturn = WSACleanup()
   If lReturn <> 0 Then
      MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
   End If
End Sub
