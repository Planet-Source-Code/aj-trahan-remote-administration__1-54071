VERSION 5.00
Begin VB.Form frmPing 
   Caption         =   "Ping"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFC0C0&
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtPingIP 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPing.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ping Results"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "IP Number to Ping"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPing_Click()
List1.Clear
Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Integer
Call Ping(txtPingIP.Text, ECHO)
List1.AddItem GetStatusCode(ECHO.status)
List1.AddItem ECHO.Address
List1.AddItem ECHO.RoundTripTime & " ms"
List1.AddItem ECHO.DataSize & " bytes"
If Left$(ECHO.Data, 1) <> Chr$(0) Then
    pos = InStr(ECHO.Data, Chr$(0))
    List1.AddItem Left$(ECHO.Data, pos - 1)
End If
List1.AddItem ECHO.DataPointer
End Sub

Private Sub txtPingIP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdPing = True
End Sub
