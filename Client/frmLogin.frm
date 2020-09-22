VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Login"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2415
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please wait while Authorization is gotten from the Remote Computer."
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
If txtUserName.Text = "" Then
    MsgBox "You Must Give The Remote Server A Name To Identify Yourself", vbOKOnly, "NO NAME GIVEN"
    txtUserName.SetFocus
    Exit Sub
End If
If txtUserName.Text = " " Then
    MsgBox "You Must Give The Remote Server A Name To Identify Yourself", vbOKOnly, "NO NAME GIVEN"
    txtUserName.Text = ""
    txtUserName.SetFocus
    Exit Sub
End If
Dim QWE As String
QWE = "|LOGIN|" & txtUserName.Text
frmMain.SockMain.SendData QWE
Me.BorderStyle = 0
Label1.Visible = False
txtUserName.Visible = False
cmdLogin.Visible = False
Label2.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub txtUserName_Change()
If txtUserName.Text = frmMain.txtUser.Text Then cmdLogin = True
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdLogin = True
End Sub

