VERSION 5.00
Begin VB.Form frmSysInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local System Information"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmSysInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Info"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Label lblWindowsDir 
         Caption         =   "Windows Dir"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lblRootDrive 
         Caption         =   "Root Drive"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label lblSystemDrive 
         Caption         =   "System Drive"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblProcessorID 
         Caption         =   "Processor Identifier"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblNumberOfProcessors 
         Caption         =   "Number Of Processors"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblComputerName 
         Caption         =   "Computer Name"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "Windows Directory"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "User Name:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "Root Drive:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "System Drive:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "Processor Identifier:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "Number Of Processors:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         Caption         =   "Computer Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblComputerName.Caption = (Environ("COMPUTERNAME"))
lblNumberOfProcessors.Caption = (Environ("NUMBER_OF_PROCESSORS"))
lblProcessorID.Caption = (Environ("PROCESSOR_IDENTIFIER"))
lblSystemDrive.Caption = (Environ("SystemDrive"))
lblRootDrive.Caption = (Environ("SystemRoot"))
lblUserName.Caption = (Environ("UserName"))
lblWindowsDir.Caption = (Environ("WinDir"))
End Sub

