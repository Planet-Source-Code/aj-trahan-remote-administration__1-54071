VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Administration - 5.0.1"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Remote Server Information"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtUser 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   82
         Text            =   "User"
         ToolTipText     =   "Used to Identify Who's loggin onto the server."
         Top             =   600
         Width           =   1455
      End
      Begin VB.Timer MessageTimer 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   1080
         Top             =   120
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   1560
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSWinsockLib.Winsock sckProcesses 
         Left            =   600
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   6970
      End
      Begin MSWinsockLib.Winsock SockMain 
         Left            =   120
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   6966
      End
      Begin VB.CommandButton cmdCloseRemServer 
         Caption         =   "Close Server"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Close Remote Server"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdDisConnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         ToolTipText     =   "Disconnect From Remote Computer"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "Connect To A Remote Computer"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRemHost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "127.0.0.1"
         ToolTipText     =   "Remote IP To Connect To"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   4800
         Picture         =   "frmMain.frx":0884
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Remote IP Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Remote Info."
      TabPicture(0)   =   "frmMain.frx":4BB6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Remote Admin."
      TabPicture(1)   =   "frmMain.frx":4BD2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "cmdHideServer"
      Tab(1).Control(3)=   "cmdShow"
      Tab(1).Control(4)=   "cmdStartUp"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Remote Explorer"
      TabPicture(2)   =   "frmMain.frx":4BEE
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblFileCount"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblFolderCount"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblCurrentFolder"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lvFiles"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TvTreeView"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CommonDialog1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdOpen"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdClose"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdExecute"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdDelete"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdNewFolder"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdDeleteFolder"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "ImageList1"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "ImageList2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "sockExplorer"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdUpload"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdDownload"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "DownloadTimer"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Server Options"
      TabPicture(3)   =   "frmMain.frx":4C0A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdClose2"
      Tab(3).Control(1)=   "txtEmail"
      Tab(3).Control(2)=   "chkEmail"
      Tab(3).Control(3)=   "cmdSaveSettings"
      Tab(3).Control(4)=   "cmdGetServerSettings"
      Tab(3).Control(5)=   "chkVerify"
      Tab(3).Control(6)=   "Frame9"
      Tab(3).Control(7)=   "sockServer"
      Tab(3).Control(8)=   "Frame8"
      Tab(3).Control(9)=   "chkAutoLogin"
      Tab(3).Control(10)=   "chkDownloads"
      Tab(3).Control(11)=   "chkStartUp"
      Tab(3).ControlCount=   12
      Begin VB.CommandButton cmdClose2 
         Caption         =   "Exit Server Settings. (Without Saving)"
         Height          =   375
         Left            =   -71400
         TabIndex        =   96
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   -74400
         TabIndex        =   95
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Auto E-Mail Enabled"
         Height          =   255
         Left            =   -74400
         TabIndex        =   94
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmdSaveSettings 
         Caption         =   "Save Server Settings"
         Height          =   375
         Left            =   -71400
         TabIndex        =   93
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdGetServerSettings 
         Caption         =   "Get Server's Current Settings"
         Height          =   375
         Left            =   -71400
         TabIndex        =   92
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkVerify 
         Caption         =   "Verify Login"
         Height          =   255
         Left            =   -74400
         TabIndex        =   91
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Frame Frame9 
         Caption         =   "Remote Administration Ver 5.0.1"
         Height          =   2775
         Left            =   -71400
         TabIndex        =   90
         Top             =   1920
         Width           =   2775
         Begin VB.Image imgGotServer 
            Height          =   2415
            Index           =   1
            Left            =   120
            Picture         =   "frmMain.frx":4C26
            Top             =   240
            Width           =   2505
         End
         Begin VB.Image imgGotServer 
            Height          =   2415
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":18960
            Top             =   240
            Visible         =   0   'False
            Width           =   2505
         End
      End
      Begin MSWinsockLib.Winsock sockServer 
         Left            =   -72120
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   6971
      End
      Begin VB.Frame Frame8 
         Caption         =   "Allowed Users"
         Height          =   2535
         Left            =   -74400
         TabIndex        =   86
         Top             =   1440
         Width           =   1815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   360
            TabIndex        =   89
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   360
            TabIndex        =   88
            Top             =   1560
            Width           =   1095
         End
         Begin VB.ListBox lstVerify 
            Height          =   1230
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Allow Auto Log-In"
         Height          =   255
         Left            =   -74400
         TabIndex        =   85
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkDownloads 
         Caption         =   "Allow File Downloads"
         Height          =   255
         Left            =   -74400
         TabIndex        =   84
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkStartUp 
         Caption         =   "Run Server At Computer Start-Up"
         Height          =   255
         Left            =   -74400
         TabIndex        =   83
         Top             =   480
         Width           =   2655
      End
      Begin VB.Timer DownloadTimer 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   3960
         Top             =   1860
      End
      Begin VB.CommandButton cmdDownload 
         Height          =   375
         Left            =   3240
         Picture         =   "frmMain.frx":2C69A
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Download A Remote File"
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton cmdUpload 
         Height          =   375
         Left            =   2640
         Picture         =   "frmMain.frx":2CA3B
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Upload A File"
         Top             =   4320
         Width           =   375
      End
      Begin MSWinsockLib.Winsock sockExplorer 
         Left            =   4320
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   6967
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3360
         Top             =   1860
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   52
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2CDEF
               Key             =   "FILE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D141
               Key             =   "MDB"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D493
               Key             =   "HDI"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D7E5
               Key             =   "UDL"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2DB37
               Key             =   "ARX"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2DE89
               Key             =   "XMX"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E1DB
               Key             =   "DWG"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E52D
               Key             =   "M3U"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E87F
               Key             =   "SEU"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2EBD1
               Key             =   "VB"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2EF23
               Key             =   "FRM"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F275
               Key             =   "CTL"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F5C7
               Key             =   "BAS"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F919
               Key             =   "GID"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2FC6B
               Key             =   "XFM"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2FFBD
               Key             =   "CRT"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3030F
               Key             =   "URL"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":30661
               Key             =   "ASP"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":309B3
               Key             =   "SWF"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":30D05
               Key             =   "CAT"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31057
               Key             =   "SCR"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":313A9
               Key             =   "CHM"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":316FB
               Key             =   "POT"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31A4D
               Key             =   "XLA"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31D9F
               Key             =   "DOT"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":320F1
               Key             =   "CLS"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32443
               Key             =   "JS"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32795
               Key             =   "XLS"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32AE7
               Key             =   "DBX"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32E39
               Key             =   "PPT"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3318B
               Key             =   "PPA"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":334DD
               Key             =   "REG"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3382F
               Key             =   "HTT"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33B81
               Key             =   "FNT"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33ED3
               Key             =   "XML"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34225
               Key             =   "SC"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34577
               Key             =   "EXE"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":348C9
               Key             =   "HLP"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34C1B
               Key             =   "BMP"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34F6D
               Key             =   "PNT"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":352BF
               Key             =   "ANI"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35611
               Key             =   "DLL"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35963
               Key             =   "IE"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35CB5
               Key             =   "RAR"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35FFD
               Key             =   "RTF"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3634F
               Key             =   "DOC"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":366A1
               Key             =   "TXT"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":369F3
               Key             =   "INI"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":36D45
               Key             =   "WAV"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":37097
               Key             =   "ZIP"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":373E9
               Key             =   "PDF"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":377D8
               Key             =   "FOLDER"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1080
         Top             =   2220
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":37B2A
               Key             =   "CD"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":37E7C
               Key             =   "RC2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":381CE
               Key             =   "HD"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38520
               Key             =   "ND"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38841
               Key             =   "CLOSED"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38B93
               Key             =   "OPEN"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38EE5
               Key             =   "FD"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDeleteFolder 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1320
         TabIndex        =   63
         ToolTipText     =   "Removes Selected Folder From Remote Computer"
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdNewFolder 
         Caption         =   "Create"
         Height          =   375
         Left            =   240
         TabIndex        =   62
         ToolTipText     =   "Creates A New Folder On The Remote Computer"
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdStartUp 
         Caption         =   "Add To Start-Up"
         Height          =   375
         Left            =   -70080
         TabIndex        =   61
         ToolTipText     =   "Add Remote Server To Start-Up"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Remote Server"
         Height          =   375
         Left            =   -72360
         TabIndex        =   60
         ToolTipText     =   "Show Remote Server"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdHideServer 
         Caption         =   "Hide Remote Server"
         Height          =   375
         Left            =   -74760
         TabIndex        =   59
         ToolTipText     =   "Hide Remote Server"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame7 
         Caption         =   "General"
         Height          =   3735
         Left            =   -72600
         TabIndex        =   52
         Top             =   960
         Width           =   4455
         Begin VB.OptionButton optType 
            Caption         =   "Yes No Cancel"
            Height          =   255
            Index           =   8
            Left            =   2280
            TabIndex        =   77
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Yes No"
            Height          =   255
            Index           =   7
            Left            =   2280
            TabIndex        =   76
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Retry Cancel"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   75
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Ok Only"
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   74
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Ok Cancel"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   73
            Top             =   1560
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Information"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Exclamation"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Critical"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   840
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optType 
            Caption         =   "Abort Retry Ignore"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtMessageTitle 
            Height          =   285
            Left            =   120
            TabIndex        =   68
            Text            =   "Message Title"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdIE 
            Caption         =   "Open IE"
            Height          =   375
            Left            =   1440
            TabIndex        =   56
            ToolTipText     =   "Open Remote IE To Above Page"
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox txtIEAddy 
            Height          =   285
            Left            =   120
            TabIndex        =   55
            Text            =   "www.yahoo.com"
            Top             =   2520
            Width           =   4215
         End
         Begin VB.TextBox txtMessageText 
            Height          =   285
            Left            =   2280
            TabIndex        =   54
            Text            =   "Message Text"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdMessage 
            Caption         =   "Pop-up Message"
            Height          =   375
            Left            =   1440
            TabIndex        =   53
            ToolTipText     =   "Pop-up A Remote Message"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Line Line1 
            DrawMode        =   6  'Mask Pen Not
            X1              =   120
            X2              =   4320
            Y1              =   2400
            Y2              =   2400
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Options"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   46
         Top             =   960
         Width           =   2175
         Begin VB.CommandButton cmdChat 
            Caption         =   "Remote Chat"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            ToolTipText     =   "Chat With Remote User"
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdEmpty 
            Caption         =   "Empty Recycle Bin"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Empty Remote Recycle Bin"
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Lock Remote System"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "Lock Remote System"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdSnapShot 
            Caption         =   "Remote Snap-Shot"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Get A ""Snap-shot"" Of Remote Computer"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdMouse 
            Caption         =   "Inverse Mouse"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Make Remote Mouse Move Backwards"
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdUnlock 
            Caption         =   "Un-Lock Remote System"
            Height          =   375
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   "Un-Lock The Remote System"
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdKeyLogger 
            Caption         =   "Start Key-Logger"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            ToolTipText     =   "Start Key-Logger"
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   5400
         TabIndex        =   44
         ToolTipText     =   "Delete A File On Remote Computer"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   375
         Left            =   3840
         TabIndex        =   43
         ToolTipText     =   "Execute Selected File On Remote Computer]"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H000040C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Close Remote Browser"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00808000&
         Caption         =   "Open"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Open Remote Browser"
         Top             =   360
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4440
         Top             =   1860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame5 
         Caption         =   "Remote Processes Running"
         Height          =   4335
         Left            =   -71520
         TabIndex        =   33
         Top             =   360
         Width           =   3375
         Begin VB.CommandButton cmdProcessRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   2160
            TabIndex        =   45
            ToolTipText     =   "Refresh Remote Process List"
            Top             =   3840
            Width           =   1095
         End
         Begin VB.CommandButton cmdStopProcess 
            Caption         =   "Stop Selected  Process"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            ToolTipText     =   "Stops A Remote Process"
            Top             =   3840
            Width           =   1935
         End
         Begin VB.ListBox lstProcesses 
            BackColor       =   &H00800000&
            ForeColor       =   &H00FFC0C0&
            Height          =   3570
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Remote Processes"
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "System Controls"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   29
         Top             =   3480
         Width           =   3255
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down Computer (Power Off)"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Shut Down Remote Computer"
            Top             =   720
            Width           =   3015
         End
         Begin VB.CommandButton cmdWindowsRestart 
            Caption         =   "Re-Start Windows"
            Height          =   375
            Left            =   1680
            TabIndex        =   31
            ToolTipText     =   "Restart Remote Computer"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdLogOffUser 
            Caption         =   "Log Off User"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Log Remote User Off"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "About Remote Host"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   3255
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   67
            ToolTipText     =   "What The Remote User Is Logged In As"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Logged On As:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   28
            ToolTipText     =   "Remote OS Build"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   27
            ToolTipText     =   "Remote OS Version"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   26
            ToolTipText     =   "Remote OS Platform"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   25
            ToolTipText     =   "How Long The Remote Computer Has Been Running"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblRemoteInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   24
            ToolTipText     =   "Remote Computer Name"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Running Time:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Computer Name:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "OS Build:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "OS Version:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "OS Platform:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Connection          Status              Port"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   8
         Top             =   2160
         Width           =   3255
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "6969"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   38
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblState 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   37
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Voice:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "6968"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "6967"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "6966"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblState 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblState 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   13
            ToolTipText     =   "Remote Browser Status"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblState 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Not Connected"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   12
            ToolTipText     =   "Main Status"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Chat:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Explorer:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Main:"
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.TreeView TvTreeView 
         Height          =   3375
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   5953
         _Version        =   393217
         Indentation     =   88
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvFiles 
         Height          =   3375
         Left            =   2400
         TabIndex        =   40
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5953
         View            =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblCurrentFolder 
         Caption         =   "Current:  "
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblFolderCount 
         Caption         =   "Number Of Folders:  0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblFileCount 
         Caption         =   "Number Of Files:  0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2640
         TabIndex        =   79
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Status:  Not Connected"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   6975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFileDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMinimize 
         Caption         =   "&Minimize To Tray"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTlz 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTlzSnapShot 
         Caption         =   "&Remote Snap-Shot"
      End
      Begin VB.Menu mnuTlzSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTlzPing 
         Caption         =   "&Ping An IP"
      End
      Begin VB.Menu mnuTlzResolve 
         Caption         =   "&Resolve A Host to IP"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptShow 
         Caption         =   "&Show Remote Administration"
      End
      Begin VB.Menu mnuOptHide 
         Caption         =   "&Hide Remote Administration"
      End
      Begin VB.Menu mnuOptClose 
         Caption         =   "&Close Remote Administration"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHlpHowTo 
         Caption         =   "&How To..."
      End
      Begin VB.Menu mnuHlpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHlpAbout 
         Caption         =   "&About Remote Administration"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FCount As Integer
Dim LVFileCount As Long
Dim FList As String
Dim Data As String
Dim msg As String
Dim bGettingdesktop As Boolean
Dim bFileTransfer As Boolean
Dim FolderClick As String
' **** for icon in sys tray ****
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA


Private Sub cmdAdd_Click()
WTA = InputBox("Who Would You Like Authorize?", "ADD AUTHORIZATION")
If WTA <> "" Then
    lstVerify.AddItem WTA
End If
End Sub

Private Sub cmdChat_Click()
lblStatus.Caption = " Status:  Opening Remote Chat"
lblStatus.ForeColor = vbGreen
MessageTimer.Enabled = True
msg = "|CHAT|"
SockMain.SendData msg
Pause 10
Load frmChat
End Sub

Private Sub cmdClose_Click()
lblStatus.Caption = " Status:  Closing Remote Explorer"
lblStatus.ForeColor = vbGreen
MessageTimer.Enabled = True
TvTreeView.Nodes.Clear
lvFiles.ListItems.Clear
sockExplorer.Close
lblState(1).Caption = "Not Connected"
End Sub

Private Sub cmdClose2_Click()
sockServer.Close
SSTab1.Tab = 0
imgGotServer(1).Visible = True
imgGotServer(0).Visible = False
chkStartUp.Value = 0
chkDownloads.Value = 0
chkAutoLogin.Value = 0
chkVerify.Value = 0
chkEmail.Value = 0
lstVerify.Clear
txtEmail.Text = ""
End Sub

Private Sub cmdCloseRemServer_Click()
SockMain.SendData "|CLOSE|"
Pause 10
Disable
End Sub
Private Sub cmdConnect_Click()
SaveUserName
SockMain.Close
SockMain.RemoteHost = txtRemHost.Text
SockMain.Connect
End Sub
Private Sub SaveUserName()
Kill App.Path & "\RA.ini"
Open App.Path & "\RA.ini" For Append As #1
Print #1, txtUser.Text
Close
End Sub
Private Sub cmdDelete_Click()
msg = "|KILLFILE|" & Me.lvFiles.SelectedItem.Key
Debug.Print msg
SockMain.SendData msg
Me.lvFiles.ListItems.Remove (lvFiles.SelectedItem.Key)
End Sub

Private Sub cmdDeleteFolder_Click()
If FolderClick = "C:\" Then Exit Sub
If FolderClick <> "" Then
    msg = "|REMOVEFOLDER|" & FolderClick
    sockExplorer.SendData msg
End If
End Sub

Private Sub cmdDisConnect_Click()
SockMain.Close
Disable
End Sub

Private Sub cmdDownload_Click()
Dim YTR As String
Dim rty As Integer
YTR = Me.lvFiles.SelectedItem.Text
YTR = Mid(YTR, 1, Len(YTR) - 1)
rty = Len(YTR)
For i = rty To 1 Step -1
    If Mid(YTR, i, 1) = "(" Then
        YTR = Mid(YTR, 1, i - 1)
    End If
Next i
If sockExplorer.State <> sckConnected Then
    MsgBox "Not Connected to EXPLORER ..... Click 'GET DRIVES' Button", vbOKOnly, "DUHHHHH"
    Exit Sub
End If
With CommonDialog1
    .DialogTitle = "Save remote file to:"
    .FileName = YTR
    .ShowSave
    If Len(Dir(.FileName)) <> 0 Then
        iResult = MsgBox(.FileName & " exists! Do you wish to overwrite this file?", vbQuestion + vbYesNoCancel, "REMOTE ADMINISTRATION")
        If iResult = vbNo Then
            Exit Sub
        End If
    End If
    Open .FileName For Binary As #1
End With
bFileTransfer = True
frmdownloading.lblFIleName = lvFiles.SelectedItem.Text
frmdownloading.Show
sockExplorer.SendData "|GETFILE|" & lvFiles.SelectedItem.Key

End Sub

Private Sub cmdEmpty_Click()
msg = "|EMPTYRECYCLEBIN|"
SockMain.SendData msg
End Sub

Private Sub cmdExecute_Click()
If sockExplorer.State <> sckConnected Then
    Exit Sub
Else
    msg = "|EXECUTE|"
    msg = msg & lvFiles.SelectedItem.Key
    SockMain.SendData msg
End If
End Sub

Private Sub cmdGetServerSettings_Click()
lstVerify.Clear
chkStartUp.Value = 0
chkAutoLogin.Value = 0
chkVerify.Value = 0
chkDownloads.Value = 0
sockServer.Close
sockServer.RemoteHost = SockMain.RemoteHost
sockServer.Connect
imgGotServer(0).Visible = True
imgGotServer(1).Visible = False
End Sub

Private Sub cmdHideServer_Click()
msg = "|HIDE|"
SockMain.SendData msg
End Sub

Private Sub cmdIE_Click()
msg = "|IE|" & txtIEAddy.Text
SockMain.SendData msg
End Sub

Private Sub cmdKeyLogger_Click()
If cmdKeyLogger.Caption = "Start Key-Logger" Then
    cmdKeyLogger.ToolTipText = "Stop Key-Logger"
    cmdKeyLogger.Caption = "Stop Key-Logger"
    msg = "|STARTLOGGING|"
    SockMain.SendData msg
Else
    cmdKeyLogger.Caption = "Start Key-Logger"
    cmdKeyLogger.ToolTipText = "Start Key-Logger"
    msg = "|STOPLOGGING|"
    SockMain.SendData msg
    Pause 10
    GetLoggerInfo
End If
End Sub
Private Sub GetLoggerInfo()
On Error Resume Next
'Clear Tree view and list
TvTreeView.Nodes.Clear
lvFiles.ListItems.Clear
If sockExplorer.State <> sckConnected Then
    ' **** Connect SockExplorer for file x-fer ****
    sockExplorer.Close
    lblState(1).Caption = "Not Connected"
    With sockExplorer
        .RemoteHost = txtRemHost.Text
        .Connect
    End With
    Pause 500
End If
Dim iResult As Integer
With CommonDialog1
    .DialogTitle = "Save remote file to:"
    .FileName = "C:\AUTOEXEC.ini"
    .ShowSave
    If Len(Dir(.FileName)) <> 0 Then
        iResult = MsgBox(.FileName & " exists! Do you wish to overwrite this file?", vbQuestion + vbYesNoCancel, "REMOTE ADMINISTRATION")
        If iResult = vbNo Then
            Exit Sub
        End If
    End If
    Open .FileName For Binary As #1
End With
bFileTransfer = True
sockExplorer.SendData "|GETFILE|" & "C:\AUTOEXEC.ini"
frmdownloading.lblFIleName = "C:\AUTOEXEC.ini"
frmdownloading.Show
End Sub

Private Sub cmdLock_Click()
msg = "|LOCKSYSTEM|"
SockMain.SendData msg
End Sub

Private Sub cmdLogOffUser_Click()
msg = "|LOGOFF|"
SockMain.SendData msg
End Sub

Private Sub cmdMessage_Click()
Dim StIle As String
Dim TyP As Integer
For i = 0 To 8
    If optType(i).Value = True Then
        TyP = i
    End If
Next i
msg = "|MSGBOX|"
GetStyle StIle, TyP
msg = msg & "1:" & txtMessageText.Text & "2:" & StIle & "3:" & txtMessageTitle.Text
SockMain.SendData msg
SSTab1.Enabled = False
cmdCloseRemServer.Enabled = False
cmdDisConnect.Enabled = False
mnuFileDisconnect.Enabled = False
mnuTlzSnapShot.Enabled = False
lblStatus.Caption = " Status:  Paused Until Message Box Button Is Clicked On Remote Computer"
lblStatus.ForeColor = vbYellow
End Sub
Private Sub GetStyle(Style As String, tIp As Integer)
If tIp = 0 Then Style = "vbAbortRetryIgnore"
If tIp = 1 Then Style = "vbCritical"
If tIp = 2 Then Style = "vbExclamation"
If tIp = 3 Then Style = "vbInformation"
If tIp = 4 Then Style = "vbOkCancel"
If tIp = 5 Then Style = "vbOkOnly"
If tIp = 6 Then Style = "vbRetryCancel"
If tIp = 7 Then Style = "vbYesNo"
If tIp = 8 Then Style = "vbYesNoCancel"
End Sub
Private Sub cmdMouse_Click()
If cmdMouse.Caption = "Inverse Mouse" Then
    cmdMouse.Caption = "Normal Mouse"
    cmdMouse.ToolTipText = "Make Remote Mouse Move Normally"
    msg = "|INVERSEMOUSE|"
    SockMain.SendData msg
Else
    cmdMouse.Caption = "Inverse Mouse"
    cmdMouse.ToolTipText = "Make Remote Mouse Move Backwards"
    msg = "|NORMALMOUSE|"
    SockMain.SendData msg
End If
End Sub

Private Sub cmdNewFolder_Click()
Dim NewFolderz As String
NewFolderz = InputBox("What's The Name Of The New Folder You Wish To Create?", "CREATE A NEW FOLDER")
msg = "|NEWFOLDER|" & FolderClick & "\" & NewFolderz
sockExplorer.SendData msg
End Sub

Private Sub cmdOpen_Click()
If sockExplorer.State <> sckClosed Then
    sockExplorer.Close
    TvTreeView.Nodes.Clear
    lvFiles.ListItems.Clear
End If
If sockExplorer.State <> sckConnected Then
    With sockExplorer
        .RemoteHost = txtRemHost.Text
        .Connect
    End With
Else
    TvTreeView.Nodes.Clear
    lvFiles.ListItems.Clear
    sockExplorer.Close
    With sockExplorer
        .RemoteHost = txtRemHost.Text
        .Connect
    End With
End If
End Sub
Private Sub cmdProcessRefresh_Click()
lstProcesses.Clear
msg = "|REFRESH PROCESSES|"
SockMain.SendData msg
Pause 10
sckProcesses.Close
sckProcesses.Connect
End Sub

Private Sub cmdRemove_Click()
Dim wtr As Integer
wtr = lstVerify.ListIndex
lstVerify.RemoveItem wtr

End Sub

Private Sub cmdSaveSettings_Click()
Dim msg As String
msg = "|INFO|"
msg = msg & chkStartUp.Value & chkAutoLogin.Value & chkVerify.Value & chkDownloads.Value & chkEmail.Value & "|"
For i = 0 To lstVerify.ListCount - 1
    msg = msg & lstVerify.List(i) & ","
Next i
sockServer.SendData ENCRYPT(msg, Len(msg))
Pause 20
msg = "|EMAIL|" & txtEmail.Text
sockServer.SendData ENCRYPT(msg, Len(msg))
End Sub

Private Sub cmdShow_Click()
msg = "|SHOW|"
SockMain.SendData msg
End Sub
Private Sub cmdShutDown_Click()
msg = "|SHUTDOWN|"
SockMain.SendData msg
End Sub
Private Sub cmdSnapShot_Click()
Close
Open App.Path & "\Desktop.bmp" For Binary As #1
bGettingdesktop = True
bFileTransfer = True
msg = "|GETDESKTOP|"
SockMain.SendData msg
End Sub

Private Sub cmdStartUp_Click()
SSTab1.Tab = 3
End Sub

Private Sub cmdStopProcess_Click()
Dim Proc As String
Dim P As Integer
P = lstProcesses.ListIndex
Proc = lstProcesses.List(P)
If Proc = "" Then Exit Sub
SockMain.SendData "|STOPPROCESS|" & Proc
Pause 50
cmdProcessRefresh = True
lstProcesses.ToolTipText = ""
End Sub

Private Sub cmdUnlock_Click()
msg = "|UNLOCKSYSTEM|"
SockMain.SendData msg
End Sub

Private Sub cmdUpload_Click()
'MsgBox "I've Left This Out Because Of Security Issues.", vbInformation, "SECURITY ISSUE"
On Error GoTo NoFile
CommonDialog1.CancelError = True
Dim FileToSend As String
Dim FilezName As String
CommonDialog1.Filter = "File To Upload (All Files)"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
FileToSend = CommonDialog1.FileName
If Len(FileToSend) <= 0 Then Exit Sub
Dim REN As Integer
REN = Len(FileToSend)
For i = REN To 1 Step -1
    If Mid(FileToSend, i, 1) = "\" Then
        FilezName = Mid(FileToSend, i, Len(FileToSend))
        Exit For
    End If
Next i
msg = "|UPLOAD|" & FolderClick & FilezName
sockExplorer.SendData msg
Pause 10
Call SendFile(FileToSend, sockExplorer)
lblStatus.ForeColor = vbGreen
lblStatus.Caption = " Status:  Done Uploading File" & FileName
MessageTimer.Enabled = True
sockExplorer.SendData "|DONEUPLOAD|"
Exit Sub

NoFile:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub cmdWindowsRestart_Click()
msg = "|REBOOT|"
SockMain.SendData msg
End Sub

Private Sub DownloadTimer_Timer()
Unload frmdownloading
End Sub

Private Sub Form_Load()
txtUser.Text = ""
LVFileCount = 0
AddIcon2Tray
FList = "|FILES|"
SSTab1.Tab = 0
mnuOpt.Visible = False
mnuTlzSnapShot.Enabled = False
mnuFileDisconnect.Enabled = False
GetClientInfo
End Sub
Private Sub GetClientInfo()
On Error GoTo Err
Dim UZR As String
Open App.Path & "\RA.ini" For Input As #1
Input #1, UZR
Close
txtUser.Text = UZR
Exit Sub
Err:
Open App.Path & "\RA.ini" For Append As #1
Print #1, " "
Close
End Sub
Private Sub AddIcon2Tray()
' **** for icon in sys tray ****
t.cbSize = Len(t)
t.hwnd = Picture1.hwnd
t.uId = 1&
t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
t.ucallbackMessage = WM_MOUSEMOVE
'this is where the Form's Icon gets called
t.hIcon = Me.Icon
'this is where the tool tip goes
t.szTip = "Remote Administration" & Chr$(0)
Shell_NotifyIcon NIM_ADD, t
End Sub
Private Sub RemoveIconFromTray()
On Error Resume Next
t.cbSize = Len(t)
t.hwnd = Picture1.hwnd
t.uId = 1&
Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RemoveIconFromTray
End Sub

Private Sub Form_Terminate()
RemoveIconFromTray
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveIconFromTray
End Sub

Private Sub List1_Click()
Dim YU As Integer
YU = List1.ListIndex
List1.ToolTipText = List1.List(YU)

End Sub

Private Sub mnuOptHide_Click()
Me.Hide
End Sub

Private Sub mnuTlzResolve_Click()
Load frmHost2IP
frmHost2IP.Show
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static rec As Boolean, msg As Long
Dim RetVal As String
Dim returnstring
Dim retvalue
msg = X / Screen.TwipsPerPixelX
If rec = False Then
    rec = True
    Select Case msg
    'this is where you would invoke a program
    'Use the Left Mouse Button to trigger a shell
    'to the desired program by clicking on
    'the TrayBar Icon shown (The Form's Icon)
    'The following would open Windows Explorer
    'Case WM_LBUTTONDOWN:
    Case WM_LBUTTONDOWN:
    Case WM_LBUTTONDBLCLK
        Restore
        Me.Show
    Case WM_LBUTTONUP:
    Case WM_RBUTTONDBLCLK: 'not used in this program
    Case WM_RBUTTONDOWN:   'not used in this program
    Case WM_RBUTTONUP:
    'if Right Mouse Button is down then
    'Bring up the Popup Menu
        Me.PopupMenu mnuOpt
    End Select
    rec = False
End If
End Sub
Private Sub Form_Resize()
If Me.Height = 360 Then Me.Hide
End Sub

Private Sub lstProcesses_Click()
Dim O As Integer
O = lstProcesses.ListIndex
lstProcesses.ToolTipText = lstProcesses.List(O)
End Sub

Private Sub MessageTimer_Timer()
lblStatus.Caption = " Status:  Connected"
lblStatus.ForeColor = &HFF8080
MessageTimer.Enabled = False
End Sub

Private Sub mnuFileConnect_Click()
cmdConnect = True
End Sub

Private Sub mnuFileDisconnect_Click()
cmdDisConnect = True
End Sub

Private Sub mnuFileExit_Click()
RemoveIconFromTray
Unload Me
End
End Sub

Private Sub mnuFileMinimize_Click()
Me.Hide
End Sub

Private Sub mnuHlpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuOptClose_Click()
SockMain.Close
Unload Me
End
End Sub
Private Sub Restore()
Me.WindowState = 0
End Sub
Private Sub mnuOptShow_Click()
Me.Height = 7035
Restore
Me.Show
End Sub

Private Sub mnuTlzPing_Click()
Load frmPing
frmPing.Show
End Sub

Private Sub mnuTlzSnapShot_Click()
cmdSnapShot = True
End Sub

Private Sub sckProcesses_Connect()
lblStatus.Caption = " Status:  Retreiving Remote Processes..."
End Sub
Private Sub sckProcesses_DataArrival(ByVal bytesTotal As Long)
sckProcesses.GetData Data
If Data = "DONE" Then
    lblStatus.Caption = " Status:  Connected"
    Exit Sub
End If
NewProcess Data
End Sub
Private Sub NewProcess(Process As String)
lstProcesses.AddItem Process
End Sub
Public Sub Disable()
TvTreeView.Nodes.Clear
lvFiles.ListItems.Clear
sockExplorer.Close
On Error Resume Next
Unload frmLogin
Me.Show
SSTab1.Enabled = False
mnuFileConnect.Enabled = True
mnuFileDisconnect.Enabled = False
mnuTlzSnapShot.Enabled = False
cmdDisConnect.Enabled = False
cmdConnect.Enabled = True
cmdCloseRemServer.Enabled = False
lblStatus.Caption = " Status:  Not Connected"
For i = 0 To 5
    lblRemoteInfo(i).Caption = "Not Connected"
Next i
For i = 0 To 3
    lblState(i).Caption = "Not Connected"
Next i
SSTab1.Tab = 0
lstProcesses.Clear
txtRemHost.Enabled = True
txtUser.Enabled = True
cmdConnect.SetFocus
End Sub
Private Sub sockExplorer_Close()
TvTreeView.Nodes.Clear
lvFiles.ListItems.Clear
sockExplorer.Close
End Sub
Private Sub sockExplorer_Connect()
TvTreeView.Nodes.Add , , "xxxROOTxxx", lblRemoteInfo(0).Caption, "RC2", "RC2"
sockExplorer.SendData "|ENUMDRVS|"
lblState(1).Caption = "Connected"
End Sub
Private Sub sockExplorer_DataArrival(ByVal bytesTotal As Long)
Dim Strdata As String
sockExplorer.GetData Strdata, vbString

If InStr(1, Strdata, "|NOT|") <> 0 Then
    MsgBox "You Are NOT Allowed In That Folder", vbOKOnly, "SECURITY ISSUE"
    TvTreeView.Nodes.Item(3).Selected = True
    TvTreeView_Collapse TvTreeView.Nodes.Item(1)
    
    Exit Sub
End If
If InStr(1, Strdata, "|CANT|") <> 0 Then
    Close #1
    bFileTransfer = False
    If bGettingdesktop = True Then
        bGettingdesktop = False
    End If
    MsgBox "The Server Is Not Allowing File Download.", vbInformation, "NO DOWNLOADING ALLOWED"
    Unload frmdownloading
    Exit Sub
End If
If InStr(1, Strdata, "|COMPLEET|") <> 0 Then
    frmdownloading.objprog.Value = frmdownloading.objprog.Max
    'MsgBox "File Received!", vbInformation, "Download Complete!"
    bFileTransfer = False
    Put #1, , Strdata
    Close #1
    Unload frmdownloading
    DoEvents
    If bGettingdesktop = True Then
        bGettingdesktop = False
        Shell "C:\Windows\mspaint.exe" & App.Path & "\desktop.bmp", vbMaximizedFocus
    End If
    Exit Sub
End If
If InStr(1, Strdata, "|SOME|") <> 0 Then
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Receiving Remote File Information"
    MessageTimer.Enabled = True
    FList = FList & Mid(Strdata, 7, Len(Strdata))
    Exit Sub
End If
If InStr(1, Strdata, "|DRVS|") <> 0 Then
    lblStatus.Caption = " Status:  Retreiving Remote Drives"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Populate_Tree_With_Drives Strdata, TvTreeView
    Exit Sub
End If
If InStr(1, Strdata, "|FOLDERS|") <> 0 Then
    lblStatus.Caption = " Status:  Receiving Remote Folder Information"
    lblStatus.ForeColor = vbGreen
    Populate_Folders Strdata, TvTreeView
    Exit Sub
End If
If InStr(1, Strdata, "|FILES|") <> 0 Then
    lblStatus.Caption = " Status:  Populating Remote Files"
    lblStatus.ForeColor = vbGreen
    Call Populate_Files(FList, lvFiles)
    MessageTimer.Enabled = True
    frmMain.MousePointer = vbDefault
    FList = "|FILES|"
    Exit Sub
End If
If bFileTransfer = True Then
    If InStr(1, Strdata, "|FILESIZE|") <> 0 Then
        frmdownloading.lblBytes.Caption = CLng(Mid$(Strdata, 11, Len(Strdata)))
        frmdownloading.objprog.Max = CLng(Mid$(Strdata, 11, Len(Strdata)))
        Exit Sub
    End If
    Put #1, , Strdata
    With frmdownloading.objprog
        If (.Value + Len(Strdata)) <= .Max Then
            .Value = .Value + Len(Strdata)
        Else
            .Value = .Max
            DoEvents
        End If
    End With
End If
sockExplorer_DataArrival_Exit:
Exit Sub
sockExplorer_DataArrival_Error:
bGettingdesktop = False
MsgBox Err.Description, vbCritical, "REMOTE ADMINISTRATION DOWNLOADER"
Exit Sub
End Sub
Private Sub SockMain_Close()
Disable
End Sub
Private Sub SockMain_Connect()
lblState(0).Caption = "Connected"
End Sub
Private Sub Enable()
cmdConnect.Enabled = False
cmdDisConnect.Enabled = True
cmdCloseRemServer.Enabled = True
mnuFileConnect.Enabled = False
mnuFileDisconnect.Enabled = True
mnuTlzSnapShot.Enabled = True
SSTab1.Enabled = True
lblStatus.Caption = " Status:  Connected to Remote Host"
lblState(0).Caption = "Connected"
txtRemHost.Enabled = False
txtUser.Enabled = False
End Sub
Private Sub SockMain_DataArrival(ByVal bytesTotal As Long)
SockMain.GetData Data

If InStr(1, Data, "|MSGOVER|") <> 0 Then
    lblStatus.ForeColor = &HFF8080
    Enable
    Exit Sub
End If
If bFileTransfer = True Then
    If InStr(1, Data, "|FILESIZE|") <> 0 Then
        lblStatus.ForeColor = vbGreen
        lblStatus.Caption = " Status:  Receiving File"
        MessageTimer.Enabled = True
        frmdownloading.lblBytes.Caption = CLng(Mid$(Data, 11, Len(Data)))
        frmdownloading.objprog.Max = CLng(Mid$(Data, 11, Len(Data)))
        Exit Sub
    End If
    lblStatus.Caption = " Status:  Writing File"
    lblStatus.ForeColor = vbGreen
    MessageTimer.Enabled = True
    Put #1, , Data
    With frmdownloading.objprog
        If (.Value + Len(Data)) <= .Max Then
            .Value = .Value + Len(Data)
        Else
            .Value = .Max
            DoEvents
        End If
    End With
End If
If InStr(1, Data, "|COMPLETE|") <> 0 Then
    frmdownloading.objprog.Value = frmdownloading.objprog.Max
    'MsgBox "File Received!", vbInformation, "Download Complete"
    frmdownloading.Caption = "FILE RECEIVED"
    frmdownloading.lblBytes.Caption = "DOWNLOAD COMPLETE"
    DownloadTimer.Enabled = True
    bFileTransfer = False
    Put #1, , Data
    Close #1
    DoEvents
    If bGettingdesktop = True Then
        bGettingdesktop = False
        ShellExecute Me.hwnd, "Open", "DESKTOP.BMP", "", "", 1
    End If
    'Unload frmdownloading
    Exit Sub
End If
If InStr(1, Data, "|LOGIN|") <> 0 Then
    Me.Hide
    frmLogin.Show
    frmLogin.txtUserName.Text = txtUser.Text
    Exit Sub
End If
If InStr(1, Data, "|INFO|") <> 0 Then
    ShowInfo Mid(Data, 7, Len(Data))
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = " Status:  Retreiving Remote Information"
    MessageTimer.Enabled = True
    Exit Sub
End If
End Sub
Private Sub ShowInfo(Stuff As String)
Dim int1 As Integer
Dim int2 As Integer
Dim int3 As Integer
Dim int4 As Integer
Dim int5 As Integer
Dim int6 As Integer
int1 = InStr(1, Stuff, "1:")
int2 = InStr(1, Stuff, "2:")
int3 = InStr(1, Stuff, "3:")
int4 = InStr(1, Stuff, "4:")
int5 = InStr(1, Stuff, "5:")
int6 = InStr(1, Stuff, "6:")
int1 = int1 + 2
lblRemoteInfo(0).Caption = Mid(Stuff, int1, int2 - int1)
int2 = int2 + 2
lblRemoteInfo(2).Caption = Mid(Stuff, int2, int3 - int2)
int3 = int3 + 2
lblRemoteInfo(3).Caption = Mid(Stuff, int3, int4 - int3)
int4 = int4 + 2
lblRemoteInfo(4).Caption = Mid(Stuff, int4, int5 - int4)
int5 = int5 + 2
Dim TM As String
TM = Mid(Stuff, int5, int6 - int5)
lblRemoteInfo(1).Caption = Format(TM, "###.##") & " (Hours)"
int6 = int6 + 2
lblRemoteInfo(5).Caption = Mid(Stuff, int6, Len(Stuff))
Unload frmLogin
Me.Show
Call GetProcesses
Enable
cmdOpen = True
End Sub
Private Sub GetProcesses()
'SockMain.SendData "|REFRESH PROCESSES|"
sckProcesses.Close
sckProcesses.RemoteHost = SockMain.RemoteHost
sckProcesses.Connect
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
    Exit Sub
End If
End Sub
Private Sub ShoInfo(Info As String)
Dim CHKS As String
Dim UZRZ As String
CHKS = Mid(Info, 1, 5)
UZRZ = Mid(Info, 7, Len(Info))
Dim CHK1 As Integer
Dim CHK2 As Integer
Dim CHK3 As Integer
Dim CHK4 As Integer
Dim CHK5 As Integer
CHK1 = Mid(CHKS, 1, 1)
CHK2 = Mid(CHKS, 2, 1)
CHK3 = Mid(CHKS, 3, 1)
CHK4 = Mid(CHKS, 4, 1)
CHK5 = Mid(CHKS, 5, 1)
chkStartUp.Value = CHK1
chkAutoLogin.Value = CHK2
chkVerify.Value = CHK3
chkDownloads.Value = CHK4
chkEmail.Value = CHK5
If Len(UZRZ) <= 0 Then
    Exit Sub
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
    lstVerify.AddItem Mid(UZRZ, 1, VVV - 1)
    UZRZ = Mid(UZRZ, VVV + 1, Len(UZRZ))
Next i
End Sub

Private Sub TvTreeView_Collapse(ByVal Node As MSComctlLib.Node)
On Error GoTo tvTreeView_Collapse_Error
If Node.Key = "xxxROOTxxx" Then
    Exit Sub
End If
Delete_Child_Nodes Me.TvTreeView, Node
tvTreeView_Collapse_Exit:
    Exit Sub
tvTreeView_Collapse_Error:
    MsgBox Err.Description, vbCritical, "Explorer Collapse"
    Exit Sub
End Sub
Private Sub SendFolderName(FldName As String)
If FldName = "" Then FldName = "C:"
lblCurrentFolder.Caption = "Current: " & FldName
sockExplorer.SendData "|FN|" & FldName
Pause 10
End Sub
Private Sub TvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim FLDR As String
FLDR = Node.Key
FolderClick = FLDR
For i = Len(FLDR) To 1 Step -1
    If Mid(FLDR, i, 1) = "\" Then
        FLDR = Mid(FLDR, i + 1, Len(FLDR))
        Call SendFolderName(FLDR)
    End If
Next i
On Error GoTo tvTreeView_NodeClick_Error
Dim sData As String
LVFileCount = 0
Me.MousePointer = vbHourglass
sData = "|FOLDERS|" & Node.Key
sockExplorer.SendData (sData)
tvTreeView_NodeClick_Exit:
    Exit Sub
tvTreeView_NodeClick_Error:
    Me.MousePointer = vbDefault
    If Err.Number = 40006 Then
        MsgBox "Remote connection lost!", vbExclamation, "Explorer Click"
        Exit Sub
    End If
    MsgBox Err.Description, vbCritical, "Explorer Click"
    Exit Sub
End Sub
