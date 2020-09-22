VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdownloading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Admin (Downloading)"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmdownloading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame a 
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3780
      Begin MSComctlLib.ProgressBar objprog 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "File Size (Bytes):"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "File Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblFIleName 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label lblBytes 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmdownloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If lblFIleName.Caption = "" Then
    lblFIleName.Caption = "Remote Snapshot"
End If
Me.Show
Me.Refresh
DoEvents
End Sub
