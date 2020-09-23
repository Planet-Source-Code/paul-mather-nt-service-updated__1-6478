VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NT Service Example (NTServ)"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheckRun 
      Caption         =   "Check if Running"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdCheckInstall 
      Caption         =   "Check if Installed"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Service"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Service"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblAdd 
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label lblRemove 
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblRun 
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblCheckRun 
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label lblCheckInstall 
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    lblAdd.Caption = InstallService(App.Title, e_ServiceType_Automatic, App.Path & "\" & App.EXEName & ".exe")
End Sub
Private Sub cmdRemove_Click()
    lblRemove.Caption = RemoveService(App.Title)
End Sub
Private Sub cmdCheckInstall_Click()
    lblCheckInstall.Caption = CheckServiceInstalled(App.Title)
End Sub
Private Sub cmdCheckRun_Click()
Dim serviceRunning As e_ServiceState
    If CheckServiceRunning(App.Title, serviceRunning) = True Then
        If serviceRunning <> e_ServiceState_Stopped Then
            lblCheckRun.Caption = "True"
        Else
            lblCheckRun.Caption = "False"
        End If
    Else
        lblCheckRun.Caption = "False"
    End If
End Sub
Private Sub cmdRun_Click()
    ' I can never get this to work.
    lblRun.Caption = RunService(App.Title)
End Sub
