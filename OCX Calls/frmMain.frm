VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NT Service Example (NTServ)"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin NTService.NTService NTService1 
      Left            =   3240
      Top             =   1560
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
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
   Begin VB.Label lblRun 
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label lblStop 
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label lblAdd 
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label lblRemove 
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    NTService1.DisplayName = App.Title
    NTService1.StartMode = svcStartAutomatic
    NTService1.ControlsAccepted = svcCtrlPauseContinue
    lblRun.Caption = NTService1.StartService
End Sub

Private Sub cmdAdd_Click()
    lblAdd.Caption = NTService1.Install
End Sub
Private Sub cmdRemove_Click()
    lblRemove.Caption = NTService1.Uninstall
End Sub
Private Sub cmdRun_Click()
    lblRun.Caption = NTService1.StartService
End Sub
Private Sub cmdStop_Click()
    Call NTService1.StopService
    lblStop.Caption = True
End Sub

