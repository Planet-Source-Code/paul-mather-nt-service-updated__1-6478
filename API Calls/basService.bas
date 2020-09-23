Attribute VB_Name = "basService"
Option Explicit
' This code was taken from the MSDN CDs and modified
' to allow for easier use.
' MSDN Topic: INFO: Running Visual Basic Applications as Windows NT Services
' MSDN Topic: HOWTO: Query an NT Service for Status and Configuration

Private Const SERVICE_WIN32_OWN_PROCESS = &H10&
Private Const SERVICE_WIN32_SHARE_PROCESS = &H20&
Private Const SERVICE_WIN32 = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS

Private Const SC_MANAGER_CONNECT = &H1
Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SC_MANAGER_LOCK = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Private Const SERVICE_CONTROL_CONTINUE = &H3
Private Const SERVICE_CONTROL_INTERROGATE = &H4
Private Const SERVICE_CONTROL_PAUSE = &H2
Private Const SERVICE_CONTROL_SHUTDOWN = &H5
Private Const SERVICE_CONTROL_STOP = &H1

Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7
Private Const SERVICE_ACCEPT_STOP = &H1
Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
   
Private Const SERVICE_DISABLED As Long = &H4
Private Const SERVICE_DEMAND_START As Long = &H3
Private Const SERVICE_AUTO_START  As Long = &H2
Private Const SERVICE_SYSTEM_START As Long = &H1
Private Const SERVICE_BOOT_START As Long = &H0

Private Const GENERIC_READ = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122


Public Enum e_ServiceType
    e_ServiceType_Disabled = 4
    e_ServiceType_Manual = 3
    e_ServiceType_Automatic = 2
    e_ServiceType_SystemStart = 1
    e_ServiceType_BootTime = 0
End Enum

Private Const SERVICE_ERROR_NORMAL As Long = &H1

Public Enum e_ServiceControl
   e_ServiceControl_Stop = &H1
   e_ServiceControl_Pause = &H2
   e_ServiceControl_Continue = &H3
   e_ServiceControl_Interrogate = &H4
   e_ServiceControl_Shutdown = &H5
End Enum

Public Enum e_ServiceState
   e_ServiceState_Stopped = &H1
   e_ServiceState_StartPending = &H2
   e_ServiceState_StopPending = &H3
   e_ServiceState_Running = &H4
   e_ServiceState_ContinuePending = &H5
   e_ServiceState_PausePending = &H6
   e_ServiceState_Paused = &H7
End Enum

Private Type SERVICE_TABLE_ENTRY
   lpServiceName As String
   lpServiceProc As Long
   lpServiceNameNull As Long
   lpServiceProcNull As Long
End Type

Private Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
End Type

Private Type QUERY_SERVICE_CONFIG
   dwServiceType As Long
   dwStartType As Long
   dwErrorControl As Long
   lpBinaryPathName As Long 'String
   lpLoadOrderGroup As Long ' String
   dwTagId As Long
   lpDependencies As Long 'String
   lpServiceStartName As Long 'String
   lpDisplayName As Long  'String
End Type

Private Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
Private Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long
Private Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As Byte, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
   

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long

Private hServiceStatus As Long
Private ServiceStatus As SERVICE_STATUS

Dim SERVICE_NAME As String

Public Function InstallService(ByVal serviceName As String, ByVal serviceType As e_ServiceType, Optional ByVal servicePath As String) As Boolean
Dim hSCManager As Long
Dim hService As Long
Dim cmd As String
Dim lServiceType As Long
    
   Const SERVICE_DISABLED As Long = &H4
   Const SERVICE_DEMAND_START As Long = &H3
   Const SERVICE_AUTO_START  As Long = &H2
   Const SERVICE_SYSTEM_START As Long = &H1
   Const SERVICE_BOOT_START As Long = &H0
    
    Select Case serviceType
        Case e_ServiceType_Automatic
            lServiceType = SERVICE_AUTO_START
        Case e_ServiceType_BootTime
            lServiceType = SERVICE_BOOT_START
        Case e_ServiceType_Disabled
            lServiceType = SERVICE_DISABLED
        Case e_ServiceType_Manual
            lServiceType = SERVICE_DEMAND_START
        Case e_ServiceType_SystemStart
            lServiceType = SERVICE_SYSTEM_START
    End Select
    
    If servicePath = "" Then
        servicePath = App.Path & "\" & App.EXEName & ".exe"
    End If
    
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    hService = CreateService(hSCManager, serviceName, serviceName, SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, lServiceType, SERVICE_ERROR_NORMAL, servicePath, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    If hService = 0 Then
        InstallService = False
    Else
        InstallService = True
    End If
    CloseServiceHandle hService
    CloseServiceHandle hSCManager
End Function
Public Function RemoveService(serviceName As String) As Boolean
Dim hSCManager As Long
Dim hService As Long
Dim ret As Long
Dim cmd As String
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    hService = OpenService(hSCManager, serviceName, SERVICE_ALL_ACCESS)
    ret = DeleteService(hService)
    If ret = 0 Then
        RemoveService = False
    Else
        RemoveService = True
    End If
    CloseServiceHandle hService
    CloseServiceHandle hSCManager
End Function
Public Function RunService(serviceName As String) As Boolean
Dim ServiceTableEntry As SERVICE_TABLE_ENTRY
Dim ret As Long
Dim servicePath As String

    If CheckServiceInstalled(serviceName) = False Then
        RunService = False
        Exit Function
    End If
    If CheckServiceRunning(serviceName, , , servicePath) = False Then
        RunService = False
        Exit Function
    End If
    If Dir(servicePath) = "" Then
        RunService = False
        Exit Function
    End If
        
    ServiceTableEntry.lpServiceName = serviceName
    SERVICE_NAME = serviceName
    ServiceTableEntry.lpServiceProc = FncPtr(AddressOf ServiceMain)
    ret = StartServiceCtrlDispatcher(ServiceTableEntry)
    If ret = 0 Then
        RunService = False
    Else
        RunService = True
    End If
End Function
Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
      Dim b As Boolean

      'Set initial state
      ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS
      ServiceStatus.dwCurrentState = SERVICE_START_PENDING
      ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_PAUSE_CONTINUE Or SERVICE_ACCEPT_SHUTDOWN
      ServiceStatus.dwWin32ExitCode = 0
      ServiceStatus.dwServiceSpecificExitCode = 0
      ServiceStatus.dwCheckPoint = 0
      ServiceStatus.dwWaitHint = 0

      hServiceStatus = RegisterServiceCtrlHandler(SERVICE_NAME, AddressOf Handler)
      ServiceStatus.dwCurrentState = SERVICE_START_PENDING
      b = SetServiceStatus(hServiceStatus, ServiceStatus)

      ServiceStatus.dwCurrentState = SERVICE_RUNNING
      b = SetServiceStatus(hServiceStatus, ServiceStatus)

   End Sub

Private Sub Handler(ByVal fdwControl As Long)
    Dim b As Boolean
    
    Select Case fdwControl
        Case SERVICE_CONTROL_PAUSE
            ServiceStatus.dwCurrentState = SERVICE_PAUSED
        Case SERVICE_CONTROL_CONTINUE
            ServiceStatus.dwCurrentState = SERVICE_RUNNING
        Case SERVICE_CONTROL_STOP
            ServiceStatus.dwWin32ExitCode = 0
            ServiceStatus.dwCurrentState = SERVICE_STOP_PENDING
            ServiceStatus.dwCheckPoint = 0
            ServiceStatus.dwWaitHint = 0
            b = SetServiceStatus(hServiceStatus, ServiceStatus)
            ServiceStatus.dwCurrentState = SERVICE_STOPPED
        Case SERVICE_CONTROL_INTERROGATE
        Case Else
    End Select
    b = SetServiceStatus(hServiceStatus, ServiceStatus)
End Sub

Function FncPtr(ByVal fnp As Long) As Long
    FncPtr = fnp
End Function
Public Function CheckServiceInstalled(ByVal serviceName As String) As Boolean
    If GetSetting(serviceName, "", "ImagePath", "Nothing", HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services") = "Nothing" Then
        CheckServiceInstalled = False
    Else
        CheckServiceInstalled = True
    End If
End Function

Public Function CheckServiceRunning(ByVal serviceName As String, Optional ByRef serviceRunning As e_ServiceState, Optional ByRef serviceStartType As e_ServiceType, Optional servicePath As String) As Boolean
Dim hSCM  As Long
Dim hSVC As Long
Dim pSTATUS As SERVICE_STATUS
Dim udtConfig As QUERY_SERVICE_CONFIG
Dim lRet As Long
Dim lBytesNeeded As Long
Dim sTemp As String
Dim pFileName As Long

    CheckServiceRunning = True
    ' Open The Service Control Manager
    '
    hSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If hSCM = 0 Then
       CheckServiceRunning = False
    End If

    ' Open the specific Service to obtain a handle
    '
    hSVC = OpenService(hSCM, Trim(serviceName), GENERIC_READ)
    If hSVC = 0 Then
        CheckServiceRunning = False
       'MsgBox "Error - " & Err.LastDllError
       GoTo CloseHandles
    End If

    ' Fill the Service Status Structure
    '
    lRet = QueryServiceStatus(hSVC, pSTATUS)
    If lRet = 0 Then
       CheckServiceRunning = False
       GoTo CloseHandles
    End If

    ' Report the Current State
    '
    Select Case pSTATUS.dwCurrentState
    Case SERVICE_STOP
       serviceRunning = e_ServiceState_Stopped
    Case SERVICE_START
       serviceRunning = e_ServiceState_StartPending
    Case SERVICE_STOP_PENDING
       serviceRunning = e_ServiceState_StopPending
    Case SERVICE_RUNNING
       serviceRunning = e_ServiceState_Running
    Case SERVICE_CONTINUE_PENDING
       serviceRunning = e_ServiceState_ContinuePending
    Case SERVICE_PAUSE_PENDING
       serviceRunning = e_ServiceState_PausePending
    Case SERVICE_PAUSED
       serviceRunning = e_ServiceState_Paused
    Case SERVICE_ACCEPT_STOP
       serviceRunning = e_ServiceState_Stopped
    Case SERVICE_ACCEPT_PAUSE_CONTINUE
       serviceRunning = e_ServiceState_Paused
    Case SERVICE_ACCEPT_SHUTDOWN
       serviceRunning = e_ServiceState_StopPending
    End Select

    ' Call QueryServiceConfig with 1 byte buffer to generate an error
    ' that returns the size of a buffer we need.
    '
    ReDim abConfig(0) As Byte
    lRet = QueryServiceConfig(hSVC, abConfig(0), 0&, lBytesNeeded)
    If lRet = 0 And Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
       CheckServiceRunning = False
    End If

    ' Redim our byte array to the size necessary and call
    ' QueryServiceConfig again
    '
    ReDim abConfig(lBytesNeeded) As Byte
    lRet = QueryServiceConfig(hSVC, abConfig(0), lBytesNeeded, _
       lBytesNeeded)
    If lRet = 0 Then
       CheckServiceRunning = False
       GoTo CloseHandles
    End If

    ' Fill our Service Config User Defined Type.
    '
    CopyMemory udtConfig, abConfig(0), Len(udtConfig)

    serviceStartType = udtConfig.dwStartType

    sTemp = Space(255)

    ' Now use the pointer obtained to copy the path into the temporary
    ' String Variable
    '
    lRet = lstrcpy(sTemp, udtConfig.lpBinaryPathName)
    servicePath = Trim(sTemp)

CloseHandles:
    ' Close the Handle to the Service
    '
    CloseServiceHandle (hSVC)
    
    ' Close the Handle to the Service Control Manager
    '
    CloseServiceHandle (hSCM)
End Function
