Attribute VB_Name = "NTService"
''
'NT Service module.
'@author Original code from Sergey Merzlikin, modified by Carlos Garcés

Option Explicit


Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Private ServiceStatus As SERVICE_STATUS
Private hServiceStatus As Long
Private ServiceName() As Byte, ServiceNamePtr As Long
Private hStopEvent As Long, hStartEvent As Long


''
' The FncPtr function returns function pointer.
'
'@param fnp Address of the original function
'@return Pointer to the function
Private Function FncPtr(ByVal fnp As Long) As Long
    FncPtr = fnp
End Function


''
' The ServiceThread sub starts the service.
' This sub returns control only after service termination.
'
'@param dummy Descripción_del_parámetro
Private Sub ServiceThread(ByVal dummy As Long)
    Dim ServiceTableEntry As SERVICE_TABLE
    ServiceTableEntry.lpServiceName = ServiceNamePtr
    ServiceTableEntry.lpServiceProc = FncPtr(AddressOf ServiceMain)
    StartServiceCtrlDispatcherW ServiceTableEntry
End Sub

''
' It initializes service, sets event hStartEvent, and waits hStopEvent event.<BR>
' When hStopEvent fires, this sub exits and service stops.
Private Sub ServiceMain()
'Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
    ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS
    ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP _
                                    Or SERVICE_ACCEPT_SHUTDOWN
    ServiceStatus.dwWin32ExitCode = 0&
    ServiceStatus.dwServiceSpecificExitCode = 0&
    ServiceStatus.dwCheckPoint = 0&
    ServiceStatus.dwWaitHint = 0&
    hServiceStatus = RegisterServiceCtrlHandlerW(ServiceNamePtr, _
                           AddressOf Handler)
    SetServiceState SERVICE_START_PENDING
    ' Set hStartEvent. It unlocks main application thread
    ' which allows to do some work in it
    SetEvent hStartEvent
    ' Wait until hStopEvent fires
    WaitForSingleObject hStopEvent, INFINITE
End Sub
   
''
' The Handler sub processes commands from Service Dispatcher.
' It sets event hStopEvent when processes command
' SERVICE_CONTROL_STOP or SERVICE_CONTROL_SHUTDOWN.
'
'@param fdwControl Command to set the service (SERVICE_CONTROL_STOP or SERVICE_CONTROL_SHUTDOWN)
Private Sub Handler(ByVal fdwControl As Long)
    Select Case fdwControl
        Case SERVICE_CONTROL_SHUTDOWN, SERVICE_CONTROL_STOP
            SetServiceState SERVICE_STOP_PENDING
            SetEvent hStopPendingEvent
        Case Else
            SetServiceState
    End Select
End Sub



''
' The SetServiceState sub changes service state.
' If parameter omitted, it confirms previous state.
'
'@param NewState New service state
Private Sub SetServiceState(Optional ByVal NewState As SERVICE_STATE = 0&)
    If NewState <> 0& Then ServiceStatus.dwCurrentState = NewState
    SetServiceStatus hServiceStatus, ServiceStatus
End Sub


''
'Start the NT service of Pop2owa
Public Sub StartService()
On Error GoTo ErrHandler

Dim bIsNT           As Boolean
Dim h(0 To 1)       As Long
Dim hnd             As Long

Const WAIT_TIMEOUT = 258&

'Check OS
GetWindowsVersion 0, 0, 0, 0, bIsNT
If bIsNT Then

    'Events for NT Service
    hStopEvent = CreateEventW(0&, 1&, 0&, 0&)
    hStopPendingEvent = CreateEventW(0&, 1&, 0&, 0&)
    hStartEvent = CreateEventW(0&, 1&, 0&, 0&)
    
    ServiceName = StrConv(Service_Name, vbFromUnicode)
    ServiceNamePtr = StrPtr(Service_Name)

    'Create the service
    hnd = CreateThread(0&, 0&, AddressOf ServiceThread, 0&, 0&, 0&)
    
    h(0) = hnd
    h(1) = hStartEvent
    'Waiting for one of two events: successful service start (1) or Termination of service thread (0)
    IsNTService = MsgWaitObj(INFINITE, h(0), 2&) = 1&
    If Not IsNTService Then
        CloseHandle hnd
        MessageBox 0&, "This program must be started as a service.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
    End If
Else
    MessageBox 0&, "This program is only for Windows NT/2000/XP/2003.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
End If

If IsNTService Then
    App.LogEvent "Starting: " & Service_Name
    Set oPOP3 = New clsPOP3
    'Run the NT Service
    SetServiceState SERVICE_RUNNING
    App.LogEvent "Running Service: " & Service_Name
    Do
        DoEvents
    Loop While MsgWaitObj(100, hStopPendingEvent, 1&) = WAIT_TIMEOUT
    Set oPOP3 = Nothing
    SetServiceState SERVICE_STOPPED
    WriteLog "Ask for Stoping Service", Information
    App.LogEvent "Stoping Service: " & Service_Name
    SetEvent hStopEvent
    ' Waiting for service thread termination
    MsgWaitObj INFINITE, hnd, 1&
    CloseHandle hnd
    WriteLog "Service Stopped", Information
End If
CloseHandle hStopEvent
CloseHandle hStartEvent
CloseHandle hStopPendingEvent
Exit Sub
ErrHandler:
    WriteLog Err.Source & vbTab & Err.Description, Fail
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
    Err.Raise vbObjectError + 1, "Starting NT Service", "Error starting NT Service"
End Sub
