Attribute VB_Name = "modGlobal"
''
'Store al public variables, constants and functions.

Option Explicit
Public strUser          As String
Public strPassWord      As String
Public objDOMMsg        As DOMDocument
Public objDOMInbox      As DOMDocument
Public strExchSvrName   As String

Public bSaveinsent      As Boolean
Public bAuthentication  As Boolean
Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"

Public Const RETARDO_BUCLE = 30000

Public Const INFINITE = -1&      '  Timeout infinito
Private Const WAIT_TIMEOUT = 258&
Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400

Public Const VER_PLATFORM_WIN32_NT = 2&
Private Const STATUS_TIMEOUT = &H102&
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)



Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type

Public Const Service_Name  As String = "POP2OWA"

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long

''
'API para provocar un retardo del sistema
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent As Long
Public IsNT As Boolean, IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long


''
'Write a pop2owa.err file with errors.
'
'@param strText Test to write in the log
Public Sub WriteLog(ByVal strText As String)
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\pop2owa.log" For Append As #intFile
    Print #intFile, Now & vbTab & strText
Close #intFile
End Sub

Public Sub Main()
On Error GoTo ErrHandler
Dim hnd As Long
Dim h(0 To 1) As Long
Dim strParametro As String
Dim hProcess As Long
Dim RetVal As Long
Dim objFrame As Frame
Dim lngInterval As Long
Dim oPOP3 As clsPOP3

    'Check OS
    If CheckIsNT() Then
    
        'Events for NT Service
        hStopEvent = CreateEventW(0&, 1&, 0&, 0&)
        hStopPendingEvent = CreateEventW(0&, 1&, 0&, 0&)
        hStartEvent = CreateEventW(0&, 1&, 0&, 0&)
        
        ServiceName = StrConv(Service_Name, vbFromUnicode)
        ServiceNamePtr = StrPtr(Service_Name)
        'ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))

        'Create the service
        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        'Waiting for one of two events: sucsessful service start (1) or Terminaton of service thread (0)
        IsNTService = MsgWaitObj(INFINITE, h(0), 2&) = 1&
        
    Else
        IsNTService = False
    End If
    
If IsNTService Then
    App.LogEvent "Starting: " & Service_Name
    Set oPOP3 = New clsPOP3
    
    'Run the NT Service
    SetServiceState SERVICE_RUNNING
    App.LogEvent "Running Service: " & Service_Name
    lngInterval = 10000
    Do
        oPOP3.Refresh
        If oPOP3.isActive Then
            lngInterval = 100
        Else
            lngInterval = 10000
        End If
        DoEvents
    Loop While MsgWaitObj(lngInterval, hStopPendingEvent, 1&) = WAIT_TIMEOUT
    WriteLog "Fin bucle"
    Set oPOP3 = Nothing
    WriteLog "Objeto destruido"
    SetServiceState SERVICE_STOPPED
    WriteLog "Peticion de parada de servicio"
    App.LogEvent "Stoping Service: " & Service_Name
    SetEvent hStopEvent
    ' Waiting for service thread termination
    MsgWaitObj INFINITE, hnd, 1&
    CloseHandle hnd
    WriteLog "Servicio parado"
Else
    hStopPendingEvent = 0
    With frmMain
        .Caption = .Caption & " " & App.Major & "." & App.Minor
        'Redraw controls
        For Each objFrame In .Frame1
            objFrame.Move .TabStrip1.clientLeft, _
                        .TabStrip1.clientTop, _
                        .TabStrip1.clientWidth, _
                        .TabStrip1.clientHeight
        Next
        .Move .Left, .Top, 2700, 3150
        Set objFrame = Nothing
        .Init
        .Show
    End With
End If
CloseHandle hStopEvent
CloseHandle hStartEvent
CloseHandle hStopPendingEvent
Exit Sub
ErrHandler:
    If IsNTService Then
        WriteLog Err.Source & vbTab & Err.Description
    Else
        MsgBox Err.Description, vbCritical, "POP2OWA: " & Err.Source
        Unload frmMain
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End Sub


''
' Check if O.S. is NT compatible
'
Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function


' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
' Unlike these functions, it
' doesn't block thread messages processing.
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.

Public Function MsgWaitObj(Interval As Long, _
            Optional hObj As Long = 0&, _
            Optional nObj As Long = 0&) As Long
    Dim T As Long, T1 As Long
    If Interval <> INFINITE Then
        T = GetTickCount()
        On Error Resume Next
        T = T + Interval
        ' Overflow prevention
        If Err <> 0& Then
            If T > 0& Then
                T = ((T + &H80000000) _
                + Interval) + &H80000000
            Else
                T = ((T - &H80000000) _
                + Interval) - &H80000000
            End If
        End If
        On Error GoTo 0
        ' T contains now absolute time of the end of interval
    Else
        T1 = INFINITE
    End If
    Do
        If Interval <> INFINITE Then
            T1 = GetTickCount()
            On Error Resume Next
         T1 = T - T1
            ' Overflow prevention
            If Err <> 0& Then
                If T > 0& Then
                    T1 = ((T + &H80000000) _
                    - (T1 - &H80000000))
                Else
                    T1 = ((T - &H80000000) _
                    - (T1 + &H80000000))
                End If
            End If
            On Error GoTo 0
            ' T1 contains now the remaining interval part
            If IIf((T1 Xor Interval) > 0&, _
                T1 > Interval, T1 < 0&) Then
                ' Interval expired
                ' during DoEvents
                MsgWaitObj = STATUS_TIMEOUT
                Exit Function
            End If
        End If
        ' Wait for event, interval expiration
        ' or message appearance in thread queue
        MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
                hObj, 0&, T1, QS_ALLINPUT)
        'Checks for hStopPendingEvent event
        If Not hStopPendingEvent = 0 Then
            If WaitForSingleObject(hStopPendingEvent, 0&) = 0& Then Exit Function
        End If
        ' Let's message be processed
        DoEvents
        If MsgWaitObj <> nObj Then Exit Function
        ' It was message - continue to wait
    Loop
End Function





