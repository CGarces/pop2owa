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


Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type

Public Const VER_PLATFORM_WIN32_NT = 2&
Public Const Service_Name  As String = "POP2OWA"

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


''
'API para provocar un retardo del sistema
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent
Public IsNT As Boolean, IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long


''
'Write a pop2owa.err file with errors.
'
'@param strText Test to write in the log
Public Sub WriteLog(ByVal strText As String)
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\pop2owa.err" For Append As #intFile
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
        hStopEvent = CreateEvent(0, 1, 0, vbNullString)
        hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
        hStartEvent = CreateEvent(0, 1, 0, vbNullString)
        ServiceName = StrConv(Service_Name, vbFromUnicode)
        ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))

        'Create the service
        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        'Waiting for one of two events: sucsessful service start (1) or Terminaton of service thread (0)
        IsNTService = WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&
    Else
        IsNTService = False
    End If
    
If IsNTService Then
    App.LogEvent "Tratando de iniciar: " & Service_Name
    
    'Get info from registry
    Dim c As cRegistry
    Set c = New cRegistry
    Set oPOP3 = New clsPOP3
    'Values are stored in (HKEY_CURRENT_USER\Software\pop2owa)
    With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\pop2owa"
        .ValueType = REG_SZ
        WriteLog "strExchSvrName =" & strExchSvrName
        'Exchange server
        .ValueKey = "ExchangeServer"
        strExchSvrName = .Value
        WriteLog "strExchSvrName =" & strExchSvrName
        
        'IP to listen
        .ValueKey = "IP"
        oPOP3.IP = .Value
        
        'Port values are integers
        .ValueType = REG_DWORD
        'POP3 Port
        .ValueKey = "POP3"
        oPOP3.Port(0) = .Value
        WriteLog "Port(0) =" & .Value
        'SMTP Port
        .ValueKey = "SMTP"
        oPOP3.Port(1) = .Value
        WriteLog "Port(1) =" & .Value
        'Leave a copy in send folder
        .ValueKey = "Saveinsent"
        oPOP3.Saveinsent = (.Value = vbChecked)
        'Form-Based-Authentication on/off
        .ValueKey = "FormBasedAuth"
        oPOP3.FormBasedAuthentication = (.Value = vbChecked)
        App.LogEvent "Start: " & Service_Name
        oPOP3.Start
        App.LogEvent "Start valido: " & Service_Name
    End With
 
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
    Loop While WaitForSingleObject(hStopPendingEvent, lngInterval) = WAIT_TIMEOUT
    oPOP3.Destroy
    Set oPOP3 = Nothing
    SetServiceState SERVICE_STOPPED
    App.LogEvent "Stoping Service: " & Service_Name
    SetEvent hStopEvent
    ' Waiting for service thread termination
    WaitForSingleObject hnd, INFINITE
    CloseHandle hnd
Else
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
    WriteLog Err.Source & vbTab & Err.Description
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
