Attribute VB_Name = "modGlobal"
''
'Store al public variables, constants and functions.

Option Explicit
Public strUser          As String
Public strPassWord      As String
Public objDOMInbox      As DOMDocument
Public strExchSvrName   As String

Private intVerbosity As Integer
Public Enum Verbosity
    Error = 0
    Warning = 1
    Information = 2
    Paranoid = 3
End Enum

Public bSaveinsent      As Boolean
Public bAuthentication  As Boolean
Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"

Public Const INFINITE = -1&      '  Timeout infinito
Private Const WAIT_TIMEOUT = 258&

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
''
' API go get the OS version
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long

''
'API para provocar un retardo del sistema
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent As Long
Public IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''
'variable constant to match if the mutex exists
Private Const ERROR_ALREADY_EXISTS = 183&


Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
'-------------------------------------------------------
Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

''
'Write a pop2owa.log file with errors, warnings and log messages.
'
'@param strText Test to write in the log
'@param intVerbose Versosity
Public Sub WriteLog(ByVal strText As String, Optional ByVal intVerbose As Verbosity = Error)
Dim intFile As Integer
If intVerbose <= intVerbosity Then
    intFile = FreeFile
    Open App.Path & "\pop2owa.log" For Append As #intFile
        Print #intFile, Now & vbTab & strText
    Close #intFile
    Debug.Print Now & "-> " & strText
End If
End Sub


''
'Main function.
Public Sub Main()
On Error GoTo ErrHandler
Dim hnd             As Long
Dim h(0 To 1)       As Long
Dim strParametro    As String
Dim hProcess        As Long
Dim RetVal          As Long
Dim lngInterval     As Long
Dim oPOP3           As clsPOP3
Dim bIsNT           As Boolean
Dim mutexvalue      As Long

'Create an individual mutex value for the application
mutexvalue = CreateMutex(ByVal 0&, 1, App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision)

'If an error occured creating the mutex, that means it
'must have already existed, therefore your application
'is already running
If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
    'Terminate the application via the reference to it, its hObject value
    CloseHandle mutexvalue
    If InStr(LCase$(Command$), "-quit") Then
        Shell "net stop " & Service_Name, vbHide
        KillProcess App.EXEName
    Else
        'Inform the user of running the same app twice
        MsgBox App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " is already running."
    End If
    Exit Sub
End If

intVerbosity = 0
Call ParseCommandLine(Command$)
WriteLog "Inicio v " & intVerbosity & " NT " & IsNTService, Information
If IsNTService Then
    'Check OS
    GetWindowsVersion 0, 0, 0, 0, bIsNT
    If bIsNT Then
    
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
        lngInterval = 100
        Do
            DoEvents
        Loop While MsgWaitObj(lngInterval, hStopPendingEvent, 1&) = WAIT_TIMEOUT
        Set oPOP3 = Nothing
        SetServiceState SERVICE_STOPPED
        WriteLog "Peticion de parada de servicio", Information
        App.LogEvent "Stoping Service: " & Service_Name
        SetEvent hStopEvent
        ' Waiting for service thread termination
        MsgWaitObj INFINITE, hnd, 1&
        CloseHandle hnd
        WriteLog "Servicio parado", Information
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
Else
    WriteLog "Init", Information
    hStopPendingEvent = 0
    frmMain.Init
End If
Exit Sub
ErrHandler:
    WriteLog Err.Source & vbTab & Err.Description, Error
    If IsNTService Then
        CloseHandle hStopEvent
        CloseHandle hStartEvent
        CloseHandle hStopPendingEvent
    Else
        'MsgBox Err.Description, vbCritical, "POP2OWA: " & Err.Source
        MessageBox 0&, Err.Description, App.Title & " " & Err.Source, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
        Unload frmMain
    End If
End Sub



''
'Get the Windows version instaled.
'
'@param lMajor  Major version number
'@param lMinor  Minor version number
'@param lRevision Revision number
'@param lBuildNumber Bulid number
'@param bIsNT True if have NT kernel
Public Sub GetWindowsVersion( _
      Optional ByRef lMajor As Integer = 0, _
      Optional ByRef lMinor As Integer = 0, _
      Optional ByRef lRevision As Integer = 0, _
      Optional ByRef lBuildNumber As Integer = 0, _
      Optional ByRef bIsNT As Boolean = False _
   )
Dim lR As Long
   lR = GetVersion()
   lBuildNumber = (lR And &H7F000000) \ &H1000000
   If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
   lRevision = (lR And &HFF0000) \ &H10000
   lMinor = (lR And &HFF00&) \ &H100
   lMajor = (lR And &HFF)
   bIsNT = ((lR And &H80000000) = 0)
End Sub

''
' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
' Unlike these functions, it
' doesn't block thread messages processing.
'
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.
'
'@param Interval Milliseconds
'@param hObj
'@param nObj
'@return
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


''
'Function to parse de command line pased by the user.
'
'@param cmdline String with the command line
Private Sub ParseCommandLine(ByVal cmdline As String)
    Dim c As String
    c = Trim$(cmdline)

    Dim s As String
    Do Until c = ""
        s = GetNextBlock(c)
        Select Case s
        Case "-v", "/v"
            intVerbosity = GetNextBlock(c)
        Case "-NT", "/NT"
            IsNTService = True
        End Select
    Loop

End Sub

Private Function GetNextBlock(ByRef c As String) As String
    Dim ret As String
    c = LTrim$(c)
    ret = ""
    If Left$(c, 1) = """" Then
        c = Right$(c, Len(c) - 1)
        Do Until Left$(c, 1) = """" Or c = ""
            ret = ret & Left$(c, 1)
            c = Right$(c, Len(c) - 1)
        Loop
        If c <> "" Then c = Right$(c, Len(c) - 1)    '   in case the user didn't enter ending quote
        c = LTrim$(c)

    Else
        Do Until Left$(c, 1) = " " Or c = ""
            ret = ret & Left$(c, 1)
            c = Right$(c, Len(c) - 1)
        Loop
        c = LTrim$(c)
    End If
    GetNextBlock = ret
End Function    '   GetNextBlock


''
'Kill the process that match with the name pased.
'
'@param NameProcess Name of the process to kill
Private Sub KillProcess(NameProcess As String)
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&

Dim uProcess        As PROCESSENTRY32
Dim RProcessFound   As Long
Dim hSnapshot       As Long
Dim SzExename       As String
Dim ExitCode        As Long
Dim MyProcess       As Long
Dim AppKill         As Boolean
Dim AppCount        As Integer
Dim i               As Integer
       
If NameProcess <> "" Then
   AppCount = 0

   uProcess.dwSize = Len(uProcess)
   hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
   RProcessFound = ProcessFirst(hSnapshot, uProcess)
   Do
     i = InStr(1, uProcess.szexeFile, Chr(0))
     SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
     If Left$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
        AppCount = AppCount + 1
        MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(MyProcess, ExitCode)
        Call CloseHandle(MyProcess)
     End If
     RProcessFound = ProcessNext(hSnapshot, uProcess)
   Loop While RProcessFound
   Call CloseHandle(hSnapshot)
End If

End Sub
