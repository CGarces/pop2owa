Attribute VB_Name = "modGlobal"
''
'Store al public variables, constants and common functions.
Option Explicit

Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"
Public Const INFINITE = -1&      '  Timeout infinito
Public Const Service_Name  As String = "POP2OWA"

Public Const OK    As String = "+OK "
Public Const Error As String = "-ERR "

''
'Object to handle POP3/STMP commands.
Public oPOP3            As clsPOP3
''
'Cookies used for authentication
Public strCookies  As String

Public hStopPendingEvent As Long

Public intVerbosity    As Integer
Public Enum Verbosity
    Fail = 0
    Warning = 1
    Information = 2
    Paranoid = 3
End Enum


Public Config As clsConfig

Public IsNTService      As Boolean

''
' API go get the OS version
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

''
'API para provocar un retardo del sistema
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type PROCESSENTRY32
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
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

''
'Write a pop2owa.log file with errors, warnings and log messages.
'
'@param strText Text to write in the log
'@param intVerbose Verbosity
Public Sub WriteLog(ByVal strText As String, Optional ByVal intVerbose As Verbosity = Fail)
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

Dim mutexvalue      As Long

''
'variable constant to match if the mutex exists
Const ERROR_ALREADY_EXISTS = 183&


'Test environment to detect if is running into VB
If bIsEXE Then
    'Create an individual mutex value for the application
    mutexvalue = CreateMutex(ByVal 0&, 1, App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision)
    
    'If an error occurred creating the mutex, that means it
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
End If

intVerbosity = 0
Call ParseCommandLine(Command$)
WriteLog "Inicio v " & intVerbosity & " NT " & IsNTService, Information
If IsNTService Then
    StartService
Else
    frmMain.Init
End If
Exit Sub
ErrHandler:
    WriteLog "Main ->" & Err.Source & vbTab & Err.Description, Fail
    If Not IsNTService Then
        Unload frmMain
    End If
End Sub

''
'Get the Windows version installed.
'
'@param lMajor  Major version number
'@param lMinor  Minor version number
'@param lRevision Revision number
'@param lBuildNumber Build number
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

Const STATUS_TIMEOUT = &H102&
Const QS_KEY = &H1&
Const QS_MOUSEMOVE = &H2&
Const QS_MOUSEBUTTON = &H4&
Const QS_POSTMESSAGE = &H8&
Const QS_TIMER = &H10&
Const QS_PAINT = &H20&
Const QS_SENDMESSAGE = &H40&
Const QS_HOTKEY = &H80&
Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

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

''
'Function to parse Command Line string.
'
'@param c   String with the Command Line
'@return String First parameter in the command line
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
Dim i               As Integer
       
If NameProcess <> "" Then

   uProcess.dwSize = Len(uProcess)
   hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
   RProcessFound = ProcessFirst(hSnapshot, uProcess)
   Do
     i = InStr(1, uProcess.szexeFile, vbNullChar)
     SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
     If Left$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
        MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        Call TerminateProcess(MyProcess, ExitCode)
        Call CloseHandle(MyProcess)
     End If
     RProcessFound = ProcessNext(hSnapshot, uProcess)
   Loop While RProcessFound
   Call CloseHandle(hSnapshot)
End If

End Sub

''
'Evaluate if the process is running under IDE environment.
'
'@return Return False if is under IDE
Private Function bIsEXE() As Boolean
On Error GoTo IDEInUse
    'division by zero error will only run in IDE when using Debug.Print
    Debug.Print 1 \ 0
    bIsEXE = True
Exit Function
IDEInUse:
     'division by zero error got us here so we must be running under the IDE
     bIsEXE = False
End Function
