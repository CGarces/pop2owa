VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPOP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Class to emulate a POP3 server.
' It's a proxy between e-mail client and OWA class.
'
'@author Carlos B
'@version 1.0
'@date 08/10/2005

Option Explicit

Private lngPop3Read As Long
Private lngPop3Write As Long
Private lngSMTPRead As Long
Private lngSMTPWrite As Long

Private objOWA As clsOWA

Private m_strIP As String
Private strSMTPState As String
Private strMailData As String
Private vPorts(0 To 1) As Integer

Property Let IP(ByVal strIP As String)
    m_strIP = strIP
End Property


Property Let Port(ByVal intPort As Integer, ByVal intValue As Integer)
    vPorts(intPort) = intValue
End Property

Property Let Saveinsent(ByVal bSave As Boolean)
    bSaveinsent = bSave
End Property

''
'Creates a new socket.
'
'@param intPort por lo lisen
'@return The number of the socket
Private Function fncNewSokect(ByVal intPort As Integer) As Long
    Dim lngAddressFamily    As Long
    Dim lngRetValue         As Long
    Dim lngSocket           As Long
    Dim lngSocketType       As Long
    Dim lngProtocol         As Long
    '

    '
    'Get values for arguments to pass to the vbSocket function
    lngAddressFamily = CLng(AF_INET)
    lngSocketType = CLng(SOCK_STREAM)
    lngProtocol = CLng(IPPROTO_TCP)
    '
    'Call the vbSocket function in order to create a new socket
    lngSocket = modWinsockAPI.vbSocket(lngAddressFamily, lngSocketType, lngProtocol)
    '
    If lngSocket = INVALID_SOCKET Then
        '
        'If the function has returned the INVALID_SOCKET
        'value the socket was not created.
        Err.Raise Err.LastDllError, GetErrorDescription(Err.LastDllError)
        '
    Else
        'Call the vbBind function to bind the socket to a new address
        lngRetValue = vbBind(lngSocket, m_strIP, intPort)
        
        If lngRetValue = SOCKET_ERROR Then
            '
            'If an error was occurred, raise an error message box
            Err.Raise Err.LastDllError, , GetErrorDescription(Err.LastDllError)
        Else
            'Call the vbListen function
            lngRetValue = vbListen(lngSocket)
            'Check a value returned by the vbListen function
            If lngRetValue = SOCKET_ERROR Then
                'An error was occurred - raise the error message
                Err.Raise Err.LastDllError, GetErrorDescription(Err.LastDllError)
            End If
            '
        End If
    End If
    fncNewSokect = lngSocket
End Function

''
'Close al sockets and call to WSACleanup API
Public Sub Destroy()
    If lngPop3Read <> 0 Then
        Call closesocket(lngPop3Read)
    End If
    If lngPop3Write <> 0 Then
        Call closesocket(lngPop3Write)
    End If
    If lngSMTPRead <> 0 Then
        Call closesocket(lngSMTPRead)
    End If
    If lngSMTPWrite <> 0 Then
        Call closesocket(lngSMTPWrite)
    End If
    Set objOWA = Nothing
    Call WSACleanup
End Sub

''
'Sub to search incomings calls
Public Sub Refresh()

    Dim lngRetValue     As Long
    Dim udtRead_fds     As fd_set
    Dim udtWrite_fds    As fd_set
    Dim udtError_fds    As fd_set
    Dim lngSocketCount  As Long
    Dim lngSocket       As Long
    Dim i               As Integer
        '
        'Prepare sockets arrays to pass to
        'the select Winsock API function
        '
            '
            'Get socket's handle
            '
        'Put that handle into each array
    If lngPop3Read <> 0 Then
        lngSocketCount = lngSocketCount + 1
        udtRead_fds.fd_array(lngSocketCount) = lngPop3Read
        udtWrite_fds.fd_array(lngSocketCount) = lngPop3Read
        udtError_fds.fd_array(lngSocketCount) = lngPop3Read
    End If
    
    If lngPop3Write <> 0 Then
        lngSocketCount = lngSocketCount + 1
        udtRead_fds.fd_array(lngSocketCount) = lngPop3Write
        udtWrite_fds.fd_array(lngSocketCount) = lngPop3Write
        udtError_fds.fd_array(lngSocketCount) = lngPop3Write
    End If
    If lngSMTPRead <> 0 Then
        lngSocketCount = lngSocketCount + 1
        udtRead_fds.fd_array(lngSocketCount) = lngSMTPRead
        udtWrite_fds.fd_array(lngSocketCount) = lngSMTPRead
        udtError_fds.fd_array(lngSocketCount) = lngSMTPRead
    End If
    If lngSMTPWrite <> 0 Then
        lngSocketCount = lngSocketCount + 1
        udtRead_fds.fd_array(lngSocketCount) = lngSMTPWrite
        udtWrite_fds.fd_array(lngSocketCount) = lngSMTPWrite
        udtError_fds.fd_array(lngSocketCount) = lngSMTPWrite
    End If
      
    'Initialize the fd_count member of the
    'fd_set structure for each set of sockets
    udtRead_fds.fd_count = lngSocketCount
    udtWrite_fds.fd_count = lngSocketCount
    udtError_fds.fd_count = lngSocketCount
    '
    'Call the select function
    lngRetValue = vbselect(0&, udtRead_fds, udtWrite_fds, udtError_fds, 0&)
    '
        If lngRetValue = SOCKET_ERROR Then
            '
            'If the function returned value of SOCKET_ERROR
            'just show a message box with error description
            '
            'Call ShowErrorMessage(Err.LastDllError)
            Err.Raise Err.LastDllError, , GetErrorDescription(Err.LastDllError)
            '
        ElseIf lngRetValue > 0 Then
            '
            'If returned value grater than 0 that means that
            'the select function has found some sockets that
            'can be writable, readable or in error state
            '
            'Check for writable sockets
            If udtWrite_fds.fd_count > 0 Then
                '
                For i = 1 To udtWrite_fds.fd_count
                    '
                    'Get the socket handle
                    lngSocket = udtWrite_fds.fd_array(i)
                    'Call recv(lngSocket)
                    '
                    If lngSocket = lngSMTPRead Then
                        Received lngSocket
                    End If
                Next i
                '
            End If
            '
            'Check for readable sockets
            If udtRead_fds.fd_count > 0 Then
                '
                For i = 1 To udtRead_fds.fd_count
                    '
                    'Get the socket handle
                    lngSocket = udtRead_fds.fd_array(i)
                    If lngSocket = lngPop3Read And lngPop3Write = 0 Then
                        lngPop3Write = Accept(lngPop3Read)
                        SendPOP3 Empty
                    End If
                    If lngSocket = lngSMTPRead And lngSMTPWrite = 0 Then
                        lngSMTPWrite = Accept(lngSMTPRead)
                        SendSMTP Empty
                    End If
                    If lngSocket = lngPop3Write Then
                        Received lngSocket
                    End If
                    If lngSocket = lngSMTPWrite Then
                        Received lngSocket
                    End If
                Next i
                '
            End If
            '
            'Check for sockets in error state
'            If udtError_fds.fd_count > 0 Then
'                '
'                For i = 1 To udtError_fds.fd_count
'                    '
'                    'Get the socket handle
'                    lngSocket = udtError_fds.fd_array(i)
'                    '
''                    'Mark the corresponding listview item with red color
''                    lvSockets.ListItems("S" & lngSocket).ForeColor = vbRed
'                    '
'                Next i
'                '
'            End If
            '
        End If
End Sub


''
'Handle received data in a socket.
'
'@param lngSocket The descriptor that identifies a connected socket.
Private Sub Received(ByVal lngSocket As Long)
    '
    Dim strData             As String
    Dim lngBytesReceived    As Long
        '
    'Call the vbRecv function to read data
    lngBytesReceived = vbRecv(lngSocket, strData)
    '
    If lngBytesReceived > 0 Then
        '
        If lngSocket = lngPop3Write Then
            SendPOP3 strData
        Else
            SendSMTP strData
        End If
    ElseIf lngBytesReceived = SOCKET_ERROR Then
        '
        'An error was occurred
        Err.Raise Err.LastDllError, , GetErrorDescription(Err.LastDllError)
    End If
End Sub


''
'Accept a call of email client.
'
'@param lngSocketRead The descriptor that identifies a connected socket.
'
'@return The descriptor that identifies a connected socket, 0 if not is acepted
Private Function Accept(ByVal lngSocketRead As Long) As Long
Dim lngSocket As Long
    '
    'Call the vbAccept function in order to accept the
    'connection request and create a new socket
    lngSocket = vbAccept(lngSocketRead)
    '
    If lngSocket = INVALID_SOCKET Then
        '
        'An error was occurred - raise the error message
        Err.Raise Err.LastDllError, , GetErrorDescription(Err.LastDllError)
        Accept = 0
    Else
        '
        'Accept the connection request
        Accept = lngSocket
    End If  '--> If lngRetValue = INVALID_SOCKET Then
    '
End Function

''
'Evaluate a POP3 command and
'and send a reply to the email client
'
'@param strDataRecived A valid POP3 command ((<A HREF='http://www.hunnysoft.com/rfc/rfc.html' target='_blank'>RFC1939</A>))
Private Sub SendPOP3(ByVal strDataRecived As String)
On Error GoTo GestionErrores
      
    Dim oElement As IXMLDOMElement
    Dim oElements As IXMLDOMNodeList
    Dim strDataToSend As String
    Dim intMsg As Integer
    Dim intTotal As Integer
    Dim bCloseSocket As Boolean
    '
    If Left(strDataRecived, 4) <> "PASS" Then
        Debug.Print strDataRecived
    End If
    Select Case Left(strDataRecived, 4)
    Case Empty
        strDataToSend = "+OK POP3 server ready"
    Case "USER"
        strUser = Mid(strDataRecived, 6, Len(strDataRecived) - 7)
        strDataToSend = "+OK Password required for " & strUser
    Case "PASS"
        strPassWord = Mid(strDataRecived, 6, Len(strDataRecived) - 7)
        If objOWA.GetFoldersURL Then
            strDataToSend = "+OK mailbox for " & strUser & " ready"
        Else
            strDataToSend = "-ERR Unable to conect with mailbox "
        End If
    Case "STAT"
        If objOWA.GetMsgList Then
            strDataToSend = "+OK " & objOWA.intMsgCount & " 0"
        Else
            strDataToSend = "-ERR Unable to conect with mailbox "
        End If
    Case "RETR"
        strDataToSend = getMsg(objDOMInbox.childNodes(1).childNodes(CInt(Mid(strDataRecived, 6)) - 1).selectSingleNode("a:propstat/a:prop/a:href").Text)
        bCloseSocket = (Left(strDataToSend, 4) = "-ERR")
    Case "QUIT"
        strDataToSend = "+OK " & m_strIP & " server signing off, 0 messages deleted"
        bCloseSocket = True
    Case "LIST"
        SendMsg "+OK scan listing follows ", lngPop3Write
        intTotal = objOWA.intMsgCount
        For intMsg = 1 To intTotal
            strDataToSend = strDataToSend & intMsg & " 1 " & vbCrLf
        Next
        SendMsg strDataToSend, lngPop3Write
        strDataToSend = "."
    Case "NOOP"
        strDataToSend = "+OK "
    Case "UIDL"
        Set oElements = objDOMInbox.selectNodes(XMLPATH & "a:id")
        If Len(strDataRecived) > 6 Then
            intMsg = CInt(Mid(strDataRecived, 6))
            strDataToSend = "+OK " & intMsg & " " & oElements(intMsg - 1).Text & vbCrLf
        Else
            strDataToSend = "+OK " & vbCrLf
            For Each oElement In oElements
                intMsg = intMsg + 1
                strDataToSend = strDataToSend & intMsg & " " & oElement.Text & " " & vbCrLf
            Next
            strDataToSend = strDataToSend & "."
        End If
    Case "CAPA"
        SendMsg "+OK Capability list follows", lngPop3Write
        SendMsg "UIDL", lngPop3Write
        strDataToSend = "."
    Case "DELE"
        intMsg = CInt(Mid(strDataRecived, 6))
        If objOWA.Delete(intMsg) Then
            strDataToSend = "+OK message " & intMsg & " deleted"
            'Additional Sleep, My Outlook 2000 fails if delete msg faster
            Sleep 500
        Else
            strDataToSend = "-ERR deleting message " & Mid(strDataRecived, 6)
            bCloseSocket = True
        End If
        
    Case "AUTH"
        strDataToSend = "-ERR " & Left(strDataRecived, 4) & " unsuported feature "
    Case Else
        Debug.Assert False
        strDataToSend = "-ERR " & Left(strDataRecived, 4) & " unsuported feature "
    End Select
    SendMsg strDataToSend, lngPop3Write
    If bCloseSocket Then
        'Call the closesocket Winsock API function
        'in order to close the socket
         Call closesocket(lngPop3Write)
        lngPop3Write = 0
    End If
Exit Sub
GestionErrores:
    Debug.Assert False
    Call closesocket(lngPop3Write)
    lngPop3Write = 0
End Sub

''
'Send a reply to the email client.
'
'@param strDataToSend Strings to send
'@param lngSocket The descriptor that identifies a connected socket
Private Sub SendMsg(ByVal strDataToSend As String, lngSocket)
    'Add CRLF if needed
    If Right(strDataToSend, 2) <> vbCrLf Then
        strDataToSend = strDataToSend & vbCrLf
    End If
    '
    'Call the vbSend function in order to send data
    If vbSend(lngSocket, strDataToSend) = SOCKET_ERROR Then
        '
        'If the vbSend function has returned a value of
        'SOCKET_ERROR, just display a message box with the
        'error description
        Call closesocket(lngSocket)
       
        Debug.Print GetErrorDescription(Err.LastDllError)
'        Debug.Print vbSend(lngPop3Write, "-ERR " & Err.LastDllError)
        'Debug.Assert False
    End If
    Sleep 100
    DoEvents
End Sub


''
'Initialize Winsock code.
'
Public Sub Start()
    'value returned by the InitializeWinsock function
    Dim lngRetValue As Long
    Call WSACleanup
    '
    'innitialize the Winsock service
    lngRetValue = modWinsockAPI.InitializeWinsock(SOCKET_VERSION_22)
    '
    If lngRetValue = 0 Then
        '
        'if the Winsock service was initialized
        'successfully, initialize the controls
        Set objOWA = New clsOWA
        lngPop3Read = fncNewSokect(vPorts(0))
        If vPorts(1) <> 0 Then
            lngSMTPRead = fncNewSokect(vPorts(1))
        End If
    Else
        '
        'if the Winsock service was not initialized
        'successfully, show message box with the
        'error description
        
        '
        'as we cannot initialize the Winsock service
        'we have nothing to do with this application
        Err.Raise lngRetValue, , GetErrorDescription(lngRetValue)
    End If
End Sub

''
'Retrieve a OWA message and convert it into valid POP3 reply
'
'@param strUrl URL of the .EML file
'@return Strig to send to e-mail client (rfc1939).
Private Function getMsg(ByVal strUrl As String) As String
On Error GoTo GestionErrores
    Dim strDataToSend   As String
    Dim strBoundary     As String
    Dim strBody         As String
    Dim bHasAttachment  As Boolean
    Dim strContentType  As String
    Dim oElements       As IXMLDOMNodeList
    Dim oElement        As IXMLDOMElement
    Dim intPos          As Integer

    'strUrl = objDOMInbox.childNodes(1).childNodes(intMsg - 1).selectSingleNode("a:propstat/a:prop/a:href").Text
    objOWA.LoadMsg strUrl
    If objDOMMsg.selectSingleNode(XMLPATH & "e:hasattachment") Is Nothing Then
        bHasAttachment = True
    Else
        bHasAttachment = (objDOMMsg.selectSingleNode(XMLPATH & "e:hasattachment").Text = 1)
'        If Not (bHasAttachment Or objDOMMsg.selectSingleNode(XMLPATH & "a:contentclass") Is Nothing) Then
'            bHasAttachment = (objDOMMsg.selectSingleNode(XMLPATH & "a:contentclass").Text = "urn:content-classes:calendarmessage")
'        End If
    End If
    If Not objDOMMsg.selectSingleNode(XMLPATH & "d:content-type") Is Nothing Then
        strContentType = "Content-Type: " & objDOMMsg.selectSingleNode(XMLPATH & "d:content-type").Text
    Else
        'In some case the Content-Type not is correct.
        'MSExchange sent text/plain instead text/html.
        'Force to read the HTML content to work OK in all cases
        strContentType = "Content-Type: multipart/related;boundary=""--------------" & Right(String(23, "0") & 1, 23) & """"
    End If
    If InStr(strContentType, "html") > 0 And Not (bHasAttachment) Then
        'I have a problem with HTML without attachment
        'try to emulate a multi part msg.
        strContentType = "Content-Type: multipart/related;boundary=""--------------" & Right(String(23, "0") & 1, 23) & """"
    End If
    SendMsg "+OK", lngPop3Write
    'Header
    'Search the boundary strings
    intPos = InStr(LCase(strContentType), "boundary")
    If intPos > 0 Then
        'It's a multi part msg.
        strBoundary = Mid(strContentType, intPos + Len("boundary") + 2)
        strBoundary = Left(strBoundary, Len(strBoundary) - 1)
        strDataToSend = objOWA.getMsgHeader() & _
                       strContentType & vbCrLf & vbCrLf
        SendMsg strDataToSend, lngPop3Write
        'Msg Body (HTML data)
        If Not objDOMMsg.selectSingleNode(XMLPATH & "e:htmldescription") Is Nothing Then
            strBody = objDOMMsg.selectSingleNode(XMLPATH & "e:htmldescription").Text
            strBody = Replace(strBody, "onload=""initTemplate()""", Empty)
        End If
        
        strDataToSend = vbCrLf & "This is a multi-part message in MIME format." & vbCrLf & _
                    "--" & strBoundary & vbCrLf & _
                    "Content-Type: text/html; charset=ISO-8859-1" & vbCrLf & _
                    "Content-Transfer-Encoding: 8bit " & vbCrLf & _
                    "Content-Disposition: inline" & vbCrLf & _
                    vbCrLf & _
                    strBody & vbCrLf & _
                    "" & vbCrLf
        
        
        SendMsg strDataToSend, lngPop3Write
        If bHasAttachment Then
            SendAttachments strUrl, strBoundary
        End If
        'End of multipart msg
        getMsg = "--" & strBoundary & "--" & vbCrLf & "."
    Else
        'NOT It's a multi-part msg.
        If Not Right(strContentType, 1) = ";" Then strContentType = strContentType & ";"
        strDataToSend = objOWA.getMsgHeader() & _
                        strContentType & vbCrLf & vbCrLf
        SendMsg strDataToSend, lngPop3Write
        If InStr(LCase(strContentType), "html") > 0 Then
            'HTML
            If Not objDOMMsg.selectSingleNode(XMLPATH & "e:htmldescription") Is Nothing Then
                strBody = objDOMMsg.selectSingleNode(XMLPATH & "e:htmldescription").Text
                strBody = Replace(strBody, "onload=""initTemplate()""", Empty)
            End If
        Else
            'Plain text without attachments
            If Not objDOMMsg.selectSingleNode(XMLPATH & "e:textdescription") Is Nothing Then
                strBody = objDOMMsg.selectSingleNode(XMLPATH & "e:textdescription").Text
            End If
        End If
        getMsg = vbCrLf & strBody & vbCrLf & "" & vbCrLf & "."
    End If
    'End of msg
Exit Function
GestionErrores:
    Debug.Assert False
    getMsg = "-ERR Unable to conect with mailbox "
End Function


''
'Evaluate a STMP command and
'and send a reply to the email client
'
'@param strDataRecived A valid SMTP command (<A HREF='http://www.faqs.org/rfcs/rfc822.html' target='_blank'>RFC822</A>)
Private Sub SendSMTP(ByVal strDataRecived As String)
On Error GoTo GestionErrores
    Dim strSubURL As String
    Dim strAlias As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strBody As String
    Dim bResult As Boolean
      
    Dim oElement As IXMLDOMElement
    Dim oElements As IXMLDOMNodeList
    Dim strDataToSend As String
    Dim intMsg As Integer
    Dim intTotal As Integer
    Dim bCloseSocket As Boolean
    '
    If Left(strDataRecived, 4) <> "PASS" And strSMTPState <> "DATA" Then
        Debug.Print strDataRecived
    End If
    If strSMTPState = "LOGIN" Or strSMTPState = "PASSWORD" Then
        strDataRecived = "AUTH " & strDataRecived
    End If
    Select Case Left(strDataRecived, 4)
    Case Empty
        strDataToSend = "220 " & m_strIP & " Simple Mail Transfer Service Ready"
    Case "HELO"
        strDataToSend = "250 " & m_strIP
    Case "EHLO"
'        strDataToSend = strDataToSend & vbNewLine & "250-" & m_strIP
'        strDataToSend = strDataToSend & vbNewLine & "250-8BITMIME"
'        If strPassWord = Empty Or strUser = Empty Then
'            strDataToSend = strDataToSend & vbNewLine & "250-AUTH=LOGIN PLAIN"
'            strDataToSend = strDataToSend & vbNewLine & "250-AUTH LOGIN PLAIN"
'        End If
'        strDataToSend = strDataToSend & vbNewLine & "250 Size 51200000"
        strDataToSend = "500 Syntax error, command unrecognized"
    Case "AUTH"
        Dim objBase As New Base64
        Dim strReply As String
        If strDataRecived = "AUTH LOGIN" & vbCrLf Then
            'cDataToSend.Add "334 " & objBase.Encode(strUser & "@" & Mid(strExchSvrName, InStr(strExchSvrName, "//") + 2))
            strDataToSend = "334 " & objBase.Encode("Username:")
            strSMTPState = "LOGIN"
        ElseIf Left(strDataRecived, 10) = "AUTH PLAIN" Then
            strReply = objBase.Decode(Mid(strDataRecived, 12))
            If Len(strReply) > 1 Then
                strUser = Mid(strReply, 2, InStr(2, strReply, Chr(0)) - 2)
                strPassWord = Mid(strReply, InStr(2, strReply, Chr(0)) + 1)
                strDataToSend = "235 AUTHENTICATION SUCCESSFUL"
            Else
                strDataToSend = "533  Authentication required [AUTH]"
            End If
        Else
            If strSMTPState = "LOGIN" Then
                strUser = objBase.Decode(Mid(strDataRecived, 6))
                strSMTPState = "PASSWORD"
                strDataToSend = "334 " & objBase.Encode("Password:")
            ElseIf strSMTPState = "PASSWORD" Then
                strPassWord = objBase.Decode(Mid(strDataRecived, 6))
                strDataToSend = "235 AUTHENTICATION SUCCESSFUL"
                strSMTPState = ""
            Else
                strDataToSend = "556 CLIENT AUTHENTICATION REQUIRED. USE ESMTP EHLO AND AUTH"
            End If
        End If
    Case "MAIL"
        If strPassWord = Empty Or strUser = Empty Then
            strDataToSend = "556 CLIENT AUTHENTICATION REQUIRED. USE ESMTP EHLO AND AUTH"
            bCloseSocket = True
        Else
            strMailData = strDataRecived
            strDataToSend = "250 OK"
        End If
    Case "RCPT"
        strMailData = strMailData & strDataRecived
        strDataToSend = "250 OK"
    Case "RSET"
        strDataToSend = "250 OK"
    Case "DATA"
        strDataToSend = "354 Start mail input; end with <CRLF>.<CRLF>"
        strSMTPState = "DATA"
        strMailData = strMailData & vbCrLf
    Case "QUIT"
        strDataToSend = "221 " & m_strIP & " Service closing transmission channel"
        bCloseSocket = True
    Case Else
        If strSMTPState = "DATA" Then
            strMailData = strMailData & strDataRecived
            If Right(strDataRecived, 3) = "." & vbCrLf Then
                
                'Encoding the subject in QP format
                'strDataRecived = EncodeSubject(strDataRecived)
                If objOWA.SendMail(strMailData) Then
                     strDataToSend = "250 OK"
                 Else
                     strDataToSend = "500 Syntax error, command unrecognized"
                End If
                strSMTPState = Empty
                strMailData = Empty
            End If
        Else
            Debug.Assert False
        End If
    End Select
    If strDataToSend <> Empty Then
        SendMsg strDataToSend, lngSMTPWrite
    End If
    If bCloseSocket Then
        'Call the closesocket Winsock API function
        'in order to close the socket
        Call closesocket(lngSMTPWrite)
        lngSMTPWrite = 0
    End If
Exit Sub
GestionErrores:
    Debug.Assert False
    Call closesocket(lngSMTPWrite)
    lngPop3Write = 0
End Sub

Private Function EncodeSubject(ByVal strData As String) As String
Dim intPosIni As Integer
Dim intPosFin As Integer
Dim strSubject As String
Dim objQP As Base64
Set objQP = New Base64
intPosIni = InStr(strData, "Subject:") + Len("Subject:")
intPosFin = InStr(intPosIni, strData, vbNewLine)
strSubject = Mid(strData, intPosIni, intPosFin - intPosIni)
If Left(strSubject, 16) <> " =?ISO-8859-1?Q?" Then
    'Subject: =?ISO-8859-1?Q?Le=F3n=2C_Cami=F3n_y_Espa=F1a?=
    EncodeSubject = Left(strData, intPosIni) & objQP.QP(strSubject) & Mid(strData, intPosFin)
Else
    EncodeSubject = strData
End If
Set objQP = Nothing
End Function

Private Sub SendAttachments(ByVal strUrl As String, ByVal strBoundary As String)
On Error GoTo GestionErrores

Dim strDataToSend   As String
Dim oElements       As IXMLDOMNodeList
Dim oElement        As IXMLDOMElement
Dim strContentType          As String
Dim strContentDisposition   As String
Dim strFilename As String

'Send mail Attachments
'Attachment
Set oElements = objOWA.GetAttachmentsList(strUrl)
For Each oElement In oElements
    strDataToSend = vbCrLf & "--" & strBoundary & vbCrLf & _
                    "Content-Transfer-Encoding: base64 " & vbCrLf

    strUrl = oElement.selectSingleNode("a:href").Text
    strContentType = "application/octet-stream"
    If Not oElement.selectSingleNode("a:propstat/a:prop/d:x370e001f") Is Nothing Then
        strContentType = oElement.selectSingleNode("a:propstat/a:prop/d:x370e001f").Text
    End If
    
    'PR_ATTACH_METHOD
    If oElement.selectSingleNode("a:propstat/a:prop/d:x37050003").Text = 5 Then
        'ATTACH_EMBEDDED_MSG
        strFilename = oElement.selectSingleNode("a:propstat/a:prop/f:cn").Text
    Else
        'x3704001f  PR_ATTACH_FILENAME
        'x3707001f  PR_ATTACH_LONG_FILENAME
        If oElement.selectSingleNode("a:propstat/a:prop/e:attachmentfilename") Is Nothing Then
            Debug.Assert False
            strFilename = oElement.selectSingleNode("a:propstat/a:prop/d:x3704001f").Text
        Else
            strFilename = oElement.selectSingleNode("a:propstat/a:prop/e:attachmentfilename").Text
        End If
        
    End If
    
    'x3716001f  PR_ATTACH_CONTENT_DISPOSITION
    If oElement.selectSingleNode("a:propstat/a:prop/d:x3716001f") Is Nothing Then
        strContentDisposition = "attachment"
    Else
        strContentDisposition = oElement.selectSingleNode("a:propstat/a:prop/d:x3716001f").Text
    End If
    
    'PR_ATTACH_CONTENT_ID
    If oElement.selectSingleNode("a:propstat/a:prop/d:x3712001f") Is Nothing Then
        strDataToSend = strDataToSend & "Content-Type: " & strContentType & "; name=""" & strFilename & """" & vbCrLf & _
                    "Content-Disposition: " & strContentDisposition & "; fileName = """ & strFilename & """" & vbCrLf & vbCrLf & vbCrLf
    Else
        'If has PR_ATTACH_CONTENT_ID
        strDataToSend = strDataToSend & "Content-Type: IMAGE/JPEG; name=""" & strFilename & """" & vbCrLf & _
                    "Content-Disposition: " & strContentDisposition & "; fileName = """ & strFilename & """" & vbCrLf & _
                    "Content-ID: <" & oElement.selectSingleNode("a:propstat/a:prop/d:x3712001f").Text & ">" & vbCrLf & vbCrLf & vbCrLf
        'oElement.selectSingleNode("a:propstat/a:prop/d:x3716001f").Text
    End If
    'Send attachment header
    SendMsg strDataToSend, lngPop3Write
    'Send attachment data
    SendMsg objOWA.GetAttachment(strUrl) & vbCrLf & vbCrLf, lngPop3Write
Next
Exit Sub
GestionErrores:
    Debug.Assert False
    Err.Raise Err.Number
End Sub
