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
' Class to emulate a PO3 sever.
' It's a proxy between e-mail clien and OWA sever.
'
'@author Carlos B
'@version 1.0
'@date 08/10/2005
Option Explicit
#Const DEBUG_MODE = 0

Private lngPop3Read As Long
Private lngPop3Write As Long
Private lngSMTPRead As Long
Private lngSMTPWrite As Long
Private bSaveinsent As Boolean

Private strUser As String
Private strPassWord As String

Private objDOMFolders   As DOMDocument
Private objDOMInbox     As DOMDocument
Private objDOMMsg       As DOMDocument
''
'Translations between XML tags of OWA and valid headers of a POP3 message
'@see initArrays
Private dicTags   As Dictionary
Private Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"
Private m_strIP As String
Private strExchSvrName  As String
Private strSMTPState As String
Private strMailData As String
Private vPorts(0 To 1) As Integer

Property Let IP(ByVal strIP As String)
    m_strIP = strIP
End Property

Property Let ServerName(ByVal strServerName As String)
    strExchSvrName = strServerName
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
'Create the Socket to senddata to client
'
'@return Number of the Socket, 0 if not is acepted
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
'Evaluate the data received
'and send a reply to the email client
'
'@param strDataRecived A valid POP3 command (rfc1939)
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
        If GetFoldersURL Then
            strDataToSend = "+OK mailbox for " & strUser & " ready"
        Else
            strDataToSend = "-ERR Unable to conect with mailbox "
        End If
    Case "STAT"
        If GetMsgList Then
            strDataToSend = "+OK " & objDOMInbox.childNodes(1).childNodes.Length & " 0"
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
        intTotal = objDOMInbox.childNodes(1).childNodes.Length
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
            Set oElements = objDOMInbox.selectNodes(XMLPATH & "a:id")
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
        Set oElement = objDOMInbox.childNodes(1).childNodes(intMsg - 1)
        If Delete(oElement.selectSingleNode("a:propstat/a:prop/a:href").Text) Then
            strDataToSend = "+OK message " & intMsg & " deleted"
            'Additional Sleep, My Outlook 2000 fails if delete msg faster
            Sleep 500
        Else
            strDataToSend = "-ERR deleting message " & Mid(strDataRecived, 6)
            bCloseSocket = True
        End If
        
    Case "AUTH", "CAPA"
        strDataToSend = "-ERR " & Left(strDataRecived, 4) & " unsuported feature "
        bCloseSocket = True
    Case Else
        Debug.Assert False
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
        initArrays
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
'Get the headers of a message.
'
'@return Header of the mail
Private Function getMsgHeader() As String
Dim strDataToSend   As String
Dim vKeys           As Variant
Dim Key             As Variant
Dim strValue        As String
Dim intPos1         As Integer
Dim intPos2         As Integer
Dim strEmail        As String
Dim strName         As String
Dim strReturn       As String

vKeys = dicTags.Keys
For Each Key In vKeys
    If Not objDOMMsg.selectSingleNode(XMLPATH & dicTags.Item(Key)) Is Nothing Then
        strValue = objDOMMsg.selectSingleNode(XMLPATH & dicTags.Item(Key)).Text
        Select Case Key
        Case "Importance"
            Select Case strValue
            Case 1
                strDataToSend = strDataToSend & "Importance: Normal" & vbCrLf
                strDataToSend = strDataToSend & "X-Priority: 3 (Normal)" & vbCrLf
            Case 2
                strDataToSend = strDataToSend & "Importance: High" & vbCrLf
                strDataToSend = strDataToSend & "X-Priority:  1 (Highest)" & vbCrLf
            Case Else
                strDataToSend = strDataToSend & "Importance: Low" & vbCrLf
                strDataToSend = strDataToSend & "X-Priority:  5 (Low)" & vbCrLf
            End Select
        Case "Date"
            strValue = GetPOPDate(strValue)
            strDataToSend = strDataToSend & Key & ": " & strValue & vbCrLf
        Case "To", "Cc", "Bcc", "From"
            strReturn = Empty
            Do
                strValue = strValue & "  "
                intPos1 = InStr(strValue, ">")
                If intPos1 > 0 Then
                    intPos2 = InStr(strValue, "<")
                    strName = Trim(Mid(strValue, 1, intPos2 - 1))
                    strEmail = Mid(strValue, intPos2, intPos1 - intPos2 + 1)
                    If Left(strName, 1) <> """" Then
                        strName = """" & strName & """"
                    End If
                    strReturn = strReturn & strName & " " & strEmail & ", "
                    strValue = Mid(strValue, intPos1 + 2)
                End If
            Loop While intPos1 <> 0
            strDataToSend = strDataToSend & Key & ": " & Mid(strReturn, 1, Len(strReturn) - 2) & vbCrLf
        Case Else
            strDataToSend = strDataToSend & Key & ": " & strValue & vbCrLf
        End Select
    ElseIf Key = "Content-Transfer-Encoding" Then
        strDataToSend = strDataToSend & "Content-Transfer-Encoding: 8bit " & vbCrLf
    End If
Next
getMsgHeader = strDataToSend
End Function

''
'Init a public dictionary with the trlanslation of XML tags of OWA
' and valid headers of a POP3 message
'@see dicTags
Private Sub initArrays()
    Set dicTags = New Dictionary
    dicTags.Add "Date", "e:date"
    dicTags.Add "From", "e:from"
    dicTags.Add "To", "e:to"
    dicTags.Add "Cc", "e:cc"
    dicTags.Add "Bcc", "e:bcc"
    dicTags.Add "Subject", "e:subject"
    dicTags.Add "Importance", "e:importance"
    'dicTags.Add "Received", "d:received"
    'dicTags.Add "Content-Type", "d:content-type"
    dicTags.Add "Content-Transfer-Encoding", "d:content-transfer-encoding"
    dicTags.Add "MIME-Version", "d:mime-version"
    dicTags.Add "Message-ID", "d:message-id"
    
    dicTags.Add "X-Mailer", "d:x-mailer"
End Sub


''
'Get the date of a message, converting OWA format in a standar date (rfc1939).
'
'@param strDate Date in OWA format
'@return Standar date (rfc1939)
Private Function GetPOPDate(ByVal strDate As String) As String
Dim vDate As Variant
Dim vDays(1 To 7) As String
Dim vMonths(1 To 12) As String
Dim strReturn As String
vDays(1) = "Mon"
vDays(2) = "Tue"
vDays(3) = "Wed"
vDays(4) = "Thu"
vDays(5) = "Fri"
vDays(6) = "Sat"
vDays(7) = "Sun"

vMonths(1) = "Jan"
vMonths(2) = "Feb"
vMonths(3) = "Mar"
vMonths(4) = "Apr"
vMonths(5) = "May"
vMonths(6) = "Jun"
vMonths(7) = "Jul"
vMonths(8) = "Aug"
vMonths(9) = "Sep"
vMonths(10) = "Oct"
vMonths(11) = "Nov"
vMonths(12) = "Dec"
    vDate = Split(strDate, "T")
    strReturn = vDays(DatePart("w", vDate(0)))
    strReturn = strReturn & ", " & DatePart("d", vDate(0))
    strReturn = strReturn & " " & vMonths(DatePart("m", vDate(0)))
    strReturn = strReturn & " " & DatePart("yyyy", vDate(0))
    GetPOPDate = strReturn & " " & FormatDateTime(Left(vDate(1), 8), vbLongTime) & " UT"
End Function

''
'Load an OWA mesagge into a XML object.
'
'@param strUrl Url of the .EML file
'@see objDOMMsg
Private Sub LoadMsg(ByVal strUrl As String)
Dim oXMLHTTP As XMLHTTP
' Create an instance of the xmlHTTP object
Set oXMLHTTP = CreateObject("microsoft.xmlhttp")
With oXMLHTTP
    ' Open a request to a particular url
    .open "PROPFIND", strUrl, False, strUser, strPassWord
    ' Set the header type
    .setRequestHeader "Content-type:", "text/xml"
    .setRequestHeader "Depth", "1"
    ' Send the request
    .Send
    While .readyState <> 4
        Sleep 10
        DoEvents
        'Debug.Assert oXMLHTTP.readyState <> 3
    Wend
    If (.Status >= 200 And .Status < 300) Then
        Set objDOMMsg = .responseXML
        #If DEBUG_MODE = 1 Then
            objDOMInbox.save "C:\mensaje.xml"
        #End If
    Else
        Err.Raise 51
    End If
End With
Set oXMLHTTP = Nothing
End Sub

''
'Get all files asociated to a message.
'
'@param strUrl  Url of the .EML file
'@return
Private Function GetAttachmentsList(ByVal strUrl As String) As IXMLDOMNodeList
Dim oXMLHTTP As XMLHTTP

Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    .open "X-MS-ENUMATTS", strUrl, False, strUser, strPassWord
    .setRequestHeader "Depth", "1"
    .setRequestHeader "Content-type", "xml"
    .Send
    If .Status = 207 Then
        #If DEBUG_MODE = 1 Then
            objDOMInbox.save "c:\ATTS2.xml"
        #End If
        'Set GetAttachmentsList = oXMLHTTP.responseXML.selectNodes("a:multistatus/a:response/a:href")
        Set GetAttachmentsList = .responseXML.selectNodes("a:multistatus/a:response")
    End If
End With
Set oXMLHTTP = Nothing
End Function
''
'Retrieve the attachment of a message.
'
'@param strUrl  Url of the .EML file
'@return Binary content encoded in Base64
Private Function GetAttachment(ByVal strUrl As String) As String
Dim oXMLHTTP    As XMLHTTP
Dim objMIME     As Base64
Set objMIME = New Base64
Set oXMLHTTP = New XMLHTTP
oXMLHTTP.open "GET", strUrl, False, strUser, strPassWord
oXMLHTTP.Send
GetAttachment = objMIME.EncodeArr(oXMLHTTP.responseBody)
Set objMIME = Nothing
Set oXMLHTTP = Nothing
End Function

''
'Delete a message
'
'@param strUrl  Url of the .EML file
Private Function Delete(ByVal strUrl As String) As Boolean
On Error GoTo GestionErrores
Dim oXMLHTTP    As XMLHTTP
Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    .open "DELETE", strUrl, False, strUser, strPassWord
    .setRequestHeader "Depth", "infinity"
    .Send
'    Debug.Print .Status
End With
Set oXMLHTTP = Nothing
Delete = True
Exit Function
GestionErrores:
Delete = False
End Function


Private Function GetAttachmentName(ByVal strUrl As String) As String
Dim oXMLHTTP As XMLHTTP
Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    .open "PROPFIND", strUrl, False, strUser, strPassWord
    .Send
    'GetAttachment = oXMLHTTP.responseText
    If .Status = 207 Then
        #If DEBUG_MODE = 1 Then
            objDOMInbox.save "c:\ATT.xml"
        #End If
        GetAttachmentName = .responseXML.selectSingleNode("a:multistatus/a:response/a:propstat/a:prop/a:displayname").Text
    End If
End With
Set oXMLHTTP = Nothing
End Function

''
'Retrieve a OWA message and convert it into valid POP3 reply
'
'@param intMsg Nuember of the mesagge to get
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
    LoadMsg strUrl
    If objDOMMsg.selectSingleNode(XMLPATH & "e:hasattachment") Is Nothing Then
        bHasAttachment = True
    Else
        bHasAttachment = (objDOMMsg.selectSingleNode(XMLPATH & "e:hasattachment").Text = 1)
    End If
    If Not objDOMMsg.selectSingleNode(XMLPATH & "d:content-type") Is Nothing Then
        strContentType = "Content-Type: " & objDOMMsg.selectSingleNode(XMLPATH & "d:content-type").Text
    Else
        If bHasAttachment Then
            'Has Attachment, it's a multi part msg.
            strContentType = "Content-Type: multipart/related;boundary=""--------------" & Right(String(23, "0") & 1, 23) & """"
        Else
            'Ony have plain text
            strContentType = "Content-Type: text/plain;charset=""iso-8859-1"";"
        End If
    End If
    'Patch to get HTML msg without attachments
    If InStr(strContentType, "html") > 0 And Not (bHasAttachment) Then
        strContentType = "Content-Type: multipart/related;boundary=""--------------" & Right(String(23, "0") & 1, 23) & """"
    End If
    SendMsg "+OK", lngPop3Write
    'Header
    'Search the boundary strings
    intPos = InStr(strContentType, "boundary")
    If intPos > 0 Then
        'It's a multi part msg.
        strBoundary = Mid(strContentType, intPos + Len("boundary") + 2)
        strBoundary = Left(strBoundary, Len(strBoundary) - 1)
        strDataToSend = getMsgHeader() & _
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
        strDataToSend = getMsgHeader() & _
                        strContentType & vbCrLf & vbCrLf
        SendMsg strDataToSend, lngPop3Write
        If InStr(strContentType, "html") > 0 Then
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
'Get the list of al messages in the inbox folder.
'
'@return Reply to e-mail client
Private Function GetMsgList() As Boolean
On Error GoTo GestionErrores

    Dim strUrl As String
    Dim oXMLHTTP As XMLHTTP
    strUrl = objDOMFolders.selectSingleNode(XMLPATH & "d:inbox").Text & "/"
    ' Create an instance of the xmlHTTP object
    Set oXMLHTTP = CreateObject("microsoft.xmlhttp")
    With oXMLHTTP
        ' Open a request to a particular url
        .open "SEARCH", strUrl, False, strUser, strPassWord
        ' Set the header type - we want XML
        .setRequestHeader "Content-type:", "text/xml"
        .setRequestHeader "Depth", "1"
        ' Send the request
        Sleep 100
        Call .Send("<?xml version='1.0' ?>" & _
                 "<a:searchrequest xmlns:a='DAV:'><a:sql>" & _
                  "SELECT" & _
                  " ""DAV:href""" & _
                  ",""urn:schemas:httpmail:subject""" & _
                  ",""DAV:id""" & _
                  " FROM scope('shallow traversal of """ & strUrl & """')" & _
                 " WHERE ""DAV:ishidden""=False" & _
                 " AND ""DAV:isfolder""=False" & _
                "</a:sql></a:searchrequest>")
        While .readyState <> 4
            Sleep 10
            DoEvents
            'Debug.Assert oXMLHTTP.readyState <> 3
        Wend
        Set objDOMInbox = .responseXML
        #If DEBUG_MODE = 1 Then
            objDOMInbox.save "C:\OWA2.XML"
        #End If
    End With
    GetMsgList = True
Exit Function
GestionErrores:
    Debug.Assert False
    Debug.Print Err.Description
    GetMsgList = False
End Function
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
'        cDataToSend.Add "250-" & m_strIP
'        cDataToSend.Add "250-8BITMIME"
'        If strPassWord = Empty Or strUser = Empty Then
'            cDataToSend.Add "250-AUTH=LOGIN PLAIN"
'            cDataToSend.Add "250-AUTH LOGIN PLAIN"
'        End If
'        cDataToSend.Add "250 Size 51200000"
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
                If SendMail() Then
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


Private Function GetFoldersURL() As Boolean

On Error GoTo ErrHandler
Dim query As String
Dim strUrl As String
Dim oXMLHTTP    As XMLHTTP

query = "<?xml version='1.0'?>"
query = query + "<a:propfind xmlns:a='DAV:'>"
query = query + "<a:prop xmlns:m='urn:schemas:httpmail:'>"
query = query + "<m:sendmsg/>"
query = query + "<m:inbox/>"
query = query + "</a:prop>"
query = query + "</a:propfind>"

Set oXMLHTTP = New XMLHTTP
strUrl = strExchSvrName & "/exchange/" & strUser & "/"
With oXMLHTTP
    .open "PROPFIND", strUrl, False, strUser, strPassWord
    .setRequestHeader "Content-Type", "text/xml"
    .setRequestHeader "Depth", "0"
    
    .Send query
    
    ' process the result
    If (.Status >= 200 And .Status < 300) Then
       Set objDOMFolders = .responseXML
       GetFoldersURL = True
    Else
       GetFoldersURL = False
    End If
End With
Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    GetFoldersURL = False
End Function

Function SendMail() As Boolean

Dim strText
Dim strSubURL
strSubURL = objDOMFolders.selectSingleNode(XMLPATH & "d:sendmsg").Text
'         To use MSXML 4.0 use the followinf DIM/SET statements
         Dim xmlReq As XMLHTTP
         Set xmlReq = New XMLHTTP

         On Error GoTo ErrHandler

          'Create the DAV PUT request.

         xmlReq.open "PUT", strSubURL, False, strUser, strPassWord
         If strMailData <> "" Then
            xmlReq.setRequestHeader "Content-Type", "message/rfc821"
            If Not bSaveinsent Then
                xmlReq.setRequestHeader "Saveinsent", "f"
            End If
            xmlReq.Send strMailData
         
         End If

         'Process the results.
         If (xmlReq.Status >= 200 And xmlReq.Status < 300) Then
            SendMail = True
         Else
            SendMail = False
         End If
ErrExit:
      Set xmlReq = Nothing
      Exit Function
ErrHandler:
    Debug.Print Err.Description
      SendMail = False
End Function

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
Set oElements = GetAttachmentsList(strUrl)
For Each oElement In oElements
    strDataToSend = vbCrLf & "--" & strBoundary & vbCrLf & _
                    "Content-Transfer-Encoding: base64 " & vbCrLf

    strUrl = oElement.selectSingleNode("a:href").Text
    strContentType = oElement.selectSingleNode("a:propstat/a:prop/d:x370e001f").Text
    
    'PR_ATTACH_METHOD
    If oElement.selectSingleNode("a:propstat/a:prop/d:x37050003").Text = 5 Then
        'ATTACH_EMBEDDED_MSG
        strFilename = oElement.selectSingleNode("a:propstat/a:prop/f:cn").Text
    Else
        strFilename = oElement.selectSingleNode("a:propstat/a:prop/e:attachmentfilename").Text
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
    SendMsg GetAttachment(strUrl) & vbCrLf & vbCrLf, lngPop3Write
Next
Exit Sub
GestionErrores:
    Debug.Assert False
    Err.Raise Err.Number
End Sub
