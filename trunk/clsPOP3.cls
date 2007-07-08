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

''
'STMP Socket object
Dim WithEvents skSMTP As CSocketMaster
Attribute skSMTP.VB_VarHelpID = -1
''
'POP3 Socket object
Dim WithEvents skPOP3 As CSocketMaster
Attribute skPOP3.VB_VarHelpID = -1

''
'Handle WEBDAW command with OWA server
Private objOWA As clsOWA

Private m_strIP As String
Private strSMTPState As String
Private strMailData As String
Private vMailData() As String

Private vPorts(0 To 1) As Long


''
'Close al sockets and call to WSACleanup API
Public Sub Disconnect()
If Not (skPOP3 Is Nothing) Then
    skPOP3.CloseSck
End If
If Not (skSMTP Is Nothing) Then
    skSMTP.CloseSck
End If

End Sub




''
'Initialize Winsock code.
'
Public Sub Start()
On Error GoTo ErrHandler

Dim lngRetValue As Long
Dim objXML As DOMDocument
Dim c As cRegistry
Set c = New cRegistry

'Read values form registry
    
'Values are stored in (HKEY_CURRENT_USER\Software\pop2owa)
With c
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\pop2owa"
    .ValueType = REG_SZ
    'Exchange server
    .ValueKey = "ExchangeServer"
    strExchSvrName = .Value
    If Right$(strExchSvrName, 1) = "/" Then
        strExchSvrName = Left$(strExchSvrName, Len(strExchSvrName) - 1)
    End If
    'IP to listen
    .ValueKey = "IP"
    m_strIP = .Value
    
    'Port values are integers
    .ValueType = REG_DWORD
    'POP3 Port
    .ValueKey = "POP3"
    vPorts(0) = .Value
    'SMTP Port
    .ValueKey = "SMTP"
    vPorts(1) = .Value
    'Leave a copy in send folder
    .ValueKey = "Saveinsent"
    bSaveinsent = .Value
    'Form-Based-Authentication on/off
    .ValueKey = "FormBasedAuth"
    bAuthentication = .Value
End With

Set skPOP3 = New CSocketMaster
skPOP3.Bind vPorts(0), m_strIP
skPOP3.Listen
If vPorts(1) <> 0 Then
    Set skSMTP = New CSocketMaster
    skSMTP.Bind vPorts(1), m_strIP
    skSMTP.Listen
End If

Set c = Nothing

Exit Sub
ErrHandler:
    Me.Disconnect
    Err.Raise Err.Number, "Start " & Err.Source, Err.Description
End Sub

''
'Validate the current user & password, using GetFoldersURL
'
'@return SMTP comand with the result of the operation
'
'@see clsOWA.GetFoldersURL
'@date 2006/03/31
Private Function ValidateSMTPAUTH() As String
    Select Case objOWA.GetFoldersURL
    Case 200 To 299
        ValidateSMTPAUTH = "235 AUTHENTICATION SUCCESSFUL"
    Case 401
        'Wrong password
        ValidateSMTPAUTH = "535 5.7.0 authentication failed"
    Case Else
        'Unknow error
        ValidateSMTPAUTH = "533  Authentication required [AUTH]"
    End Select
End Function

Private Sub Class_Initialize()
    Me.Disconnect
    Me.Start
End Sub

Private Sub Class_Terminate()
    Me.Disconnect
End Sub

Private Sub skPOP3_ConnectionRequest(ByVal requestID As Long)
    skPOP3.CloseSck
    skPOP3.Accept requestID
    skPOP3.SendData "+OK POP3 server ready"
End Sub
''
'Evaluate a POP3 command  (<A HREF='http://www.faqs.org/rfcs/rfc1939.html' target='_blank'>RFC1939</A>)
'and send a reply to the email client
'
'@param bytesTotal bytes recived
Private Sub skPOP3_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
      
    Dim oElement        As IXMLDOMElement
    Dim oElements       As IXMLDOMNodeList
    Dim strDataToSend   As String
    Dim intMsg          As Integer
    Dim intTotal        As Integer
    Dim bCloseSocket    As Boolean
    Dim lngSize         As Long
    Dim strDataRecived  As String
    Dim bFBALoginFail   As Boolean
    
    skPOP3.GetData strDataRecived
    '
    If Left$(strDataRecived, 4) <> "PASS" Then
        Debug.Print strDataRecived
    End If
    Select Case Left$(strDataRecived, 4)
    Case Empty
        strDataToSend = "+OK POP3 server ready"
    Case "USER"
        strUser = Mid$(strDataRecived, 6, Len(strDataRecived) - 7)
        strDataToSend = "+OK Password required for " & strUser
    Case "PASS"
        If objOWA Is Nothing Then
            Set objOWA = New clsOWA
        End If
        strPassWord = Mid$(strDataRecived, 6, Len(strDataRecived) - 7)
        Select Case objOWA.GetFoldersURL
        Case 200 To 299
            strDataToSend = "+OK mailbox for " & strUser & " ready"
        Case 401
            strDataToSend = "-ERR Invalid Password for " & strUser
            bCloseSocket = True
        Case Else
            strDataToSend = "-ERR Unable to conect with mailbox "
            bCloseSocket = True
        End Select
        
        
    Case "STAT"
        If objOWA.GetMsgList Then
            Set oElements = objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
            For Each oElement In oElements
                'strDataToSend = strDataToSend & intMsg & " 1 " & vbCrLf
                lngSize = lngSize + oElement.Text
            Next
            strDataToSend = "+OK " & objOWA.intMsgCount & " " & lngSize
        Else
            strDataToSend = "-ERR Unable to conect with mailbox "
        End If
    Case "RETR", "TOP "
        'strDataToSend = GetMsg(objDOMInbox.childNodes(1).childNodes(CInt(Mid$(strDataRecived, 6)) - 1).selectSingleNode("a:propstat/a:prop/a:href").Text)
        If Left$(strDataRecived, 4) = "RETR" Then
            intMsg = CInt(Mid$(strDataRecived, 6))
        Else
            intMsg = CInt(Mid$(strDataRecived, 5, InStr(5, strDataRecived, " ") - 5))
        End If
        With objDOMInbox.childNodes(1).childNodes(intMsg - 1)
            skPOP3.SendData "+OK " & .selectSingleNode("a:propstat/a:prop/a:getcontentlength").Text & " octets"
            
            bCloseSocket = Not (objOWA.GetMsg(.selectSingleNode("a:propstat/a:prop/a:href").Text, skPOP3))
        End With
        strDataToSend = vbCrLf & "."
    Case "QUIT"
        strDataToSend = "+OK " & m_strIP & " server signing off, 0 messages deleted"
        bCloseSocket = True
        
    Case "LIST"
        Set oElements = objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
        If Len(strDataRecived) > 6 Then
            intMsg = CInt(Mid$(strDataRecived, 6))
            strDataToSend = "+OK " & intMsg & " " & oElements(intMsg - 1).Text & vbCrLf
        Else
            intTotal = objOWA.intMsgCount
            'For intMsg = 1 To intTotal
            For Each oElement In oElements
                'strDataToSend = strDataToSend & intMsg & " 1 " & vbCrLf
                intMsg = intMsg + 1
                lngSize = lngSize + oElement.Text
                strDataToSend = strDataToSend & intMsg & " " & oElement.Text & vbCrLf
            Next
            skPOP3.SendData "+OK " & intTotal & " messages (" & lngSize & ") octets"
            Debug.Assert intTotal = oElements.Length
            skPOP3.SendData strDataToSend
            strDataToSend = "."
        End If
    Case "NOOP"
        strDataToSend = "+OK "
    Case "UIDL"
        Set oElements = objDOMInbox.selectNodes(XMLPATH & "a:id")
        If Len(strDataRecived) > 6 Then
            intMsg = CInt(Mid$(strDataRecived, 6))
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
        strDataToSend = "+OK Capability list follows" & vbCrLf & _
                        "USER" & vbCrLf & _
                        "UIDL" & vbCrLf & _
                        "."
    Case "DELE"
        intMsg = CInt(Mid$(strDataRecived, 6))
        If objOWA.Delete(intMsg) Then
            strDataToSend = "+OK message " & intMsg & " deleted"
            'Additional Sleep, My Outlook 2000 fails if delete msg faster
            MsgWaitObj 500
        Else
            strDataToSend = "-ERR deleting message " & Mid$(strDataRecived, 6)
            bCloseSocket = True
        End If
        
    Case "AUTH"
        strDataToSend = "-ERR " & Left$(strDataRecived, 4) & " unsuported feature "
    Case Else
        Debug.Assert False
        strDataToSend = "-ERR Syntax error"
    End Select
    skPOP3.SendData strDataToSend
    If bCloseSocket Then
        Set objDOMInbox = Nothing
        Set objOWA = Nothing
        'Reopen the socket to put it in listen mode
        'Maybe I nned to find other way
        skPOP3.CloseSck
        skPOP3.Bind vPorts(0), m_strIP
        skPOP3.Listen
    End If
    
Set oElement = Nothing
Set oElements = Nothing
Exit Sub
ErrHandler:
    Debug.Assert False
    skPOP3.CloseSck
    Err.Raise Err.Number, "SendPOP3: " & Err.Source, Err.Description

End Sub

Private Sub skPOP3_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skPOP3_Error " & Number & "->" & Description, Error
End Sub

Private Sub skSMTP_ConnectionRequest(ByVal requestID As Long)
    skSMTP.CloseSck
    skSMTP.Accept requestID
    skSMTP.SendData "220 " & m_strIP & " Simple Mail Transfer Service Ready"
End Sub
''
'Evaluate a STMP command (<A HREF='http://www.faqs.org/rfcs/rfc822.html' target='_blank'>RFC822</A>)
'and send a reply to the email client
'
'@param bytesTotal Bytes recived
Private Sub skSMTP_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
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
    Dim lngTotal As Long
    Dim bCloseSocket As Boolean
    Dim strDataRecived As String
    skSMTP.GetData strDataRecived
    '
    If Left$(strDataRecived, 4) <> "PASS" And strSMTPState <> "DATA" Then
        Debug.Print strDataRecived
    End If
    If strSMTPState = "LOGIN" Or strSMTPState = "PASSWORD" Then
        strDataRecived = "AUTH " & strDataRecived
    End If
    Select Case Left$(strDataRecived, 4)
    Case Empty
        strDataToSend = "220 " & m_strIP & " Simple Mail Transfer Service Ready"
    Case "HELO"
        strDataToSend = "250 " & m_strIP
    Case "EHLO"
        strDataToSend = "250-" & m_strIP
        strDataToSend = strDataToSend & vbNewLine & "250-AUTH LOGIN PLAIN"
        strDataToSend = strDataToSend & vbNewLine & "250 HELP"
    Case "AUTH"
        Dim objBase As New Base64
        Dim strReply As String
        If objOWA Is Nothing Then
            Set objOWA = New clsOWA
        End If
        If strDataRecived = "AUTH LOGIN" & vbCrLf Then
            'cDataToSend.Add "334 " & objBase.Encode(strUser & "@" & Mid(strExchSvrName, InStr(strExchSvrName, "//") + 2))
            strDataToSend = "334 " & objBase.Encode("Username:")
            strSMTPState = "LOGIN"
        ElseIf Left$(strDataRecived, 10) = "AUTH PLAIN" Then
            strReply = objBase.Decode(Mid$(strDataRecived, 12))
            If Len(strReply) > 1 Then
                strUser = Mid$(strReply, 2, InStr(2, strReply, Chr(0)) - 2)
                strPassWord = Mid$(strReply, InStr(2, strReply, Chr(0)) + 1)
                strDataToSend = ValidateSMTPAUTH
            Else
                strDataToSend = "533  Authentication required [AUTH]"
            End If
        Else
            Select Case strSMTPState
            Case "LOGIN"
                strUser = objBase.Decode(Mid$(strDataRecived, 6))
                strSMTPState = "PASSWORD"
                strDataToSend = "334 " & objBase.Encode("Password:")
            Case "PASSWORD"
                strPassWord = objBase.Decode(Mid$(strDataRecived, 6))
                strDataToSend = ValidateSMTPAUTH
                strSMTPState = ""
            Case Else
                strDataToSend = "504 Unrecognized authentication type."
                'strDataToSend = "556 CLIENT AUTHENTICATION REQUIRED. USE ESMTP EHLO AND AUTH"
            End Select
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
        ReDim vMailData(1 To 1)
        strMailData = strMailData & vbCrLf
    Case "QUIT"
        strDataToSend = "221 " & m_strIP & " Service closing transmission channel"
        bCloseSocket = True
    Case Else
        If strSMTPState = "DATA" Then
            
            If Right$(strDataRecived, 3) = "." & vbCrLf Then
                'strMailData = strMailData & Left$(strDataRecived, Len(strDataRecived) - 3)
                lngTotal = UBound(vMailData)
                vMailData(lngTotal) = Left$(strDataRecived, Len(strDataRecived) - 3)
                strMailData = strMailData & Join(vMailData, Empty)
                Erase vMailData
                If objOWA.SendMail(strMailData) Then
                     strDataToSend = "250 OK"
                 Else
                     strDataToSend = "500 Syntax error, command unrecognized"
                End If
                strSMTPState = Empty
                strMailData = Empty
            Else
                'strMailData = strMailData & strDataRecived
                lngTotal = UBound(vMailData)
                vMailData(lngTotal) = strDataRecived
                ReDim Preserve vMailData(1 To lngTotal + 1)
            End If
        Else
            Debug.Assert False
        End If
    End Select
    If strDataToSend <> Empty Then
        skSMTP.SendData strDataToSend
    End If
    If bCloseSocket Then
        'Reopen the socket to put it in listen mode
        'Maybe I nned to find other way
        Set objOWA = Nothing
        skSMTP.CloseSck
        skSMTP.Bind vPorts(1), m_strIP
        skSMTP.Listen
    End If
Exit Sub
ErrHandler:
    WriteLog "SendSMTP " & Err.Description, Error
    Debug.Assert False
    skSMTP.CloseSck
    'lngPop3Write = 0
End Sub

Private Sub skSMTP_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skPOP3_Error " & Number & "->" & Description, Error
End Sub
