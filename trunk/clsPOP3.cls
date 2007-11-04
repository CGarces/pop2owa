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
Private WithEvents skSMTP As CSocketMaster
Attribute skSMTP.VB_VarHelpID = -1
''
'POP3 Socket object
Private WithEvents skPOP3 As CSocketMaster
Attribute skPOP3.VB_VarHelpID = -1

''
'Handle WEBDAW commands with OWA server
Private objOWA As clsOWA

Private m_strIP As String
Private intState As States

Private Enum States
    None = 0
    POP3 = 1
    SMTP = 2
    BOTH = 3
End Enum

''
'Close al sockets and call to WSACleanup API
Private Sub Disconnect()
    If Not (skPOP3 Is Nothing) Then
        skPOP3.CloseSck
    End If
    If Not (skSMTP Is Nothing) Then
        skSMTP.CloseSck
    End If
    Set objOWA = Nothing
End Sub

''
'Initialize Winsock code.
'
Private Sub Start()
On Error GoTo ErrHandler

Dim vPorts(0 To 1)  As Long
Dim c               As cRegistry

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
    .ValueKey = "SMTPEnabled"
    If .Value = 0 Then
        vPorts(1) = 0
    Else
        .ValueKey = "SMTP"
        vPorts(1) = .Value
    End If
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
    Disconnect
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
    Disconnect
    Start
End Sub

Private Sub Class_Terminate()
    Disconnect
End Sub

Private Sub skPOP3_CloseSck()
    'Reopen the socket to put it in listen mode
    'Maybe I need to find other way
    skPOP3.CloseSck
    skPOP3.Bind skPOP3.LocalPort, m_strIP
    skPOP3.Listen
    Select Case intState
    Case POP3
        strPassWord = Empty
        strUser = Empty
        intState = None
        Set objOWA = Nothing
    Case BOTH
        intState = SMTP
    End Select
End Sub

Private Sub skPOP3_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler
    skPOP3.CloseSck
    skPOP3.Accept requestID
    skPOP3.SendData "+OK POP3 server ready"
    WriteLog "+OK POP3 server ready", Paranoid
Exit Sub
ErrHandler:
WriteLog "POP3 ConnectionRequest ->" & Err.Source & vbTab & Err.Description, Error
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
    Dim lngSize         As Long
    Dim strDataRecived  As String
    WriteLog "skPOP3_DataArrival." & bytesTotal & "Bytes", Paranoid
    skPOP3.GetData strDataRecived
    '
    If Left$(strDataRecived, 4) <> "PASS" Then
        WriteLog strDataRecived, Paranoid
    End If
    Select Case Left$(strDataRecived, 4)
    Case Empty
        strDataToSend = "+OK POP3 server ready"
    Case "USER"
        strUser = Mid$(strDataRecived, 6, Len(strDataRecived) - 7)
        strDataToSend = "+OK Password required for " & strUser
    Case "PASS"
        Select Case intState
        Case SMTP
            intState = BOTH
        Case None
            intState = POP3
        End Select
        If objOWA Is Nothing Then
            Set objOWA = New clsOWA
        End If
        strPassWord = Mid$(strDataRecived, 6, Len(strDataRecived) - 7)
        Select Case objOWA.GetFoldersURL
        Case 200 To 299
            strDataToSend = "+OK mailbox for " & strUser & " ready"
        Case 401
            strDataToSend = "-ERR Invalid Password for " & strUser
        Case Else
            strDataToSend = "-ERR Unable to connect with mailbox "
        End Select
    Case "STAT"
        If objOWA.GetMsgList Then
            Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
            For Each oElement In oElements
                'strDataToSend = strDataToSend & intMsg & " 1 " & vbCrLf
                lngSize = lngSize + oElement.Text
            Next
            strDataToSend = "+OK " & objOWA.intMsgCount & " " & lngSize
        Else
            strDataToSend = "-ERR Unable to connect with mailbox "
        End If
    Case "RETR", "TOP "
        'strDataToSend = GetMsg(objDOMInbox.childNodes(1).childNodes(CInt(Mid$(strDataRecived, 6)) - 1).selectSingleNode("a:propstat/a:prop/a:href").Text)
        If Left$(strDataRecived, 4) = "RETR" Then
            intMsg = CInt(Mid$(strDataRecived, 6))
        Else
            intMsg = CInt(Mid$(strDataRecived, 5, InStr(5, strDataRecived, " ") - 5))
        End If
        With objOWA.objDOMInbox.childNodes(1).childNodes(intMsg - 1)
            skPOP3.SendData "+OK " & .selectSingleNode("a:propstat/a:prop/a:getcontentlength").Text & " octets"
            objOWA.GetMsg .selectSingleNode("a:href").Text, skPOP3
        End With
        strDataToSend = vbCrLf & "."
    Case "QUIT"
        strDataToSend = "+OK " & m_strIP & " server signing off, 0 messages deleted"
    Case "LIST"
        Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
        If Len(strDataRecived) > 6 Then
            intMsg = CInt(Mid$(strDataRecived, 6))
            strDataToSend = "+OK " & intMsg & " " & oElements(intMsg - 1).Text & vbCrLf
        Else
            strDataToSend = "+OK " & intMsg & " messages (" & lngSize & ") octets" & vbCrLf
            For Each oElement In oElements
                'strDataToSend = strDataToSend & intMsg & " 1 " & vbCrLf
                intMsg = intMsg + 1
                lngSize = lngSize + oElement.Text
                strDataToSend = strDataToSend & intMsg & " " & oElement.Text & vbCrLf
            Next
            strDataToSend = strDataToSend & vbCrLf & "." & vbCrLf
        End If
    Case "NOOP"
        strDataToSend = "+OK "
    Case "UIDL"
        Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:id")
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
        End If
    Case "AUTH"
        strDataToSend = "-ERR " & Left$(strDataRecived, 4) & " unsuported feature "
    Case Else
        Debug.Assert False
        strDataToSend = "-ERR Syntax error"
    End Select
    If skPOP3.State = sckConnected Then skPOP3.SendData strDataToSend
    
    Set oElement = Nothing
    Set oElements = Nothing
Exit Sub
ErrHandler:
    Debug.Assert False
    skPOP3.CloseSck
    Err.Raise Err.Number, "SendPOP3: " & Err.Source, Err.Description

End Sub

Private Sub skPOP3_Error(ByVal Number As Long, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skPOP3_Error " & Number & "->" & Description, Error
End Sub


Private Sub skSMTP_CloseSck()
    Debug.Assert False
    'Socket closed by email client
    Select Case intState
    Case SMTP
        intState = None
        strPassWord = Empty
        strUser = Empty
        Set objOWA = Nothing
    Case BOTH
        intState = POP3
    End Select
    'Reopen the socket to put it in listen mode
    'Maybe I need to find other way
    skSMTP.CloseSck
    skSMTP.Bind skSMTP.LocalPort, m_strIP
    skSMTP.Listen
End Sub

Private Sub skSMTP_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler
    skSMTP.CloseSck
    skSMTP.Accept requestID
    skSMTP.SendData "220 " & m_strIP & " Simple Mail Transfer Service Ready"
    WriteLog "220 " & m_strIP & " Simple Mail Transfer Service Ready", Paranoid
Exit Sub
ErrHandler:
    WriteLog "skSMTP_ConnectionRequest ->" & Err.Source & vbTab & Err.Description, Error
End Sub
''
'Evaluate a STMP command (<A HREF='http://www.faqs.org/rfcs/rfc2821.html' target='_blank'>RFC2821</A>)
'and send a reply to the email client
'
'@param bytesTotal Bytes recived
Private Sub skSMTP_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
    
    Const ENDMAIL   As String = vbCrLf & "." & vbCrLf
    Const CODEOK    As String = "250 OK"
    
    Dim strDataToSend   As String
    Dim lngTotal        As Long
    Dim strDataRecived  As String
    
    Static strSMTPState As String
    Static strMailData  As String
    Static vMailData()  As String

    skSMTP.GetData strDataRecived
    '
    If Left$(strDataRecived, 4) <> "PASS" And strSMTPState <> "DATA" Then
        WriteLog strDataRecived, Paranoid
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
        
        Select Case intState
        Case POP3
            intState = BOTH
        Case None
            intState = SMTP
        End Select

        If objOWA Is Nothing Then
            Set objOWA = New clsOWA
        End If
        If strDataRecived = "AUTH LOGIN" & vbCrLf Then
            'cDataToSend.Add "334 " & objBase.Encode(strUser & "@" & Mid(strExchSvrName, InStr(strExchSvrName, "//") + 2))
            strDataToSend = "334 " & base64_encode("Username:")
            strSMTPState = "LOGIN"
        ElseIf Left$(strDataRecived, 10) = "AUTH PLAIN" Then
            strDataRecived = base64_decode(Mid$(strDataRecived, 12))
            If Len(strDataRecived) > 1 Then
                strUser = Mid$(strDataRecived, 2, InStr(2, strDataRecived, Chr(0)) - 2)
                strPassWord = Mid$(strDataRecived, InStr(2, strDataRecived, Chr(0)) + 1)
                strDataToSend = ValidateSMTPAUTH
            Else
                strDataToSend = "533  Authentication required [AUTH]"
            End If
        Else
            Select Case strSMTPState
            Case "LOGIN"
                strUser = base64_decode(Mid$(strDataRecived, 6))
                strSMTPState = "PASSWORD"
                strDataToSend = "334 " & base64_encode("Password:")
            
            Case "PASSWORD"
                strPassWord = base64_decode(Mid$(strDataRecived, 6))
                strDataToSend = ValidateSMTPAUTH
                strSMTPState = ""
            
            Case Else
                strDataToSend = "504 Unrecognized authentication type."
            End Select
        End If
    Case "MAIL"
        If strPassWord = Empty Or strUser = Empty Then
            strDataToSend = "556 CLIENT AUTHENTICATION REQUIRED. USE ESMTP EHLO AND AUTH"
        Else
            strMailData = strDataRecived
            strDataToSend = CODEOK
        End If
    Case "RCPT"
        strMailData = strMailData & strDataRecived
        strDataToSend = CODEOK
    Case "RSET"
        strSMTPState = ""
        strMailData = ""
        Erase vMailData
        strDataToSend = CODEOK
    Case "DATA"
        strDataToSend = "354 Start mail input; end with <CRLF>.<CRLF>"
        strSMTPState = "DATA"
        ReDim vMailData(1 To 1)
        strMailData = strMailData & vbCrLf
    Case "QUIT"
        strSMTPState = ""
        strMailData = ""
        Erase vMailData
        strDataToSend = "221 " & m_strIP & " Service closing transmission channel"
    Case Else
        If strSMTPState = "DATA" Then
            
            If Right$(strDataRecived, 5) = ENDMAIL Then
                'strMailData = strMailData & Left$(strDataRecived, Len(strDataRecived) - 3)
                lngTotal = UBound(vMailData)
                vMailData(lngTotal) = Left$(strDataRecived, Len(strDataRecived) - 5)
                strMailData = strMailData & Join(vMailData, Empty)
                Erase vMailData
                Debug.Assert Not (objOWA Is Nothing)
                If objOWA.SendMail(strMailData) Then
                     strDataToSend = CODEOK
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
    If strDataToSend <> Empty And skSMTP.State = sckConnected Then
        skSMTP.SendData strDataToSend
    End If
Exit Sub
ErrHandler:
    WriteLog "skSMTP_DataArrival ->" & Err.Source & vbTab & Err.Description, Error
    Debug.Assert False
    skSMTP.CloseSck
    'lngPop3Write = 0
End Sub

Private Sub skSMTP_Error(ByVal Number As Long, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skPOP3_Error " & Number & "->" & Description, Error
End Sub

''
'Encode one string in Base64 Format
'
'@param strString String to be formated
'@return Output String in Base64 format
Private Function base64_encode(ByVal strString As String) As String
Dim objXMLDoc   As DOMDocument
Dim bArray()    As Byte

    Set objXMLDoc = New DOMDocument
    bArray() = StrConv(strString, vbFromUnicode)
    
    With objXMLDoc
        .async = False
        .validateOnParse = False
        .loadXML "<xml/>"
        .documentElement.dataType = "bin.base64"
        .documentElement.nodeTypedValue = bArray
        base64_encode = .Text
    End With
    Set objXMLDoc = Nothing
End Function

''
'Decode one string in Base64 Format
'
'@param strString String in Base64 format
'@return Output Decoded string
Private Function base64_decode(ByVal strString As String) As String
Dim objXMLDoc   As DOMDocument

    Set objXMLDoc = New DOMDocument
    With objXMLDoc
        .async = False
        .validateOnParse = False
        .loadXML "<xml/>"
        .documentElement.dataType = "bin.base64"
        .documentElement.Text = strString
        base64_decode = StrConv(.documentElement.nodeTypedValue, vbUnicode)
    End With
    Set objXMLDoc = Nothing
End Function
