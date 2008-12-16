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

Private bBusy As Boolean

Private Const ENDMAIL As String = "." & vbNewLine

''
'Close all sockets and clean OWA object
Private Sub Disconnect()
    Set skPOP3 = Nothing
    Set skSMTP = Nothing
    Set objOWA = Nothing
End Sub

''
'Initialize sockets.
Private Sub Start()
On Error GoTo ErrHandler
Set Config = New clsConfig
Config.ReadConfig
Set skPOP3 = New CSocketMaster
skPOP3.Bind Config.intPOP3Port, Config.strIP
skPOP3.Listen
If Config.bSMTPPort Then
    Set skSMTP = New CSocketMaster
    skSMTP.Bind Config.intSMTPPort, Config.strIP
    skSMTP.Listen
End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, "Start " & Err.Source, Err.Description
End Sub

''
'Validate the current user & password, using GetFoldersURL
'
'@return SMTP command with the result of the operation
'
'@see clsOWA.GetFoldersURL
Private Function ValidateSMTPAUTH() As String
    SetOWAObject
    Select Case objOWA.GetFoldersURL
    Case 200 To 299
        ValidateSMTPAUTH = "235 AUTHENTICATION SUCCESSFUL"
    Case 401
        'Wrong password
        ValidateSMTPAUTH = "535 5.7.0 Authentication failed"
    Case Else
        'Unknown error
        ValidateSMTPAUTH = "533 Authentication required [AUTH]"
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
    Socket_CloseSck skPOP3
End Sub

Private Sub skPOP3_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler
    Socket_ConnectionRequest skPOP3, requestID, OK & "POP3 server ready"
Exit Sub
ErrHandler:
WriteLog "POP3 ConnectionRequest ->" & Err.Source & vbTab & Err.Description, Fail
End Sub
''
'Evaluate one POP3 command  (<A HREF='http://www.faqs.org/rfcs/rfc1939.html' target='_blank'>RFC1939</A>)
'and send a reply to the email client
'
'@param bytesTotal bytes received
Private Sub skPOP3_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
      
Dim oElement        As IXMLDOMElement
Dim oElements       As IXMLDOMNodeList
Dim strDataToSend   As String
Dim intMsg          As Integer
Dim strDataRecived  As String
Dim vData()         As String
Dim strCommand      As String

    Const UNABLETOCONNECT As String = "Unable to connect with mailbox "
    
    skPOP3.GetData strDataRecived
    'Strip last chars (13+10)
    strDataRecived = Mid$(strDataRecived, 1, Len(strDataRecived) - 2)
    If Len(strDataRecived) = 0 Then
        strCommand = ""
    Else
        vData = Split(strDataRecived, " ")
        strCommand = vData(0)
    End If
        
    If strCommand = "PASS" Then
        WriteLog "PASS **********", Paranoid
    Else
        WriteLog strDataRecived, Paranoid
    End If
    Select Case strCommand
    Case "USER"
        Config.Profile.strUser = vData(1)
        strDataToSend = OK & "Password required for " & Config.Profile.strUser
    Case "PASS"
        SetOWAObject
        Config.Profile.strPassWord = vData(1)
        Select Case objOWA.GetFoldersURL
        Case 200 To 299
            strDataToSend = OK & "mailbox for " & Config.Profile.strUser & " ready"
        Case 401
            strDataToSend = Error & "Invalid Password for " & Config.Profile.strUser
        Case -2147012889
            strDataToSend = Error & "The server name or address could not be resolved"
        Case Else
            strDataToSend = Error & UNABLETOCONNECT
        End Select
    Case "STAT"
        If objOWA.GetMsgList Then
            Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
            For Each oElement In oElements
                'Calcutate the size of the entire inbox folder.
                objOWA.lngSize = objOWA.lngSize + oElement.Text
            Next
            strDataToSend = OK & objOWA.intMsgCount & " " & objOWA.lngSize
        Else
            strDataToSend = Error & UNABLETOCONNECT
        End If
    Case "RETR", "TOP"
        intMsg = CInt(vData(1))
        With objOWA.objDOMInbox.childNodes(1).childNodes(intMsg - 1)
            skPOP3.SendData OK & .selectSingleNode("a:propstat/a:prop/a:getcontentlength").Text & " octets"
            objOWA.GetMsg .selectSingleNode("a:href").Text, skPOP3
        End With
    Case "QUIT"
        strDataToSend = OK & Config.strIP & " server signing off, 0 messages deleted"
    Case "NOOP"
        strDataToSend = OK
    Case "UIDL", "LIST"
        If strCommand = "UIDL" Then
            Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:uid")
        Else
            Set oElements = objOWA.objDOMInbox.selectNodes(XMLPATH & "a:getcontentlength")
        End If
        If Len(strDataRecived) > 5 Then
            strDataToSend = OK & vData(1) & " " & oElements(vData(1) - 1).Text & vbNewLine
        Else
            strDataToSend = OK & objOWA.intMsgCount & " messages (" & objOWA.lngSize & ") octets"
            For Each oElement In oElements
                intMsg = intMsg + 1
                strDataToSend = strDataToSend & (vbNewLine & intMsg & " " & oElement.Text)
            Next
            strDataToSend = strDataToSend & vbNewLine & ENDMAIL
        End If
    Case "CAPA"
        strDataToSend = OK & "Capability list follows" & vbNewLine & _
                        "USER" & vbNewLine & _
                        "UIDL" & vbNewLine & _
                        ENDMAIL
    Case "DELE"
        intMsg = CInt(vData(1))
        If objOWA.Delete(intMsg) Then
            strDataToSend = OK & "message " & intMsg & " deleted"
            'Additional Sleep, My Outlook 2000 fails if delete msg faster
            MsgWaitObj 500
        Else
            strDataToSend = Error & "deleting message " & intMsg
        End If
    Case "AUTH"
        strDataToSend = Error & strCommand & " unsupported feature "
    Case Else
        Debug.Assert False
        WriteLog "Unknown data: " & strDataRecived, Warning
        strDataToSend = Error & "Syntax error"
    End Select
    If LenB(strDataToSend) > 0 And skPOP3.State = sckConnected Then skPOP3.SendData strDataToSend
    Set oElement = Nothing
    Set oElements = Nothing
Exit Sub
ErrHandler:
    Debug.Assert False
    skPOP3.CloseSck
    Err.Raise Err.Number, "SendPOP3: " & Err.Source, Err.Description

End Sub

Private Sub skPOP3_Error(ByVal Number As Long, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skPOP3_Error " & Number & "->" & Description, Fail
End Sub


Private Sub skSMTP_CloseSck()
    Socket_CloseSck skSMTP
End Sub

Private Sub skSMTP_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler
    Socket_ConnectionRequest skSMTP, requestID, "220 " & Config.strIP & " Simple Mail Transfer Service Ready"
Exit Sub
ErrHandler:
    WriteLog "skSMTP_ConnectionRequest ->" & Err.Source & vbTab & Err.Description, Fail
End Sub
''
'Evaluate one STMP command (<A HREF='http://www.faqs.org/rfcs/rfc2821.html' target='_blank'>RFC2821</A>)
'and send a reply to the email client
'
'@param bytesTotal Bytes received
Private Sub skSMTP_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
    
    Const CODEOK    As String = "250 OK"
    
    Dim strDataToSend   As String
    Dim lngTotal        As Long
    Dim strDataRecived  As String
    Dim strSTMPCommmad  As String
    Static strSMTPState As String
    Static vMailData()  As String

    skSMTP.GetData strDataRecived
    strSTMPCommmad = Left$(strDataRecived, 4)
    '
    If strSTMPCommmad <> "PASS" And strSMTPState <> "DATA" Then
        WriteLog strDataRecived, Paranoid
    End If
    
    Select Case strSMTPState
    'TODO Check the correct value of strSMTPState at RFC
    Case "LOGIN"
        Config.Profile.strUser = base64_handler(strDataRecived)
        strSMTPState = "PASSWORD"
        '"Password:" base64 encoded
        strDataToSend = "334 UGFzc3dvcmQ6"
    
    Case "PASSWORD"
        Config.Profile.strPassWord = base64_handler(strDataRecived)
        strDataToSend = ValidateSMTPAUTH
        strSMTPState = vbNullString
    
    Case "DATA"
        lngTotal = UBound(vMailData)
        'If Msg end with CRLF+"." + CRLF, send the mail to the server
        If Right$(strDataRecived, 5) = vbNewLine & ENDMAIL Then
            vMailData(lngTotal) = Left$(strDataRecived, Len(strDataRecived) - 5)
            Debug.Assert Not (objOWA Is Nothing)
            If objOWA.SendMail(Join(vMailData, Empty)) Then
                 strDataToSend = CODEOK
             Else
                 strDataToSend = "500 Syntax error, command unrecognized"
            End If
            Erase vMailData
            strSMTPState = vbNullString
        Else
            'Store Msg in the buffer
            vMailData(lngTotal) = strDataRecived
            ReDim Preserve vMailData(1 To lngTotal + 1)
        End If
    Case Else
        Select Case strSTMPCommmad
        Case Empty
            strDataToSend = "220 " & Config.strIP & " Simple Mail Transfer Service Ready"
        Case "HELO"
            strDataToSend = "250 " & Config.strIP
        Case "EHLO"
            strDataToSend = "250-" & Config.strIP & vbNewLine & _
                            "250-AUTH LOGIN PLAIN" & vbNewLine & _
                            "250 HELP"
        Case "AUTH"
            If strDataRecived = "AUTH LOGIN" & vbNewLine Then
                '"Username:" as base64 encoded
                strDataToSend = "334 VXNlcm5hbWU6"
                strSMTPState = "LOGIN"
            ElseIf Left$(strDataRecived, 10) = "AUTH PLAIN" Then
                If Len(strDataRecived) = 12 Then
                    strDataToSend = "534 Authentication mechanism is too weak"
                Else
                    strDataRecived = base64_handler(Mid$(strDataRecived, 12))
                    If Len(strDataRecived) > 1 Then
                        Config.Profile.strUser = Mid$(strDataRecived, 2, InStr(2, strDataRecived, vbNullChar) - 2)
                        Config.Profile.strPassWord = Mid$(strDataRecived, InStr(2, strDataRecived, vbNullChar) + 1)
                        strDataToSend = ValidateSMTPAUTH
                    Else
                        strDataToSend = "534 Authentication mechanism is too weak"
                    End If
                End If
            Else
                strDataToSend = "504 Unrecognized authentication type."
            End If
        Case "MAIL"
            If Config.Profile.strPassWord = Empty Or Config.Profile.strUser = Empty Then
                strDataToSend = "530 Authentication required"
            Else
                ReDim vMailData(1 To 4)
                vMailData(1) = strDataRecived
                strDataToSend = CODEOK
            End If
        Case "RCPT"
            vMailData(2) = vMailData(2) & strDataRecived
            strDataToSend = CODEOK
        Case "RSET"
            strSMTPState = vbNullString
            Erase vMailData
            strDataToSend = CODEOK
        Case "DATA"
            strDataToSend = "354 Start mail input; end with <CRLF>.<CRLF>"
            strSMTPState = "DATA"
            vMailData(3) = vbNewLine
        Case "QUIT"
            strSMTPState = vbNullString
            Erase vMailData
            strDataToSend = "221 " & Config.strIP & " Service closing transmission channel"
        Case Else
            Debug.Assert False
        End Select
    End Select
    If LenB(strDataToSend) > 0 And skSMTP.State = sckConnected Then
        skSMTP.SendData strDataToSend
    End If
Exit Sub
ErrHandler:
    WriteLog "skSMTP_DataArrival ->" & Err.Source & vbTab & Err.Description, Fail
    Debug.Assert False
    skSMTP.CloseSck
End Sub

Private Sub skSMTP_Error(ByVal Number As Long, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "skSMTP_Error " & Number & "->" & Description, Fail
End Sub
''
'Encode/Decode one string in Base64 Format
'
'@param strString String to handle
'@param bEncode True to encode the string, false to decode
'
'@return Output Decoded/Encoded string
Private Function base64_handler(ByVal strString As String) As String
Dim objXMLDoc   As DOMDocument
Dim bArray()    As Byte

Set objXMLDoc = New DOMDocument

With objXMLDoc
    .async = False
    .validateOnParse = False
    .loadXML "<xml/>"
    .documentElement.dataType = "bin.base64"
    .documentElement.Text = strString
    base64_handler = StrConv(.documentElement.nodeTypedValue, vbUnicode)
End With
Set objXMLDoc = Nothing

End Function



''
'Initialize the OWA object
'
'@see clsOWA
Private Sub SetOWAObject()
    If objOWA Is Nothing Then
        Set objOWA = New clsOWA
    End If
End Sub

Private Sub Socket_ConnectionRequest(ByRef oSocket As CSocketMaster, ByVal requestID As Long, ByVal strMessage As String)
On Error GoTo ErrHandler

    If oSocket.State = sckListening And Not bBusy Then
        oSocket.CloseSck
        oSocket.Accept requestID
        oSocket.SendData strMessage
        WriteLog strMessage, Paranoid
        bBusy = True
    Else
        Debug.Assert False
        WriteLog "server busy", Fail
    End If
Exit Sub
ErrHandler:
    WriteLog "Socket_ConnectionRequest ->" & Err.Source & vbTab & Err.Description, Fail
End Sub

Private Sub Socket_CloseSck(ByRef oSocket As CSocketMaster)
    'Reopen the socket to put it in listen mode
    'Maybe I need to find other way
    oSocket.CloseSck
    oSocket.Bind oSocket.LocalPort, Config.strIP
    oSocket.Listen
    If bBusy Then
        Config.Profile.strPassWord = vbNullString
        Config.Profile.strUser = vbNullString
        bBusy = False
        Set objOWA = Nothing
    Else
        Debug.Assert False
    End If
End Sub

