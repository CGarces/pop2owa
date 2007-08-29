VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOWA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
'WebDav code to comunicate with MS Exchange server
'
'@author Carlos B
'@version 1.0
'@date 26/11/2005

Option Explicit


Private strCookies  As String
Private strURLSend  As String
Private strURLInbox As String


Public intMsgCount As Integer


''
'Get the URL of inbox messages
'
'@param intMsg  Message position in the inbox folder
Private Function GetURL(ByVal intMsg As Integer) As String
    GetURL = objDOMInbox.childNodes(1).childNodes(intMsg - 1).selectSingleNode("a:href").Text

End Function



''
'Retrieve message in RFC 821 format and send it to the client.
'
'@param strUrl  Url of the .EML file
'@param objPOP3 Socket object to send data
'@return True if sucess.
Public Function GetMsg(ByVal strURL As String, ByRef objPOP3 As CSocketMaster) As Boolean
On Error GoTo ErrHandler

Dim oXMLHTTP    As XMLHTTP
Dim lngTimer    As Long
Dim strMsg      As String

Set oXMLHTTP = New XMLHTTP

With oXMLHTTP
    '.setOption 3, ""
    .open "GET", strURL & "?ID=" & Timer, True, strUser, strPassWord
    .setRequestHeader "Content-type", "text/xml"
    .setRequestHeader "Depth", "infinity"
    .setRequestHeader "Translate", "f"
    
    If strCookies <> Empty Then
        .setRequestHeader "Cookies", "Necessary according to Q234486"
        .setRequestHeader "Cookies", strCookies
    End If
    WriteLog "Getting " & vbTab & strURL, Information
    .send
    lngTimer = Timer
    While Not .readyState = 4
        MsgWaitObj 100
        If Timer - lngTimer > 30 Then
            WriteLog "Donwloading readyState  " & .readyState & vbTab & strURL, Warning
            objPOP3.SendData "X-POP2OWA-Header: Header to avoid timeout" & vbCrLf
            lngTimer = Timer
'            Debug.Assert False
        End If
    Wend
    WriteLog "Donwloaded Status " & .Status & vbTab & strURL, Information
    If (.Status >= 200 And .Status < 300) Then
        strMsg = StrConv(.responseBody, vbUnicode)
        'Fix errors in inline images
        ParseAtachment strMsg
        'Fix OWA errors in recived messages
        If InStr(strMsg, "X-MimeOLE: Produced By Microsoft Exchange") > 0 Then
            ParseHeader strMsg
        End If
        objPOP3.SendData strMsg
        GetMsg = True
    Else
        WriteLog "GetMsg.Status " & .statusText, Warning
        objPOP3.SendData "-ERR " & .statusText
        GetMsg = False
    End If
End With
Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    WriteLog "GetMsg " & Err.Description, Error
    objPOP3.SendData "-ERR " & Err.Description
    GetMsg = False
End Function

''
'Delete a message
'
'@param intMsg  Position of the message in the inbox folder
Public Function Delete(ByVal intMsg As Integer) As Boolean
On Error GoTo ErrHandler
Dim oXMLHTTP    As XMLHTTP
    
    Set oXMLHTTP = New XMLHTTP
    With oXMLHTTP
        '.setOption 3, ""
        .open "DELETE", GetURL(intMsg), True, strUser, strPassWord
        If strCookies <> Empty Then
            .setRequestHeader "Cookies", "Necessary according to Q234486"
            .setRequestHeader "Cookies", strCookies
        End If
        .setRequestHeader "Depth", "infinity"
        .send
        While Not .readyState = 4
            MsgWaitObj 100
        Wend

    '    Debug.Print .Status
    End With
    Delete = True
Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    Delete = False
    Set oXMLHTTP = Nothing
End Function
''
'Get the list of al messages in the inbox folder.
'
'@return Reply to e-mail client
Public Function GetMsgList() As Boolean
On Error GoTo ErrHandler
Dim oXMLHTTP As XMLHTTP
' Create an instance of the xmlHTTP object
Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    ' Open a request to a particular url
    .open "SEARCH", strURLInbox, True, strUser, strPassWord
    ' Set the header type - we want XML
    .setRequestHeader "Content-type", "text/xml"
    .setRequestHeader "Depth", "1"
    If strCookies <> Empty Then
        .setRequestHeader "Cookies", "Necessary according to Q234486"
        .setRequestHeader "Cookies", strCookies
    End If
    ' Send the request
    MsgWaitObj 100
    WriteLog "Geting msg list ", Information
    Call .send("<?xml version='1.0' ?>" & _
                "<a:searchrequest xmlns:a='DAV:'><a:sql>" & _
                 "SELECT" & _
                 " ""DAV:id""" & _
                 ",""DAV:getcontentlength""" & _
                 " FROM scope('shallow traversal of """ & strURLInbox & """')" & _
                 " WHERE ""DAV:ishidden""=False" & _
                 " AND ""DAV:isfolder""=False" & _
                 "</a:sql></a:searchrequest>")
    
    While Not .readyState = 4
        MsgWaitObj 100
    Wend
    
    If .Status = 207 Then

        Set objDOMInbox = .responseXML
        objDOMInbox.setProperty "SelectionLanguage", "XPath"
        objDOMInbox.setProperty "SelectionNamespaces", "xmlns:a='DAV:'"
        intMsgCount = objDOMInbox.childNodes(1).childNodes.Length
        WriteLog "msg list fetch " & intMsgCount & " messages", Information
        #If DEBUG_MODE = 1 Then
            objDOMInbox.save App.Path & "\listado.XML"
        #End If
    Else
        WriteLog "Reciving msg list " & .statusText, Warning
        Err.Raise vbObjectError + 1, "Reciving msg", .statusText
    End If
End With
Set oXMLHTTP = Nothing

GetMsgList = True
Exit Function
ErrHandler:
    Debug.Assert False
    Set oXMLHTTP = Nothing
    GetMsgList = False
End Function


''
'Get Well-Known Mailbox Folder URLs that are retrieved from a users root mailbox folder.
'
'@return True if the URLs are retrieved
'@remarks <A HREF='http://msdn2.microsoft.com/en-us/library/ms992623.aspx' target='_blank'>Microsoft Exchange 2000 SDK June 2005</A>
Public Function GetFoldersURL() As Long

On Error GoTo ErrHandler
Dim query           As String
Dim strURL          As String
Dim bFBALoginFail   As Boolean
Dim oXMLHTTP        As XMLHTTP
Dim strAccount      As String
Dim objDOMFolders   As DOMDocument

If InStr(strUser, "^") > 0 Then
    strAccount = Mid(strUser, InStr(strUser, "^") + 1)
    strUser = Left(strUser, InStr(strUser, "^") - 1)
Else
    strAccount = strUser
End If

If bAuthentication Then
    WriteLog "GetFoldersURL : Trying LogingFBA ", Information
    bFBALoginFail = Not (LogingFBA)
End If
If bFBALoginFail Then
    GetFoldersURL = -1
Else
    Set oXMLHTTP = New XMLHTTP
    strURL = strExchSvrName & "/exchange/" & strAccount & "/"

    With oXMLHTTP
        WriteLog "Getting Folders : " & strURL
        '.setOption 3, ""
        .open "PROPFIND", strURL, True, strUser, strPassWord
        .setRequestHeader "Content-Type", "text/xml"
        .setRequestHeader "Depth", "0"
        If strCookies <> Empty Then
            .setRequestHeader "Cookies", "Necessary according to Q234486"
            .setRequestHeader "Cookies", strCookies
        End If
        .send "<?xml version='1.0'?>" & _
            "<a:propfind xmlns:a='DAV:'>" & _
            "<a:prop xmlns:m='urn:schemas:httpmail:'>" & _
            "<m:sendmsg/>" & _
            "<m:inbox/>" & _
            "</a:prop>" & _
            "</a:propfind>"

        While Not .readyState = 4
            MsgWaitObj 100
        Wend
    
        ' process the result
        If (.Status >= 200 And .Status < 300) Then
            Set objDOMFolders = .responseXML
            objDOMFolders.setProperty "SelectionLanguage", "XPath"
            objDOMFolders.setProperty "SelectionNamespaces", "xmlns:a='DAV:'"
            objDOMFolders.setProperty "SelectionNamespaces", "xmlns:a='DAV:' xmlns:d='urn:schemas:httpmail:'"
            WriteLog "GetFoldersURL: " & vbCrLf & objDOMFolders.xml & vbCrLf, Paranoid
            strURLSend = objDOMFolders.selectSingleNode(XMLPATH & "d:sendmsg").Text
            strURLInbox = objDOMFolders.selectSingleNode(XMLPATH & "d:inbox").Text & "/"
            Set objDOMFolders = Nothing
        Else
            Debug.Assert False
            WriteLog "GetFoldersURL: " & .Status & vbTab & .statusText, Warning
        End If
        GetFoldersURL = .Status
    End With
End If
Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    WriteLog "GetFoldersURL: " & Err.Number & vbTab & Err.Description, Error
    GetFoldersURL = -1
    Debug.Assert False
End Function

''
'Access the Exchange store via WebDAV with Form-Based-Authentication turned on.
'
'@return Retrun True if the login is correct.
Private Function LogingFBA() As Boolean
On Error GoTo ErrHandler
Dim vHeaders()  As String
Dim intCounter  As Integer
Dim intIndex    As Integer
Dim strURL      As String
Dim oXMLHTTP    As XMLHTTP
Dim strUsername As String
Dim strHeaders  As String
Dim strVersion  As String


strUsername = strUser
If InStr(strUser, "@") > 0 Then
    strUsername = Left(strUsername, InStr(strUser, "@") - 1)
End If
strURL = strExchSvrName & "/exchweb/bin/auth/owaauth.dll"
WriteLog "LogingFBA -> strURL: " & strURL, Information

Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    '.setOption 3, ""
    .open "POST", strURL, True
    .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    'Generate the body for FBA login
    .send "destination=" & strExchSvrName & "/exchange/&trusted=4&username=" & strUsername & "&password=" & strPassWord
    While Not .readyState = 4
        MsgWaitObj 100
    Wend
    strHeaders = .getAllResponseHeaders()
End With
Set oXMLHTTP = Nothing

WriteLog "LogingFBA -> Headers: " & vbCrLf & strHeaders, Paranoid
vHeaders = Split(strHeaders, vbNewLine)

For intIndex = 0 To UBound(vHeaders) - 1
    If Left$(vHeaders(intIndex), 10) = "Set-Cookie" Then
        intCounter = intCounter + 1
        strCookies = strCookies & Mid$(vHeaders(intIndex), 12) & ";"
    End If
    If Left$(vHeaders(intIndex), 13) = "X-OWA-Version" Then
        strVersion = Mid$(vHeaders(intIndex), 16)
    End If
Next

If intCounter > 1 Then
    LogingFBA = True
Else
    Select Case True
    Case strVersion = Empty
        WriteLog "LogingFBA -> X-OWA-Version headder missing", Error
        LogingFBA = True
    Case Int(Left$(strVersion, InStr(strVersion, "."))) >= 8
        WriteLog "Error LogingFBA: " & strVersion, Paranoid
        LogingFBA = True
    Case Else
        WriteLog "LogingFBA -> Error getting headers: " & vbCrLf & strHeaders, Error
        LogingFBA = False
    End Select
End If
Exit Function
ErrHandler:
    WriteLog "GetFoldersURL: " & Err.Number & vbTab & Err.Description, Error
    Set oXMLHTTP = Nothing
    LogingFBA = False
End Function

''
'Send one email in standar rfc821 format.
'
'@param strMailData Message string (rfc821)
'@return <B>True</B> if the mail is sent.
Public Function SendMail(ByVal strMailData As String) As Boolean

On Error GoTo ErrHandler
Dim oXMLHTTP As XMLHTTP

    Set oXMLHTTP = New XMLHTTP
     
    'Create the DAV PUT request.
    'xmlReq.setOption 3, ""
    With oXMLHTTP
        .open "PUT", strURLSend, True, strUser, strPassWord
        If strCookies <> Empty Then
            .setRequestHeader "Cookies", "Necessary according to Q234486"
            .setRequestHeader "Cookies", strCookies
        End If
        
        .setRequestHeader "Content-Type", "message/rfc821"
        If Not bSaveinsent Then
            .setRequestHeader "Saveinsent", "f"
        End If
        .send strMailData
        While Not oXMLHTTP.readyState = 4
             MsgWaitObj 100
        Wend
        
        'Process the results.
        If (.Status >= 200 And .Status < 300) Then
           SendMail = True
        Else
           WriteLog "SendMail: " & .Status & vbTab & .statusText, Warning
           SendMail = False
        End If
    End With
    Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    WriteLog "SendMail: " & Err.Number & vbTab & Err.Description, Error
    Set oXMLHTTP = Nothing
    SendMail = False
End Function

''
'Parse the mail header to fix incorrect email adress with non US chars.
'
'@param strMsg Email message in RFC 821 format
Private Sub ParseHeader(ByRef strMsg As String)
Dim strHeader       As String
Dim lngPos          As Long
Dim lngPosHeader    As Long
Dim lngPosEnd       As Long
Dim bModified       As Boolean

lngPosHeader = InStr(1, strMsg, vbCrLf & vbCrLf)
If lngPosHeader > 0 Then
    'Extract the header, faster than work with the entire msg
    strHeader = Left$(strMsg, lngPosHeader)
    lngPos = InStr(1, strHeader, "?Q", vbTextCompare)
    While lngPos > 0
        If Not Mid$(strHeader, lngPos + 3, 3) = "=22" Then
            lngPosEnd = InStr(lngPos + 3, strHeader, "?=", vbTextCompare)
            If InStr(1, Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3), "=2C", vbTextCompare) > 0 Then
                strHeader = Left$(strHeader, lngPos + 2) & "=22" & Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3) & "=22" & Mid$(strHeader, lngPosEnd)
                bModified = True
            End If
        End If
        lngPos = InStr(lngPos + 1, strHeader, "?Q", vbTextCompare)
    Wend
    If bModified Then
        strMsg = strHeader & Mid$(strMsg, lngPosHeader)
    End If
End If

End Sub



''
'Parse the mail attachments to fix image files mark as application/octet-stream.
'
'@param strMsg Email message in RFC 821 format
Private Sub ParseAtachment(ByRef strMsg As String)

Dim lngPos As Long
Dim lngPosEnd As Long
Dim strBoundary As String
Dim lngContentTypeStart As Long
    
lngPos = InStr(1, strMsg, "boundary=""", vbTextCompare)
If lngPos > 0 Then

    lngPosEnd = InStr(lngPos + 10, strMsg, """", vbTextCompare)
    'Get the Boundary string to identify the attachments
    strBoundary = Mid$(strMsg, lngPos + 10, lngPosEnd - lngPos - 10)

    lngPos = InStr(lngPosEnd, strMsg, strBoundary, vbTextCompare)
    While lngPos > 0
        lngContentTypeStart = InStr(lngPos, strMsg, "application/octet-stream", vbTextCompare)
        If lngContentTypeStart > 0 Then
            lngPos = lngContentTypeStart
            lngPosEnd = InStr(lngContentTypeStart, strMsg, vbCrLf & vbCrLf, vbTextCompare)
            If lngPosEnd > 0 And Mid$(strMsg, lngPosEnd + 4, 4) = "/9j/" Then
                'JPEG file
                strMsg = Left$(strMsg, lngContentTypeStart - 1) & "image/jpeg" & Mid$(strMsg, lngContentTypeStart + 24)
            End If
            If lngPosEnd > 0 And Mid$(strMsg, lngPosEnd + 4, 4) = "R0lG" Then
                'GIF file
                strMsg = Left$(strMsg, lngContentTypeStart - 1) & "image/gif" & Mid$(strMsg, lngContentTypeStart + 24)
            End If

        End If
        lngPos = InStr(lngPos + 1, strMsg, strBoundary, vbTextCompare)
    Wend

End If
End Sub

