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
'WebDav code to comunicate with Exchange server
'
'@author Carlos B
'@version 1.0
'@date 26/11/2005

Option Explicit

Private objDOMFolders   As DOMDocument

Private strCookies As String

Public intMsgCount As Integer


''
'Get the URL of one inbox message
'
'@param intMsg  Position of the message in the inbox folder
Private Function GetURL(ByVal intMsg As Integer) As String
    GetURL = objDOMInbox.childNodes(1).childNodes(intMsg - 1).selectSingleNode("a:propstat/a:prop/a:href").Text
End Function



''
'Retrieve message in RFC 821 format and send it to the client.
'
'@param strUrl  Url of the .EML file
'@param lngSocket Socket to send data
'@return True if sucess.
Public Function GetMsg(ByVal strURL As String, ByRef objPOP3 As CSocketMaster) As Boolean
On Error GoTo ErrHandler
Dim oXMLHTTP    As XMLHTTP
Dim lngTimer    As Long
Dim strMsg      As String


Set oXMLHTTP = New XMLHTTP

With oXMLHTTP
    '.setOption 3, ""
    .open "GET", strURL, True, strUser, strPassWord
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
        'Fix OWA errors in recived messages
        If InStr(strMsg, "X-MimeOLE: Produced By Microsoft Exchange") > 0 Then
            strMsg = ParseMsg(strMsg)
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
    objPOP3.SendData "-ERR "
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
Dim strURL As String
Dim oXMLHTTP As XMLHTTP
strURL = objDOMFolders.selectSingleNode(XMLPATH & "d:inbox").Text & "/"
' Create an instance of the xmlHTTP object
Set oXMLHTTP = New XMLHTTP
With oXMLHTTP
    ' Open a request to a particular url
    .open "SEARCH", strURL, True, strUser, strPassWord
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
                 " ""DAV:href""" & _
                 ",""DAV:id""" & _
                 ",""DAV:getcontentlength""" & _
                 " FROM scope('shallow traversal of """ & strURL & """')" & _
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

If bAuthentication Then
    WriteLog "GetFoldersURL : Trying LogingFBA ", Information
    bFBALoginFail = Not (LogingFBA)
End If
If bFBALoginFail Then
    GetFoldersURL = -1
Else
    query = "<?xml version='1.0'?>"
    query = query + "<a:propfind xmlns:a='DAV:'>"
    query = query + "<a:prop xmlns:m='urn:schemas:httpmail:'>"
    query = query + "<m:sendmsg/>"
    query = query + "<m:inbox/>"
    query = query + "</a:prop>"
    query = query + "</a:propfind>"
     
    Set oXMLHTTP = New XMLHTTP
    strURL = strExchSvrName & "/exchange/" & strUser & "/"

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
        .send query
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
'@author Carlos B
'@version 1.0
'@date 20/01/2006
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
    WriteLog "LogingFBA -> Headers: " & vbCrLf & strHeaders, Paranoid
    vHeaders = Split(strHeaders, vbNewLine)
End With
Set oXMLHTTP = Nothing
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
    If Int(Left$(strVersion, InStr(strVersion, "."))) >= 8 Then
        WriteLog "Error LogingFBA: " & strVersion, Paranoid
        LogingFBA = True
    Else
        WriteLog "LogingFBA -> Error getting headers: " & vbCrLf & strHeaders, Error
        LogingFBA = False
    End If
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
Dim strSubURL As String
Dim xmlReq As XMLHTTP

    
    Set xmlReq = New XMLHTTP
    strSubURL = objDOMFolders.selectSingleNode(XMLPATH & "d:sendmsg").Text
     
     'Create the DAV PUT request.
    'xmlReq.setOption 3, ""
    xmlReq.open "PUT", strSubURL, True, strUser, strPassWord
    If strMailData <> "" Then
       xmlReq.setRequestHeader "Content-Type", "message/rfc821"
       If Not bSaveinsent Then
           xmlReq.setRequestHeader "Saveinsent", "f"
       End If
       xmlReq.send strMailData
        While Not xmlReq.readyState = 4
            MsgWaitObj 100
        Wend
    End If
    
    'Process the results.
    If (xmlReq.Status >= 200 And xmlReq.Status < 300) Then
       SendMail = True
    Else
       WriteLog "SendMail: " & xmlReq.Status & vbTab & xmlReq.statusText, Warning
       SendMail = False
    End If
    Set xmlReq = Nothing

Exit Function
ErrHandler:
    WriteLog "SendMail: " & Err.Number & vbTab & Err.Description, Error
    Set xmlReq = Nothing
    SendMail = False
End Function

Private Function ParseMsg(ByVal strMsg As String) As String
Dim strHeader As String
Dim lngPosFin As Long
lngPosFin = InStr(1, strMsg, vbCrLf & vbCrLf)
If lngPosFin > 0 Then
 '   intPosFin = InStr(intPosIni + 1, strMsg, vbCrLf & vbCrLf, vbTextCompare)
    strHeader = Left$(strMsg, lngPosFin)
    ParseHeader strHeader
    ParseMsg = strHeader & Mid$(strMsg, lngPosFin)
Else
    ParseMsg = strMsg
End If

End Function
Private Sub ParseHeader(ByRef strHeader As String)
Dim lngPos As Long
Dim lngPosEnd As Long
'=2C
lngPos = InStr(1, strHeader, "?Q", vbTextCompare)
While lngPos > 0
    If Not Mid$(strHeader, lngPos + 3, 3) = "=22" Then
        lngPosEnd = InStr(lngPos + 3, strHeader, "?=", vbTextCompare)
        If InStr(1, Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3), "=2C", vbTextCompare) > 0 Then
            strHeader = Left$(strHeader, lngPos + 2) & "=22" & Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3) & "=22" & Mid$(strHeader, lngPosEnd)
        End If
    End If
    lngPos = InStr(lngPosEnd + 1, strHeader, "?Q", vbTextCompare)
Wend
End Sub


