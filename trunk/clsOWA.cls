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

Private vTypes(3, 1)    As String


Public objDOMInbox As DOMDocument
Public intMsgCount As Integer

''
'Retrieve message in RFC 821 format and send it to the client.
'
'@param strUrl  Url of the .EML file
'@param objPOP3 Socket object to send data
Public Sub GetMsg(ByVal strURL As String, ByRef objPOP3 As CSocketMaster)
On Error GoTo ErrHandler

'Dim oXMLHTTP    As XMLHTTP
Dim oXMLHTTP    As ServerXMLHTTP
Dim lngTimer    As Long
Dim strMsg      As String
Const ERRCOMMAND As String = "-ERR no such message"
'Set oXMLHTTP = New XMLHTTP
Set oXMLHTTP = New ServerXMLHTTP

With oXMLHTTP
    .setOption 3, ""
    .open "GET", strURL & "?ID=" & (Timer * 100), True, Config.Profile.strUser, Config.Profile.strPassWord
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
    WriteLog "Download Status " & .Status & vbTab & strURL, Information
    If (.Status >= 200 And .Status < 300) Then
        strMsg = StrConv(.responseBody, vbUnicode)
        'Fix errors in inline images
        ParseAtachment strMsg
        'Fix OWA errors in recived messages
        If InStr(strMsg, "X-MimeOLE: Produced By Microsoft") <> 0 Then
            ParseHeader strMsg
        End If
        objPOP3.SendData strMsg
    Else
        WriteLog "GetMsg.Status " & .statusText, Warning
        objPOP3.SendData ERRCOMMAND

    End If
End With
Set oXMLHTTP = Nothing
Exit Sub
ErrHandler:
    WriteLog "GetMsg " & Err.Description, Fail
    objPOP3.SendData ERRCOMMAND
End Sub

''
'Delete a message
'
'@param intMsg  Position of the message in the inbox folder
Public Function Delete(ByVal intMsg As Integer) As Boolean
On Error GoTo ErrHandler
'Dim oXMLHTTP    As XMLHTTP
Dim oXMLHTTP    As ServerXMLHTTP
    
    'Set oXMLHTTP = New XMLHTTP
    Set oXMLHTTP = New ServerXMLHTTP
    With oXMLHTTP
        .setOption 3, ""
        .open "DELETE", objDOMInbox.childNodes(1).childNodes(intMsg - 1).selectSingleNode("a:href").Text, True, Config.Profile.strUser, Config.Profile.strPassWord
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
'Dim oXMLHTTP As XMLHTTP
Dim oXMLHTTP As ServerXMLHTTP


'Set oXMLHTTP = New XMLHTTP
Set oXMLHTTP = New ServerXMLHTTP
With oXMLHTTP
    .setOption 3, ""
    ' Open a request to a particular url
    .open "SEARCH", Config.Profile.URLInbox, True, Config.Profile.strUser, Config.Profile.strPassWord
    ' Set the header type - we want XML
    .setRequestHeader "Content-type", "text/xml"
    .setRequestHeader "Depth", "1"
    If strCookies <> Empty Then
        .setRequestHeader "Cookies", "Necessary according to Q234486"
        .setRequestHeader "Cookies", strCookies
    End If
    ' Send the request
    MsgWaitObj 100
    WriteLog "Getting msg list ", Information
    Call .send("<?xml version='1.0' ?>" & _
                "<a:searchrequest xmlns:a='DAV:'><a:sql>" & _
                 "SELECT" & _
                 " ""DAV:id""" & _
                 ",""DAV:getcontentlength""" & _
                 " FROM scope('shallow traversal of """ & Config.Profile.URLInbox & """')" & _
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
'Urls are stored in config object
'@return WebDav status code, -1 with unespected error
'@remarks <A HREF='http://msdn2.microsoft.com/en-us/library/ms992623.aspx' target='_blank'>Microsoft Exchange 2000 SDK June 2005</A>
'@see Config

Public Function GetFoldersURL() As Long

On Error GoTo ErrHandler
Dim query           As String
Dim bFBALoginFail   As Boolean
'Dim oXMLHTTP        As XMLHTTP
Dim oXMLHTTP        As ServerXMLHTTP
Dim strAccount      As String
Dim objDOMFolders   As DOMDocument

Select Case Config.Profile.Authentication
Case basic
    'Do nothing, is the standard type
Case fba
    WriteLog "GetFoldersURL : Trying LogingFBA ", Information
    bFBALoginFail = Not (LogingFBA)
'Case NTLM
'    bFBALoginFail = Not (LogingNTLM)
End Select


If bFBALoginFail Then
    GetFoldersURL = -1
Else
    Set oXMLHTTP = New ServerXMLHTTP
    'strURL = strExchSvrName & "/exchange/" & strAccount & "/"

    With oXMLHTTP
        WriteLog "Getting Folders : " & Config.Profile.URLUser
        .setOption 3, ""
        .open "PROPFIND", Config.Profile.URLUser, True, Config.Profile.strUser, Config.Profile.strPassWord
        .setRequestHeader "Content-Type", "text/xml"
        .setRequestHeader "Depth", "0"
        .setRequestHeader "Accept-Encoding", "gzip,deflate"
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
            Config.Profile.URLSend = objDOMFolders.selectSingleNode(XMLPATH & "d:sendmsg").Text
            Config.Profile.URLInbox = objDOMFolders.selectSingleNode(XMLPATH & "d:inbox").Text & "/"
            Set objDOMFolders = Nothing
        Else
            WriteLog "GetFoldersURL: " & .Status & vbTab & .statusText, Warning
        End If
        GetFoldersURL = .Status
    End With
End If
Set oXMLHTTP = Nothing
Exit Function
ErrHandler:
    WriteLog "GetFoldersURL: " & Err.Number & vbTab & Err.Description, Fail
    GetFoldersURL = -1
    Debug.Assert False
End Function

''
'Access the Exchange store via WebDAV with Form-Based-Authentication turned on.
'
'@return Return True if the login is correct.
Private Function LogingFBA() As Boolean
On Error GoTo ErrHandler
Dim vHeaders()  As String
Dim intCounter  As Integer
Dim intIndex    As Integer
'Dim oXMLHTTP    As XMLHTTP
Dim oXMLHTTP    As ServerXMLHTTP
Dim strUsername As String
Dim strHeaders  As String
Dim strVersion  As String


If InStr(Config.Profile.strUser, "@") = 0 Then
    strUsername = Config.Profile.strUser
Else
    strUsername = Left(strUsername, InStr(Config.Profile.strUser, "@") - 1)
End If

WriteLog "LogingFBA -> Config.AuthPage: " & Config.Profile.AuthPage, Information

Set oXMLHTTP = New ServerXMLHTTP
With oXMLHTTP
    .setOption 3, ""
    .open "POST", Config.Profile.AuthPage, True
    .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    'Generate the body for FBA login
    .send Config.Profile.HiddenField & "=" & Config.Profile.HiddenValue & "&trusted=0&username=" & URLEncode(strUsername) & "&password=" & URLEncode(Config.Profile.strPassWord)
    While Not .readyState = 4
        MsgWaitObj 100
    Wend
    If (.Status >= 200 And .Status < 300) Then
        strHeaders = .getAllResponseHeaders()
        WriteLog "LogingFBA -> Headers: " & vbCrLf & strHeaders, Paranoid
        vHeaders = Split(strHeaders, vbNewLine)
        For intIndex = 0 To UBound(vHeaders) - 1
            If Left$(vHeaders(intIndex), 10) = "Set-Cookie" Then
                intCounter = intCounter + 1
                strCookies = strCookies & Mid$(vHeaders(intIndex), 12) & ";"
            ElseIf Left$(vHeaders(intIndex), 13) = "X-OWA-Version" Then
                strVersion = Mid$(vHeaders(intIndex), 16)
            End If
        Next
        If intCounter > 1 Then
            LogingFBA = True
        Else
            Select Case True
            Case strVersion = Empty
                WriteLog "LogingFBA -> X-OWA-Version headder missing", Paranoid
                LogingFBA = True
            Case Int(Left$(strVersion, InStr(strVersion, "."))) >= 8
                WriteLog "Error LogingFBA: " & strVersion, Paranoid
                LogingFBA = True
            Case Else
                WriteLog "LogingFBA -> Error getting headers: " & vbCrLf & strHeaders, Error
                LogingFBA = False
            End Select
        End If
    Else
       WriteLog "LogingFBA: " & .Status & vbTab & .statusText, Warning
       LogingFBA = False
    End If
End With
Set oXMLHTTP = Nothing

Exit Function
ErrHandler:
    WriteLog "GetFoldersURL: " & Err.Number & vbTab & Err.Description, Fail
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
'Dim oXMLHTTP As XMLHTTP
Dim oXMLHTTP As ServerXMLHTTP

    'Set oXMLHTTP = New XMLHTTP
    Set oXMLHTTP = New ServerXMLHTTP
     
    'Create the DAV PUT request.
    With oXMLHTTP
        .setOption 3, ""
        .open "PUT", Config.Profile.URLSend, True, Config.Profile.strUser, Config.Profile.strPassWord
        If strCookies <> Empty Then
            .setRequestHeader "Cookies", "Necessary according to Q234486"
            .setRequestHeader "Cookies", strCookies
        End If
        
        .setRequestHeader "Content-Type", "message/rfc821"
        If Not Config.Profile.bSaveinsent Then
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
    WriteLog "SendMail: " & Err.Number & vbTab & Err.Description, Fail
    Set oXMLHTTP = Nothing
    SendMail = False
End Function

''
'Parse the mail header to fix incorrect email adress with non US chars.
'For example:<BR>
'=?iso-8859-1?Q?   Martin_Sacrist=E1n=2C_C=2E   ?= <XXXX>
'Should be:<BR>
'=?iso-8859-1?Q?=22Martin_Sacrist=E1n=2C_C=2E=22?= <XXXX>
'Quotes are missing and some email clientes fails.
'@param strMsg Email message in RFC 821 format
Private Sub ParseHeader(ByRef strMsg As String)
Dim strHeader       As String
Dim strQuotedText   As String
Dim lngPos          As Long
Dim lngPosHeader    As Long
Dim lngPosEnd       As Long
Dim bModified       As Boolean
Const STARTQUOTED As String = "?q"

lngPosHeader = InStr(1, strMsg, vbCrLf & vbCrLf)
If lngPosHeader <> 0 Then
    'Extract the header, faster than work with the entire msg
    strHeader = Left$(strMsg, lngPosHeader)
    'Look for quoted string
    lngPos = InStr(1, LCase(strHeader), STARTQUOTED)
    While lngPos <> 0
        'Look for end quoted string
        lngPosEnd = InStr(lngPos + 3, strHeader, "?=")
        strQuotedText = Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3)
        If InStr(1, LCase(strQuotedText), "=2c") <> 0 Then
            If Not Left$(strQuotedText, 3) = "=22" Then
                strHeader = Left$(strHeader, lngPos + 2) & "=22" & strQuotedText & "=22" & Mid$(strHeader, lngPosEnd)
                bModified = True
            End If
        End If
        Debug.Assert InStr(lngPos + 1, LCase(strHeader), STARTQUOTED) = InStr(lngPosEnd + 1, LCase(strHeader), STARTQUOTED)
        lngPos = InStr(lngPosEnd + 1, LCase(strHeader), STARTQUOTED)
        
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

Dim lngPos              As Long
Dim lngPosEnd           As Long
Dim strBoundary         As String
Dim lngBoundary         As Long
Dim lngContentTypeStart As Long
Dim intType             As Integer

lngPos = InStr(1, LCase(strMsg), "boundary=""")
    
While lngPos <> 0
    lngPosEnd = InStr(lngPos + 10, strMsg, """")
    'Get the Boundary string to identify the attachments
    strBoundary = LCase(Mid$(strMsg, lngPos + 10, lngPosEnd - lngPos - 10))
    lngBoundary = InStr(lngPosEnd, LCase(strMsg), strBoundary)
    lngPos = lngBoundary
    While lngPos <> 0
        lngContentTypeStart = InStr(lngPos, LCase(strMsg), "application/octet-stream")
        If lngContentTypeStart > 0 Then
            lngPos = lngContentTypeStart
            lngPosEnd = InStr(lngContentTypeStart, strMsg, vbCrLf & vbCrLf)
            If lngPosEnd <> 0 Then
                For intType = 0 To UBound(vTypes, 1)
                    If Mid$(strMsg, lngPosEnd + 4, Len(vTypes(intType, 1))) = vTypes(intType, 1) Then
                        strMsg = Left$(strMsg, lngContentTypeStart - 1) & vTypes(intType, 0) & Mid$(strMsg, lngContentTypeStart + 24)
                        Exit For
                    End If
                Next
            End If

        End If
        lngPos = InStr(lngPos + 1, LCase(strMsg), strBoundary)
    Wend
    lngPos = InStr(lngBoundary + 1, LCase(strMsg), "boundary=""")
Wend
End Sub

Private Sub Class_Initialize()

vTypes(0, 0) = "image/jpeg"
vTypes(0, 1) = "/9j/"
vTypes(1, 0) = "image/gif"
vTypes(1, 1) = "R0lG"
vTypes(2, 0) = "image/bmp"
vTypes(2, 1) = "Qk"
vTypes(3, 0) = "image/png"
vTypes(3, 1) = "iVBORw0K"

End Sub

Private Sub Class_Terminate()
Set objDOMInbox = Nothing
Erase vTypes
End Sub

''
'Encode one string to send it as url parameter.
'
'@param Text String to encode
'@return Encoded string
Function URLEncode(ByVal Text As String) As String
Dim i As Integer
Dim acode As Integer

URLEncode = Text

For i = Len(URLEncode) To 1 Step -1
    acode = Asc(Mid$(URLEncode, i, 1))
    Select Case acode
    Case 48 To 57, 65 To 90, 97 To 122
        ' don't touch alphanumeric chars
    Case 32
        ' replace space with "+"
        Mid$(URLEncode, i, 1) = "+"
    Case Else
        ' replace punctuation chars with "%hex"
        URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$(URLEncode, i + 1)
    End Select
Next
    
End Function

