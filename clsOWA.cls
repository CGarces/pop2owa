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

''
'Array for image types
Private vTypes(3, 1)    As String

''
'Dom object with inbox messages
Public objDOMInbox  As DOMDocument
''
'Number of messages in inbox folder
Public intMsgCount  As Integer
''
'Size of messages in inbox folder
Public lngSize As Long

Private oXMLHTTP    As clsXMLRequest

Const TWO_NEWLINES = vbNewLine & vbNewLine
''
'Retrieve message in RFC 821 format and send it to the client.
'
'@param strUrl  Url of the .EML file
'@param objPOP3 Socket object to send data
Public Sub GetMsg(ByVal strURL As String, ByRef objPOP3 As CSocketMaster)
On Error GoTo ErrHandler

Dim strMsg      As String
Const ERRCOMMAND As String = Error & "no such message"

With oXMLHTTP
    .OpenXML "GET", strURL & "?ID=" & (Timer * 100)
    .setRequestHeader "Content-type", "text/xml"
    .setRequestHeader "Depth", "infinity"
    .setRequestHeader "Translate", "f"
    WriteLog "Getting " & vbTab & strURL, Information
    .send
    While Not .readyState = 4
        If Not .waitForResponse(25) Then
            WriteLog "Downloading readyState  " & .readyState & vbTab & strURL, Warning
            objPOP3.SendData "X-POP2OWA-Header: Header to avoid timeout" & vbNewLine
        End If
    Wend
    WriteLog "Download Status " & .Status & vbTab & strURL, Information
    If (.Status >= 200 And .Status < 300) Then
        strMsg = StrConv(.responseBody, vbUnicode)
        'Fix errors in inline images
        ParseAtachment strMsg
        'Fix OWA errors in received messages
        If InStr(1, strMsg, "X-MimeOLE: Produced By Microsoft", vbTextCompare) <> 0 Then
            ParseHeader strMsg
        End If
        If InStr(strMsg, vbNewLine & "." & vbNewLine) > 0 Then
            strMsg = Replace(strMsg, vbNewLine & "." & vbNewLine, vbNewLine & ".." & vbNewLine)
        End If
        objPOP3.SendData strMsg
        objPOP3.SendData "."
        strMsg = vbNullString
    Else
        .WriteWarning "GetMsg.Status"
        objPOP3.SendData ERRCOMMAND
    End If
End With
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

With oXMLHTTP
    .OpenXML "DELETE", objDOMInbox.childNodes(1).childNodes(intMsg - 1).selectSingleNode("a:href").Text
    .setRequestHeader "Depth", "infinity"
    .send
    .waitUntilReady
End With
Delete = True
Exit Function
ErrHandler:
    Delete = False
End Function
''
'Get the list of messages in the inbox folder.
'
'@return Reply to e-mail client, true if success
Public Function GetMsgList() As Boolean
On Error GoTo ErrHandler

With oXMLHTTP
    .OpenXML "SEARCH", Config.Profile.URLInbox
    ' Set the header type - we want XML
    .setRequestHeader "Content-type", "text/xml"
    .setRequestHeader "Depth", "1"
    ' Send the request
    WriteLog "Getting msg list ", Information
    Call .send("<?xml version='1.0' ?>" & _
                "<a:searchrequest xmlns:a='DAV:'><a:sql>" & _
                 "SELECT" & _
                 " ""DAV:uid""" & _
                 ",""DAV:getcontentlength""" & _
                 " FROM scope('shallow traversal of """ & Config.Profile.URLInbox & """')" & _
                 " WHERE ""DAV:ishidden""=False" & _
                 " AND ""DAV:isfolder""=False" & _
                 "</a:sql></a:searchrequest>")
    .waitUntilReady
    If .Status = 207 Then
        Set objDOMInbox = .responseXML
        objDOMInbox.setProperty "SelectionLanguage", "XPath"
        objDOMInbox.setProperty "SelectionNamespaces", "xmlns:a='DAV:'"
        intMsgCount = objDOMInbox.childNodes(1).childNodes.Length
        WriteLog "msg list fetch " & intMsgCount & " messages", Information
    Else
        .WriteWarning "Receiving msg list"
        Err.Raise vbObjectError + 1, "Receiving msg", .statusText
    End If
End With

GetMsgList = True
Exit Function
ErrHandler:
    Debug.Assert False
    GetMsgList = False
End Function


''
'Get Well-Known Mailbox Folder URLs that are retrieved from a users root mailbox folder.
'Urls are stored in config object
'@return WebDav status code, -1 with unexpected error
'@remarks <A HREF='http://msdn2.microsoft.com/en-us/library/ms992623.aspx' target='_blank'>Microsoft Exchange 2000 SDK June 2005</A>
'@see Config

Public Function GetFoldersURL() As Long

On Error GoTo ErrHandler
Dim bFBALoginFail   As Boolean
Dim objDOMFolders   As DOMDocument

Select Case Config.Profile.Authentication
Case basic
    'Do nothing, is the standard type
Case fba
    WriteLog "GetFoldersURL : Trying LogingFBA ", Information
    bFBALoginFail = Not LogingFBA
#If test = 1 Then
Case NTLM
'    bFBALoginFail = Not (LogingNTLM)
#End If
End Select


If bFBALoginFail Then
    GetFoldersURL = -1
Else
    WriteLog "Getting Folders : " & Config.Profile.URLUser
    'strURL = strExchSvrName & "/exchange/" & strAccount & "/"

    With oXMLHTTP
        .OpenXML "PROPFIND", Config.Profile.URLUser
        .setRequestHeader "Content-Type", "text/xml"
        .setRequestHeader "Depth", "0"
        .send "<?xml version='1.0'?>" & _
            "<a:propfind xmlns:a='DAV:'>" & _
            "<a:prop xmlns:m='urn:schemas:httpmail:'>" & _
            "<m:sendmsg/>" & _
            "<m:inbox/>" & _
            "</a:prop>" & _
            "</a:propfind>"
        .waitUntilReady
        ' process the result
        If (.Status >= 200 And .Status < 300) Then
            Set objDOMFolders = .responseXML
            objDOMFolders.setProperty "SelectionLanguage", "XPath"
            objDOMFolders.setProperty "SelectionNamespaces", "xmlns:a='DAV:'"
            objDOMFolders.setProperty "SelectionNamespaces", "xmlns:a='DAV:' xmlns:d='urn:schemas:httpmail:'"
            
            Config.Profile.URLSend = objDOMFolders.selectSingleNode(XMLPATH & "d:sendmsg").Text
            Config.Profile.URLInbox = objDOMFolders.selectSingleNode(XMLPATH & "d:inbox").Text & "/"
            Set objDOMFolders = Nothing
        Else
            Debug.Assert False
            .WriteWarning "GetFoldersURL"
        End If
        GetFoldersURL = .Status
    End With
End If

Exit Function
ErrHandler:
    WriteLog "GetFoldersURL: " & Err.Number & vbTab & Err.Description, Fail
    GetFoldersURL = Err.Number
    Debug.Assert False
End Function

''
'Access the Exchange store via WebDAV with Form-Based-Authentication turned on.
'
'@return Return True if the login is correct.
Private Function LogingFBA() As Boolean
On Error GoTo ErrHandler

Dim strUsername As String

'If InStr(Config.Profile.strUser, "@") = 0 Then
    strUsername = Config.Profile.strUser
'Else
'    strUsername = Left$(Config.Profile.strUser, InStr(Config.Profile.strUser, "@") - 1)
'End If

WriteLog "LogingFBA -> Config.AuthPage: " & Config.Profile.AuthPage, Information

With oXMLHTTP
    .OpenXML "POST", Config.Profile.AuthPage, False
    .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    'Generate the body for FBA login
    .send Config.Profile.HiddenField & "=" & Config.Profile.HiddenValue & "&trusted=0&username=" & URLEncode(strUsername) & "&password=" & URLEncode(Config.Profile.strPassWord)
    WriteLog "LogingFBA -> Send info: " & Config.Profile.HiddenField & "=" & Config.Profile.HiddenValue & "&trusted=0&username=" & URLEncode(strUsername) & "&password=******", Information
    .waitUntilReady
    WriteLog "GetCookies -> Headers: " & vbNewLine & .getAllResponseHeaders(), Paranoid
    If (.Status >= 200 And .Status < 300) Then
        .GetCookies
        LogingFBA = True
    Else
        .WriteWarning "LogingFBA"
        LogingFBA = False
    End If
End With

Exit Function
ErrHandler:
    WriteLog "LogingFBA: " & Err.Number & vbTab & Err.Description, Fail
    LogingFBA = False
End Function

''
'Send one email in standard rfc821 format.
'
'@param strMailData Message string (rfc821)
'@return <B>True</B> if the mail is sent.
Public Function SendMail(ByVal strMailData As String) As Boolean

On Error GoTo ErrHandler
 
'Create the DAV PUT request.
With oXMLHTTP
    .OpenXML "PUT", Config.Profile.URLSend
    .setRequestHeader "Content-Type", "message/rfc821"
    If Not Config.Profile.bSaveinsent Then
        .setRequestHeader "Saveinsent", "f"
    End If
    .send strMailData
    .waitUntilReady
'Process the results.
    If (.Status >= 200 And .Status < 300) Then
        SendMail = True
        WriteLog "Mail send OK", Paranoid
    Else
        .WriteWarning "SendMail Fails!"
        SendMail = False
    End If
End With
Exit Function
ErrHandler:
    WriteLog "SendMail: " & Err.Number & vbTab & Err.Description, Fail
    SendMail = False
End Function

''
'Parse the mail header to fix incorrect email address with non US chars.
'For example:<BR>
'=?iso-8859-1?Q?   Martin_Sacrist=E1n=2C_C=2E   ?= <XXXX>
'Should be:<BR>
'=?iso-8859-1?Q?=22Martin_Sacrist=E1n=2C_C=2E=22?= <XXXX>
'Quotes are missing and some email clients fails parsing it.
'@param strMsg Email message in RFC 821 format
#If test = 1 Then
Public Sub ParseHeader(ByRef strMsg As String)
#Else
Private Sub ParseHeader(ByRef strMsg As String)
#End If
Dim strHeader       As String
Dim strQuotedText   As String
Dim lngPos          As Long
Dim lngPosHeader    As Long
Dim lngPosEnd       As Long
Dim bModified       As Boolean
Const START_ENCODED   As String = "?q"
Const ENCODED_QUOTE   As String = "=22"

lngPosHeader = InStr(1, strMsg, TWO_NEWLINES)
If lngPosHeader <> 0 Then
    'Extract the header, faster than work with the entire msg
    strHeader = Left$(strMsg, lngPosHeader - 1)
    'Look for quoted string
    lngPos = InStr(1, LCase$(strHeader), START_ENCODED)
    While lngPos <> 0
        'Look for end quoted string
        lngPosEnd = InStr(lngPos + 3, strHeader, "?=")
        strQuotedText = Mid$(strHeader, lngPos + 3, lngPosEnd - lngPos - 3)
        'Look for encoded "," that cause the problem
        If InStr(1, LCase$(strQuotedText), "=2c") <> 0 Then
            If Not Left$(strQuotedText, 3) = ENCODED_QUOTE Then
                strHeader = Left$(strHeader, lngPos + 2) & (ENCODED_QUOTE & strQuotedText & ENCODED_QUOTE) & Mid$(strHeader, lngPosEnd)
                bModified = True
            End If
        End If
        lngPos = InStr(lngPosEnd + 1, LCase$(strHeader), START_ENCODED)
    Wend
    If bModified Then
        strMsg = strHeader & Mid$(strMsg, lngPosHeader)
    End If
    strHeader = vbNullString
End If

End Sub

''
'Parse the mail attachments to fix image files mark as application/octet-stream.
'
'@param strMsg Email message in RFC 821 format
#If test = 1 Then
Public Sub ParseAtachment(ByRef strMsg As String)
#Else
Private Sub ParseAtachment(ByRef strMsg As String)
#End If
On Error GoTo ErrHandler

Dim lngPos              As Long
Dim lngPosEnd           As Long
Dim strBoundary         As String
Dim lngBoundary         As Long
Dim lngContentTypeStart As Long
Dim intType             As Integer
Dim strType             As String
Const BOUNDARY          As String = "boundary="""
Const STREAM          As String = "application/octet-stream"

lngBoundary = InStr(1, LCase$(strMsg), BOUNDARY)
    
While lngBoundary <> 0
    lngPos = lngBoundary + 10
    lngPosEnd = InStr(lngPos, strMsg, """")
    'Get the Boundary string to identify the attachments
    strBoundary = LCase$(Mid$(strMsg, lngPos, lngPosEnd - lngPos))
    lngPos = InStr(lngPosEnd, LCase$(strMsg), strBoundary)
    While lngPos <> 0
        'Find one header with application/octet-stream object
        lngContentTypeStart = InStr(lngPos, LCase$(strMsg), STREAM)
        If lngContentTypeStart > 0 Then
            lngPos = lngContentTypeStart
            lngPosEnd = InStr(lngContentTypeStart, strMsg, TWO_NEWLINES)
            If lngPosEnd <> 0 Then
                For intType = 0 To UBound(vTypes, 1)
                    strType = vTypes(intType, 1)
                    If Mid$(strMsg, lngPosEnd + 4, Len(strType)) = strType Then
                        strMsg = Left$(strMsg, lngContentTypeStart - 1) & vTypes(intType, 0) & Mid$(strMsg, lngContentTypeStart + 24)
                        Exit For
                    End If
                Next
            End If
        End If
        lngPos = InStr(lngPos + 1, LCase$(strMsg), STREAM)
    Wend
    lngBoundary = InStr(lngBoundary + 1, LCase$(strMsg), BOUNDARY)
Wend
Exit Sub
ErrHandler:
    Err.Raise Err.Number, "ParseAtachment->" & Err.Source, Err.Description
End Sub

Private Sub Class_Initialize()
Set oXMLHTTP = New clsXMLRequest

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
Set oXMLHTTP = Nothing
Erase vTypes
End Sub

''
'Encode one string to send it as url parameter.
'
'@param Text String to encode
'@return Encoded string
Function URLEncode(ByVal Text As String) As String
Dim i       As Integer
Dim acode   As Integer

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
        URLEncode = Left$(URLEncode, i - 1) & ("%" & Hex$(acode)) & Mid$(URLEncode, i + 1)
    End Select
Next
    
End Function
