VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
'Encapsulate Msxml2 objects
'
'@author Carlos B

Option Explicit
''
'Msxml2 object
Private objXMLRequest As Object
''
'True if ServerXMLHTTP is used
Private bServerObject As Boolean
''
'Cookies used for authentication
Private strCookies  As String


''
'Open one WebDav connection.
'
'@param bstrMethod  Method used in open command
'@param bstrUrl     URL used in open command
'@param bAuth       User and password from configuration
'@return            WebDav object
Public Sub OpenXML(ByVal bstrMethod As String, ByVal bstrUrl As String, Optional ByVal bAuth As Boolean = True)


If Not Config.Profile.strProxyUser = Empty Then
    Set objXMLRequest = CreateObject("MSXML2.ServerXMLHTTP.4.0")
    bServerObject = True
ElseIf Config.Profile.bForceXMLHTTP Then
    Set objXMLRequest = CreateObject("MSXML2.ServerXMLHTTP")
    bServerObject = True
Else
    Set objXMLRequest = CreateObject("MSXML2.XMLHTTP")
End If
If bServerObject Then
    objXMLRequest.setOption 0, 65001
'    objXMLRequest.setOption 2, ""
    objXMLRequest.setOption 3, ""
End If

If bAuth Then
    objXMLRequest.open bstrMethod, bstrUrl, True, Config.Profile.strUser, Config.Profile.strPassWord
Else
    objXMLRequest.open bstrMethod, bstrUrl, True
End If
objXMLRequest.setRequestHeader "Accept-Encoding", "gzip,deflate"
'objXMLRequest.setRequestHeader "Man", bstrMethod & " " & bstrUrl & " HTTP/1.1"
If LenB(strCookies) <> 0 Then
    objXMLRequest.setRequestHeader "Cookie", "Necessary according to Q234486"
    objXMLRequest.setRequestHeader "Cookie", strCookies
End If
With Config.Profile
    If .strProxyUser <> Empty Then
        objXMLRequest.setProxyCredentials .strProxyUser, .strProxyPassword
    End If
End With
End Sub

''
'Wraper to MSXML setRequestHeader function.
'
'@param bstrHeader  Header to set
'@param bstrValue   Value to set in the header
Public Sub setRequestHeader(ByVal bstrHeader As String, ByVal bstrValue As String)
    objXMLRequest.setRequestHeader bstrHeader, bstrValue
End Sub

''
'Wraper to MSXML send function.
'
'@param strBody Message to send
Public Sub send(Optional ByVal strBody As Variant = Empty)
    objXMLRequest.send strBody
End Sub
''
'Wraper to MSXML readyState function.
'
'@return    Connection state
Public Function readyState() As Long
    readyState = objXMLRequest.readyState
End Function

''
'Wraper to MSXML Status function.
'
'@return    Connection Status
Public Function Status() As Long
    Status = objXMLRequest.Status
End Function

''
'Wraper to MSXML responseBody function.
'
'@return    MSXML response body
Public Function responseBody() As Variant
    responseBody = objXMLRequest.responseBody
End Function

''
'Wraper to MSXML statusText function.
'
'@return    Connection Status described
Public Function statusText() As String
    statusText = objXMLRequest.statusText
End Function

''
'Wraper to MSXML responseXML function.
'
'@return    MSXML response in XML format
Public Function responseXML() As Object
    Set responseXML = objXMLRequest.responseXML
End Function

''
'Wraper to MSXML getAllResponseHeaders function.
'
'@return    Message headers
Public Function getAllResponseHeaders() As String
    getAllResponseHeaders = objXMLRequest.getAllResponseHeaders
End Function

'TODO check if (objXMLRequest.readyState = 4) and waitForResponse do the same
''
'Wraper to MSXML waitForResponse function using ServerXMLHTTP.</BR>
'If pop2owa is configured to use XMLHTTP, emulate waitForResponse
'using While ... Wend
'
'@param lngTimeout  Time (in seconds) to wait of the response
'
'@return    True when receive a valid response.
Public Function waitForResponse(ByVal lngTimeout As Long) As Boolean
Dim lngTimer As Long
    If bServerObject Then
        waitForResponse = objXMLRequest.waitForResponse(lngTimeout)
    Else
        lngTimer = Timer
        While Not ((objXMLRequest.readyState = 4) Or lngTimeout < (Timer - lngTimer))
            MsgWaitObj 1000
        Wend
        waitForResponse = (objXMLRequest.readyState = 4)
    End If
End Function

''
'Wait until readyState is equal to 4. Otherwise should fail by MSXML timeout
Public Sub waitUntilReady()
On Error GoTo ErrHandler
While Not objXMLRequest.readyState = 4
    If bServerObject Then
        objXMLRequest.waitForResponse 30
    Else
        MsgWaitObj 1000
    End If
Wend
Exit Sub
ErrHandler:
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''
'Get Cookies from the response in Form Based Autentication mode.
Public Sub GetCookies()
Dim vHeaders()  As String
Dim intIndex    As Integer
strCookies = vbNullString
vHeaders = Split(objXMLRequest.getAllResponseHeaders(), vbNewLine)
For intIndex = 0 To UBound(vHeaders) - 1
    If Left$(vHeaders(intIndex), 10) = "Set-Cookie" Then
        strCookies = strCookies & Mid$(vHeaders(intIndex), 12) & ";"
    End If
Next

End Sub
''
'Write one warning with the status of MSXML message.
'@param strSource Source of the warning to identify the problem.
Public Sub WriteWarning(ByVal strSource As String)
    WriteLog strSource & ": " & objXMLRequest.Status & vbTab & objXMLRequest.statusText, Warning
End Sub

''
'Clean the object used by the object
Private Sub Class_Terminate()
Set objXMLRequest = Nothing
End Sub
