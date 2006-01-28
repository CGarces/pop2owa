Attribute VB_Name = "modGlobal"
Public strUser          As String
Public strPassWord      As String
Public objDOMMsg        As DOMDocument
Public objDOMInbox      As DOMDocument
Public strExchSvrName  As String

Public bSaveinsent      As Boolean
Public bAuthentication  As Boolean
Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"


Public Function WriteLog(ByVal strText As String)
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\pop2owa.err" For Append As #intFile
    Print #intFile, Now & vbTab & strText
Close #intFile
End Function
