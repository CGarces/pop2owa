Attribute VB_Name = "modGlobal"
''
'Store al public variables, constants and functions.

Option Explicit
Public strUser          As String
Public strPassWord      As String
Public objDOMMsg        As DOMDocument
Public objDOMInbox      As DOMDocument
Public strExchSvrName   As String

Public bSaveinsent      As Boolean
Public bAuthentication  As Boolean
Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"


''
'Write a pop2owa.err file with the logs of erros.
'
'@param strText Test to write in the log
Public Sub WriteLog(ByVal strText As String)
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\pop2owa.err" For Append As #intFile
    Print #intFile, Now & vbTab & strText
Close #intFile
End Sub
