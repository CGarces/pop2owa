Attribute VB_Name = "modGlobal"
Public strUser          As String
Public strPassWord      As String
Public objDOMMsg        As DOMDocument
Public objDOMInbox      As DOMDocument
Public strExchSvrName  As String
Public bSaveinsent As Boolean

Public Const XMLPATH As String = "a:multistatus/a:response/a:propstat/a:prop/"


