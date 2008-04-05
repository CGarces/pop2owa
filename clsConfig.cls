VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strIP            As String
Public intPOP3Port      As Integer
Public intSMTPPort      As Integer
Public bSMTPPort        As Boolean
Public strActiveProfile As String
Private vProfiles()     As String

Private colProfiles  As Collection


Public Property Get Profile() As clsProfile
Set Profile = colProfiles(strActiveProfile)
End Property

Public Property Get Profiles() As Variant
    Profiles = vProfiles
End Property

Public Sub ReadConfig()
On Error GoTo ErrHandler
Dim objDOMConfig    As DOMDocument
Dim objNodeElement  As IXMLDOMElement
Dim objNodeProfile  As IXMLDOMElement
Dim objNodeList     As IXMLDOMNodeList
Dim objProfileList  As IXMLDOMNodeList
Dim strProfile      As String
Dim objProfile      As clsProfile
Dim intCont         As Integer

Set objDOMConfig = New DOMDocument
Set colProfiles = New Collection
objDOMConfig.async = False
objDOMConfig.Load App.Path & "\config.xml"
If objDOMConfig.parseError.errorCode <> 0 Then
    Err.Raise 51, "clsOwa", "Error parsing config file"
End If
    
Set objNodeElement = objDOMConfig.selectSingleNode("xml/pop2owa")
strIP = objNodeElement.getAttribute("ip")
intPOP3Port = objNodeElement.getAttribute("pop3")
bSMTPPort = (objNodeElement.getAttribute("smtp_enabled") = "yes")
If bSMTPPort Then
    intSMTPPort = objNodeElement.getAttribute("smtp")
End If
Set objProfileList = objDOMConfig.selectNodes("xml/profile")
For Each objNodeProfile In objProfileList
    strProfile = objNodeProfile.getAttribute("name")
    If strProfile <> "template" Then
        ReDim Preserve vProfiles(0 To intCont)
        vProfiles(intCont) = strProfile
        intCont = intCont + 1
        Set objProfile = New clsProfile
        objProfile.strExchSvrName = objNodeProfile.getElementsByTagName("server").Item(0).Text
        If Right$(objProfile.strExchSvrName, 1) <> "/" Then
            objProfile.strExchSvrName = objProfile.strExchSvrName & "/"
        End If
        Set objNodeElement = objDOMConfig.selectSingleNode("xml/profile[@name='" & strProfile & "']/options")
        objProfile.bSaveinsent = (objNodeElement.getAttribute("savesent") = "yes")
        objProfile.Authentication = objNodeElement.getAttribute("auth")
        objProfile.bForceXMLHTTP = (objNodeElement.getAttribute("ForceNewVersion") = "yes")

        #If test = 1 Then
            objProfile.strUser = objNodeElement.getAttribute("user")
            objProfile.strPassWord = objNodeElement.getAttribute("password")
        #End If
        objProfile.AuthPage = objProfile.strExchSvrName & objDOMConfig.selectSingleNode("/xml/profile[@name='" & strProfile & "']/Form_Based_Auth/AuthPage").Text
        
        Set objNodeElement = objDOMConfig.selectSingleNode("xml/profile[@name='" & strProfile & "']/Form_Based_Auth/Main_Hidden_Field")
        objProfile.HiddenField = objNodeElement.getAttribute("Name")
        objProfile.HiddenValue = objNodeElement.getAttribute("Value")
        Set objNodeList = objDOMConfig.selectNodes("xml/profile[@name='" & strProfile & "']/Form_Based_Auth/Hidden_Field")
        For Each objNodeElement In objNodeList
            objProfile.HiddenValue = objProfile.HiddenValue & "&" & objNodeElement.getAttribute("Name") & "=" & objNodeElement.getAttribute("Value")
        Next
        objProfile.strProxyUser = objDOMConfig.selectSingleNode("xml/profile[@name='" & strProfile & "']/proxy/@user").Text
        objProfile.strProxyPassword = objDOMConfig.selectSingleNode("xml/profile[@name='" & strProfile & "']/proxy/@password").Text
        
        colProfiles.Add objProfile, strProfile
    End If
Next

Set objProfile = Nothing
Set objNodeElement = Nothing
Set objNodeProfile = Nothing
Set objNodeList = Nothing
Set objProfileList = Nothing

Set objDOMConfig = Nothing
Exit Sub
ErrHandler:
    Err.Raise 51, "clsOwa", "Error parsing config file, " & Err.Description
End Sub

Public Sub WriteConfig()

Dim objDOMConfig    As DOMDocument
Dim objNodeElement  As IXMLDOMElement

Set objDOMConfig = New DOMDocument
objDOMConfig.async = False
objDOMConfig.Load App.Path & "\config.xml"
If objDOMConfig.parseError.errorCode <> 0 Then
    Err.Raise 51, "clsOwa", "Error parsing config file"
End If
    
Set objNodeElement = objDOMConfig.selectSingleNode("xml/pop2owa")
objNodeElement.setAttribute "ip", strIP
objNodeElement.setAttribute "pop3", intPOP3Port
objNodeElement.setAttribute "smtp_enabled", IIf(bSMTPPort, "yes", "no")
If bSMTPPort Then
    objNodeElement.setAttribute "smtp", intSMTPPort
End If
With Me.Profile
    Set objNodeElement = objDOMConfig.selectSingleNode("xml/profile [@name='default']/server")
    objNodeElement.Text = .strExchSvrName
    Set objNodeElement = objDOMConfig.selectSingleNode("xml/profile [@name='default']/options")
    objNodeElement.setAttribute "savesent", IIf(.bSaveinsent, "yes", "no")
    objNodeElement.setAttribute "auth", .Authentication
    objDOMConfig.save App.Path & "\config.xml"
End With
Set objNodeElement = Nothing
Set objDOMConfig = Nothing

End Sub

Private Sub Class_Initialize()
strActiveProfile = "default"
End Sub

Private Sub Class_Terminate()
Erase vProfiles()
Set colProfiles = Nothing
End Sub
