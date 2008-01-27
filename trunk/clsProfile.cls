VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public URLSend      As String
Public URLInbox     As String
Public URLUser     As String

Public AuthPage     As String
Public HiddenField  As String
Public HiddenValue  As String
Public strAccount   As String
Public strPassWord  As String
Public strExchSvrName   As String
Public bSaveinsent      As Boolean
Public Authentication  As AuthTypes

Private m_strUser      As String


Public Enum AuthTypes
    basic = 0
    fba = 1
    NTLM = 2
End Enum

Property Get strUser() As String
    strUser = m_strUser
End Property

Property Let strUser(ByVal user As String)
    m_strUser = user
    If InStr(m_strUser, "^") <> 0 Then
        strAccount = Mid$(m_strUser, InStr(m_strUser, "^") + 1)
        m_strUser = Left$(m_strUser, InStr(m_strUser, "^") - 1)
    Else
        strAccount = m_strUser
    End If
    
    If Right(LCase(Me.strExchSvrName), Len("/exchange/")) = "/exchange/" Then
        Me.URLUser = Me.strExchSvrName & Me.strAccount & "/"
    Else
       Me.URLUser = Me.strExchSvrName & "exchange/" & Me.strAccount & "/"
    End If
End Property


