VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================
' Class:    cRegistry
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'
' A nice class wrapper around the registry functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Updated 29 April 1998 for VB5.
'   * Fixed GPF in EnumerateValues
'   * Added support for all registry types, not just strings
'   * Put all declares in local class
'   * Added VB5 Enums
'   * Added CreateKey and DeleteKey methods
'
' Updated 2 January 1999
'   * The CreateExeAssociation method failed to set up the
'     association correctly if the optional document icon
'     was not provided.
'   * Added new parameters to CreateExeAssociation to set up
'     other standard handlers: Print, Add, New
'   * Provided the CreateAdditionalEXEAssociations method
'     to allow non-standard menu items to be added (for example,
'     right click on a .VBP file.  VB installs Run and Make
'     menu items).
'
' Updated 8 February 2000
'   * Ensure CreateExeAssociation and related items sets up the
'     registry keys in the
'           HKEY_LOCAL_MACHINE\SOFTWARE\Classes
'     branch as well as the HKEY_CLASSES_ROOT branch.
'
' ---------------------------------------------------------------------------
' vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' =========================================================

'Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const ERROR_NO_MORE_ITEMS = 259


'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long
   

' Other declares:
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)


Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Private m_hClassKey     As Long
Private m_sSectionKey   As String
Private m_sValueKey     As String
Private m_vValue        As Variant
Private m_vDefault      As Variant
Private m_eValueType    As ERegistryValueTypes


Public Property Get Value() As Variant
Dim vValue As Variant
Dim cData As Long, sData As String, ordType As Long, e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
    'ApiRaiseIf e

    e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        Value = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, iData, cData)
        vValue = CLng(iData)
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, dwData, cData)
        vValue = SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = ExpandEnvStr(sData)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hKey, m_sValueKey, 0&, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    Value = vValue
    
End Property
Public Property Let Value( _
        ByVal vValue As Variant _
    )
Dim ordType As Long
Dim c       As Long
Dim hKey    As Long
Dim e       As Long
Dim lCreate As Long
Dim tSA As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (varType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) - 1
                e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (varType(vValue) = vbInteger) Or (varType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = REG_DWORD
                e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
        
        'Close the key
        RegCloseKey hKey
    
    End If
    
End Property

Public Property Get ValueType() As ERegistryValueTypes
    ValueType = m_eValueType
End Property
Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property
Public Property Get ClassKey() As ERegistryClassConstants
    ClassKey = m_hClassKey
End Property
Public Property Let ClassKey( _
        ByVal eKey As ERegistryClassConstants _
    )
    m_hClassKey = eKey
End Property
Public Property Get SectionKey() As String
    SectionKey = m_sSectionKey
End Property
Public Property Let SectionKey( _
        ByVal sSectionKey As String _
    )
    m_sSectionKey = sSectionKey
End Property
Public Property Get ValueKey() As String
    ValueKey = m_sValueKey
End Property
Public Property Let ValueKey( _
        ByVal sValueKey As String _
    )
    m_sValueKey = sValueKey
End Property
Public Property Get Default() As Variant
    Default = m_vDefault
End Property
Public Property Let Default( _
        ByVal vDefault As Variant _
    )
    m_vDefault = vDefault
End Property
Private Function SwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function
Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    ' Get the length
    s = "" ' Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sData, s, c)
    ' Expand the string
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function
