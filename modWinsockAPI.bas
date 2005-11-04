Attribute VB_Name = "modWinsockAPI"
Option Explicit

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const INVALID_SOCKET = -1


'/*
' * All Windows Sockets error constants are biased by WSABASEERR from
' * the "normal"
' */
Public Const WSABASEERR = 10000
'/*
' * Windows Sockets definitions of regular Microsoft C error constants
' */
Public Const WSAEINTR = (WSABASEERR + 4)
Public Const WSAEBADF = (WSABASEERR + 9)
Public Const WSAEACCES = (WSABASEERR + 13)
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEMFILE = (WSABASEERR + 24)

'/*
' * Windows Sockets definitions of regular Berkeley error constants
' */
Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Public Const WSAEINPROGRESS = (WSABASEERR + 36)
Public Const WSAEALREADY = (WSABASEERR + 37)
Public Const WSAENOTSOCK = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Public Const WSAEMSGSIZE = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Public Const WSAEADDRINUSE = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSAENETUNREACH = (WSABASEERR + 51)
Public Const WSAENETRESET = (WSABASEERR + 52)
Public Const WSAECONNABORTED = (WSABASEERR + 53)
Public Const WSAECONNRESET = (WSABASEERR + 54)
Public Const WSAENOBUFS = (WSABASEERR + 55)
Public Const WSAEISCONN = (WSABASEERR + 56)
Public Const WSAENOTCONN = (WSABASEERR + 57)
Public Const WSAESHUTDOWN = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)
Public Const WSAETIMEDOUT = (WSABASEERR + 60)
Public Const WSAECONNREFUSED = (WSABASEERR + 61)
Public Const WSAELOOP = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Public Const WSAENOTEMPTY = (WSABASEERR + 66)
Public Const WSAEPROCLIM = (WSABASEERR + 67)
Public Const WSAEUSERS = (WSABASEERR + 68)
Public Const WSAEDQUOT = (WSABASEERR + 69)
Public Const WSAESTALE = (WSABASEERR + 70)
Public Const WSAEREMOTE = (WSABASEERR + 71)

'/*
' * Extended Windows Sockets error constant definitions
' */
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAEDISCON = (WSABASEERR + 101)
Public Const WSAENOMORE = (WSABASEERR + 102)
Public Const WSAECANCELLED = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED = (WSABASEERR + 111)
Public Const WSAEREFUSED = (WSABASEERR + 112)

Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004

Public Const FD_SETSIZE = 64

Public Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Public Type servent
    s_name    As Long
    s_aliases As Long
    s_port    As Integer
    s_proto   As Long
End Type

Public Type protoent
    p_name    As String 'Official name of the protocol
    p_aliases As Long   'Null-terminated array of alternate names
    p_proto   As Long   'Protocol number, in host byte order
End Type

Public Type sockaddr_in
    sin_family       As Integer
    sin_port         As Integer
    sin_addr         As Long
    sin_zero(1 To 8) As Byte
End Type

Public Type timeval
  tv_sec  As Long   'seconds
  tv_usec As Long   'and microseconds
End Type

Public Type fd_set
  fd_count                  As Long '// how many are SET?
  fd_array(1 To FD_SETSIZE) As Long '// an array of SOCKETs
End Type

Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long

Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long

Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long

Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long

Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long

Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long

Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long

Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long

Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long

Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long

Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long

Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
    
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
    
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long

Public Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
  
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
                  
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
                  
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
                  
Public Declare Function vbselect Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As Long) As Long

Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
                  
Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
                  
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long

Public Declare Function Accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'/*
'* Address families.
'*/
Public Enum AddressFamily
    '
    AF_UNSPEC = 0      '/* unspecified */
'/*
' * Although  AF_UNSPEC  is  defined for backwards compatibility, using
' * AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
' * DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
' * depends  on the actual address family chosen.  As environments grow
' * to  include  more  and  more  address families that use overlapping
' * protocol  values  there  is  more  and  more  chance of choosing an
' * undesired address family when AF_UNSPEC is used.
' */
    AF_UNIX = 1        '/* local to host (pipes, portals) */
    AF_INET = 2        '/* internetwork: UDP, TCP, etc. */
    AF_IMPLINK = 3     '/* arpanet imp addresses */
    AF_PUP = 4         '/* pup protocols: e.g. BSP */
    AF_CHAOS = 5       '/* mit CHAOS protocols */
    AF_NS = 6          '/* XEROX NS protocols */
    AF_IPX = AF_NS     '/* IPX protocols: IPX, SPX, etc. */
    AF_ISO = 7         '/* ISO protocols */
    AF_OSI = AF_ISO    '/* OSI is ISO */
    AF_ECMA = 8        '/* european computer manufacturers */
    AF_DATAKIT = 9     '/* datakit protocols */
    AF_CCITT = 10      '/* CCITT protocols, X.25 etc */
    AF_SNA = 11        '/* IBM SNA */
    AF_DECnet = 12     '/* DECnet */
    AF_DLI = 13        '/* Direct data link interface */
    AF_LAT = 14        '/* LAT */
    AF_HYLINK = 15     '/* NSC Hyperchannel */
    AF_APPLETALK = 16  '/* AppleTalk */
    AF_NETBIOS = 17    '/* NetBios-style addresses */
    AF_VOICEVIEW = 18  '/* VoiceView */
    AF_FIREFOX = 19    '/* Protocols from Firefox */
    AF_UNKNOWN1 = 20   '/* Somebody is using this! */
    AF_BAN = 21        '/* Banyan */
    AF_ATM = 22        '/* Native ATM Services */
    AF_INET6 = 23      '/* Internetwork Version 6 */
    AF_CLUSTER = 24    '/* Microsoft Wolfpack */
    AF_12844 = 25      '/* IEEE 1284.4 WG AF */
    AF_MAX = 26
    '
End Enum
'
'Socket types
'
Public Enum SocketType
    SOCK_STREAM = 1    ' /* stream socket */
    SOCK_DGRAM = 2     ' /* datagram socket */
    SOCK_RAW = 3       ' /* raw-protocol interface */
    SOCK_RDM = 4       ' /* reliably-delivered message */
    SOCK_SEQPACKET = 5 ' /* sequenced packet stream */
End Enum

'/*
' * Protocols
' */
Public Enum SocketProtocol
    IPPROTO_IP = 0             '/* dummy for IP */
    IPPROTO_ICMP = 1           '/* control message protocol */
    IPPROTO_IGMP = 2           '/* internet group management protocol */
    IPPROTO_GGP = 3            '/* gateway^2 (deprecated) */
    IPPROTO_TCP = 6            '/* tcp */
    IPPROTO_PUP = 12           '/* pup */
    IPPROTO_UDP = 17           '/* user datagram protocol */
    IPPROTO_IDP = 22           '/* xns idp */
    IPPROTO_ND = 77            '/* UNOFFICIAL net disk proto */
    IPPROTO_RAW = 255          '/* raw IP packet */
    IPPROTO_MAX = 256
End Enum
'
'Maximum queue length specifiable by listen.
Public Const SOMAXCONN = &H7FFFFFFF
'
Public Enum WinsockVersion
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum
'
Public Enum IPEndPointFields
    LOCAL_HOST
    LOCAL_HOST_IP
    LOCAL_PORT
    REMOTE_HOST
    REMOTE_HOST_IP
    REMOTE_PORT
End Enum
'
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767


Public Function UnsignedToLong(Value As Double) As Long
    '
    'The function takes a Double containing a value in the 
    'range of an unsigned Long and returns a Long that you 
    'can pass to an API that requires an unsigned Long
    '
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
    '
End Function

Public Function LongToUnsigned(Value As Long) As Double
    '
    'The function takes an unsigned Long from an API and 
    'converts it to a Double for display or arithmetic purposes
    '
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
    '
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
    '
    'The function takes a Long containing a value in the range 
    'of an unsigned Integer and returns an Integer that you 
    'can pass to an API that requires an unsigned Integer
    '
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
    '
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    '
    'The function takes an unsigned Integer from and API and 
    'converts it to a Long for display or arithmetic purposes
    '
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
    '
End Function

Public Function StringFromPointer(ByVal lPointer As Long) As String
    '
    Dim strTemp As String
    Dim lRetVal As Long
    '
    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    '
    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    '
    'return a string
    If lRetVal Then StringFromPointer = strTemp
    '
End Function

Public Function GetAddressLong(ByVal strHostName As String) As Long
    '
    'pointer to HOSTENT structure returned by
    'the gethostbyname function
    Dim lngPtrToHOSTENT As Long
    '
    'structure which stores all the host info
    Dim udtHostent      As HOSTENT
    '
    'pointer to the IP address' list
    Dim lngPtrToIP      As Long
    '
    Dim lngAddress As Long
    '
    lngAddress = inet_addr(strHostName)
    '
    If lngAddress = INADDR_NONE Then
        '
        lngPtrToHOSTENT = gethostbyname(strHostName)
        '
        If lngPtrToHOSTENT <> 0 Then
            '
            'The gethostbyname function has found the address
            '
            'Copy retrieved data to udtHostent structure
            RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
            '
            'Now udtHostent.hAddrList member contains
            'an array of IP addresses
            '
            'Get a pointer to the first address
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
            '
            'Get the address
            RtlMoveMemory lngAddress, lngPtrToIP, udtHostent.hLength
            '
        Else
            '
            lngAddress = INADDR_NONE
            '
        End If
        '
    End If
    '
    GetAddressLong = lngAddress
    '
End Function

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
    '
    Dim strDesc As String
    '
    Select Case lngErrorCode
        '
        Case WSAEACCES
            strDesc = "Permission denied."
        Case WSAEADDRINUSE
            strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
            strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
            strDesc = "Operation already in progress."
        Case WSAECONNABORTED
            strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
            strDesc = "Connection refused."
        Case WSAECONNRESET
            strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            strDesc = "Destination address required."
        Case WSAEFAULT
            strDesc = "Bad address."
        Case WSAEHOSTDOWN
            strDesc = "Host is down."
        Case WSAEHOSTUNREACH
            strDesc = "No route to host."
        Case WSAEINPROGRESS
            strDesc = "Operation now in progress."
        Case WSAEINTR
            strDesc = "Interrupted function call."
        Case WSAEINVAL
            strDesc = "Invalid argument."
        Case WSAEISCONN
            strDesc = "Socket is already connected."
        Case WSAEMFILE
            strDesc = "Too many open files."
        Case WSAEMSGSIZE
            strDesc = "Message too long."
        Case WSAENETDOWN
            strDesc = "Network is down."
        Case WSAENETRESET
            strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
            strDesc = "Network is unreachable."
        Case WSAENOBUFS
            strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
            strDesc = "Bad protocol option."
        Case WSAENOTCONN
            strDesc = "Socket is not connected."
        Case WSAENOTSOCK
            strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
            strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
            strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
            strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
            strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
            strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
            strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
            strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            strDesc = "Host not found."
        Case WSANOTINITIALISED
            strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
            strDesc = "System call failure."
        Case WSASYSNOTREADY
            strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
            strDesc = "Graceful shutdown in progress."
        Case Else
            strDesc = "Unknown error."
    End Select
    '
    GetErrorDescription = strDesc
    '
End Function

Public Function InitializeWinsock(Version As WinsockVersion) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Purpose   :Initializes the Winsock service
'Returns   :Zero if successful. Otherwise, it returns error code.
'           The description of the error can be retrieved with
'           the GetErrorDescription function.
'Arguments :Version - Winsock's version: 1.1 or 2.2
'********************************************************************************
    '
    Dim udtWinsockData  As WSAData
    Dim lngRetValue     As Long
    '
    'start up winsock service
    lngRetValue = WSAStartup(Version, udtWinsockData)
    '
    'assign returned value
    InitializeWinsock = lngRetValue
    '
End Function


Public Function vbSocket(ByVal AdrFamily As AddressFamily, ByVal SckType As SocketType, ByVal SckProtocol As SocketProtocol) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Purpose   :Creates a new socket
'Returns   :The socket handle if successful, otherwise - INVALID_SOCKET
'Arguments :
'********************************************************************************
    '
    On Error GoTo vbSocket_Err_Handler
    '
    Dim lngRetValue As Long 'value returned by the socket API function
    '
    'Call the socket Winsock API function
    'in order create a new socket
    lngRetValue = socket(AdrFamily, SckType, SckProtocol)
    '
    'Assign returned value
    vbSocket = lngRetValue
    '
EXIT_LABEL:
    Exit Function

vbSocket_Err_Handler:
    '
    vbSocket = INVALID_SOCKET
    '
End Function

Public Function vbConnect(ByVal lngSocket As Long, ByVal strRemoteHost As String, ByVal intRemotePort As Integer) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :05-Oct-2001
'Purpose   :Establishes connection to the remote host.
'Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
'Arguments :lngSocket     - the socket to establish the connection for.
'           strRemoteHost - name or IP address of the host to connect to
'           intRemotePort - the port number to connect to
'********************************************************************************
    '
    Dim udtSocketAddress As sockaddr_in
    Dim lngReturnValue   As Long
    Dim lngAddress       As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    vbConnect = SOCKET_ERROR
    '
    'Check the socket handle
    If Not lngSocket > 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the socket handle is invalid one
        '
        Exit Function
        '
    End If
    '
    'Check the remote host address argument
    If Len(strRemoteHost) = 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the strRemoteHost argument can't be empty
        '
        Exit Function
        '
    End If
    '
    'Check the port number
    If Not intRemotePort > 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the intRemotePort must be a positive value
        '
        Exit Function
        '
    End If
    '
    'Prepare the sockaddr_in structure to pass to the
    'connect Winsock API function
    '
    'The sin_family member of the structure needs
    'the address family value that we can retieve
    'with CProtocol class
    '
    Dim objProtocol As New CProtocol
    Dim lngAdrFamily As Long
    '
    Call objProtocol.GetBySocketHandle(lngSocket)
    '
    lngAdrFamily = objProtocol.AddressFamily
    '
    Set objProtocol = Nothing
    '
    'the strRemoteHost may contain the host name
    'or IP address - GetAddressLong returns a valid
    'value anyway
    lngAddress = GetAddressLong(strRemoteHost)
    '
    If lngAddress = INADDR_NONE Then
        '
        Exit Function
        '
    End If
    '
    With udtSocketAddress
        '
        .sin_addr = lngAddress
        '
        'convert the port number to the network byte ordering
        .sin_port = htons(UnsignedToInteger(CLng(intRemotePort)))
        '
        .sin_family = lngAdrFamily
        '
    End With
    '
    vbConnect = connect(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
    '
EXIT_LABEL:
    Exit Function
    '
ERROR_HANDLER:
    '
    If Not objProtocol Is Nothing Then
        Set objProtocol = Nothing
    End If
    '
    vbConnect = SOCKET_ERROR
    '
End Function

Public Function vbBind(ByVal lngSocket As Long, ByVal strLocalHost As String, ByVal lngLocalPort As Long) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :02-Nov-2001
'Purpose   :Binds the socket to the local address
'Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
'Arguments :lngSocket    - the socket to bind
'           strLocalHost - name or IP address of the local host to bind to
'           lngLocalPort - the port number to bind to
'********************************************************************************
    '
    Dim udtSocketAddress As sockaddr_in
    Dim lngReturnValue   As Long
    Dim lngAddress       As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    vbBind = SOCKET_ERROR
    '
    'Check the socket handle
    If Not lngSocket > 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the socket handle is invalid one
        '
        Exit Function
        '
    End If
    '
    'Check the local host address argument
    If Len(strLocalHost) = 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the strLocalHost argument can't be empty
        '
        Exit Function
        '
    End If
    '
    'Check the port number
    If Not lngLocalPort > 0 Then
        '
        'TO DO: Inform the user or the calling procedure
        '       that the lngLocalPort must be a positive value
        '
        Exit Function
        '
    End If
    '
    'Prepare the sockaddr_in structure to pass to the
    'bind Winsock API function
    '
    'The sin_family member of the structure needs
    'the address family value that we can retieve
    'with CProtocol class
    '
    Dim objProtocol As New CProtocol
    Dim lngAdrFamily As Long
    '
    Call objProtocol.GetBySocketHandle(lngSocket)
    '
    lngAdrFamily = objProtocol.AddressFamily
    '
    Set objProtocol = Nothing
    '
    'the strLocalHost may contain the host name
    'or IP address - GetAddressLong returns a valid
    'value anyway
    lngAddress = GetAddressLong(strLocalHost)
    '
    If lngAddress = INADDR_NONE Then
        '
        Exit Function
        '
    End If
    '
    With udtSocketAddress
        '
        .sin_addr = lngAddress
        '
        'convert the port number to the network byte ordering
        .sin_port = htons(UnsignedToInteger(lngLocalPort))
        '
        .sin_family = lngAdrFamily
        '
    End With
    '
    vbBind = bind(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
    '
EXIT_LABEL:
    Exit Function
    '
ERROR_HANDLER:
    '
    If Not objProtocol Is Nothing Then
        Set objProtocol = Nothing
    End If
    '
    vbBind = SOCKET_ERROR
    '
End Function

Public Function GetIPEndPointField(ByVal lngSocket As Long, _
                                   ByVal EndPointField As IPEndPointFields) As Variant
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :21.10.01
'Purpose   :Retrieves IP address or host name or port number of
'           an end-point of the connection established
'           on the socket - lngSocket
'
'Return    :If no errors occures, the function returns the value
'           requested by the EndPointField argument.
'           Otherwise, it returns the value of SOCKET_ERROR
'
'Arguments :
'       lngSocket -  socket's handle. The socket must be connected to the remote host.
'       EndPointField - specifies the value to return:
'               LOCAL_HOST
'               LOCAL_HOST_IP
'               LOCAL_PORT
'               REMOTE_HOST
'               REMOTE_HOST_IP
'               REMOTE_PORT
'********************************************************************************
    '
    Dim udtSocketAddress    As sockaddr_in
    Dim lngReturnValue      As Long
    Dim lngPtrToAddress     As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    Select Case EndPointField
        Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT
            '
            'If the info of a local end-point of the connection is
            'requested, call the getsockname Winsock API function
            lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
            '
        Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT
            '
            'If the info of a remote end-point of the connection is
            'requested, call the getpeername Winsock API function
            lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
            '
    End Select '->EndPointField
    '
    If lngReturnValue = 0 Then
        '
        'If no errors were occurred, the getsockname or getpeername
        'function returns 0.
        '
        Select Case EndPointField
            Case LOCAL_PORT, REMOTE_PORT
                '
                'If the port number is requested, retrieve that value
                'from the sin_port member of the udtSocketAddress
                'structure, and change the byte order of that value from
                'the network to host byte order.
                GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))
                '
            Case LOCAL_HOST_IP, REMOTE_HOST_IP
                '
                'The host address is stored in the sin_addr member of the structure
                'as 4-byte value.
                '
                'To get an IP address of the host:
                '
                'Get pointer to the string that contains the IP address
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                '
                'Retrieve that string by the pointer
                GetIPEndPointField = StringFromPointer(lngPtrToAddress)
                '
            Case LOCAL_HOST, REMOTE_HOST
                '
                'The same procedure as for an IP address.
                'But here is the GetHostNameByAddress function call
                'to retrieve host name by IP address.
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                GetIPEndPointField = GetHostNameByAddress(StringFromPointer(lngPtrToAddress))
                '
        End Select  '->
        '
    Else
        '
        GetIPEndPointField = SOCKET_ERROR
        '
    End If  '->lngReturnValue = 0
    '
EXIT_LABEL:
    Exit Function

ERROR_HANDLER:
    GetIPEndPointField = SOCKET_ERROR
    
End Function

Private Function GetHostNameByAddress(strIpAddress As String) As String
    '
    Dim lngInetAdr As Long
    Dim lngPtrHostEnt As Long
    Dim lngPtrHostName As Long
    Dim strHostName As String
    Dim udtHostent As HOSTENT
    '
    strIpAddress = Trim$(strIpAddress)
    '
    'Valid IP address contains at least 7 characters
    If Len(strIpAddress) > 6 Then
        '
        'Convert the IP address string to Long
        lngInetAdr = inet_addr(strIpAddress)
        '
        '## Retrieve host name
        '
        'Get the pointer to the HostEnt structure
        lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)
        '
        'Copy data into the HostEnt structure
        RtlMoveMemory udtHostent, ByVal lngPtrHostEnt, LenB(udtHostent)
        '
        'Prepare the buffer to receive a string
        strHostName = String(256, 0)
        '
        'Copy the host name into the strHostName variable
        RtlMoveMemory ByVal strHostName, ByVal udtHostent.hName, 256
        '
        'Cut received string by first chr(0) character
        strHostName = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
        '
        'Return the found value
        GetHostNameByAddress = strHostName
        '
    End If
    '
End Function

Public Function vbSend(ByVal lngSocket As Long, strData As String) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :27-Nov-2001
'Purpose   :Sends data to the remote host with connected socket
'Returns   :Number of bytes sent to the remote host
'Arguments :lngSocket    - the socket connected to the remote host
'           strData      - data to send
'********************************************************************************
    '
    Dim arrBuffer()     As Byte
    Dim lngBytesSent    As Long
    Dim lngBufferLength As Long
    '
    lngBufferLength = Len(strData)
    '
    If IsConnected(lngSocket) And lngBufferLength > 0 Then
        '
        'Convert the data string to a byte array
        arrBuffer() = StrConv(strData, vbFromUnicode)
        '
        'Call the send Winsock API function in order to send data
        lngBytesSent = Send(lngSocket, arrBuffer(0), lngBufferLength, 0&)
        '
        vbSend = lngBytesSent
        '
    Else
        '
        vbSend = SOCKET_ERROR
        '
    End If
    '
End Function

Public Function vbRecv(ByVal lngSocket As Long, strBuffer As String) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :27-Nov-2001
'Purpose   :Retrieves data from the Winsock buffer.
'Returns   :Number of bytes received.
'Arguments :lngSocket    - the socket connected to the remote host
'           strBuffer    - buffer to read data to
'********************************************************************************
    '
    Const MAX_BUFFER_LENGTH As Long = 8192
    '
    Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
    Dim lngBytesReceived                    As Long
    Dim strTempBuffer                       As String
    '
    'Check the socket for readabilty with
    'the IsDataAvailable function
    If IsDataAvailable(lngSocket) Then
        '
        'Call the recv Winsock API function in order to read data from the buffer
        lngBytesReceived = recv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
        '
        If lngBytesReceived > 0 Then
            '
            'If we have received some data, convert it to the Unicode
            'string that is suitable for the Visual Basic String data type
            strTempBuffer = StrConv(arrBuffer, vbUnicode)
            '
            'Remove unused bytes
            strBuffer = Left$(strTempBuffer, lngBytesReceived)
            '
        End If
        '
        'If lngBytesReceived is equal to 0 or -1, we have nothing to do with that
        '
        vbRecv = lngBytesReceived
        '
    Else
        '
        'Something wrong with the socket.
        'Maybe the lngSocket argument is not a valid socket handle at all
        vbRecv = SOCKET_ERROR
        '
    End If
    '
End Function

Public Function vbListen(ByVal lngSocketHandle As Long) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :15-Dec-2001
'Purpose   :Turns a socket into a listening state.
'Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
'Arguments :lngSocketHandle - the socket to turn into a listening state.
'********************************************************************************
    '
    Dim lngRetValue As Long
    '
    lngRetValue = listen(lngSocketHandle, SOMAXCONN)
    '
    vbListen = lngRetValue
    '
    'We have nothing to do as vbListen function returns the same value
    'as the listen Winsock API function.
    '
End Function

Public Function vbAccept(ByVal lngSocketHandle As Long) As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :15-Dec-2001
'Purpose   :Accepts a connection request, and creates a new socket.
'Return    :If no error occurs, returns the new socket's handle. Otherwise, it returns INVALID_SOCKET.
'Arguments :lngSocketHandle - the listening socket.
'********************************************************************************
    '
    Dim lngRetValue         As Long
    Dim udtSocketAddress    As sockaddr_in
    Dim lngBufferSize       As Long
    '
    'Calculate the buffer size
    lngBufferSize = LenB(udtSocketAddress)
    '
    'Call the accept Winsock API function in order to create a new socket
    lngRetValue = Accept(lngSocketHandle, udtSocketAddress, lngBufferSize)
    '
    vbAccept = lngRetValue
    '
End Function

Public Function IsConnected(ByVal lngSocket As Long) As Boolean
    '
    Dim udtRead_fd      As fd_set
    Dim udtWrite_fd     As fd_set
    Dim udtError_fd     As fd_set
    Dim lngSocketCount  As Long
    '
    udtWrite_fd.fd_count = 1
    udtWrite_fd.fd_array(1) = lngSocket
    '
    lngSocketCount = vbselect(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    '
    IsConnected = CBool(lngSocketCount)
    '
End Function

Public Function IsDataAvailable(ByVal lngSocket As Long) As Boolean
    '
    Dim udtRead_fd As fd_set
    Dim udtWrite_fd As fd_set
    Dim udtError_fd As fd_set
    Dim lngSocketCount As Long
    '
    udtRead_fd.fd_count = 1
    udtRead_fd.fd_array(1) = lngSocket
    '
    lngSocketCount = vbselect(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    '
    IsDataAvailable = CBool(lngSocketCount)
    '
End Function


