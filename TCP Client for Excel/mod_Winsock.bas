Attribute VB_Name = "mod_winsock"
Public Type FD_SET
    fd_count As Long
    fd_array(1 To 64) As Long
End Type

Public Type TIMEVAL
    tv_sec  As Long
    tv_usec As Long
End Type

Public Declare Function wselect Lib "ws2_32.dll" Alias "select" (ByVal Reserved As Long, _
                                        ByRef ReadFds As FD_SET, ByRef WriteFds As FD_SET, _
                                        ByRef ExceptFds As FD_SET, ByRef timeout As TIMEVAL) As Long
                                        
'Public Declare Function recv Lib "ws2_32.dll" (ByVal hSocket As Long, ByRef buffer As Any, _
                                        ByVal BufferLength As Long, ByVal Flags As Long) As Long

Public Declare Function recvstr Lib "ws2_32.dll" Alias "recv" (ByVal hSocket As Long, ByVal buffer As String, _
                                        ByVal BufferLength As Long, ByVal Flags As Long) As Long

'This is the Winsock API definition file for Visual Basic
'Setup the variable type 'hostent' for the WSAStartup command

Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As String * 2
    h_length As String * 2
    h_addr_list As Long
End Type

Public Const SZHOSTENT = 16

'Set the Internet address type to a long integer (32-bit)
Type in_addr
   s_addr As Long
End Type

'A note to those familiar with the C header file for Winsock
'Visual Basic does not permit a user-defined variable type
'to be used as a return structure.  In the case of the
'variable definition below, sin_addr must
'be declared as a long integer rather than the user-defined
'variable type of in_addr.
Type sockaddr_in
   sin_family As Integer
   sin_port As Integer
   sin_addr As Long
   sin_zero As String * 8
End Type

Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128
Public Const WSA_DescriptionSize = WSADESCRIPTION_LEN + 1
Public Const WSA_SysStatusSize = WSASYS_STATUS_LEN + 1

'Setup the structure for the information returned from
'the WSAStartup() function.
Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription As String * WSA_DescriptionSize
   szSystemStatus As String * WSA_SysStatusSize
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpVendorInfo As String * 200
End Type

'Define socket return codes
Public Const INVALID_SOCKET = &HFFFF
Public Const SOCKET_ERROR = -1

'Define socket types
Public Const SOCK_STREAM = 1           'Stream socket
Public Const SOCK_DGRAM = 2            'Datagram socket
Public Const SOCK_RAW = 3              'Raw data socket
Public Const SOCK_RDM = 4              'Reliable Delivery socket
Public Const SOCK_SEQPACKET = 5        'Sequenced Packet socket


'Define address families
Public Const AF_UNSPEC = 0             'unspecified
Public Const AF_UNIX = 1               'local to host (pipes, portals)
Public Const AF_INET = 2               'internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK = 3            'arpanet imp addresses
Public Const AF_PUP = 4                'pup protocols: e.g. BSP
Public Const AF_CHAOS = 5              'mit CHAOS protocols
Public Const AF_NS = 6                 'XEROX NS protocols
Public Const AF_ISO = 7                'ISO protocols
Public Const AF_OSI = AF_ISO           'OSI is ISO
Public Const AF_ECMA = 8               'european computer manufacturers
Public Const AF_DATAKIT = 9            'datakit protocols
Public Const AF_CCITT = 10             'CCITT protocols, X.25 etc
Public Const AF_SNA = 11               'IBM SNA
Public Const AF_DECnet = 12            'DECnet
Public Const AF_DLI = 13               'Direct data link interface
Public Const AF_LAT = 14               'LAT
Public Const AF_HYLINK = 15            'NSC Hyperchannel
Public Const AF_APPLETALK = 16         'AppleTalk
Public Const AF_NETBIOS = 17           'NetBios-style addresses
Public Const AF_MAX = 18               'Maximum # of address families


'Setup sockaddr data type to store Internet addresses
Type sockaddr
  sa_family As Integer
  sa_data As String * 14
End Type

Public Const SADDRLEN = 16

'Declare Socket functions
Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long

Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr_in, ByVal namelen As Long) As Long

Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer

Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

Public Declare Function recvB Lib "wsock32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long

Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAData) As Long

Public Declare Function WSACleanup Lib "wsock32.dll" () As Long

Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
