Attribute VB_Name = "mod_Framewrk"
Public Const COMMAND_ERROR = -1
Public Const RECV_ERROR = -1
Public Const NO_ERROR = 0

Public socketId As Long

Global State As Integer

Sub CloseConnection()

x = closesocket(socketId)
    
If x = SOCKET_ERROR Then
    MsgBox ("ERROR: closesocket = " + Str$(x))
    Exit Sub
End If

End Sub

Sub EndIt()

x = WSACleanup() 'Shutdown Winsock DLL

End Sub

Sub StartIt()

Dim StartUpInfo As WSAData
    
'Get WinSock version
'Version 1.1 (1*256 + 1) = 257
'version 2.0 (2*256 + 0) = 512
version = 257
    
'Initialize Winsock DLL
x = WSAStartup(version, StartUpInfo)

End Sub
 
Function OpenSocket(ByVal Hostname As String, ByVal PortNumber As Integer) As Integer
Dim I_SocketAddress As sockaddr_in
Dim ipAddress As Long
    
ipAddress = inet_addr(Hostname)

'Create a new socket
socketId = Socket(AF_INET, SOCK_STREAM, 0)
If socketId = SOCKET_ERROR Then
    MsgBox ("ERROR: socket = " + Str$(socketId))
    OpenSocket = COMMAND_ERROR
    Exit Function
End If

'Open a connection to a server
I_SocketAddress.sin_family = AF_INET
I_SocketAddress.sin_port = htons(PortNumber)
I_SocketAddress.sin_addr = ipAddress
I_SocketAddress.sin_zero = String$(8, 0)

x = connect(socketId, I_SocketAddress, Len(I_SocketAddress))

If socketId = SOCKET_ERROR Then
    MsgBox ("ERROR: connect = " + Str$(x))
    OpenSocket = COMMAND_ERROR
    Exit Function
End If

OpenSocket = socketId

End Function

Function sendCommand(ByVal command As String) As Integer

Dim strSend As String
    
strSend = command + vbCrLf
    
count = send(socketId, ByVal strSend, Len(strSend), 0)
    
If count = SOCKET_ERROR Then
    MsgBox ("ERROR: send = " + Str$(count))
    sendCommand = COMMAND_ERROR
    Exit Function
End If
    
sendCommand = NO_ERROR

End Function

Public Function RecvStrTO(sock As Long, Optional timeout As Long = 10) As String
Dim lbuffer As String
Dim lfdr As FD_SET, lfdw As FD_SET, lfde As FD_SET
Dim lRet As Long
Dim lti As TIMEVAL

lti.tv_sec = timeout 'Time in seconds

lfdr.fd_count = 1 'A socket to check

lfdr.fd_array(1) = sock 'The socket passed as a parameter

lRet = wselect(0, lfdr, lfdw, lfde, lti) 'Is the socket ready?

If lRet > 0 Then 'If no error and delay not exceeded
   Do 'Loop as long as there is data
       If lfdr.fd_count = 1 Then 'If socket ready
           lbuffer = Space(1024) 'Receive
           lRet = recvstr(sock, lbuffer, 1024, 0)
           If lRet > 0 Then 'Adds the received data to the result
               lbuffer = Left(lbuffer, lRet)
               RecvStrTO = RecvStrTO & lbuffer
           ElseIf lRet <= 0 Then
               Exit Do
           End If
       End If
       lti.tv_sec = 0 'Checks if there is still data to be received (with timeout at zero)
       lRet = wselect(0, lfdr, lfdw, lfde, lti)
       If lRet <= 0 Then Exit Do 'If error or delay exceeded
   Loop
End If

End Function

Function RecvAscii(dataBuf As String, ByVal maxLength As Integer) As Integer
Dim c As String * 1
Dim length As Integer
    
dataBuf = ""

While length < maxLength
    DoEvents
    count = recv(socketId, c, 1, 0)
    
    If count < 1 Then
        RecvAscii = RECV_ERROR
        dataBuf = Chr$(0)
        Exit Function
    End If
        
    If c = Chr$(10) Then
        dataBuf = dataBuf + Chr$(0)
        RecvAscii = NO_ERROR
        Exit Function
    End If
        
    length = length + count
    dataBuf = dataBuf + c
Wend
    
RecvAscii = RECV_ERROR
    
End Function

Function RecvAryReal(dataBuf() As Double) As Long
'receive DOS format 64bit binary data

Dim buf As String * 20
Dim size As Long
Dim length As Long
Dim count As Long
Dim recvBuf(25616) As Byte
    
x = recv(socketId, buf, 8, 0) 'receive header info "#6NNNNNN"
size = Val(Mid$(buf, 3, 6))
    
count = 0
length = 0

Do While length < size
    DoEvents
    count = recvB(socketId, recvBuf(length), size - length, 0)
    If (count > 0) Then
        length = length + count
    End If
Loop
    
count = recv(socketId, buf, 1, 0) 'receive ending LF
    
CopyMemory dataBuf(LBound(dataBuf)), recvBuf(0), length 'copy recieved data to Single type array dataBuf()

RecvAryReal = length / 8 'dataBuf = recvBuf
    
End Function
