Attribute VB_Name = "modSocket"
Option Explicit

Private Type addrinfo
    ai_flags     As Long
    ai_family    As Long
    ai_socktype  As Long
    ai_protocol  As Long
    ai_addrlen   As Long
    ai_canonname As String
    ai_addr      As Long
    ai_next      As Long
End Type

Private Declare Function getaddrinfo Lib "ws2_32.dll" (ByVal pNodeName As String, ByVal pServiceName As String, ByRef pHints As Any, ByRef ppResult As Long) As Long
Private Declare Sub freeaddrinfo Lib "ws2_32.dll" (ByRef ai As addrinfo)

'Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" _
        (ByVal lDestIPAddr As Long, _
         ByRef lHopCount As Long, _
         ByVal lMaxHops As Long, _
         ByRef lRTT As Long) As Long

'Private Type SERVER_INFO_API
'    PlatformId As Long
'    ServerName As Long
'    Type As Long
'    VerMajor As Long
'    VerMinor As Long
'    Comment As Long
'End Type
'Private Declare Function NetServerEnum Lib "netapi32" _
'        (lpServer As Any, ByVal lLevel As Long, vBuffer As Any, _
'        lPreferedMaxLen As Long, lEntriesRead As Long, lTotalEntries As Long, _
'        ByVal lServerType As Long, ByVal sDomain$, vResume As Any) As Long
'Declare Function NetApiBufferFree Lib "netapi32" _
'       (ByVal lBuffer&) As Long

'Private Type MIB_UDPROW_OWNER_PID
'  dwLocalAddr As Long
'  dwLocalPort As Long
'  dwOwningPid As Long
'End Type
'typedef struct _MIB_TCPROW_OWNER_PID {
'  DWORD dwState;
'  DWORD dwLocalAddr;
'  DWORD dwLocalPort;
'  DWORD dwRemoteAddr;
'  DWORD dwRemotePort;
'  DWORD dwOwningPid;
'}
'Private Type MIB_TABLE_OWNER_PID
'  dwNumEntries As Long
'  table [ANY_SIZE];
'End Type

'Private Declare Function GetExtendedTcpTable Lib "IPHLPAPI.dll" (ByRef pUdpTable As Any, ByRef pdwSize As Long, ByVal Border As Long, ByVal ulAf As Long, ByVal nTableClass As Long, ByVal nReserved As Long) As Long
Private Declare Function GetExtendedUdpTable Lib "iphlpapi.dll" (ByRef pUdpTable As Any, ByRef pdwSize As Long, ByVal Border As Long, ByVal ulAf As Long, ByVal nTableClass As Long, ByVal nReserved As Long) As Long
 'ulaf = 2AF_INET| 23=AF_INET6
 'nTableClass
' typedef enum _UDP_TABLE_CLASS {
'  UDP_TABLE_BASIC,
'  1 UDP_TABLE_OWNER_PID,
'  UDP_TABLE_OWNER_MODULE
'}
 'typedef enum _TCP_TABLE_CLASS {
'  TCP_TABLE_BASIC_LISTENER,
'  TCP_TABLE_BASIC_CONNECTIONS,
'  TCP_TABLE_BASIC_ALL,
'  TCP_TABLE_OWNER_PID_LISTENER,
'  4 TCP_TABLE_OWNER_PID_CONNECTIONS,
'  5 TCP_TABLE_OWNER_PID_ALL,
'  TCP_TABLE_OWNER_MODULE_LISTENER,
'  TCP_TABLE_OWNER_MODULE_CONNECTIONS,
'  TCP_TABLE_OWNER_MODULE_ALL
'}

'Private Type prWinInetContext
'    dwExitFlag  As Long
'    dwRetCode   As Long
'    dwErrCode   As Long
'End Type
'Const INTERNET_STATUS_REQUEST_COMPLETE      As Long = 100



'Private Type WSAData
'    wVersion As Integer
'    wHighVersion As Integer
'    szDescription As String * 257 'WSADESCRIPTION_LEN
'    szSystemStatus As String * 129 'WSASYS_STATUS_LEN
'    iMaxSockets As Integer
'    iMaxUdpDg As Integer
'    lpVendorInfo As Long
'End Type
Private Type ssock
    hSocket As Long
    sa As sockaddr_in
End Type
'Private Type sockaddr_in
'    sin_family       As Integer
'    sin_port         As Integer
'    sin_addr         As Long
'    sin_zero(1 To 8) As Byte
'End Type
'Private Type HOSTENT
'    hName     As Long
'    hAliases  As Long
'    hAddrType As Integer
'    hLength   As Integer
'    hAddrList As Long
'End Type


'Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
'Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
'Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
'
'Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
'Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
'Private Declare Function api_sendto Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
'Private Declare Function api_recvfrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef from As sockaddr_in, ByRef fromlen As Long) As Long
'Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
'
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByVal buf As String, ByVal lLen As Long, ByVal Flags As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByVal buf As String, ByVal lLen As Long, ByVal Flags As Long) As Long


Public Const FD_SETSIZE = 64
Type FD_SET
    fd_count As Long
    fd_array(0 To FD_SETSIZE - 1) As Long
End Type
Type TIME_VAL
    tv_sec As Long
    tv_usec As Long
End Type

Private Declare Function sselect Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, readfds As FD_SET, writefds As FD_SET, exceptfds As FD_SET, timeOut As TIME_VAL) As Long

Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long

Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare Function api_shutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
'Private Const SD_RECEIVE As Long = &H0
'Private Const SD_SEND As Long = &H1
'Private Const SD_BOTH As Long = &H2


'Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long 'IP String->Long
'Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long 'IP Long->String

'Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
'Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'Private Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
'Private Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
'Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long




'Private Const udp_packet_size As Long = 1280
Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
'Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal lLen As Long, ByVal flags As Long) As Long
'Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal lLen As Long, ByVal flags As Long) As Long

'Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long



Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIP As Long, ByVal SrcIP As Long, ByRef pMacAddr As Long, ByRef PhyAddrLen As Long) As Long

Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type
Private Type ICMP_ECHO_REPLY
   address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   reserved        As Integer
   ptrData                 As Long
   Options        As IP_OPTION_INFORMATION
   Data            As String * 250
End Type
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal timeOut As Long) As Long
    
Private m_blnWinsockInit As Boolean
Private m_lngMaxMsgSize As Long
'Private m_MaxMsgSize As Long
Private nSockets As Long


Private RSEvents As Recordset


'====== xEVENTS ========
Function xEvent(ByVal EvTx$, ByVal hWnd&, ByVal hEv$)
If RSEvents Is Nothing Then
'    Set RSEvents = New Recordset
'    With RSEvents
'        .Fields.append "ev", adVarChar, 50
'        .Fields.append "hw", adInteger
'        .Fields.append "he", adVarChar, 255
'        .CursorLocation = adUseClient
'        .Open
'    End With
    Set RSEvents = xMain.CMatrix(0, Replace("ev$50$,hw&,he$", ",", vbTab)).GetRows(80)
End If
If RSEvents Is Nothing Then Exit Function
If (Len(EvTx) > 0) And (hWnd <> 0) Then 'ADD EVENT
    RSEvents.Filter = "ev='" & EvTx & "' AND hw=" & hWnd
    If RSEvents.RecordCount = 0 Then RSEvents.AddNew Array("ev", "hw", "he"), Array(EvTx, hWnd, hEv)
Else
    If (Len(EvTx) = 0) And (hWnd <> 0) Then  'REMOVE EVENT HWND
        RSEvents.Filter = "hw=" & hWnd
    ElseIf (Len(EvTx) > 0) And (hWnd = 0) Then 'REMOVE EVENT
        RSEvents.Filter = "ev='" & EvTx & "'"
    Else 'REMOVE ALL EVENTS
        RSEvents.Filter = 0
    End If
    While RSEvents.RecordCount
    RSEvents.MoveFirst
    RSEvents.Delete
    Wend
End If

'On Error Resume Next
'RSEvents.Filter = 0
'RSEvents.MoveFirst
'Debug.Print "xEvent=" & xMain.GetString(RSEvents, "", ",", "<", ">")
'Debug.Print
End Function
'====== xEVENTS ========



Function pIPAddress$(HostName, Optional retLong As Long)
Dim u As HOSTENT, a1&, a As sockaddr_in 'a2&, a3&,
'Dim z As Boolean
'If Not m_blnWinsockInit Then InitWinsockService: z = 1

If InitWinsockService Then
    a1 = gethostbyname("" & HostName)
    If a1 Then
        CopyMemory u, ByVal a1, LenB(u)
        HostName = StringFromPointer(u.hName)
        
        'CopyMemory a1, ByVal u.hAddrList, 4
        GetMem4 ByVal u.hAddrList, a1
        'CopyMemory a.sin_addr, ByVal a1, 4
        GetMem4 ByVal a1, a.sin_addr
        If retLong = -1 Then pIPAddress = a.sin_addr: Exit Function
        a1 = inet_ntoa(a.sin_addr)
        pIPAddress = StringFromPointer(a1)
        If Not retLong Then
            If Trim$("" & HostName) = pIPAddress Then
                a1 = gethostbyaddr(a.sin_addr, 4&, 2&)  'AF_INET=2, AF_NETBIOS=17
                If a1 Then
                    CopyMemory u, ByVal a1, LenB(u)
                    HostName = StringFromPointer(u.hName)
                End If
            End If
        End If
    End If
'If z Then CleanupWinsock
End If

End Function

Function pMACAddress$(ByVal IPAdr)
Dim b(7) As Byte, a&
'Dim z As Boolean
'If Not m_blnWinsockInit Then InitWinsockService: z = 1
a = inet_addr("" & IPAdr)
If a = -1 Then Exit Function
If SendARP(a, 0&, ByVal VarPtr(b(0)), 8) = 0 Then pMACAddress = Hex2(b(0)) & ":" & Hex2(b(1)) & ":" & Hex2(b(2)) & ":" & Hex2(b(3)) & ":" & Hex2(b(4)) & ":" & Hex2(b(5))
'If z Then CleanupWinsock
End Function
Function Hex2(ByVal the_byt As Byte) As String
Hex2 = Hex$(the_byt)
If Len(Hex2) = 1 Then Hex2 = "0" & Hex2
End Function


Private Function InitWinsockService() As Boolean
Dim wsa As WSADATA   'structure to pass to WSAStartup as an argument
If Not m_blnWinsockInit Then
    If WSAStartup(&H101, wsa) = 0 Then
        m_blnWinsockInit = True
        If wsa.iMaxUdpDg < 0 Then m_lngMaxMsgSize = wsa.iMaxUdpDg + 65536 Else m_lngMaxMsgSize = wsa.iMaxUdpDg
        'Debug.Print Hex(wsa.wVersion)
        InitWinsockService = m_lngMaxMsgSize
        'Debug.Print "wsa.iMaxUdpDg =" & m_lngMaxMsgSize, wsa.iMaxUdpDg
    End If
Else
    InitWinsockService = m_lngMaxMsgSize
End If
End Function

'Private
Sub CleanupWinsock()
If m_blnWinsockInit And nSockets < 1 Then
    Call WSACleanup
    m_blnWinsockInit = False
End If
End Sub

Function SocketClose(hSocket As Long, ByVal hWnd As Long)
If hSocket = -1 Then Exit Function
'Debug.Print "SocketClose " & hSocket
If hWnd Then WSAAsyncSelect hSocket, hWnd, 0&, 0&
nSockets = nSockets - 1
If nSockets < 0 Then nSockets = 0
api_closesocket hSocket
hSocket = 0
CleanupWinsock
End Function

Function SocketWaitData(ByVal hWnd&, hSocket&, LocalPort&, LocalHost$) As Boolean
Dim udt As sockaddr_in
If hWnd = 0 Then Exit Function
If hSocket <> -1 Then SocketClose hSocket, hWnd
If LocalPort = 0 Then Exit Function
InitWinsockService
hSocket = api_socket(2, 2, 17) 'UDT
If hSocket = -1 Then Exit Function
udt.sin_addr = inet_addr(pIPAddress(LocalHost, 1))
udt.sin_family = 2 'AF_INET
udt.sin_port = htons(LOWORD(LocalPort))
'Private Const SO_RCVBUF As Long = &H1002
'Dim n&: n = 15000
'setsockopt hSocket, &HFFFF&, &H1002&, n, Len(n) 'PACKET SIZE , SO_RCVBUF=1006, SOL_SOCKET=&HFFFF
If api_bind(hSocket, udt, LenB(udt)) = -1 Then api_closesocket hSocket: hSocket = -1: Exit Function
If WSAAsyncSelect(hSocket, hWnd, UM_SOCKET, &H1) = -1 Then api_closesocket hSocket: hSocket = -1: Exit Function
nSockets = nSockets + 1
SocketWaitData = hSocket > 0
End Function

Function pSocketPort(hs&) As Long
Dim udt As sockaddr_in
If hs < 1 Then Exit Function
getsockname hs, udt, Len(udt)
pSocketPort = ntohs(udt.sin_port)
End Function

Function SocketGetData(ByVal hSocket&, RemoteHost$, RemotePort&, Optional NoRSEvents As Long, Optional sSocket$) As String
Dim n As Long, udt As sockaddr_in 'socket address of the remote peer
If Not m_blnWinsockInit Then Exit Function
If hSocket = -1 Then Exit Function
If NoRSEvents Then
    n = 500: setsockopt hSocket, &HFFFF&, &H1006&, n, Len(n) 'Send Time out=timeOut, SO_RCVTIMEO=1006, SOL_SOCKET=&HFFFF
End If
SocketGetData = pSocketRecv(hSocket, udt)
RemoteHost = StringFromPointer(inet_ntoa(udt.sin_addr))
RemotePort = ntohs(udt.sin_port)
Dim ss As ssock, sa$: If Len(S_(sSocket)) Then ss.hSocket = hSocket: ss.sa = udt: sa = String(Len(ss), 0): CopyMemory ByVal StrPtr(sa), ss, Len(ss): sSocket = sa
If Not NoRSEvents Then If Len(SocketGetData) Then SocketGetData = S_(FireRSEvent(SocketGetData))
End Function

Private Function pSocketRecv(ByVal hSocket&, udt As sockaddr_in)
Dim n As Long
Dim b() As Byte
ReDim b(65535)
n = api_recvfrom(hSocket, b(0), 1 + UBound(b), 0&, udt, Len(udt))
If n < 1 Then n = 1
ReDim Preserve b(n - 1) ': st.Write b
If gIsCompressed(b) Then b = gDecompress(b)
pSocketRecv = StrConv(b, vbUnicode)
End Function


Function SocketSendData(txData$, RemotePort&, RemoteHost$, Optional nWaitResponse&)
Dim hSocket&, n&
Dim b() As Byte, udt As sockaddr_in
If Len(txData) = 0 Then Exit Function
If Not InitWinsockService Then Exit Function
hSocket = api_socket(2, 2, 17)
If hSocket = -1 Then Exit Function
udt.sin_addr = inet_addr(pIPAddress(RemoteHost, 1))
udt.sin_port = htons(LOWORD(RemotePort))
udt.sin_family = 2 'AF_INET

n = 500
setsockopt hSocket, &HFFFF&, &H1005&, n, Len(n) 'SEND TIME OUT=timeOut   SO_SNDTIMEO=1005 SOL_SOCKET=&HFFFF
setsockopt hSocket, &H17&, &H1&, ByVal 0&, 4 'TCP_NODELAY =NO TCP_NODELAY=1 IPPROTO_UDP=17

'n = 10000
'setsockopt hSocket, &HFFFF&, &H1001&, n, Len(n) 'PACKET SIZE , SO_SNDBUF=1001, SOL_SOCKET=&HFFFF

b = StrConv(txData, vbFromUnicode)
If UBound(b) > 1024 Then b = gCompress(b)
'n = UBound(b) + 1
'If n > 1280 Then setsockopt hSocket, &HFFFF&, &H1001&, n, Len(n)    'PACKET SIZE , SO_SNDBUF=1001, SOL_SOCKET=&HFFFF
SocketSendData = api_sendto(hSocket, b(0), 1 + UBound(b), 0&, udt, Len(udt))

If SocketSendData < 0 Then SocketSendData = 0
If nWaitResponse > 0 Then
    SocketSendData = ""
    n = nWaitResponse
    setsockopt hSocket, &HFFFF&, &H1006&, n, Len(n) 'RECV TIME OUT=timeOut   SO_RCVTIMEO=1006 SOL_SOCKET=&HFFFF
    'n = 32768: setsockopt hSocket, &HFFFF&, &H1002&, n, Len(n) 'PACKET SIZE , SO_RCVBUF=1006, SOL_SOCKET=&HFFFF
    API_DoEvents
    SocketSendData = pSocketRecv(hSocket, udt)
End If
api_closesocket hSocket
End Function

Function pSocketReply(ByVal sSocket$, txData$)
If Len(sSocket) = 0 Then Exit Function
Dim b() As Byte
Dim ss As ssock: CopyMemory ss, ByVal StrPtr(sSocket), Len(ss)
b = StrConv(txData, vbFromUnicode)
If UBound(b) > 256 Then b = gCompress(b)
If UBound(b) = -1 Then ReDim b(0)
'setsockopt ss.hSocket, &HFFFF&, &H1005&, timeOut, Len(timeOut) 'SEND TIME OUT=timeOut   SO_SNDTIMEO=1005 SOL_SOCKET=&HFFFF
'setsockopt hSocket, &H17&, &H1&, ByVal 0&, 4 'TCP_NODELAY =NO TCP_NODELAY=1 IPPROTO_UDP=17
pSocketReply = api_sendto(ByVal ss.hSocket, b(0), 1& + UBound(b), 0&, ss.sa, Len(ss.sa))
'Debug.Print "pSocketReply>>" & UBound(b) & ">>" & pSocketReply
End Function

Function FireRSEvent(ByVal evname) 'evname= name:argument
FireRSEvent = evname
If RSEvents Is Nothing Then Exit Function
On Error Resume Next
Dim rs, i&, xc As xControl, n&, code$, ev$
ev = xMain.LeftStr(S_(evname), ":")
If Len(ev) = 0 Then Exit Function
RSEvents.Filter = "ev=" & xMain.Quot(ev, 39)
If RSEvents.RecordCount Then
    RSEvents.MoveFirst
    rs = RSEvents.GetRows()
    If ArrayDims(rs) = 2 Then
        For i = 0 To UBound(rs, 2)
            Set xc = hxControl(0& + rs(1, i))
            n = 0: n = xc.hWnd
            If n <> 0 Then
                code = S_(rs(2, i))
                If Len(code) Then FireRSEvent = xc.ASE.Run(code) Else FireRSEvent = xc.hEvent(ev, xMain.RightStr(evname, ":", 1))
            Else
                xEvent "", rs(1, i), ""   'Óäàëÿåì íåäåéñòâèòåëüíóþ çàïèñü îêíà=íåòó
            End If
        Next
    End If
End If
End Function
'Public Function mPing(sAddress As String) As Long
'    ' Based on an article on 'Codeguru' by 'Bill Nolde'
'    ' Thx to this guy! It 's simple and great!
'    ' Implemented for VB by G. Wirth, Ulm,  Germany
'    ' Enjoy!
'    Dim lIPadr      As Long
'    Dim lHopsCount  As Long
'    Dim lRTT        As Long
'    Dim lMaxHops    As Long
'    lMaxHops = 20               ' should be enough ...
'    lIPadr = inet_addr(sAddress)
'    'lIPadr = inet_addr( pIPAddress(sAddress, 1))
'    mPing = (GetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT) = 1)
'End Function

Public Function mPing(sAddress As String) As Long
'nPort=0 ICMP
'nPort>0 TCP/UDP
', Optional nPort = 0

Dim hIcmp As Long
Dim lAddress As Long
Dim lTimeOut As Long
Dim StringToSend As String
Dim Reply As ICMP_ECHO_REPLY
If Len(sAddress) = 0 Then Exit Function
StringToSend = "0" 'Short string of data to send
lTimeOut = 1500 'ms' ICMP (ping) timeout
lAddress = inet_addr(pIPAddress(sAddress))  'Convert string address to a long representation.
If (lAddress <> -1) And (lAddress <> 0) Then
    hIcmp = IcmpCreateFile() 'Create the handle for ICMP requests.
    If hIcmp Then
        'Ping the destination IP address.
        Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)
        mPing = (Reply.Status = 0) 'Reply status
        IcmpCloseHandle hIcmp 'Close the Icmp handle.
'    Else
'        gDebugPrint "failure opening icmp handle."
'        mPing = 0
    End If
'Else
'    mPing = 0
End If
End Function


'
'Sub InternetStatusCallback(hInternet As Long, dwContext As Long, dwInternetStatus As Long, lpvStatusInformation As Long, dwStatusInformationLength As Long)
'Debug.Print "hInternet=" & hInternet, "dwContext=" & dwContext, "dwInternetStatus=" & dwInternetStatus, "lpvStatusInformation=" & lpvStatusInformation, "dwStatusInformationLength=" & dwStatusInformationLength
'End Sub
'
'' §§§§§§§§§§§§§§§§§§§§§§§§§§ CallBack §§§§§§§§§§§§§§§§§§§§§§§§§§
'
''--------------------------------------------------------------------------------
'' Ïðîåêò     :  OfflineClient
'' Ïðîöåäóðà  :  InternetCallbackFunc
'' Îïèñàíèå   :  Äëÿ àñèíõðîííîãî ðåæèìà â WinInet
'' Êåì ñîçäàí :  SNE
'' Äàòà-Âðåìÿ :  13.01.2005-23:34:57
''--------------------------------------------------------------------------------
'Sub InternetCallbackFunc(ByVal hInternet As Long, ByVal dwContext As Long, _
'                                                          ByVal dwInternetStatus As Long, _
'                                                          ByVal pbStatusInformation As Long, _
'                                                          ByVal dwStatusInformationLength As Long)
'Debug.Print "hInternet=" & hInternet, "dwContext=" & dwContext, "dwInternetStatus=" & dwInternetStatus, "lpvStatusInformation=" & pbStatusInformation, "dwStatusInformationLength=" & dwStatusInformationLength
'
'    Dim lng         As Long, _
'        szTmpBuffer As String * &H100
'    Dim stContext   As prWinInetContext
'
'    If Not dwContext = 0& Then Call CopyMemory(stContext, ByVal dwContext, Len(stContext))
'
'    stContext.dwRetCode = dwInternetStatus
'
'    Select Case dwInternetStatus
'        Case Is = INTERNET_STATUS_REQUEST_COMPLETE
'            Call CopyMemory(lng, ByVal pbStatusInformation + 4&, 4&)    ' dwError
'            stContext.dwErrCode = lng
'
'            If lng = 0& Then
'                Call CopyMemory(lng, ByVal pbStatusInformation, 4&)     ' bool value (is suses = true)
'                If lng = 1& Then
'
''                    Do Until InternetReadFile(hInternet, ByVal szTmpBuffer, Len(szTmpBuffer), lng) = 0&
''                        If lng = 0& Then
''                            stContext.dwExitFlag = vbNull
''                            Exit Do
''                        Else
''                            sOutBuffer = sOutBuffer & VBA.Left$(szTmpBuffer, lng)
''                        End If
''                    Loop
'                End If
'            End If
'    End Select
'
'    If Not dwContext = 0& Then Call CopyMemory(ByVal dwContext, stContext, Len(stContext))
'End Sub

Function pPingTCP(sRemoteHostPort$, Optional ByVal timeOut = 100) As Boolean
Dim udt As sockaddr_in, ar, rh$, rp$, ret&
Dim pAddrInfo As Long, tm As Double
Dim hSocket&
If Not InitWinsockService Then Exit Function
rp = xMain.SplitIndex(sRemoteHostPort, ":")
rh = Left(sRemoteHostPort, Len(sRemoteHostPort) - Len(rp) - 1)
ret = getaddrinfo(rh, rp, ByVal 0&, pAddrInfo)
If ret = 0& Then
    Dim ai As addrinfo
    CopyMemory ai, ByVal pAddrInfo, LenB(ai) 'Get the first addrinfo structure
    CopyMemory udt, ByVal ai.ai_addr, ai.ai_addrlen 'Get the sockaddr structure
    
    hSocket = api_socket(2, 1, 6)    'AF_INET/STREAM/TCP
    'Const FIONBIO As Long = &H8004667E
    ret = ioctlsocket(hSocket, &H8004667E, 1)
    ret = api_connect(hSocket, udt, Len(udt))
    
    tm = Timer
    While Timer - tm < timeOut / 1000
    'Debug.Print pSocketPort(hSocket), WSAGetLastError()
    DoEvents
    Wend
    ret = 0
    ioctlsocket hSocket, &H8004667E, 0

    ret = api_send(hSocket, rp, Len(rp), 1&)
    pPingTCP = ret = Len(rp)
    api_shutdown hSocket, 2
    api_closesocket hSocket
    'SocketClose hSocket, 0
'WSACleanup
freeaddrinfo ai  '<-- Don't forget to free the allocated structures!
End If
End Function


'Function pPingTCP(ByVal RemoteHostPort$, ByVal timeOut&) As Boolean
'pPingTCP = Routine(RemoteHostPort)
'Exit Function
'
'Dim udt As sockaddr_in, ar, rh$, rp$
'If Not InitWinsockService Then Exit Function
'Dim hSocket&
'hSocket = api_socket(2, 1, 6)  'AF_INET/STREAM/TCP
'
'If hSocket = -1 Then Exit Function
'rp = xMain.SplitIndex(RemoteHostPort, ":")
'rh = Left(RemoteHostPort, Len(rp) + 1)
''If UBound(ar) > 0 Then rh = ar(0): rp = aVal(ar(1)) Else rp = 80
'
''udt.sin_addr = pIPAddress(rh, -1)
''If udt.sin_addr = -1 Then Exit Function
''udt.sin_port = htons(LOWORD(rp))
'
'getaddrinfo rh, rp, 1, 2
'
'
'
'udt.sin_family = 2 'AF_INET
'
'Dim ret&
'
'Const FIONBIO As Long = &H8004667E
'ret = ioctlsocket(hSocket, FIONBIO, 1)
'
'ret = api_connect(hSocket, udt, Len(udt))
''
'Dim b$
''ar = StrConv(txData, vbFromUnicode)
''b = "."
''ret = 0: ret = api_send(hSocket, b, 1, 1&)
'
'Debug.Print pSocketPort(hSocket), WSAGetLastError()
'If ret = -1 Then 'ERORR
'
'Debug.Print
'
'Else
'Debug.Print
'
'End If
'
'ioctlsocket hSocket, FIONBIO, 0
'api_shutdown hSocket, 2
''SocketClose hSocket, 0
'api_closesocket hSocket
'
'End Function

'
'Function pPingTCP2(ByVal RemoteHostPort$, ByVal timeOut&, Optional OutIPAddress) As Boolean
'Dim hSocket&
'Dim udt As sockaddr_in, ar, rh$, rp&
'If Not InitWinsockService Then Exit Function
'
'''https://msdn.microsoft.com/en-us/library/windows/desktop/ms740506(v=vs.85).aspx
''AF_UNSPEC = 0 'The address family is unspecified.
''AF_INET = 2 'The Internet Protocol version 4 (IPv4) address family.
''AF_IPX = 6 'The IPX/SPX address family.
''AF_APPLETALK = 16 'The AppleTalk address family.
''AF_NETBIOS = 17 'The NetBIOS address family.
''AF_INET6 = 23 'The Internet Protocol version 6 (IPv6) address family.
''AF_IRDA = 26 'The Infrared Data Association (IrDA) address family.
''AF_BTH = 32 'The Bluetooth address family
''
''SOCK_STREAM = 1 'TCP (AF_INET or AF_INET6)
''SOCK_DGRAM = 2  'UDP (AF_INET or AF_INET6)
''SOCK_RAW = 3
''SOCK_RDM = 4
''SOCK_SEQPACKET = 5
''
''IPPROTO_ICMP = 1 'AF_UNSPEC, AF_INET, or AF_INET6 and the type parameter is SOCK_RAW
''IPPROTO_IGMP = 2 'AF_UNSPEC, AF_INET, or AF_INET6 and the type parameter is SOCK_RAW
''BTHPROTO_RFCOMM = 3 'AF_BTH and the type parameter is SOCK_STREAM
''IPPROTO_TCP = 6 'AF_INET or AF_INET6 and the type parameter is SOCK_STREAM
''IPPROTO_UDP = 17 'AF_INET or AF_INET6 and the type parameter is SOCK_DGRAM
''IPPROTO_ICMPV6 = 58 'AF_UNSPEC, AF_INET, or AF_INET6 and the type parameter is SOCK_RAW
''IPPROTO_RM = 113 'AF_INET and the type parameter is SOCK_RDM
'
'hSocket = api_socket(2, 1, 0 + 0 * 6) 'AF_INET/STREAM/TCP
'
'If hSocket = -1 Then Exit Function
'ar = Split(RemoteHostPort, ":")
'If UBound(ar) > 0 Then rh = ar(0): rp = aVal(ar(1)) Else rp = 80
'udt.sin_addr = inet_addr(pIPAddress("" & rh, 1))
'If udt.sin_addr = -1 Then Exit Function
'udt.sin_port = htons(LOWORD(rp))
'udt.sin_family = 2 'AF_INET
'
'Dim ret&
'
'Const FIONBIO As Long = &H8004667E
'ret = ioctlsocket(hSocket, FIONBIO, 1)
'
''Const WSAEWOULDBLOCK As Long = 10035
''Const SOCKET_ERROR As Long = -1
'
''nSockets = nSockets + 1
'If api_connect(hSocket, udt, Len(udt)) = -1 Then 'ERORR
'
'Debug.Print pSocketPort(hSocket), WSAGetLastError()
'
'    If WSAGetLastError() = 10035 Then
'        Dim fdsW As FD_SET, fdsR As FD_SET, fdsE As FD_SET
'        fdsW.fd_count = 1
'        fdsW.fd_array(0) = hSocket
'        Dim nTime As TIME_VAL
'        nTime.tv_sec = timeOut \ 1000
'        nTime.tv_usec = 1000 * timeOut 'timeout Mod 1000
'        Dim rc&
'        rc = sselect(0, fdsR, fdsW, fdsE, nTime)
'        pPingTCP2 = rc > 0
'    End If
'
'
'End If
'ioctlsocket hSocket, FIONBIO, 0
'api_shutdown hSocket, 2
'SocketClose hSocket, 0
'End Function

Function UDPList(Optional ByVal nPID&, Optional ByVal nRet&) As String
Dim ptr&, sz&
Dim buf() As Long
Dim ret&, i&, p&, res$, s$
ret = GetExtendedUdpTable(ptr, sz, 0, 2, 1, 0)
If sz = 0 Then Exit Function
ReDim buf(sz \ 4)
ret = GetExtendedUdpTable(buf(0), sz, 0, 2, 1, 0)
If ret Then Exit Function
If buf(0) = 0 Then Exit Function
ReDim Preserve buf(buf(0) * 3)
Dim cp As New CParam
For i = 1 To 3 * buf(0) Step 3 'IIf(bTCP, 6, 3)
    p = ntohs(LOWORD(buf(i + 1))) 'PORT NUMBER
    If (nPID = 0 Or nPID = buf(i + 2)) And p >= 999 And p < 1100 Then 'JSON PID\PORT=IPADDRESS
        If nRet = 2 Then
            res = res & IIf(Len(res), ";", "") & p
        Else
            s = StringFromPointer(inet_ntoa(buf(i))) 'IP ADDRESS
            If nRet = 1 Then cp(buf(i + 2) & "\" & p) = s Else res = res & IIf(Len(res), ";", "") & buf(i + 2) & "," & p & "," & s
        End If
    End If
Next
'Debug.Print cp.json
If nRet = 1 Then 'JSON  PID\PORT=IPADDRESS
    UDPList = cp.json
ElseIf nRet = 2 Then 'LIST UDP ports
    UDPList = res
Else 'MATRIX PID,PORT,ADDR
    UDPList = Replace(Replace("pid$,port$,addr$;" & res, ";", vbCrLf), ",", vbTab)
End If
End Function

'Public Function LANList(ByVal sCommentSeparator As String)
'    Dim nRet As Long, X As Integer
'    Dim tServerInfo As SERVER_INFO_API
'    Dim lServerInfo As Long
'    Dim lServerInfoPtr As Long
'    Dim lPreferedMaxLen As Long
'    Dim lEntriesRead As Long
'    Dim lTotalEntries As Long
'    Dim sDomain As String
'    Dim vResume As Variant
'    Dim res$
'    lPreferedMaxLen = 65536
'
'nRet = 234
'Do While (nRet = 234)
'    nRet = NetServerEnum(0, 101, lServerInfo, lPreferedMaxLen, lEntriesRead, lTotalEntries, &HFF, sDomain, vResume)
'    If (nRet <> 0 And nRet <> 234) Then Exit Do
'    X = 1
'    lServerInfoPtr = lServerInfo
'    Do While X <= lTotalEntries
'        CopyMemory tServerInfo, ByVal lServerInfoPtr, Len(tServerInfo)
'        res = res & IIf(Len(res), ";", "") & StringWFromPointer(tServerInfo.ServerName)
'        If Len(sCommentSeparator) Then res = res & sCommentSeparator & StringWFromPointer(tServerInfo.Comment)
'        X = X + 1
'        lServerInfoPtr = lServerInfoPtr + Len(tServerInfo)
'    Loop
'    NetApiBufferFree lServerInfo
'Loop
'LANList = res
'End Function

