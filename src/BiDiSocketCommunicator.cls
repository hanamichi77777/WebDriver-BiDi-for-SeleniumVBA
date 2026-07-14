VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BiDiSocketCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' WebDriver BiDi for SeleniumVBA
' https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA
' Version 2.7 / MIT License / Copyright (c) hanamichi77777
'
' BiDiSocketCommunicator owns loopback WebSocket transport, RFC 6455 framing,
' request/system IDs, event FIFO and monotonic timing. Retry classification,
' context recovery and SPA synchronization belong to BiDiCommandWrapper.
'
' Preserve these invariants:
'   - Accept only local unencrypted ws:// endpoints.
'   - One SendAndReceive call owns response correlation; do not pump Office messages.
'   - m_isReading and m_isRecovering are released only by structured cleanup.
'   - IDs [1..99] are system traffic; normal commands use [100..ID_MAX].
'   - Recovery may heal the socket but never replay the same payload/request ID.
'   - Every native Pump is non-blocking and bounded by time and work.
'   - Do not call Trace from a cleanup-critical error handler.
Private Const DEBUG_MODE As Boolean = False
Private Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' IDs [1..99] are reserved for PostScript and rotating heartbeat traffic.
' Commands use [CMD_ID_FLOOR..ID_MAX]; id-less BiDi events are not system replies.
' Reserved-band replies are consumed at both receive sites.
Private Const CMD_ID_FLOOR As Long = 100         ' command ids: [100 .. ID_MAX]
Private Const ID_MAX As Long = 2000000000        ' wrap ceiling (within VBA Long range)
Private Const POST_SCRIPT_REPLY_ID As Long = 2   ' reserved fixed id: PostScript fire-and-forget reply (discarded)
Private Const SYSTEM_PING_ID_MIN As Long = 3     ' heartbeat rotating id band floor
Private Const SYSTEM_PING_ID_MAX As Long = 99    ' heartbeat rotating id band ceiling
Private Const MAX_CONSECUTIVE_SOFT_TIMEOUTS As Long = 2 ' soft timeouts on one connection before StateError
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 100

Private m_wsTransportInitialized As Boolean
' Only symbolic localhost may fall back from [::1] to 127.0.0.1.
' Explicit numeric loopback addresses preserve the caller's address family.
Private m_WinsockEffectiveWsUrl As String
Private m_WinsockIpv4FallbackUsed As Boolean
Private m_WinsockPrimaryFailureText As String
Private Const WINSOCK_CONNECT_TIMEOUT_MS As Long = 2000
Private Const WINSOCK_HANDSHAKE_TIMEOUT_MS As Long = 5000
Private Const WINSOCK_SEND_TIMEOUT_MS As Long = 5000
Private Const WINSOCK_PUMP_MAX_MS As Long = 10
Private Const WINSOCK_PUMP_MAX_BYTES As Long = 262144
Private Const WINSOCK_PUMP_MAX_CALLS As Long = 16

Private m_LastWsUrl As String
Private m_LastPort As Long
Private m_RetryCount As Long
Private m_MaxRetryAttempts As Long
Private m_AutoReconnect As Boolean
Private m_consecutiveInternalErrors As Long

Private m_isReading As Boolean ' Non-yielding receive/Pump guard; structural release only
Private m_isSendReceiving As Boolean ' True while one command owns response correlation
Private m_isRecovering As Boolean ' Prevents nested connection recovery attempts

Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum
Private m_state As BiDiSocketState
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal StateName As String)

Private m_EventBuffer As Collection

Private m_DefaultTimeout As Long
Private m_PollingInterval As Long
Private m_DroppedEventsCount As Long
Private m_lastPollTime As Double
Private m_RegExp As RegExp
Private m_AllowUiPump As Boolean
Private m_qpcFreq As Currency ' QueryPerformanceFrequency cache (set in Class_Initialize)
Private m_lastPumpTicks As Currency ' Throttle state for PumpMessages

Private m_nextRequestId As Long
Private m_nextSystemPingId As Long ' heartbeat id rotation cursor [3..99]
Private m_consecutiveSoftTimeouts As Long ' consecutive soft timeouts on the current connection

Private Const MAX_MSG_SIZE As Long = 300000 ' Threshold to prevent VBA hanging (mirrors Wrapper constant)
Private Const DEFAULT_MAX_INCOMING_SIZE As Long = 2097152 ' 2MB: receive-side assembly cap (runtime-tunable via MaxIncomingSize)
Private m_maxIncomingSize As Long ' Receive-side cap; oversized messages are discarded with identity diagnostics

' Ws*/WS_*/m_ws_* members form the RFC 6455 engine: socket, handshake,
' framing, masking and WebSocket ping/pong. Each Pump is work-bounded; BiDi IDs,
' heartbeats, recovery and the event FIFO remain in the outer communicator layer.

Private Const WS_DEBUG_MODE As Boolean = True

Private Declare PtrSafe Function WsApi_WSAStartup Lib "ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Integer, ByRef lpWSAData As WsWSADATA) As Long
Private Declare PtrSafe Function WsApi_WSACleanup Lib "ws2_32.dll" Alias "WSACleanup" () As Long
Private Declare PtrSafe Function WsApi_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socketType As Long, ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function WsApi_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
Private Declare PtrSafe Function WsApi_connectSocket Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef socketAddress As Any, ByVal nameLength As Long) As Long
Private Declare PtrSafe Function WsApi_sendSocket Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef Buffer As Any, ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_recvSocket Lib "ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByRef Buffer As Any, ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As LongPtr, ByVal command As Long, ByRef argument As Long) As Long
Private Declare PtrSafe Function WsApi_selectSocket Lib "ws2_32.dll" Alias "select" (ByVal ignored As Long, ByRef readSet As WsFD_SET, ByRef writeSet As WsFD_SET, ByRef exceptSet As WsFD_SET, ByRef TIMEOUT As WsTIMEVAL) As Long
Private Declare PtrSafe Function WsApi_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optionName As Long, ByRef optionValue As Long, ByRef optionLength As Long) As Long
Private Declare PtrSafe Function WsApi_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optionName As Long, ByRef optionValue As Any, ByVal optionLength As Long) As Long
Private Declare PtrSafe Function WsApi_shutdownSocket Lib "ws2_32.dll" Alias "shutdown" (ByVal s As LongPtr, ByVal how As Long) As Long
Private Declare PtrSafe Function WsApi_WSAGetLastError Lib "ws2_32.dll" Alias "WSAGetLastError" () As Long
Private Declare PtrSafe Function WsApi_InetPtonW Lib "ws2_32.dll" Alias "InetPtonW" (ByVal addressFamily As Long, ByVal addressText As LongPtr, ByRef addressBuffer As Any) As Long

Private Declare PtrSafe Sub WsApi_SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function WsApi_QueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef performanceCount As Currency) As Long
Private Declare PtrSafe Function WsApi_QueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef frequency As Currency) As Long

Private Declare PtrSafe Function WsApi_BCryptGenRandom Lib "bcrypt.dll" Alias "BCryptGenRandom" (ByVal hAlgorithm As LongPtr, ByRef randomBuffer As Any, ByVal bufferSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptOpenAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptOpenAlgorithmProvider" (ByRef algorithmHandle As LongPtr, ByVal algorithmId As LongPtr, ByVal implementation As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptCloseAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptCloseAlgorithmProvider" (ByVal algorithmHandle As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptGetProperty Lib "bcrypt.dll" Alias "BCryptGetProperty" (ByVal objectHandle As LongPtr, ByVal propertyName As LongPtr, ByRef outputBuffer As Any, ByVal outputSize As Long, ByRef resultSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptCreateHash Lib "bcrypt.dll" Alias "BCryptCreateHash" (ByVal algorithmHandle As LongPtr, ByRef hashHandle As LongPtr, ByRef hashObject As Any, ByVal hashObjectSize As Long, ByVal secret As LongPtr, ByVal secretSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptHashData Lib "bcrypt.dll" Alias "BCryptHashData" (ByVal hashHandle As LongPtr, ByRef inputBuffer As Any, ByVal inputSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptFinishHash Lib "bcrypt.dll" Alias "BCryptFinishHash" (ByVal hashHandle As LongPtr, ByRef outputBuffer As Any, ByVal outputSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptDestroyHash Lib "bcrypt.dll" Alias "BCryptDestroyHash" (ByVal hashHandle As LongPtr) As Long


Private Type WsWSADATA
    version As Integer
    highVersion As Integer
    description(0 To 256) As Byte
    systemStatus(0 To 128) As Byte
    maxSockets As Integer
    maxUdpDg As Integer
    vendorInfo As LongPtr
End Type

Private Type WsSOCKADDR_IN
    family As Integer
    port As Integer
    address As Long
    zero(0 To 7) As Byte
End Type

Private Type WsSOCKADDR_IN6
    family As Integer
    port As Integer
    flowInfo As Long
    address(0 To 15) As Byte
    scopeId As Long
End Type

Private Type WsFD_SET
    count As Long
    sockets(0 To 63) As LongPtr
End Type

Private Type WsTIMEVAL
    seconds As Long
    microseconds As Long
End Type

Private Const WS_AF_INET As Long = 2
Private Const WS_AF_INET6 As Long = 23
Private Const WS_SOCK_STREAM As Long = 1
Private Const WS_IPPROTO_TCP As Long = 6
Private Const WS_TCP_NODELAY As Long = 1
Private Const WS_SOCKET_ERROR As Long = -1
Private Const WS_FIONBIO As Long = &H8004667E
Private Const WS_SOL_SOCKET As Long = &HFFFF&
Private Const WS_SO_ERROR As Long = &H1007&
Private Const WS_SD_BOTH As Long = 2

Private Const WS_WSAEWOULDBLOCK As Long = 10035
Private Const WS_WSAEINPROGRESS As Long = 10036
Private Const WS_WSAEALREADY As Long = 10037
Private Const WS_WSAEISCONN As Long = 10056
Private Const WS_WSAECONNRESET As Long = 10054
Private Const WS_WSAECONNABORTED As Long = 10053
Private Const WS_WSAENOTCONN As Long = 10057

Private Const WS_OPCODE_CONTINUATION As Long = 0
Private Const WS_OPCODE_TEXT As Long = 1
Private Const WS_OPCODE_BINARY As Long = 2
Private Const WS_OPCODE_CLOSE As Long = 8
Private Const WS_OPCODE_PING As Long = 9
Private Const WS_OPCODE_PONG As Long = 10

Private Const WS_DEFAULT_CONNECT_TIMEOUT_MS As Long = 2000
Private Const WS_DEFAULT_HANDSHAKE_TIMEOUT_MS As Long = 5000
Private Const WS_DEFAULT_MAX_INCOMING_SIZE As Long = 2097152
Private Const WS_DEFAULT_MAX_PUMP_MS As Long = 10
Private Const WS_DEFAULT_MAX_PUMP_BYTES As Long = 262144
Private Const WS_DEFAULT_MAX_PUMP_CALLS As Long = 16
Private Const WS_RECEIVE_CHUNK_SIZE As Long = 16384
Private Const WS_PARSER_COPY_CHUNK_SIZE As Long = 4096
Private Const WS_PARSER_COPY_GUARD_MS As Double = 1#
Private Const WS_MAX_COMPLETED_MESSAGES_PER_PUMP As Long = 8
Private Const WS_MAX_HTTP_HEADER_SIZE As Long = 32768

Private Const WS_BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2
Private Const WS_BCRYPT_STATUS_SUCCESS As Long = 0
Private Const WS_MASK_KEY_SIZE As Long = 4

Private Const WS_CLIENT_PING_INTERVAL_MS As Long = 15000
Private Const WS_CLIENT_PING_PAYLOAD_SIZE As Long = 4
Private Const WS_CLIENT_PING_BUSY_RETRY_MS As Long = 1000

Private Const WS_HANDSHAKE_NONCE_SIZE As Long = 16
Private Const WS_WEBSOCKET_GUID As String = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
Private Const WS_BCRYPT_SHA1_ALGORITHM As String = "SHA1"
Private Const WS_BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const WS_BCRYPT_HASH_LENGTH As String = "HashDigestLength"

Private Const WS_HANDSHAKE_ORIGIN_SCHEME As String = "http://"
Private Const WS_HANDSHAKE_USER_AGENT As String = _
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"


Private Enum WsWinsockTransportState
    WsTransportDisconnected = 0
    WsTransportConnecting = 1
    WsTransportHandshakeSending = 2
    WsTransportHandshakeReceiving = 3
    WsTransportOpen = 4
    WsTransportClosing = 5
    WsTransportFailed = 6
End Enum

Private Enum WsFrameParserState
    WsFrameReadBaseHeader = 0
    WsFrameReadExtendedLength = 1
    WsFrameReadPayload = 2
    WsFrameFinalize = 3
End Enum

Private m_ws_state As WsWinsockTransportState
Private m_ws_socket As LongPtr
Private m_ws_wsaStarted As Boolean
Private m_ws_qpcFrequency As Currency

Private m_ws_url As String
Private m_ws_host As String
Private m_ws_hostHeader As String
Private m_ws_path As String
Private m_ws_port As Long
Private m_ws_phaseDeadlineMs As Double
Private m_ws_handshakeTimeoutMs As Long

Private m_ws_handshakeClientKey As String
Private m_ws_expectedHandshakeAccept As String
Private m_ws_handshakeOrigin As String
Private m_ws_handshakeUserAgent As String
Private m_ws_handshakeValidationError As String
Private m_ws_LastErrorCode As Long
Private m_ws_lastErrorText As String

Private m_ws_sendQueue As Collection
Private m_ws_sendCurrent() As Byte
Private m_ws_sendCurrentCount As Long
Private m_ws_sendOffset As Long

Private m_ws_rawBuffer() As Byte
Private m_ws_rawCount As Long
Private m_ws_rawReadPos As Long

Private m_ws_rawTextQueue As Collection
Private m_ws_rawTextLengthQueue As Collection
Private m_ws_maxIncomingSize As Long

Private m_ws_frameState As WsFrameParserState
Private m_ws_frameFin As Boolean
Private m_ws_frameOpcode As Long
Private m_ws_frameLengthCode As Long
Private m_ws_framePayloadLength As Long
Private m_ws_framePayloadRemaining As Long
Private m_ws_frameDiscardPayload As Boolean
Private m_ws_controlPayload() As Byte
Private m_ws_controlCount As Long

Private m_ws_messageActive As Boolean
Private m_ws_messageOpcode As Long
Private m_ws_messageChunks As Collection
Private m_ws_messageCount As Long
Private m_ws_discardMessage As Boolean

Private m_ws_closeSent As Boolean

Private m_ws_currentPumpCompletedMessages As Long

Private m_ws_clientPingEnabled As Boolean
Private m_ws_nextClientPingAtMs As Double
Private m_ws_clientPingOutstanding As Boolean
Private m_ws_clientPingSequence As Long
Private m_ws_clientPingOutstandingSequence As Long

Private Sub Class_Initialize()
    If QueryPerformanceFrequency(m_qpcFreq) = 0 Or m_qpcFreq <= 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", _
                  "QueryPerformanceFrequency is unavailable on this system. QPC support is required."
    End If

    m_DefaultTimeout = DEFAULT_TIMEOUT_MS
    m_PollingInterval = DEFAULT_POLLING_INTERVAL_MS
    m_AllowUiPump = True
    m_maxIncomingSize = DEFAULT_MAX_INCOMING_SIZE

    Set m_EventBuffer = New Collection

    Set m_RegExp = New RegExp
    m_RegExp.Global = False
    m_RegExp.IgnoreCase = True

    m_MaxRetryAttempts = 5
    m_AutoReconnect = True

    m_nextRequestId = CMD_ID_FLOOR - 1
    m_nextSystemPingId = SYSTEM_PING_ID_MIN - 1

    SetState StateDisconnected
End Sub

Private Sub Class_Terminate()
    ' Teardown is non-raising. Clear communicator-owned guards before native handles
    ' so an interrupted call cannot poison later reuse.
    On Error Resume Next
    m_isSendReceiving = False
    m_isReading = False

    Disconnect
    If m_wsTransportInitialized Then
        WsTransportTerminate
        m_wsTransportInitialized = False
    End If
    Set m_RegExp = Nothing
    Set m_EventBuffer = Nothing
    On Error GoTo 0
End Sub

Public Property Get IsConnected() As Boolean
    If m_wsTransportInitialized Then
        IsConnected = WsIsOpen
    End If
End Property

' Compatibility shim: WinSock is the only transport; False is rejected and True
' is accepted only while disconnected.
Public Property Let UseWinsockTransport(ByVal v As Boolean)
    If Not v Then
        Err.Raise vbObjectError + 9105, "BiDiSocketCommunicator", _
                  "The legacy WinHTTP transport has been removed. WinSock is the only supported transport."
    End If
    If m_state <> StateDisconnected Then
        Err.Raise vbObjectError + 9102, "BiDiSocketCommunicator", _
                  "Transport mode can be confirmed only while disconnected."
    End If
End Property

Public Property Get UseWinsockTransport() As Boolean
    UseWinsockTransport = True
End Property

Public Property Get WinsockEffectiveWebSocketUrl() As String
    WinsockEffectiveWebSocketUrl = m_WinsockEffectiveWsUrl
End Property

Public Property Get WinsockIpv4FallbackUsed() As Boolean
    WinsockIpv4FallbackUsed = m_WinsockIpv4FallbackUsed
End Property

Public Property Get WinsockPrimaryFailureText() As String
    WinsockPrimaryFailureText = m_WinsockPrimaryFailureText
End Property

Public Property Let PollingInterval(ByVal v As Long): m_PollingInterval = v: End Property
Public Property Let AllowUiPump(ByVal v As Boolean): m_AllowUiPump = v: End Property
Public Property Get AllowUiPump() As Boolean: AllowUiPump = m_AllowUiPump: End Property
Public Property Get state() As BiDiSocketState: state = m_state: End Property
Public Property Let AutoReconnect(ByVal v As Boolean): m_AutoReconnect = v: End Property
Public Property Get AutoReconnect() As Boolean: AutoReconnect = m_AutoReconnect: End Property
Public Property Get ConsecutiveInternalErrors() As Long: ConsecutiveInternalErrors = m_consecutiveInternalErrors: End Property
' MaxIncomingSize is a receive-assembly byte cap and is clamped to MAX_MSG_SIZE
' so valid command responses are not rejected by configuration.
Public Property Let MaxIncomingSize(ByVal v As Long)
    If v < MAX_MSG_SIZE Then v = MAX_MSG_SIZE
    m_maxIncomingSize = v

    If m_wsTransportInitialized Then
        WsMaxIncomingSize = v
    End If

End Property
Public Property Get MaxIncomingSize() As Long: MaxIncomingSize = m_maxIncomingSize: End Property
Public Property Get DroppedEventsCount() As Long: DroppedEventsCount = m_DroppedEventsCount: End Property

Private Sub SetState(ByVal NewState As BiDiSocketState)
    If m_state = NewState Then Exit Sub
    
    m_state = NewState
    
    Dim StateName As String
    Select Case m_state
        Case StateDisconnected: StateName = "DISCONNECTED"
        Case StateConnecting:   StateName = "CONNECTING"
        Case StateConnected:    StateName = "CONNECTED"
        Case StateDisconnecting: StateName = "DISCONNECTING"
        Case StateError:         StateName = "ERROR"
    End Select
    
    If DEBUG_MODE Then Debug.Print "BiDi State Changed: " & StateName
    RaiseEvent StateChanged(m_state, StateName)
End Sub

Public Function Connect(ByVal wsUrl As String, ByVal localPort As Long) As Boolean
    If m_state = StateConnecting Or m_state = StateConnected Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect aborted. Current state is " & m_state
        Exit Function
    End If

    ' Reject anything except loopback ws:// before entering the RFC 6455 parser.
    m_RegExp.pattern = "^ws://(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?(?:/|$)"
    m_RegExp.IgnoreCase = True

    If Not m_RegExp.Test(wsUrl) Then
        Trace "[Security] Blocked: endpoint must be a ws:// loopback URL " & _
              "using localhost, 127.0.0.1, or [::1].", True
        SetState StateError
        Exit Function
    End If

    m_LastWsUrl = wsUrl
    m_LastPort = localPort

    On Error GoTo ErrorHandler

    SetState StateConnecting
    InternalCleanup

    If LenB(wsUrl) = 0 Or localPort = 0 Then
        Err.Raise 512, "BiDiSocket", "Invalid WebSocket URL or port."
    End If

    If ConnectWithWinsock(wsUrl) Then
        Connect = True
        Exit Function
    End If

    Err.Raise 521, "BiDiSocket", _
              "Winsock WebSocket connection failed: " & GetWinsockFailureText()

ErrorHandler:
    SetState StateError
    InternalCleanup
    Connect = False
End Function

Private Function ConnectWithWinsock(ByVal wsUrl As String) As Boolean
    On Error GoTo ErrorHandler

    m_WinsockEffectiveWsUrl = vbNullString
    m_WinsockIpv4FallbackUsed = False
    m_WinsockPrimaryFailureText = vbNullString

    Dim primaryUrl As String
    Dim fallbackUrl As String
    Dim originalHost As String
    originalHost = LCase$(ExtractHost(wsUrl))

    Select Case originalHost
        Case "localhost"
            ' For symbolic localhost, try [::1] first and fall back to 127.0.0.1 only if
            ' failure occurs before OPEN. Explicit numeric hosts never switch family.
            primaryUrl = ReplaceLoopbackHost(wsUrl, "[::1]")
            fallbackUrl = ReplaceLoopbackHost(wsUrl, "127.0.0.1")

        Case "::1", "127.0.0.1"
            primaryUrl = wsUrl
        Case Else
            primaryUrl = wsUrl
    End Select

    If TryConnectWithWinsockUrl(primaryUrl) Then
        ConnectWithWinsock = True
        Exit Function
    End If

    If Len(fallbackUrl) > 0 Then
        m_WinsockPrimaryFailureText = GetWinsockFailureText()
        Trace "BiDi Winsock: IPv6 loopback failed before OPEN; trying IPv4 fallback. " & _
              m_WinsockPrimaryFailureText, True

        If TryConnectWithWinsockUrl(fallbackUrl) Then
            m_WinsockIpv4FallbackUsed = True
            ConnectWithWinsock = True
            Exit Function
        End If
    End If

    Exit Function

ErrorHandler:
    m_ws_LastErrorCode = Err.Number
    m_ws_lastErrorText = "ConnectWithWinsock VBA error: " & Err.description
    ConnectWithWinsock = False
End Function

Private Function TryConnectWithWinsockUrl(ByVal candidateUrl As String) As Boolean
    On Error GoTo ErrorHandler

    Dim diagnosticStage As String
    diagnosticStage = "entry"

    If Not m_wsTransportInitialized Then
        diagnosticStage = "WsTransportInitialize"
        WsTransportInitialize
        m_wsTransportInitialized = True
    End If

    diagnosticStage = "WsMaxIncomingSize"
    WsMaxIncomingSize = m_maxIncomingSize

    diagnosticStage = "WsBeginConnect"
    If Not WsBeginConnect( _
        candidateUrl, WINSOCK_CONNECT_TIMEOUT_MS, WINSOCK_HANDSHAKE_TIMEOUT_MS) Then
        Exit Function
    End If

    Dim deadline As Double
    deadline = nowMs() + WINSOCK_CONNECT_TIMEOUT_MS + WINSOCK_HANDSHAKE_TIMEOUT_MS + 1000

    Do
        diagnosticStage = "WsPump"
        WsPump WINSOCK_PUMP_MAX_MS, _
                                WINSOCK_PUMP_MAX_BYTES, _
                                WINSOCK_PUMP_MAX_CALLS

        If WsIsOpen Then
            m_WinsockEffectiveWsUrl = candidateUrl
            FinalizeConnectedState
            Trace "BiDi Winsock: OPEN via " & WinsockEndpointFamily(candidateUrl) & _
                  " loopback. Effective URL: " & candidateUrl
            TryConnectWithWinsockUrl = True
            Exit Function
        End If

        If Wsstate = WsTransportFailed Then
            Exit Function
        End If
        If nowMs() >= deadline Then
            m_ws_LastErrorCode = 9291
            m_ws_lastErrorText = "Outer connection deadline expired while state=" & WsStateName
            Exit Function
        End If

        SleepWinAPI 0
    Loop

ErrorHandler:
    m_ws_LastErrorCode = Err.Number
    m_ws_lastErrorText = "TryConnectWithWinsockUrl VBA error at " & diagnosticStage & ": " & Err.description
    TryConnectWithWinsockUrl = False
End Function

Private Function ReplaceLoopbackHost(ByVal wsUrl As String, _
                                     ByVal replacementHost As String) As String
    Dim schemeEnd As Long
    schemeEnd = InStr(1, wsUrl, "://", vbTextCompare)
    If schemeEnd = 0 Then Exit Function

    Dim authorityStart As Long
    authorityStart = schemeEnd + 3

    Dim pathStart As Long
    pathStart = InStr(authorityStart, wsUrl, "/", vbBinaryCompare)
    If pathStart = 0 Then pathStart = Len(wsUrl) + 1

    Dim authority As String
    authority = Mid$(wsUrl, authorityStart, pathStart - authorityStart)
    If Len(authority) = 0 Then Exit Function

    Dim portPart As String
    If Left$(authority, 1) = "[" Then
        Dim closeBracket As Long
        closeBracket = InStr(1, authority, "]", vbBinaryCompare)
        If closeBracket = 0 Then Exit Function
        portPart = Mid$(authority, closeBracket + 1)
    Else
        Dim colonPos As Long
        colonPos = InStrRev(authority, ":", -1, vbBinaryCompare)
        If colonPos > 0 Then portPart = Mid$(authority, colonPos)
    End If

    ReplaceLoopbackHost = Left$(wsUrl, authorityStart - 1) & _
                          replacementHost & portPart & Mid$(wsUrl, pathStart)
End Function

Private Function WinsockEndpointFamily(ByVal wsUrl As String) As String
    If LCase$(ExtractHost(wsUrl)) = "::1" Then
        WinsockEndpointFamily = "IPv6"
    Else
        WinsockEndpointFamily = "IPv4"
    End If
End Function

Private Sub FinalizeConnectedState()
    m_DroppedEventsCount = 0
    m_lastPollTime = 0
    m_RetryCount = 0
    m_consecutiveSoftTimeouts = 0
    SetState StateConnected
End Sub

Private Function GetWinsockFailureText() As String
    If m_ws_LastErrorCode <> 0 Or Len(m_ws_lastErrorText) > 0 Then
        GetWinsockFailureText = CStr(m_ws_LastErrorCode) & " / " & m_ws_lastErrorText
    ElseIf Not m_wsTransportInitialized Then
        GetWinsockFailureText = "integrated transport was not initialized"
    Else
        GetWinsockFailureText = CStr(WsLastErrorCode) & _
                                " / " & WsLastErrorText
    End If
End Function

Private Function SendMessageWinsock(ByVal msg As String) As Boolean
    If Not m_wsTransportInitialized Then Exit Function
    If Not WsIsOpen Then Exit Function
    If Not WsQueueText(msg) Then Exit Function

    Dim deadline As Double
    deadline = nowMs() + WINSOCK_SEND_TIMEOUT_MS

    Do While WsHasPendingOutput
        WsPump WINSOCK_PUMP_MAX_MS, _
               WINSOCK_PUMP_MAX_BYTES, _
               WINSOCK_PUMP_MAX_CALLS

        If Wsstate = WsTransportFailed Then Exit Function
        If nowMs() >= deadline Then Exit Function

        SleepWinAPI 0
    Loop

    If Not WsIsOpen Then Exit Function
    SendMessageWinsock = True
End Function

' allowRecovery=False is PostScript's best-effort send-only path. Recovery may
' repair a failed socket, but SendMessage returns False and never replays the payload.
Public Function SendMessage(ByVal msg As String, Optional ByVal allowRecovery As Boolean = True) As Boolean
    If LenB(msg) = 0 Then Exit Function

    Trace ">>> [Socket] Attempting to send message..."

    If m_state = StateError Then
        If Not allowRecovery Then Exit Function
        If Not AttemptRecovery() Then Exit Function
    End If

    If m_state <> StateConnected Then
        Trace ">>> [Socket] Send aborted. Not connected.", True
        Exit Function
    End If

    If SendMessageWinsock(msg) Then
        SendMessage = True
        Exit Function
    End If

    Trace ">>> [Socket] Winsock send failed. " & GetWinsockFailureText(), True
    SetState StateError
    Disconnect keepErrorState:=True

    ' NO RESEND: the peer may already have accepted the request ID. Any safe retry is
    ' an upper-layer decision using a newly allocated ID.
    If allowRecovery Then Call AttemptRecovery
    SendMessage = False
End Function

Public Sub PostScript(ByVal msg As String)
    SendMessage msg, False
End Sub

Public Function GetNextRequestId() As Long
    m_nextRequestId = m_nextRequestId + 1
    If m_nextRequestId > ID_MAX Then m_nextRequestId = CMD_ID_FLOOR
    GetNextRequestId = m_nextRequestId
End Function

' Heartbeat IDs rotate inside [3..99], separate from normal command IDs; replies
' must return before a complete band rotation to keep outstanding IDs unique.
Private Function GetNextSystemPingId() As Long
    m_nextSystemPingId = m_nextSystemPingId + 1
    If m_nextSystemPingId > SYSTEM_PING_ID_MAX Then m_nextSystemPingId = SYSTEM_PING_ID_MIN
    GetNextSystemPingId = m_nextSystemPingId
End Function

' Only numeric root IDs in [1..CMD_ID_FLOOR) are reserved replies. An id-less
' BiDi event is asynchronous application traffic.
Private Function IsReservedSystemId(ByVal idStr As String) As Boolean
    If Len(idStr) = 0 Then Exit Function
    If Not IsNumeric(idStr) Then Exit Function
    Dim v As Double: v = val(idStr)
    IsReservedSystemId = (v >= 1 And v < CMD_ID_FLOOR)
End Function

Public Function SendCommand(ByVal method As String, _
                            ByRef params As Object, _
                            Optional ByVal Silent As Boolean = False) As String
    
    Dim requestId As Long: requestId = GetNextRequestId()
    
    Dim requestDic As Dictionary: Set requestDic = New Dictionary
    requestDic.Add "id", requestId
    requestDic.Add "method", method
    requestDic.Add "params", params
    
    Dim sendMsg As String
    sendMsg = WebJsonConverter.ConvertToJson(requestDic)
    
    Set requestDic = Nothing
    
    If Len(sendMsg) > MAX_MSG_SIZE Then
        Err.Raise vbObjectError + 1003, "BiDiSocket:SendCommand", _
                  "Outgoing payload size (" & Len(sendMsg) & " chars) exceeds MAX_MSG_SIZE (" & MAX_MSG_SIZE & "). " & _
                  "Reduce the input data size or split the operation."
    End If
    
    If DEBUG_MODE And Not Silent Then
        Dim displayMsg As String: displayMsg = sendMsg
        If Len(displayMsg) > 500 Then displayMsg = Left$(displayMsg, 497) & "..."
        LogEvent method & " (ID: " & requestId & ")", displayMsg, "[REQ]"
    End If
    
    Dim strRes As String
    strRes = SendAndReceive(sendMsg)
    
    If DEBUG_MODE And Not Silent Then
        Dim displayRes As String: displayRes = strRes
        If Len(displayRes) > 500 Then displayRes = Left$(displayRes, 497) & "..."
        LogEvent method & " (ID: " & requestId & ")", displayRes, "[RES]"
    End If
    
    SendCommand = strRes
End Function

Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    If m_state = StateDisconnecting Or m_state = StateDisconnected Then Exit Sub

    If m_DroppedEventsCount > 0 Then
        Trace ">>> [Session Summary] Connection closing. Total events dropped during this session: " & m_DroppedEventsCount, True
    End If

    If Not keepErrorState Then SetState StateDisconnecting
    
    Call InternalCleanup
    
    Set m_EventBuffer = New Collection
    
    If Not keepErrorState Then
        SetState StateDisconnected
    End If
End Sub

' InternalCleanup releases transport resources only; its caller owns the public
' BiDiSocketState transition.
Private Sub InternalCleanup()
    On Error Resume Next

    If m_wsTransportInitialized Then
        If WsIsOpen Then
            WsBeginClose 1000

            Dim closeDeadline As Double
            closeDeadline = nowMs() + 250

            Do While WsHasPendingOutput And nowMs() < closeDeadline
                WsPump WINSOCK_PUMP_MAX_MS, _
                                        WINSOCK_PUMP_MAX_BYTES, _
                                        WINSOCK_PUMP_MAX_CALLS
                SleepWinAPI 0
            Loop
        End If

        WsForceClose
    End If

    m_isRecovering = False

    On Error GoTo 0
End Sub

' SendAndReceive correlates one ID around bounded non-blocking Pump calls and
' suppresses Office message pumping for the entire command-response cycle.
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    ' Re-entry raises before the local error handler is armed. The outer in-flight call
    ' therefore retains m_isSendReceiving and cannot be mistaken for an empty response.
    If m_isSendReceiving Then
        Err.Raise 606, "BiDiSocket", _
            "SendAndReceive re-entry blocked (a command was issued while another was still in flight)."
    End If
    m_isSendReceiving = True
    
    Dim succeeded As Boolean: succeeded = False
    Dim msg As String, inId As String
    Dim cmdId As String
    
    On Error GoTo Cleanup

    If m_state <> StateConnected Then
        If Not AttemptRecovery() Then
            Err.Raise 600, "BiDiSocket", "Not connected and recovery failed."
        End If
    End If
    
    cmdId = ExtractJsonId(jsonMessage)
    
    If Not SendMessage(jsonMessage) Then
        If m_state <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost during send operation."
        Else
            Err.Raise 601, "BiDiSocket", "Send failed due to an internal Winsock transport error."
        End If
    End If
    
    ' Timeout policy:
    '   - 1x raises ambiguous 604 while preserving an otherwise-open socket.
    '   - Repeated soft timeouts escalate to StateError; a healthy response/connect resets them.
    '   - 2x is a secondary overrun fail-safe and marks the socket failed.
    '   - StateError cleanup releases native handles automatically only with AutoReconnect.
    ' All native calls and Pump work are bounded, so silence cannot trap VBA in receive.
    Dim TIMEOUT As Long: TIMEOUT = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim softDeadline As Double: softDeadline = nowMs() + TIMEOUT
    Dim hardDeadline As Double: hardDeadline = nowMs() + (TIMEOUT * 2)
    
    ' The command-time event FIFO is count-bounded; individual messages remain bounded
    ' by MaxIncomingSize. Oldest asynchronous events are dropped first.
    Const MAX_EVENT_BUFFER_SIZE As Long = 20000
    
    Do
        If m_state <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        msg = ReadFromWire(softDeadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                SendAndReceive = msg
                succeeded = True
                GoTo Cleanup
            Else
                ' Queue id-less/asynchronous events, but consume reserved PostScript/heartbeat
                ' replies internally rather than exposing them through GetMessage.
                If Not IsReservedSystemId(inId) Then
                    m_EventBuffer.Add msg
                    
                    If m_EventBuffer.count > MAX_EVENT_BUFFER_SIZE Then
                        m_EventBuffer.Remove 1
                        m_DroppedEventsCount = m_DroppedEventsCount + 1
                    End If
                End If
            End If
        End If
        
        Dim nowT As Double: nowT = nowMs()
        
        If nowT > hardDeadline Then
            SetState StateError
            Err.Raise 604, "BiDiSocket", "Socket Silence: Hard timeout (connection presumed lost) waiting for ID: " & cmdId
        End If
        
        If nowT > softDeadline Then
            m_consecutiveSoftTimeouts = m_consecutiveSoftTimeouts + 1
            If m_consecutiveSoftTimeouts >= MAX_CONSECUTIVE_SOFT_TIMEOUTS Then
                SetState StateError
                Err.Raise 604, "BiDiSocket", "Socket Silence: " & m_consecutiveSoftTimeouts & _
                          " consecutive soft timeouts (connection presumed lost) waiting for ID: " & cmdId
            Else
                Err.Raise 604, "BiDiSocket", "Socket Silence: Soft timeout waiting for ID: " & cmdId & _
                          ". Connection preserved for retry."
            End If
        End If

        ' While data is flowing, drain without sleeping so events cannot bury the matching
        ' response; yield only when the bounded Pump produced nothing.
        If Len(msg) = 0 Then YieldExecution 5, False
    Loop

Cleanup:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrDesc As String: savedErrDesc = Err.description
    Dim savedErrSrc As String: savedErrSrc = Err.source
    
    If succeeded Then
        savedErrNum = 0
        Err.Clear
        m_consecutiveSoftTimeouts = 0
    End If
    
    m_isSendReceiving = False
    
    ' A first soft timeout preserves StateConnected. Hard or repeated soft timeout sets
    ' StateError; AutoReconnect then releases stale native handles during cleanup.
    If m_state = StateError And m_AutoReconnect Then
        Call InternalCleanup
    End If
    
    msg = ""
    inId = ""
    
    If Not succeeded And savedErrNum <> 0 Then
        On Error GoTo 0
        Err.Raise savedErrNum, savedErrSrc, savedErrDesc
    End If
End Function

' Recovery acquires m_isRecovering before any yield; every precondition, success and
' error path converges on cleanup so nested COM/timer re-entry cannot retain the lock.
Private Function AttemptRecovery() As Boolean
    If m_isRecovering Then Exit Function
    m_isRecovering = True
    
    On Error GoTo ErrorHandler

    If Not m_AutoReconnect Then GoTo CleanExit
    If m_RetryCount >= m_MaxRetryAttempts Then
        Trace ">>> [Recovery] Max retry attempts reached.", True
        GoTo CleanExit
    End If

    m_RetryCount = m_RetryCount + 1
    
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    Trace ">>> [Recovery] Attempt #" & m_RetryCount & " (Waiting " & waitTime & "ms)", True
    
    YieldExecution waitTime, True
    
    If Connect(m_LastWsUrl, m_LastPort) Then
        Trace ">>> [Recovery] Success.", True
        m_RetryCount = 0
        AttemptRecovery = True
    Else
        Trace ">>> [Recovery] Failed to reconnect.", True
        AttemptRecovery = False
    End If
    
CleanExit:
    m_isRecovering = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    ' Never call Trace from this error handler: secondary file-I/O failure could bypass
    ' CleanExit and leave m_isRecovering permanently set. Debug.Print is non-raising.
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [Recovery] Critical Error during attempt: " & Err.description
    
    AttemptRecovery = False
    Resume CleanExit
End Function

' ObserveWinsockMessage applies reserved-reply inspection consistently to both
' receive paths.
Private Sub ObserveWinsockMessage(ByVal rawMsg As String)
    Dim isSuspicious As Boolean
    isSuspicious = (InStr(1, rawMsg, """error"":", vbTextCompare) > 0) Or _
                   (InStr(1, rawMsg, "chrome-error", vbTextCompare) > 0)

    If isSuspicious Then
        m_consecutiveInternalErrors = m_consecutiveInternalErrors + 1
    End If

    Trace ">>> [Winsock] Message assembled. Length: " & Len(rawMsg), isSuspicious
End Sub

Private Sub ProcessReservedSystemResponse(ByVal inId As String, _
                                          ByVal rawMsg As String, _
                                          ByVal isPostHeartbeat As Boolean)
    If inId <> CStr(POST_SCRIPT_REPLY_ID) Then
        If m_RetryCount > 0 Then m_RetryCount = 0
    End If
End Sub

' Drain completed messages before an optional bounded Pump; a real message suppresses
' heartbeat generation for that GetMessage call.
Private Function TryReadWinsockBeforeHeartbeat(ByRef realMsg As String, _
                                               ByVal allowPump As Boolean) As Boolean
    If Not m_wsTransportInitialized Then Exit Function
    If Not WsIsOpen Then Exit Function

    ' m_isReading is a same-thread guard around non-yielding Pump/dequeue work and is
    ' released only by structured cleanup or emergency teardown.
    If m_isReading Then Exit Function

    On Error GoTo ErrorHandler

    m_isReading = True

    Dim didPump As Boolean
    Dim candidate As String
    Dim inId As String

DrainCompleted:
    Do While WsTryDequeueText(candidate)
        inId = ExtractJsonId(candidate)

        If IsReservedSystemId(inId) Then
            ProcessReservedSystemResponse inId, candidate, False
        Else
            ObserveWinsockMessage candidate
            realMsg = candidate
            TryReadWinsockBeforeHeartbeat = True
            GoTo FinalCleanup
        End If
    Loop

    If Not allowPump Or didPump Then GoTo FinalCleanup

    didPump = True
    WsPump WINSOCK_PUMP_MAX_MS, _
           WINSOCK_PUMP_MAX_BYTES, _
           WINSOCK_PUMP_MAX_CALLS

    If Wsstate = WsTransportFailed Then
        SetState StateError
        GoTo FinalCleanup
    End If

    If Not WsIsOpen Then
        SetState StateError
        GoTo FinalCleanup
    End If

    GoTo DrainCompleted

FinalCleanup:
    m_isReading = False
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & _
                "] [Winsock] Pre-heartbeat read error: " & Err.description
    SetState StateError
    Resume FinalCleanup
End Function

' GetMessage serves the event FIFO, performs bounded receive work, then manages the
' reserved-band session.status heartbeat.
Public Function GetMessage() As String
    If m_isSendReceiving Then Exit Function

    If m_EventBuffer.count > 0 Then
        GetMessage = m_EventBuffer(1)
        m_EventBuffer.Remove 1
        Exit Function
    End If

    If m_state = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If
    If m_state <> StateConnected Then Exit Function

    Dim preHeartbeatMsg As String
    If TryReadWinsockBeforeHeartbeat(preHeartbeatMsg, False) Then
        GetMessage = preHeartbeatMsg
        Exit Function
    End If

    If m_state <> StateConnected Then Exit Function

    Dim nowT As Double
    nowT = nowMs()
    If (nowT - m_lastPollTime) < m_PollingInterval Then Exit Function
    m_lastPollTime = nowT

    ' Perform one bounded receive Pump before heartbeat; send heartbeat only when no real
    ' message is ready.
    preHeartbeatMsg = vbNullString
    If TryReadWinsockBeforeHeartbeat(preHeartbeatMsg, True) Then
        GetMessage = preHeartbeatMsg
        Exit Function
    End If

    If m_state <> StateConnected Then Exit Function

    Dim pingId As Long
    pingId = GetNextSystemPingId()

    Dim pingCmd As String
    pingCmd = "{""id"":" & pingId & ", ""method"":""session.status"", ""params"":{}}"
    If Not SendMessage(pingCmd) Then Exit Function

    Dim rawMsg As String
    rawMsg = ReadFromWire(nowMs() + 10)

    If LenB(rawMsg) = 0 Then
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If

    Dim inId As String
    inId = ExtractJsonId(rawMsg)

    If IsReservedSystemId(inId) Then
        ProcessReservedSystemResponse inId, rawMsg, True
        GetMessage = vbNullString
    Else
        GetMessage = rawMsg
    End If
End Function

Public Function GetParsedEvent(Optional ByVal minMsgSize As Long = 10, _
                               Optional ByVal maxMsgSize As Long = 300000) As Object
    Dim rawMsg As String: rawMsg = GetMessage()
    If Len(rawMsg) = 0 Then Exit Function
    
    If Len(rawMsg) < minMsgSize Or Len(rawMsg) > maxMsgSize Then Exit Function
    
    If InStr(rawMsg, """method"":") = 0 And InStr(rawMsg, """id"":") = 0 Then Exit Function
    
    Dim parsed As Object
    On Error Resume Next
    Set parsed = WebJsonConverter.ParseJson(rawMsg)
    If Err.Number <> 0 Or parsed Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    Set GetParsedEvent = parsed
End Function

Private Function ReadFromWinsock(Optional ByVal deadline As Double = 0) As String
    If Not m_wsTransportInitialized Then Exit Function

    ' ReadFromWinsock holds the same non-yielding read guard and releases it through a
    ' single cleanup path.
    If m_isReading Then Exit Function

    On Error GoTo ErrorHandler

    m_isReading = True

    Dim rawMsg As String
    If deadline <= 0 Then deadline = nowMs() + m_DefaultTimeout

    Do
        If WsTryDequeueText(rawMsg) Then Exit Do

        WsPump WINSOCK_PUMP_MAX_MS, _
               WINSOCK_PUMP_MAX_BYTES, _
               WINSOCK_PUMP_MAX_CALLS

        If WsTryDequeueText(rawMsg) Then Exit Do

        If Wsstate = WsTransportFailed Then
            SetState StateError
            Exit Do
        End If

        If Not WsIsOpen Then
            SetState StateError
            Exit Do
        End If

        Dim remainingMs As Double
        remainingMs = deadline - nowMs()
        If remainingMs <= 0 Then Exit Do

        If remainingMs > 20 Then
            SleepWinAPI 1
        Else
            SleepWinAPI 0
        End If
    Loop

    If Len(rawMsg) > 0 Then
        ObserveWinsockMessage rawMsg
        ReadFromWinsock = rawMsg
    End If

FinalCleanup:
    m_isReading = False
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & _
                "] [Winsock] Critical Read Error: " & Err.description
    SetState StateError
    Resume FinalCleanup
End Function

Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If m_state <> StateConnected And m_state <> StateDisconnecting Then Exit Function
    ReadFromWire = ReadFromWinsock(deadline)
End Function

' NowTicks/ElapsedMs/NowMs use cached QPC frequency and are monotonic; QPC absence
' fails Class_Initialize rather than silently changing timer semantics.
Public Function NowTicks() As Currency
    QueryPerformanceCounter NowTicks
End Function

Public Function elapsedMs(ByVal startTicks As Currency) As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    elapsedMs = (CDbl(cur - startTicks) / CDbl(m_qpcFreq)) * 1000#
End Function

Public Function nowMs() As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    nowMs = (CDbl(cur) / CDbl(m_qpcFreq)) * 1000#
End Function

Public Function AttemptAutoConnect(ByVal wsUrl As String) As Boolean
    Dim port As Long
    port = GetPortFromUrl(wsUrl)
    If wsUrl <> "" And port <> 0 Then
        AttemptAutoConnect = Connect(wsUrl, port)
    End If
End Function

Public Function GetPortFromUrl(ByVal wsUrl As String) As Long
    Dim posSession As Long, posColon As Long
    Dim portStr As String
    
    posSession = InStr(wsUrl, "/session")
    If posSession = 0 Then Exit Function
    
    posColon = InStrRev(wsUrl, ":", posSession)
    If posColon = 0 Then Exit Function
    
    portStr = Mid(wsUrl, posColon + 1, posSession - posColon - 1)
    If IsNumeric(portStr) Then
        GetPortFromUrl = CLng(portStr)
    End If
End Function

' ExtractJsonId scans only the root object so nested event payload IDs cannot be
' mistaken for the BiDi request/response ID.
Private Function ExtractJsonId(ByVal json As String) As String
    Dim s As String: s = LTrim$(json)
    If Len(s) < 10 Or Left$(s, 1) <> "{" Then Exit Function

    If InStr(1, s, """id""", vbTextCompare) = 0 Then Exit Function

    Dim depth As Long: depth = 0
    Dim inString As Boolean: inString = False
    Dim isEscaped As Boolean: isEscaped = False
    Dim i As Long
    
    Dim b() As Byte: b = s
    Dim bLen As Long: bLen = UBound(b)
    
    Const CHAR_QUOTE As Byte = 34  ' "
    Const CHAR_LBRACE As Byte = 123 ' {
    Const CHAR_RBRACE As Byte = 125 ' }
    Const CHAR_SLASH As Byte = 92  ' \
    
    For i = 0 To bLen Step 2
        ' The byte scanner reads UTF-16 LE. Test each high byte before interpreting ASCII;
        ' otherwise non-ASCII low-byte collisions can corrupt string/depth tracking. Keep the
        ' paired-byte bounds guard.
        If i + 1 > bLen Then Exit For
        If b(i + 1) <> 0 Then GoTo ContinueLoop

        If isEscaped Then
            isEscaped = False
            GoTo ContinueLoop
        End If

        Select Case b(i)
            Case CHAR_SLASH
                If inString Then isEscaped = True
                
            Case CHAR_QUOTE
                inString = Not inString
                
                If inString And depth = 1 Then
                    If i + 8 <= bLen Then
                        If b(i + 2) = 105 And b(i + 3) = 0 _
                           And b(i + 4) = 100 And b(i + 5) = 0 _
                           And b(i + 6) = 34 And b(i + 7) = 0 Then
                            
                            Dim j As Long: j = i + 8 ' Start right after "id"
                            Dim idVal As String
                            
                            Do While j + 1 <= bLen
                                If b(j + 1) = 0 Then
                                    Select Case b(j)
                                        Case 48 To 57, 45 ' '0'-'9', '-'
                                            idVal = idVal & Chr$(b(j))
                                        Case 32, 58, 9, 10, 13 ' spaces, colon, tabs, CRLF
                                            If Len(idVal) > 0 Then Exit Do
                                        Case Else
                                            Exit Do
                                    End Select
                                Else
                                    Exit Do
                                End If
                                j = j + 2
                            Loop
                            
                            If Len(idVal) > 0 Then
                                ExtractJsonId = idVal
                                Exit Function
                            End If
                            
                        End If
                    End If
                End If
                
            Case CHAR_LBRACE
                If Not inString Then depth = depth + 1
                
            Case CHAR_RBRACE
                If Not inString Then
                    depth = depth - 1
                    If depth = 0 Then Exit Function
                End If
        End Select
        
ContinueLoop:
    Next i
End Function

' PumpMessages may run DoEvents only outside SendAndReceive; throttleMs limits how
' often the Office queue is pumped.
Public Sub PumpMessages(Optional ByVal throttleMs As Long = 0)
    If m_isSendReceiving Then Exit Sub
    If throttleMs > 0 And m_lastPumpTicks <> 0 Then
        If elapsedMs(m_lastPumpTicks) < throttleMs Then Exit Sub
    End If
    DoEvents
    m_lastPumpTicks = NowTicks()
End Sub

' YieldExecution may pump Office messages only when explicitly requested and never
' while a command owns response correlation.
Public Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    Dim vMs As Long: vMs = ms
    If vMs < 0 Then vMs = 1

    If vMs > 50 Then
        Call InterruptibleSleep(vMs, allowDoEvents)
    Else
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If

        If vMs > 0 Then SleepWinAPI vMs
    End If
End Sub

' DEBUG_MODE controls Immediate output and keyword-based file detection. In production,
' file logging is driven only by the caller's explicit isCritical classification.
Public Sub Trace(ByVal msg As String, Optional ByVal isCritical As Boolean = False)
    If Not DEBUG_MODE And Not isCritical Then Exit Sub
    
    Dim logMsg As String: logMsg = "[" & Format(Now, "hh:mm:ss") & "] " & msg
    
    If DEBUG_MODE Then Debug.Print logMsg
    
    Dim shouldLogToFile As Boolean
    shouldLogToFile = isCritical
    If DEBUG_MODE Then
        shouldLogToFile = shouldLogToFile Or _
                          (InStr(1, msg, "Return Code", vbTextCompare) > 0) Or _
                          (InStr(1, msg, "failed", vbTextCompare) > 0) Or _
                          (InStr(1, msg, "error", vbTextCompare) > 0)
    End If

    If shouldLogToFile Then
        Dim localDir As String: localDir = WebShared.getFullLocalPath(ThisWorkbook.path)
        If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
        If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
        
        Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
        Dim fNum As Integer: fNum = FreeFile
        
        On Error Resume Next
        Open logPath For Append As #fNum
        Print #fNum, logMsg & " (LOG_EVENT)"
        Close #fNum
        On Error GoTo 0
    End If
End Sub

Public Sub LogEvent(ByVal Header As String, ByVal Body As String, _
                   Optional ByVal tag As String = "[INFO]", _
                   Optional ByVal LogToFile As Boolean = False)
    Const CHUNK_SIZE As Long = 250
    Dim i As Long
    Dim cleanText As String
    Dim timePrefix As String
    Dim currentTick As Currency
    
    currentTick = Me.NowTicks()
    timePrefix = "[" & Format(Now, "hh:mm:ss") & " | " & Format(currentTick, "0") & "] "

    cleanText = Replace(Body, vbNullChar, "")
    Do While Len(cleanText) > 0
        Dim lastChar As String
        lastChar = Right$(cleanText, 1)
        If lastChar = vbCr Or lastChar = vbLf Then
            cleanText = Left$(cleanText, Len(cleanText) - 1)
        Else
            Exit Do
        End If
    Loop

    Dim sfx As String: sfx = "": If Len(cleanText) > 0 Then sfx = " (Length: " & Len(cleanText) & ")"
    Dim headerLine As String: headerLine = timePrefix & tag & " " & Header & sfx
    Debug.Print headerLine

    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
        Debug.Print ""
    End If
    
    If LogToFile Then
        WriteLogToFile headerLine & vbCrLf & cleanText
    End If
End Sub

Private Sub WriteLogToFile(ByVal msg As String)
    Dim localDir As String
    On Error Resume Next
    localDir = WebShared.getFullLocalPath(ThisWorkbook.path)
    If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
    If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
    
    Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
    Dim fNum As Integer: fNum = FreeFile
    
    Open logPath For Append As #fNum
    Print #fNum, msg
    Close #fNum
    On Error GoTo 0
End Sub

' Only a confirmed healthy response should reset the consecutive internal-error
' counter.
Public Sub ClearInternalErrorCounter()
    m_consecutiveInternalErrors = 0
End Sub

Private Function ExtractHost(ByVal url As String) As String
    Dim startPos As Long
    Dim relPos As Long
    Dim authEnd As Long
    Dim qPos As Long
    Dim fPos As Long
    Dim atPos As Long
    Dim portPos As Long
    Dim closeBracket As Long
    Dim authority As String
    Dim hostPart As String

    url = Trim$(url)
    If url = "" Then Exit Function

    startPos = InStr(1, url, "://", vbTextCompare)
    relPos = InStr(1, url, "//")

    If startPos > 0 Then
        startPos = startPos + 3
    ElseIf relPos = 1 Then
        startPos = 3
    Else
        startPos = 1
    End If

    authEnd = InStr(startPos, url, "/")
    
    qPos = InStr(startPos, url, "?")
    If qPos > 0 And (authEnd = 0 Or qPos < authEnd) Then authEnd = qPos
    
    fPos = InStr(startPos, url, "#")
    If fPos > 0 And (authEnd = 0 Or fPos < authEnd) Then authEnd = fPos
    
    If authEnd = 0 Then authEnd = Len(url) + 1
    
    authority = Mid$(url, startPos, authEnd - startPos)
    If authority = "" Then Exit Function
    
    atPos = InStrRev(authority, "@")
    If atPos > 0 Then authority = Mid$(authority, atPos + 1)
    
    If Left$(authority, 1) = "[" Then
        closeBracket = InStr(1, authority, "]")
        If closeBracket > 0 Then
            hostPart = Mid$(authority, 2, closeBracket - 2)
        Else
            Exit Function
        End If
    Else
        portPos = InStrRev(authority, ":")
        If portPos > 0 Then
            hostPart = Left$(authority, portPos - 1)
        Else
            hostPart = authority
        End If
    End If
    
    If hostPart <> "" Then ExtractHost = LCase$(hostPart)
End Function

Public Function ExtractOrigin(ByVal url As String) As String
    If Len(url) = 0 Then Exit Function
    Dim schemeEnd As Long: schemeEnd = InStr(1, url, "://", vbTextCompare)
    If schemeEnd = 0 Then Exit Function
    Dim authStart As Long: authStart = schemeEnd + 3
    Dim pathStart As Long: pathStart = InStr(authStart, url, "/")
    If pathStart = 0 Then
        ExtractOrigin = LCase$(url)
    Else
        ExtractOrigin = LCase$(Left$(url, pathStart - 1))
    End If
End Function

' InterruptibleSleep pumps Office messages only when allowed and outside an active
' SendAndReceive cycle; otherwise it is a physical chunked wait.
Private Sub InterruptibleSleep(ByVal totalMs As Long, Optional ByVal allowDoEvents As Boolean = True)
    Dim startTs As Double: startTs = nowMs()
    Const CHUNK_MS As Long = 50 ' Small granularity for better responsiveness
    
    Do While (nowMs() - startTs) < totalMs
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If
        
        SleepWinAPI CHUNK_MS
    Loop
End Sub

' ForceClose skips the close handshake and performs non-raising emergency teardown.
Public Sub ForceClose()
    ' Clear I/O guards before closing native transport so interrupted operations do not
    ' poison object reuse.
    On Error Resume Next
    m_isSendReceiving = False
    m_isReading = False

    If m_wsTransportInitialized Then
        WsForceClose
    End If

    Set m_EventBuffer = New Collection
    SetState StateDisconnected
    On Error GoTo 0
End Sub

Private Sub WsTransportInitialize()
    Dim frequency As Currency
    If WsApi_QueryPerformanceFrequency(frequency) = 0 Or frequency = 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", "QPC is unavailable."
    End If
    m_ws_qpcFrequency = frequency

    Dim data As WsWSADATA
    Dim result As Long
    result = WsApi_WSAStartup(&H202, data)
    If result <> 0 Then
        Err.Raise vbObjectError + 9101, "BiDiSocketCommunicator", "WSAStartup failed: " & result
    End If
    m_ws_wsaStarted = True

    Set m_ws_sendQueue = New Collection
    Set m_ws_rawTextQueue = New Collection
    Set m_ws_rawTextLengthQueue = New Collection
    m_ws_socket = -1
    m_ws_state = WsTransportDisconnected
    m_ws_maxIncomingSize = WS_DEFAULT_MAX_INCOMING_SIZE
    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_clientPingEnabled = True

End Sub

Private Sub WsTransportTerminate()
    On Error Resume Next
    WsForceClose
    If m_ws_wsaStarted Then
        WsApi_WSACleanup
        m_ws_wsaStarted = False
    End If
    On Error GoTo 0
End Sub

Private Property Get Wsstate() As WsWinsockTransportState
    Wsstate = m_ws_state
End Property

Private Property Get WsStateName() As String
    Select Case m_ws_state
        Case WsTransportDisconnected: WsStateName = "DISCONNECTED"
        Case WsTransportConnecting: WsStateName = "CONNECTING"
        Case WsTransportHandshakeSending: WsStateName = "HANDSHAKE_SEND"
        Case WsTransportHandshakeReceiving: WsStateName = "HANDSHAKE_RECEIVE"
        Case WsTransportOpen: WsStateName = "OPEN"
        Case WsTransportClosing: WsStateName = "CLOSING"
        Case WsTransportFailed: WsStateName = "FAILED"
        Case Else: WsStateName = "UNKNOWN"
    End Select
End Property

Private Property Get WsIsOpen() As Boolean
    WsIsOpen = (m_ws_state = WsTransportOpen)
End Property

Private Property Get WsLastErrorCode() As Long
    WsLastErrorCode = m_ws_LastErrorCode
End Property

Private Property Get WsLastErrorText() As String
    WsLastErrorText = m_ws_lastErrorText
End Property

Private Property Get WsMaxIncomingSize() As Long
    WsMaxIncomingSize = m_ws_maxIncomingSize
End Property

Private Property Let WsMaxIncomingSize(ByVal value As Long)
    If value < 300000 Then value = 300000
    m_ws_maxIncomingSize = value
End Property

Private Property Get WsHasPendingOutput() As Boolean
    WsHasPendingOutput = WsHasPendingSend
End Property

Private Function WsMonotonicMs() As Double
    Dim currentTick As Currency
    WsApi_QueryPerformanceCounter currentTick
    WsMonotonicMs = (CDbl(currentTick) / CDbl(m_ws_qpcFrequency)) * 1000#
End Function

Private Function WsApplyTcpNoDelay() As Boolean
    Dim optionValue As Long
    Dim actualValue As Long
    Dim optionLength As Long
    Dim socketError As Long

    optionValue = 1

    If WsApi_setsockopt(m_ws_socket, WS_IPPROTO_TCP, WS_TCP_NODELAY, _
                  optionValue, LenB(optionValue)) = WS_SOCKET_ERROR Then
        socketError = WsApi_WSAGetLastError()
        WsSetFailure socketError, "setsockopt TCP_NODELAY failed."
        Exit Function
    End If

    optionLength = LenB(actualValue)

    If WsApi_getsockopt(m_ws_socket, WS_IPPROTO_TCP, WS_TCP_NODELAY, _
                  actualValue, optionLength) = WS_SOCKET_ERROR Then
        socketError = WsApi_WSAGetLastError()
        WsSetFailure socketError, "getsockopt TCP_NODELAY failed."
        Exit Function
    End If

    If actualValue = 0 Then
        WsSetFailure 9240, "TCP_NODELAY remained disabled."
        Exit Function
    End If

    WsApplyTcpNoDelay = True
End Function

Private Function WsBeginConnect(ByVal wsUrl As String, _
                             Optional ByVal connectTimeoutMs As Long = WS_DEFAULT_CONNECT_TIMEOUT_MS, _
                             Optional ByVal handshakeTimeoutMs As Long = WS_DEFAULT_HANDSHAKE_TIMEOUT_MS) As Boolean
    On Error GoTo ErrorHandler

    WsForceClose
    WsResetRuntimeState

    If connectTimeoutMs < 1 Then connectTimeoutMs = WS_DEFAULT_CONNECT_TIMEOUT_MS
    If handshakeTimeoutMs < 1 Then handshakeTimeoutMs = WS_DEFAULT_HANDSHAKE_TIMEOUT_MS

    If Not WsParseLoopbackUrl(wsUrl) Then
        WsSetFailure 9201, "Only ws:// loopback URLs are supported."
        Exit Function
    End If

    m_ws_url = wsUrl
    m_ws_handshakeTimeoutMs = handshakeTimeoutMs

    Dim addressFamily As Long
    If m_ws_host = "[::1]" Then
        addressFamily = WS_AF_INET6
    Else
        addressFamily = WS_AF_INET
    End If

    m_ws_socket = WsApi_socket(addressFamily, WS_SOCK_STREAM, WS_IPPROTO_TCP)
    If m_ws_socket = -1 Then
        WsSetFailure WsApi_WSAGetLastError(), "socket failed."
        Exit Function
    End If

    If Not WsApplyTcpNoDelay Then Exit Function

    Dim nonBlocking As Long
    nonBlocking = 1
    If WsApi_ioctlsocket(m_ws_socket, WS_FIONBIO, nonBlocking) = WS_SOCKET_ERROR Then
        WsSetFailure WsApi_WSAGetLastError(), "FIONBIO failed."
        Exit Function
    End If

    m_ws_state = WsTransportConnecting
    m_ws_phaseDeadlineMs = WsMonotonicMs() + connectTimeoutMs

    Dim connectResult As Long
    If addressFamily = WS_AF_INET6 Then
        Dim address6 As WsSOCKADDR_IN6
        address6.family = WS_AF_INET6
        address6.port = WsHostToNetworkShort(m_ws_port)

        If WsApi_InetPtonW(WS_AF_INET6, StrPtr("::1"), address6.address(0)) <> 1 Then
            WsSetFailure WsApi_WSAGetLastError(), "IPv6 loopback conversion failed."
            Exit Function
        End If

        connectResult = WsApi_connectSocket(m_ws_socket, address6, LenB(address6))
    Else
        Dim address4 As WsSOCKADDR_IN
        address4.family = WS_AF_INET
        address4.port = WsHostToNetworkShort(m_ws_port)

        If WsApi_InetPtonW(WS_AF_INET, StrPtr("127.0.0.1"), address4.address) <> 1 Then
            WsSetFailure WsApi_WSAGetLastError(), "IPv4 loopback conversion failed."
            Exit Function
        End If

        connectResult = WsApi_connectSocket(m_ws_socket, address4, LenB(address4))
    End If

    If connectResult = 0 Then
        WsBeginHandshake
    Else
        Dim socketError As Long
        socketError = WsApi_WSAGetLastError()

        If socketError <> WS_WSAEWOULDBLOCK And _
           socketError <> WS_WSAEINPROGRESS And _
           socketError <> WS_WSAEALREADY And _
           socketError <> WS_WSAEISCONN Then
            WsSetFailure socketError, "connect failed."
            Exit Function
        End If

        If socketError = WS_WSAEISCONN Then WsBeginHandshake
    End If

    WsBeginConnect = True
    Exit Function

ErrorHandler:
    WsSetFailure Err.Number, Err.description
End Function

Private Function WsPump(Optional ByVal maxPumpMs As Long = WS_DEFAULT_MAX_PUMP_MS, _
                     Optional ByVal maxBytes As Long = WS_DEFAULT_MAX_PUMP_BYTES, _
                     Optional ByVal maxCalls As Long = WS_DEFAULT_MAX_PUMP_CALLS) As Boolean
    m_ws_currentPumpCompletedMessages = 0

    If maxPumpMs < 1 Then maxPumpMs = 1
    If maxBytes < 1 Then maxBytes = 1
    If maxCalls < 1 Then maxCalls = 1

    Dim pumpStartedAt As Double
    Dim pumpDeadline As Double
    Dim copyGuardDeadline As Double

    pumpStartedAt = WsMonotonicMs()
    pumpDeadline = pumpStartedAt + maxPumpMs
    copyGuardDeadline = pumpDeadline

    ' Reserve the final parser millisecond only for the normal product Pump budget.
    If maxPumpMs >= WS_DEFAULT_MAX_PUMP_MS Then
        copyGuardDeadline = pumpDeadline - WS_PARSER_COPY_GUARD_MS
    End If

    Select Case m_ws_state
        Case WsTransportConnecting
            WsPumpConnect

        Case WsTransportHandshakeSending, WsTransportHandshakeReceiving
            WsPumpHandshake pumpDeadline, copyGuardDeadline, maxBytes, maxCalls

        Case WsTransportOpen
            WsMaybeQueueClientPing

            If m_ws_state = WsTransportOpen Then
                WsPumpSendInternal pumpDeadline, maxBytes, maxCalls
            End If

            If m_ws_state = WsTransportOpen And WsMonotonicMs() < pumpDeadline Then
                WsPumpReceiveInternal pumpDeadline, copyGuardDeadline, _
                                    maxBytes, maxCalls, True
            End If

        Case WsTransportClosing
            WsPumpSendInternal pumpDeadline, maxBytes, maxCalls

            If Not WsHasPendingSend Then WsForceClose
    End Select

    WsPump = (m_ws_state <> WsTransportFailed)
End Function

Private Function WsQueueText(ByVal message As String) As Boolean
    If m_ws_state <> WsTransportOpen Then Exit Function

    Dim payload() As Byte
    payload = WsEncodeUtf8(message)

    Dim frame() As Byte
    frame = WsBuildClientFrame(WS_OPCODE_TEXT, payload, True)

    If m_ws_state = WsTransportFailed Then Exit Function
    If WsByteArrayLength(frame) = 0 Then Exit Function

    WsEnqueueSendBytes frame
    WsQueueText = True
End Function

Private Function WsTryDequeueText(ByRef message As String) As Boolean
    On Error GoTo ErrorHandler

    If m_ws_rawTextLengthQueue.count <> m_ws_rawTextQueue.count Then
        WsSetFailure 9218, "Deferred text queue is inconsistent."
        Exit Function
    End If

    If m_ws_rawTextQueue.count = 0 Then Exit Function

    Dim chunks As Collection
    Dim totalBytes As Long

    Set chunks = m_ws_rawTextQueue(1)
    totalBytes = CLng(m_ws_rawTextLengthQueue(1))

    m_ws_rawTextQueue.Remove 1
    m_ws_rawTextLengthQueue.Remove 1

    message = WsDecodeUtf8Chunks(chunks, totalBytes)
    WsTryDequeueText = (m_ws_state <> WsTransportFailed)
    Exit Function

ErrorHandler:
    WsSetFailure Err.Number, "Deferred UTF-8 decode failed: " & Err.description
End Function

Private Function WsBeginClose(Optional ByVal statusCode As Long = 1000) As Boolean
    If m_ws_state <> WsTransportOpen Then Exit Function
    If m_ws_closeSent Then
        WsBeginClose = True
        Exit Function
    End If

    Dim payload(0 To 1) As Byte
    payload(0) = CByte((statusCode \ &H100&) And &HFF&)
    payload(1) = CByte(statusCode And &HFF&)

    Dim frame() As Byte
    frame = WsBuildClientFrame(WS_OPCODE_CLOSE, payload, True)

    If m_ws_state = WsTransportFailed Then Exit Function
    If WsByteArrayLength(frame) = 0 Then Exit Function

    WsEnqueueSendBytes frame

    m_ws_closeSent = True
    m_ws_state = WsTransportClosing
    WsBeginClose = True
End Function

Private Sub WsForceClose()
    On Error Resume Next

    If m_ws_socket <> -1 Then
        WsApi_shutdownSocket m_ws_socket, WS_SD_BOTH
        WsApi_closesocket m_ws_socket
        m_ws_socket = -1
    End If

    Set m_ws_sendQueue = New Collection
    Set m_ws_rawTextQueue = New Collection
    Set m_ws_rawTextLengthQueue = New Collection
    Set m_ws_messageChunks = Nothing

    Erase m_ws_sendCurrent
    Erase m_ws_rawBuffer
    Erase m_ws_controlPayload

    m_ws_sendCurrentCount = 0
    m_ws_sendOffset = 0
    m_ws_rawCount = 0
    m_ws_rawReadPos = 0
    m_ws_messageCount = 0
    m_ws_controlCount = 0
    m_ws_state = WsTransportDisconnected

    On Error GoTo 0
End Sub

Private Sub WsPumpConnect()
    If WsMonotonicMs() >= m_ws_phaseDeadlineMs Then
        WsSetFailure 9202, "Connect timeout."
        Exit Sub
    End If

    Dim readSet As WsFD_SET
    Dim writeSet As WsFD_SET
    Dim exceptSet As WsFD_SET
    Dim TIMEOUT As WsTIMEVAL

    writeSet.count = 1
    writeSet.sockets(0) = m_ws_socket
    exceptSet.count = 1
    exceptSet.sockets(0) = m_ws_socket

    Dim ready As Long
    ready = WsApi_selectSocket(0, readSet, writeSet, exceptSet, TIMEOUT)
    If ready = WS_SOCKET_ERROR Then
        WsSetFailure WsApi_WSAGetLastError(), "select failed during connect."
        Exit Sub
    End If
    If ready = 0 Then Exit Sub

    Dim socketError As Long
    Dim optionLength As Long
    optionLength = 4

    If WsApi_getsockopt(m_ws_socket, WS_SOL_SOCKET, WS_SO_ERROR, socketError, optionLength) = WS_SOCKET_ERROR Then
        WsSetFailure WsApi_WSAGetLastError(), "getsockopt(SO_ERROR) failed."
        Exit Sub
    End If

    If socketError <> 0 Then
        WsSetFailure socketError, "Non-blocking connect failed."
        Exit Sub
    End If

    WsBeginHandshake
End Sub

Private Sub WsBeginHandshake()
    m_ws_handshakeValidationError = vbNullString

    If Not WsPrepareHandshakeValues Then Exit Sub

    WsPrepareHandshakeHeaders

    Dim requestText As String
    requestText = _
        "GET " & m_ws_path & " HTTP/1.1" & vbCrLf & _
        "Host: " & m_ws_hostHeader & vbCrLf & _
        "Upgrade: websocket" & vbCrLf & _
        "Connection: Upgrade" & vbCrLf & _
        "Sec-WebSocket-Key: " & m_ws_handshakeClientKey & vbCrLf & _
        "Sec-WebSocket-Version: 13" & vbCrLf & _
        "Origin: " & m_ws_handshakeOrigin & vbCrLf & _
        "User-Agent: " & m_ws_handshakeUserAgent & vbCrLf & vbCrLf

    Dim requestBytes() As Byte
    requestBytes = WsEncodeAscii(requestText)
    WsEnqueueSendBytes requestBytes

    m_ws_phaseDeadlineMs = WsMonotonicMs() + m_ws_handshakeTimeoutMs
    m_ws_state = WsTransportHandshakeSending
End Sub

Private Sub WsPumpHandshake(ByVal pumpDeadline As Double, _
                          ByVal copyGuardDeadline As Double, _
                          ByVal maxBytes As Long, _
                          ByVal maxCalls As Long)
    If WsMonotonicMs() >= m_ws_phaseDeadlineMs Then
        WsSetFailure 9203, "WebSocket handshake timeout."
        Exit Sub
    End If

    If m_ws_state = WsTransportHandshakeSending Then
        WsPumpSendInternal pumpDeadline, maxBytes, maxCalls
        If m_ws_state = WsTransportFailed Then Exit Sub

        If Not WsHasPendingSend Then
            m_ws_state = WsTransportHandshakeReceiving
        End If
    End If

    If m_ws_state = WsTransportHandshakeReceiving And WsMonotonicMs() < pumpDeadline Then
        WsPumpReceiveInternal pumpDeadline, copyGuardDeadline, _
                            maxBytes, maxCalls, False
        If m_ws_state = WsTransportFailed Then Exit Sub

        WsTryCompleteHandshake pumpDeadline, copyGuardDeadline
    End If
End Sub

Private Sub WsTryCompleteHandshake(ByVal pumpDeadline As Double, _
                                 ByVal copyGuardDeadline As Double)
    Dim headerEnd As Long
    headerEnd = WsFindHttpHeaderEnd()
    If headerEnd = 0 Then
        If WsRawAvailable > WS_MAX_HTTP_HEADER_SIZE Then
            WsSetFailure 9204, "HTTP upgrade header is too large."
        End If
        Exit Sub
    End If

    Dim headerText As String
    headerText = WsReadAsciiFromRaw(headerEnd)

    If Not WsValidateHandshakeHeader(headerText) Then
        WsSetFailure 9205, _
                   "WebSocket upgrade response validation failed: " & _
                   m_ws_handshakeValidationError
        Exit Sub
    End If

    WsConsumeRaw headerEnd
    m_ws_state = WsTransportOpen
    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_nextClientPingAtMs = WsMonotonicMs() + WS_CLIENT_PING_INTERVAL_MS

    If WS_DEBUG_MODE Then Debug.Print "BiDiSocketCommunicator[IntegratedWinsock]: OPEN"

    WsParseAvailableFrames pumpDeadline, copyGuardDeadline
End Sub

Private Function WsValidateHandshakeHeader(ByVal headerText As String) As Boolean
    m_ws_handshakeValidationError = vbNullString

    Dim lines() As String
    lines = Split(headerText, vbCrLf)
    If UBound(lines) < 1 Then
        m_ws_handshakeValidationError = "response has no headers"
        Exit Function
    End If

    Dim firstLine As String
    firstLine = Trim$(lines(0))

    Dim statusValid As Boolean
    statusValid = _
        (InStr(1, firstLine, " 101 ", vbTextCompare) > 0) Or _
        (Right$(firstLine, 4) = " 101")

    Dim hasUpgrade As Boolean
    Dim hasConnectionUpgrade As Boolean
    Dim acceptPresent As Boolean
    Dim acceptValue As String
    Dim i As Long

    For i = 1 To UBound(lines)
        Dim colonPos As Long
        colonPos = InStr(1, lines(i), ":", vbBinaryCompare)

        If colonPos > 0 Then
            Dim headerName As String
            Dim headerValue As String

            headerName = _
                LCase$(Trim$(Left$(lines(i), colonPos - 1)))
            headerValue = Trim$(Mid$(lines(i), colonPos + 1))

            If headerName = "upgrade" Then
                hasUpgrade = _
                    (StrComp(headerValue, "websocket", _
                             vbTextCompare) = 0)
            ElseIf headerName = "connection" Then
                hasConnectionUpgrade = _
                    WsHeaderContainsToken(headerValue, "upgrade")
            ElseIf headerName = "sec-websocket-accept" Then
                acceptPresent = True
                acceptValue = headerValue
            End If
        End If
    Next

    Dim acceptMatched As Boolean
    If acceptPresent Then
        acceptMatched = _
            (StrComp(acceptValue, m_ws_expectedHandshakeAccept, _
                     vbBinaryCompare) = 0)
    End If

    If Not statusValid Then
        m_ws_handshakeValidationError = "HTTP status is not 101"
    ElseIf Not hasUpgrade Then
        m_ws_handshakeValidationError = _
            "Upgrade header is not websocket"
    ElseIf Not hasConnectionUpgrade Then
        m_ws_handshakeValidationError = _
            "Connection header has no Upgrade token"
    ElseIf Not acceptPresent Then
        m_ws_handshakeValidationError = _
            "Sec-WebSocket-Accept is missing"
    ElseIf Not acceptMatched Then
        m_ws_handshakeValidationError = _
            "Sec-WebSocket-Accept does not match"
    Else
        WsValidateHandshakeHeader = True
    End If
End Function

Private Sub WsPrepareHandshakeHeaders()
    If m_ws_port = 80 Then
        m_ws_handshakeOrigin = WS_HANDSHAKE_ORIGIN_SCHEME & m_ws_host
    Else
        m_ws_handshakeOrigin = WS_HANDSHAKE_ORIGIN_SCHEME & m_ws_hostHeader
    End If

    m_ws_handshakeUserAgent = WS_HANDSHAKE_USER_AGENT
End Sub

Private Function WsPrepareHandshakeValues() As Boolean
    Dim nonce(0 To WS_HANDSHAKE_NONCE_SIZE - 1) As Byte

    Dim status As Long
    status = WsApi_BCryptGenRandom(0, nonce(0), WS_HANDSHAKE_NONCE_SIZE, _
                             WS_BCRYPT_USE_SYSTEM_PREFERRED_RNG)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        WsSetFailure 9230, _
                   "Handshake BCryptGenRandom failed. NTSTATUS=0x" & _
                   Right$("00000000" & hex$(status), 8)
        Exit Function
    End If

    m_ws_handshakeClientKey = WsBase64EncodeBytes(nonce)

    If Not WsTryBuildExpectedAccept( _
        m_ws_handshakeClientKey, m_ws_expectedHandshakeAccept) Then
        Exit Function
    End If

    WsPrepareHandshakeValues = True
End Function

Private Function WsTryBuildExpectedAccept(ByVal clientKey As String, _
                                        ByRef acceptValue As String) As Boolean
    Dim inputBytes() As Byte
    inputBytes = WsEncodeAscii(clientKey & WS_WEBSOCKET_GUID)

    Dim digest() As Byte
    Dim failureText As String

    If Not WsTrySha1(inputBytes, digest, failureText) Then
        WsSetFailure 9231, _
                   "Handshake SHA-1 failed: " & failureText
        Exit Function
    End If

    acceptValue = WsBase64EncodeBytes(digest)
    WsTryBuildExpectedAccept = True
End Function

Private Function WsTrySha1(ByRef inputBytes() As Byte, _
                         ByRef digest() As Byte, _
                         ByRef failureText As String) As Boolean
    Dim algorithmHandle As LongPtr
    Dim hashHandle As LongPtr
    Dim status As Long
    Dim completed As Boolean

    Dim algorithmId As String
    algorithmId = WS_BCRYPT_SHA1_ALGORITHM

    status = WsApi_BCryptOpenAlgorithmProvider( _
        algorithmHandle, StrPtr(algorithmId), 0, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        failureText = _
            "BCryptOpenAlgorithmProvider NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    Dim objectProperty As String
    objectProperty = WS_BCRYPT_OBJECT_LENGTH

    Dim objectLength As Long
    Dim resultSize As Long

    status = WsApi_BCryptGetProperty( _
        algorithmHandle, StrPtr(objectProperty), _
        objectLength, 4, resultSize, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Or objectLength <= 0 Then
        failureText = _
            "ObjectLength query NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    Dim hashProperty As String
    hashProperty = WS_BCRYPT_HASH_LENGTH

    Dim hashLength As Long
    resultSize = 0

    status = WsApi_BCryptGetProperty( _
        algorithmHandle, StrPtr(hashProperty), _
        hashLength, 4, resultSize, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Or hashLength <= 0 Then
        failureText = _
            "HashDigestLength query NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    Dim hashObject() As Byte
    ReDim hashObject(0 To objectLength - 1)
    ReDim digest(0 To hashLength - 1)

    status = WsApi_BCryptCreateHash( _
        algorithmHandle, hashHandle, hashObject(0), _
        objectLength, 0, 0, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        failureText = _
            "BCryptCreateHash NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    Dim inputLength As Long
    inputLength = WsByteArrayLength(inputBytes)

    If inputLength <= 0 Then
        failureText = "SHA-1 input is empty"
        GoTo Cleanup
    End If

    status = WsApi_BCryptHashData( _
        hashHandle, inputBytes(0), inputLength, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        failureText = _
            "BCryptHashData NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    status = WsApi_BCryptFinishHash( _
        hashHandle, digest(0), hashLength, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        failureText = _
            "BCryptFinishHash NTSTATUS=0x" & _
            Right$("00000000" & hex$(status), 8)
        GoTo Cleanup
    End If

    completed = True

Cleanup:
    On Error Resume Next

    If hashHandle <> 0 Then
        WsApi_BCryptDestroyHash hashHandle
        hashHandle = 0
    End If

    If algorithmHandle <> 0 Then
        WsApi_BCryptCloseAlgorithmProvider algorithmHandle, 0
        algorithmHandle = 0
    End If

    On Error GoTo 0

    WsTrySha1 = completed
End Function

Private Function WsBase64EncodeBytes(ByRef Bytes() As Byte) As String
    Const alphabet As String = _
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    Dim count As Long
    count = WsByteArrayLength(Bytes)

    If count <= 0 Then
        WsBase64EncodeBytes = vbNullString
        Exit Function
    End If

    Dim result As String
    Dim index As Long

    Do While index < count
        Dim byte0 As Long
        Dim byte1 As Long
        Dim byte2 As Long
        Dim remaining As Long

        byte0 = Bytes(index)
        remaining = count - index

        If remaining > 1 Then
            byte1 = Bytes(index + 1)
        Else
            byte1 = 0
        End If

        If remaining > 2 Then
            byte2 = Bytes(index + 2)
        Else
            byte2 = 0
        End If

        result = result & _
            Mid$(alphabet, (byte0 \ 4) + 1, 1)

        result = result & _
            Mid$(alphabet, _
                 (((byte0 And 3) * 16) Or (byte1 \ 16)) + 1, _
                 1)

        If remaining > 1 Then
            result = result & _
                Mid$(alphabet, _
                     (((byte1 And 15) * 4) Or (byte2 \ 64)) + 1, _
                     1)
        Else
            result = result & "="
        End If

        If remaining > 2 Then
            result = result & _
                Mid$(alphabet, (byte2 And 63) + 1, 1)
        Else
            result = result & "="
        End If

        index = index + 3
    Loop

    WsBase64EncodeBytes = result
End Function

Private Function WsHeaderContainsToken(ByVal headerValue As String, ByVal token As String) As Boolean
    Dim parts() As String
    parts = Split(headerValue, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If LCase$(Trim$(parts(i))) = LCase$(token) Then
            WsHeaderContainsToken = True
            Exit Function
        End If
    Next
End Function

Private Sub WsPumpSendInternal(ByVal pumpDeadline As Double, _
                             ByVal maxBytes As Long, _
                             ByVal maxCalls As Long)
    Dim sentBytes As Long
    Dim callCount As Long

    Do While callCount < maxCalls And sentBytes < maxBytes
        If WsMonotonicMs() >= pumpDeadline Then Exit Do

        If m_ws_sendCurrentCount = 0 Then
            If Not WsLoadNextSendBuffer Then Exit Do
        End If

        Dim remaining As Long
        remaining = m_ws_sendCurrentCount - m_ws_sendOffset
        If remaining <= 0 Then
            WsClearCurrentSendBuffer
        Else
            Dim allowed As Long
            allowed = remaining
            If allowed > (maxBytes - sentBytes) Then allowed = maxBytes - sentBytes
            If allowed <= 0 Then Exit Do

            Dim result As Long
            result = WsApi_sendSocket(m_ws_socket, m_ws_sendCurrent(m_ws_sendOffset), allowed, 0)
            callCount = callCount + 1

            If result > 0 Then
                m_ws_sendOffset = m_ws_sendOffset + result
                sentBytes = sentBytes + result

                If m_ws_sendOffset >= m_ws_sendCurrentCount Then
                    WsClearCurrentSendBuffer
                End If
            ElseIf result = WS_SOCKET_ERROR Then
                Dim socketError As Long
                socketError = WsApi_WSAGetLastError()

                If socketError = WS_WSAEWOULDBLOCK Then Exit Do

                WsSetFailure socketError, "send failed."
                Exit Do
            Else
                WsSetFailure 9206, "send returned zero."
                Exit Do
            End If
        End If
    Loop
End Sub

Private Sub WsPumpReceiveInternal(ByVal pumpDeadline As Double, _
                                ByVal copyGuardDeadline As Double, _
                                ByVal maxBytes As Long, _
                                ByVal maxCalls As Long, _
                                ByVal parseFrames As Boolean)
    Dim receivedBytes As Long
    Dim callCount As Long

    ' Resume parsing buffered TCP bytes before requesting more network data.
    If parseFrames Then
        WsParseAvailableFrames pumpDeadline, copyGuardDeadline
        If m_ws_state = WsTransportFailed Then Exit Sub
        If m_ws_currentPumpCompletedMessages >= _
           WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Sub
        If WsMonotonicMs() >= pumpDeadline Then Exit Sub
    End If

    Do While callCount < maxCalls And receivedBytes < maxBytes
        If m_ws_currentPumpCompletedMessages >= _
           WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Do
        If WsMonotonicMs() >= pumpDeadline Then Exit Do

        WsCompactRawIfNeeded
        WsEnsureRawCapacity m_ws_rawCount + WS_RECEIVE_CHUNK_SIZE

        Dim allowed As Long
        allowed = WS_RECEIVE_CHUNK_SIZE
        If allowed > (maxBytes - receivedBytes) Then allowed = maxBytes - receivedBytes
        If allowed <= 0 Then Exit Do

        Dim result As Long
        result = WsApi_recvSocket(m_ws_socket, m_ws_rawBuffer(m_ws_rawCount), allowed, 0)
        callCount = callCount + 1

        If result > 0 Then
            m_ws_rawCount = m_ws_rawCount + result
            receivedBytes = receivedBytes + result

            If parseFrames Then
                WsParseAvailableFrames pumpDeadline, copyGuardDeadline
                If m_ws_state = WsTransportFailed Then Exit Do
                If m_ws_currentPumpCompletedMessages >= _
                   WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Do
                If WsMonotonicMs() >= pumpDeadline Then Exit Do
            End If
        ElseIf result = 0 Then
            If m_ws_state = WsTransportClosing Then
                WsForceClose
            Else
                WsSetFailure 9207, "Peer closed the socket."
            End If
            Exit Do
        Else
            Dim socketError As Long
            socketError = WsApi_WSAGetLastError()

            If socketError = WS_WSAEWOULDBLOCK Then Exit Do

            If socketError = WS_WSAECONNRESET Or _
               socketError = WS_WSAECONNABORTED Or _
               socketError = WS_WSAENOTCONN Then
                WsSetFailure socketError, "Connection lost during recv."
            Else
                WsSetFailure socketError, "recv failed."
            End If
            Exit Do
        End If
    Loop
End Sub

Private Sub WsParseAvailableFrames(ByVal pumpDeadline As Double, _
                                 ByVal copyGuardDeadline As Double)
    Do While m_ws_state = WsTransportOpen Or m_ws_state = WsTransportClosing
        If WsMonotonicMs() >= pumpDeadline Then Exit Do
        If m_ws_currentPumpCompletedMessages >= WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Do

        Select Case m_ws_frameState
            Case WsFrameReadBaseHeader
                If WsRawAvailable < 2 Then Exit Do

                Dim firstByte As Long
                Dim secondByte As Long
                firstByte = WsReadRawByte()
                secondByte = WsReadRawByte()

                If (firstByte And &H70&) <> 0 Then
                    WsSetFailure 9210, "RSV bits are not supported."
                    Exit Do
                End If

                m_ws_frameFin = ((firstByte And &H80&) <> 0)
                m_ws_frameOpcode = firstByte And &HF&
                m_ws_frameLengthCode = secondByte And &H7F&

                If (secondByte And &H80&) <> 0 Then
                    WsSetFailure 9211, "Masked server frame rejected."
                    Exit Do
                End If

                If Not WsValidateFrameOpcode(m_ws_frameOpcode) Then
                    WsSetFailure 9212, "Invalid WebSocket opcode."
                    Exit Do
                End If

                If m_ws_frameOpcode >= WS_OPCODE_CLOSE Then
                    If Not m_ws_frameFin Then
                        WsSetFailure 9213, "Fragmented control frame rejected."
                        Exit Do
                    End If
                End If

                If m_ws_frameLengthCode <= 125 Then
                    WsPrepareFramePayload m_ws_frameLengthCode
                Else
                    m_ws_frameState = WsFrameReadExtendedLength
                End If

            Case WsFrameReadExtendedLength
                If m_ws_frameLengthCode = 126 Then
                    If WsRawAvailable < 2 Then Exit Do

                    Dim length16 As Long
                    length16 = WsReadRawByte() * &H100& + WsReadRawByte()
                    WsPrepareFramePayload length16
                Else
                    If WsRawAvailable < 8 Then Exit Do

                    Dim lengthBytes(0 To 7) As Byte
                    Dim i As Long
                    For i = 0 To 7
                        lengthBytes(i) = CByte(WsReadRawByte())
                    Next

                    Dim length64 As Double
                    length64 = WsDecodeSafePayloadLength(lengthBytes)
                    If length64 < 0 Then
                        WsSetFailure 9214, "Payload length exceeds Phase 1A limits."
                        Exit Do
                    End If

                    WsPrepareFramePayload CLng(length64)
                End If

            Case WsFrameReadPayload
                If m_ws_framePayloadRemaining = 0 Then
                    m_ws_frameState = WsFrameFinalize
                Else
                    Dim available As Long
                    available = WsRawAvailable
                    If available <= 0 Then Exit Do

                    If WsMonotonicMs() >= copyGuardDeadline Then Exit Do

                    Dim takeCount As Long
                    takeCount = available
                    If takeCount > m_ws_framePayloadRemaining Then takeCount = m_ws_framePayloadRemaining
                    If takeCount > WS_PARSER_COPY_CHUNK_SIZE Then takeCount = WS_PARSER_COPY_CHUNK_SIZE

                    WsConsumeFramePayload takeCount
                    m_ws_framePayloadRemaining = m_ws_framePayloadRemaining - takeCount

                    If m_ws_framePayloadRemaining = 0 Then
                        m_ws_frameState = WsFrameFinalize
                    End If
                End If

            Case WsFrameFinalize
                If WsMonotonicMs() >= pumpDeadline Then Exit Do

                WsCompleteCurrentFrame
                If m_ws_state = WsTransportFailed Then Exit Do
        End Select
    Loop
End Sub

Private Sub WsPrepareFramePayload(ByVal payloadLength As Long)
    If m_ws_frameOpcode >= WS_OPCODE_CLOSE And payloadLength > 125 Then
        WsSetFailure 9215, "Control payload exceeds 125 bytes."
        Exit Sub
    End If

    If m_ws_frameOpcode = WS_OPCODE_CONTINUATION Then
        If Not m_ws_messageActive Then
            WsSetFailure 9216, "Unexpected continuation frame."
            Exit Sub
        End If
    ElseIf m_ws_frameOpcode = WS_OPCODE_TEXT Or m_ws_frameOpcode = WS_OPCODE_BINARY Then
        If m_ws_messageActive Then
            WsSetFailure 9217, "New data frame received during fragmentation."
            Exit Sub
        End If

        m_ws_messageActive = True
        m_ws_messageOpcode = m_ws_frameOpcode
        m_ws_messageCount = 0
        m_ws_discardMessage = (m_ws_frameOpcode = WS_OPCODE_BINARY)

        If m_ws_discardMessage Then
            Set m_ws_messageChunks = Nothing
        Else
            Set m_ws_messageChunks = New Collection
        End If
    End If

    m_ws_framePayloadLength = payloadLength
    m_ws_framePayloadRemaining = payloadLength
    m_ws_frameDiscardPayload = False
    m_ws_controlCount = 0
    Erase m_ws_controlPayload

    If m_ws_frameOpcode <= WS_OPCODE_BINARY Then
        If payloadLength > m_ws_maxIncomingSize Then
            m_ws_discardMessage = True
            m_ws_frameDiscardPayload = True
            Set m_ws_messageChunks = Nothing

        ElseIf Not m_ws_discardMessage Then
            If CDbl(m_ws_messageCount) + CDbl(payloadLength) > CDbl(m_ws_maxIncomingSize) Then
                m_ws_discardMessage = True
                m_ws_frameDiscardPayload = True
                Set m_ws_messageChunks = Nothing
            End If
        End If
    End If

    m_ws_frameState = WsFrameReadPayload
End Sub

Private Sub WsConsumeFramePayload(ByVal count As Long)
    If count <= 0 Then Exit Sub

    If m_ws_frameOpcode >= WS_OPCODE_CLOSE Then
        WsEnsureControlCapacity m_ws_controlCount + count

        Dim i As Long
        For i = 0 To count - 1
            m_ws_controlPayload(m_ws_controlCount + i) = CByte(WsPeekRawByte(i))
        Next
        m_ws_controlCount = m_ws_controlCount + count

    ElseIf Not m_ws_discardMessage And Not m_ws_frameDiscardPayload Then
        If m_ws_messageChunks Is Nothing Then
            WsSetFailure 9219, "Text message chunk accumulator is unavailable."
            Exit Sub
        End If

        Dim dataChunk() As Byte
        ReDim dataChunk(0 To count - 1)

        Dim j As Long
        For j = 0 To count - 1
            dataChunk(j) = CByte(WsPeekRawByte(j))
        Next

        Dim chunkValue As Variant
        chunkValue = dataChunk
        m_ws_messageChunks.Add chunkValue
        m_ws_messageCount = m_ws_messageCount + count
    End If

    WsConsumeRaw count
End Sub

Private Sub WsCompleteCurrentFrame()
    If m_ws_frameOpcode = WS_OPCODE_PING Then
        Dim pongFrame() As Byte
        pongFrame = WsBuildClientFrame(WS_OPCODE_PONG, m_ws_controlPayload, True)

        If m_ws_state = WsTransportFailed Then Exit Sub
        If WsByteArrayLength(pongFrame) = 0 Then Exit Sub

        WsEnqueueSendBytes pongFrame

    ElseIf m_ws_frameOpcode = WS_OPCODE_PONG Then
        WsHandleClientPong

    ElseIf m_ws_frameOpcode = WS_OPCODE_CLOSE Then
        If Not m_ws_closeSent Then
            Dim closeFrame() As Byte
            closeFrame = WsBuildClientFrame(WS_OPCODE_CLOSE, m_ws_controlPayload, True)

            If m_ws_state = WsTransportFailed Then Exit Sub
            If WsByteArrayLength(closeFrame) = 0 Then Exit Sub

            WsEnqueueSendBytes closeFrame
            m_ws_closeSent = True
        End If

        m_ws_state = WsTransportClosing

    ElseIf m_ws_frameOpcode = WS_OPCODE_TEXT Or _
           m_ws_frameOpcode = WS_OPCODE_BINARY Or _
           m_ws_frameOpcode = WS_OPCODE_CONTINUATION Then

        If m_ws_frameFin Then
            If m_ws_messageOpcode = WS_OPCODE_TEXT And Not m_ws_discardMessage Then
                If m_ws_messageChunks Is Nothing Then
                    Set m_ws_messageChunks = New Collection
                End If

                m_ws_rawTextQueue.Add Item:=m_ws_messageChunks
                m_ws_rawTextLengthQueue.Add m_ws_messageCount
            End If

            m_ws_currentPumpCompletedMessages = m_ws_currentPumpCompletedMessages + 1
            WsResetMessageAccumulator
        End If
    End If

    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_framePayloadLength = 0
    m_ws_framePayloadRemaining = 0
    m_ws_frameDiscardPayload = False
    m_ws_controlCount = 0
    Erase m_ws_controlPayload
End Sub

Private Function WsValidateFrameOpcode(ByVal opcode As Long) As Boolean
    Select Case opcode
        Case WS_OPCODE_CONTINUATION, WS_OPCODE_TEXT, WS_OPCODE_BINARY, WS_OPCODE_CLOSE, WS_OPCODE_PING, WS_OPCODE_PONG
            WsValidateFrameOpcode = True
    End Select
End Function

Private Function WsDecodeSafePayloadLength(ByRef Bytes() As Byte) As Double
    If (Bytes(0) And &H80) <> 0 Then
        WsDecodeSafePayloadLength = -1
        Exit Function
    End If

    If Bytes(0) <> 0 Or Bytes(1) <> 0 Or Bytes(2) <> 0 Or Bytes(3) <> 0 Then
        WsDecodeSafePayloadLength = -1
        Exit Function
    End If

    Dim value As Double
    value = CDbl(Bytes(4)) * 16777216# + _
            CDbl(Bytes(5)) * 65536# + _
            CDbl(Bytes(6)) * 256# + _
            CDbl(Bytes(7))

    If value > 2147483647# Then
        WsDecodeSafePayloadLength = -1
    Else
        WsDecodeSafePayloadLength = value
    End If
End Function

Private Sub WsMaybeQueueClientPing()
    If Not m_ws_clientPingEnabled Then Exit Sub
    If m_ws_state <> WsTransportOpen Then Exit Sub

    Dim nowMs As Double
    nowMs = WsMonotonicMs()

    If m_ws_nextClientPingAtMs <= 0 Then
        m_ws_nextClientPingAtMs = nowMs + WS_CLIENT_PING_INTERVAL_MS
        Exit Sub
    End If

    If nowMs < m_ws_nextClientPingAtMs Then Exit Sub

    If m_ws_clientPingOutstanding Then
        m_ws_nextClientPingAtMs = nowMs + WS_CLIENT_PING_INTERVAL_MS
        Exit Sub
    End If

    If WsHasPendingSend Then
        m_ws_nextClientPingAtMs = nowMs + WS_CLIENT_PING_BUSY_RETRY_MS
        Exit Sub
    End If

    If m_ws_clientPingSequence >= &H7FFFFFFF Then
        m_ws_clientPingSequence = 1
    Else
        m_ws_clientPingSequence = m_ws_clientPingSequence + 1
    End If

    Dim payload(0 To WS_CLIENT_PING_PAYLOAD_SIZE - 1) As Byte
    WsEncodePingSequence m_ws_clientPingSequence, payload

    Dim pingFrame() As Byte
    pingFrame = WsBuildClientFrame(WS_OPCODE_PING, payload, True)

    If m_ws_state = WsTransportFailed Then Exit Sub
    If WsByteArrayLength(pingFrame) = 0 Then Exit Sub

    WsEnqueueSendBytes pingFrame

    m_ws_clientPingOutstanding = True
    m_ws_clientPingOutstandingSequence = m_ws_clientPingSequence
    m_ws_nextClientPingAtMs = nowMs + WS_CLIENT_PING_INTERVAL_MS
End Sub

Private Sub WsHandleClientPong()
    If m_ws_controlCount <> WS_CLIENT_PING_PAYLOAD_SIZE Then Exit Sub

    Dim sequence As Long
    If Not WsTryDecodePingSequence(m_ws_controlPayload, sequence) Then Exit Sub

    If Not m_ws_clientPingOutstanding Or _
       sequence <> m_ws_clientPingOutstandingSequence Then
        Exit Sub
    End If

    Dim nowMs As Double
    nowMs = WsMonotonicMs()

    m_ws_clientPingOutstanding = False
    m_ws_clientPingOutstandingSequence = 0
    m_ws_nextClientPingAtMs = nowMs + WS_CLIENT_PING_INTERVAL_MS
End Sub

Private Sub WsEncodePingSequence(ByVal sequence As Long, _
                               ByRef payload() As Byte)
    payload(0) = CByte((sequence \ &H1000000) And &HFF&)
    payload(1) = CByte((sequence \ &H10000) And &HFF&)
    payload(2) = CByte((sequence \ &H100&) And &HFF&)
    payload(3) = CByte(sequence And &HFF&)
End Sub

Private Function WsTryDecodePingSequence(ByRef payload() As Byte, _
                                       ByRef sequence As Long) As Boolean
    Dim value As Double
    value = CDbl(payload(0)) * 16777216# + _
            CDbl(payload(1)) * 65536# + _
            CDbl(payload(2)) * 256# + _
            CDbl(payload(3))

    If value > 2147483647# Then Exit Function

    sequence = CLng(value)
    WsTryDecodePingSequence = True
End Function

Private Function WsBuildClientFrame(ByVal opcode As Long, _
                                  ByRef payload() As Byte, _
                                  ByVal finalFrame As Boolean) As Byte()
    Dim payloadLength As Long
    payloadLength = WsByteArrayLength(payload)

    Dim extendedLengthBytes As Long
    Dim lengthCode As Long

    If payloadLength <= 125 Then
        lengthCode = payloadLength
    ElseIf payloadLength <= 65535 Then
        lengthCode = 126
        extendedLengthBytes = 2
    Else
        lengthCode = 127
        extendedLengthBytes = 8
    End If

    Dim headerLength As Long
    headerLength = 2 + extendedLengthBytes + 4

    Dim frame() As Byte
    ReDim frame(0 To headerLength + payloadLength - 1)

    Dim firstByte As Long
    firstByte = opcode And &HF&
    If finalFrame Then firstByte = firstByte Or &H80&
    frame(0) = CByte(firstByte)
    frame(1) = CByte(&H80& Or lengthCode)

    Dim position As Long
    position = 2

    If lengthCode = 126 Then
        frame(position) = CByte((payloadLength \ &H100&) And &HFF&)
        frame(position + 1) = CByte(payloadLength And &HFF&)
        position = position + 2
    ElseIf lengthCode = 127 Then
        frame(position) = 0
        frame(position + 1) = 0
        frame(position + 2) = 0
        frame(position + 3) = 0
        frame(position + 4) = CByte((payloadLength \ &H1000000) And &HFF&)
        frame(position + 5) = CByte((payloadLength \ &H10000) And &HFF&)
        frame(position + 6) = CByte((payloadLength \ &H100&) And &HFF&)
        frame(position + 7) = CByte(payloadLength And &HFF&)
        position = position + 8
    End If

    Dim mask(0 To WS_MASK_KEY_SIZE - 1) As Byte
    If Not WsGenerateMaskKey(mask) Then Exit Function

    Dim i As Long
    For i = 0 To WS_MASK_KEY_SIZE - 1
        frame(position + i) = mask(i)
    Next
    position = position + WS_MASK_KEY_SIZE

    For i = 0 To payloadLength - 1
        frame(position + i) = payload(i) Xor mask(i Mod WS_MASK_KEY_SIZE)
    Next

    WsBuildClientFrame = frame
End Function

Private Function WsGenerateMaskKey(ByRef mask() As Byte) As Boolean
    Dim status As Long
    status = WsApi_BCryptGenRandom(0, mask(0), WS_MASK_KEY_SIZE, _
                             WS_BCRYPT_USE_SYSTEM_PREFERRED_RNG)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        WsSetFailure 9220, _
                   "BCryptGenRandom failed. NTSTATUS=0x" & _
                   Right$("00000000" & hex$(status), 8)
        Exit Function
    End If

    WsGenerateMaskKey = True
End Function
Private Sub WsEnqueueSendBytes(ByRef Bytes() As Byte)
    Dim value As Variant
    value = Bytes
    m_ws_sendQueue.Add value
End Sub

Private Function WsLoadNextSendBuffer() As Boolean
    If m_ws_sendQueue.count = 0 Then Exit Function

    Dim value As Variant
    value = m_ws_sendQueue(1)
    m_ws_sendQueue.Remove 1

    m_ws_sendCurrent = value
    m_ws_sendCurrentCount = WsByteArrayLength(m_ws_sendCurrent)
    m_ws_sendOffset = 0

    WsLoadNextSendBuffer = (m_ws_sendCurrentCount > 0)
End Function

Private Property Get WsHasPendingSend() As Boolean
    WsHasPendingSend = (m_ws_sendCurrentCount > 0 Or m_ws_sendQueue.count > 0)
End Property

Private Sub WsClearCurrentSendBuffer()
    Erase m_ws_sendCurrent
    m_ws_sendCurrentCount = 0
    m_ws_sendOffset = 0
End Sub

Private Sub WsEnsureRawCapacity(ByVal requiredCount As Long)
    Dim capacity As Long
    capacity = WsByteArrayCapacity(m_ws_rawBuffer)
    If capacity >= requiredCount Then Exit Sub

    If capacity < WS_RECEIVE_CHUNK_SIZE Then capacity = WS_RECEIVE_CHUNK_SIZE
    Do While capacity < requiredCount
        capacity = capacity * 2
    Loop

    If WsByteArrayCapacity(m_ws_rawBuffer) = 0 Then
        ReDim m_ws_rawBuffer(0 To capacity - 1)
    Else
        ReDim Preserve m_ws_rawBuffer(0 To capacity - 1)
    End If
End Sub

Private Sub WsCompactRawIfNeeded()
    If m_ws_rawReadPos = 0 Then Exit Sub

    If m_ws_rawReadPos >= m_ws_rawCount Then
        m_ws_rawReadPos = 0
        m_ws_rawCount = 0
        Exit Sub
    End If

    If m_ws_rawReadPos < WS_RECEIVE_CHUNK_SIZE And m_ws_rawReadPos < (m_ws_rawCount \ 2) Then Exit Sub

    Dim unreadCount As Long
    unreadCount = m_ws_rawCount - m_ws_rawReadPos

    Dim i As Long
    For i = 0 To unreadCount - 1
        m_ws_rawBuffer(i) = m_ws_rawBuffer(m_ws_rawReadPos + i)
    Next

    m_ws_rawCount = unreadCount
    m_ws_rawReadPos = 0
End Sub

Private Property Get WsRawAvailable() As Long
    WsRawAvailable = m_ws_rawCount - m_ws_rawReadPos
End Property

Private Function WsReadRawByte() As Long
    WsReadRawByte = m_ws_rawBuffer(m_ws_rawReadPos)
    m_ws_rawReadPos = m_ws_rawReadPos + 1
    WsNormalizeRawPositions
End Function

Private Function WsPeekRawByte(ByVal offset As Long) As Long
    WsPeekRawByte = m_ws_rawBuffer(m_ws_rawReadPos + offset)
End Function

Private Sub WsConsumeRaw(ByVal count As Long)
    m_ws_rawReadPos = m_ws_rawReadPos + count
    WsNormalizeRawPositions
End Sub

Private Sub WsNormalizeRawPositions()
    If m_ws_rawReadPos >= m_ws_rawCount Then
        m_ws_rawReadPos = 0
        m_ws_rawCount = 0
    End If
End Sub

Private Function WsFindHttpHeaderEnd() As Long
    Dim i As Long
    For i = m_ws_rawReadPos To m_ws_rawCount - 4
        If m_ws_rawBuffer(i) = 13 And _
           m_ws_rawBuffer(i + 1) = 10 And _
           m_ws_rawBuffer(i + 2) = 13 And _
           m_ws_rawBuffer(i + 3) = 10 Then
            WsFindHttpHeaderEnd = (i - m_ws_rawReadPos) + 4
            Exit Function
        End If
    Next
End Function

Private Function WsReadAsciiFromRaw(ByVal count As Long) As String
    Dim text As String
    Dim i As Long

    text = String$(count, vbNullChar)
    For i = 0 To count - 1
        Mid$(text, i + 1, 1) = Chr$(m_ws_rawBuffer(m_ws_rawReadPos + i))
    Next

    WsReadAsciiFromRaw = text
End Function

Private Sub WsEnsureControlCapacity(ByVal requiredCount As Long)
    Dim capacity As Long
    capacity = WsByteArrayCapacity(m_ws_controlPayload)
    If capacity >= requiredCount Then Exit Sub

    If requiredCount < 1 Then requiredCount = 1
    If capacity = 0 Then
        ReDim m_ws_controlPayload(0 To requiredCount - 1)
    Else
        ReDim Preserve m_ws_controlPayload(0 To requiredCount - 1)
    End If
End Sub

Private Sub WsResetMessageAccumulator()
    m_ws_messageActive = False
    m_ws_messageOpcode = 0
    m_ws_messageCount = 0
    m_ws_discardMessage = False
    Set m_ws_messageChunks = Nothing
End Sub

Private Sub WsResetRuntimeState()
    m_ws_LastErrorCode = 0
    m_ws_lastErrorText = vbNullString

    Set m_ws_sendQueue = New Collection
    Set m_ws_rawTextQueue = New Collection
    Set m_ws_rawTextLengthQueue = New Collection
    Set m_ws_messageChunks = Nothing

    Erase m_ws_sendCurrent
    Erase m_ws_rawBuffer
    Erase m_ws_controlPayload

    m_ws_sendCurrentCount = 0
    m_ws_sendOffset = 0
    m_ws_rawCount = 0
    m_ws_rawReadPos = 0
    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_framePayloadLength = 0
    m_ws_framePayloadRemaining = 0
    m_ws_frameDiscardPayload = False
    m_ws_controlCount = 0
    m_ws_closeSent = False
    m_ws_nextClientPingAtMs = 0
    m_ws_clientPingOutstanding = False
    m_ws_clientPingOutstandingSequence = 0
    m_ws_handshakeClientKey = vbNullString
    m_ws_expectedHandshakeAccept = vbNullString
    m_ws_handshakeOrigin = vbNullString
    m_ws_handshakeUserAgent = vbNullString
    m_ws_handshakeValidationError = vbNullString

    WsResetMessageAccumulator
End Sub

Private Sub WsSetFailure(ByVal errorCode As Long, ByVal description As String)
    m_ws_LastErrorCode = errorCode
    m_ws_lastErrorText = description

    On Error Resume Next
    If m_ws_socket <> -1 Then
        WsApi_shutdownSocket m_ws_socket, WS_SD_BOTH
        WsApi_closesocket m_ws_socket
        m_ws_socket = -1
    End If
    On Error GoTo 0

    m_ws_state = WsTransportFailed

    If WS_DEBUG_MODE Then
        Debug.Print "BiDiSocketCommunicator[IntegratedWinsock] FAILED " & errorCode & ": " & description
    End If
End Sub

Private Function WsParseLoopbackUrl(ByVal wsUrl As String) As Boolean
    If Len(wsUrl) = 0 Then Exit Function
    If LCase$(Left$(wsUrl, 5)) <> "ws://" Then Exit Function

    Dim authorityStart As Long
    authorityStart = 6

    Dim pathStart As Long
    pathStart = InStr(authorityStart, wsUrl, "/", vbBinaryCompare)

    Dim authority As String
    If pathStart = 0 Then
        authority = Mid$(wsUrl, authorityStart)
        m_ws_path = "/"
    Else
        authority = Mid$(wsUrl, authorityStart, pathStart - authorityStart)
        m_ws_path = Mid$(wsUrl, pathStart)
    End If

    If Len(authority) = 0 Then Exit Function

    Dim portText As String

    If Left$(authority, 1) = "[" Then
        Dim closeBracket As Long
        closeBracket = InStr(1, authority, "]", vbBinaryCompare)
        If closeBracket = 0 Then Exit Function

        m_ws_host = LCase$(Left$(authority, closeBracket))
        If m_ws_host <> "[::1]" Then Exit Function

        If Len(authority) > closeBracket Then
            If Mid$(authority, closeBracket + 1, 1) <> ":" Then Exit Function
            portText = Mid$(authority, closeBracket + 2)
        End If
    Else
        Dim colonPos As Long
        colonPos = InStrRev(authority, ":", -1, vbBinaryCompare)

        If colonPos > 0 Then
            m_ws_host = LCase$(Left$(authority, colonPos - 1))
            portText = Mid$(authority, colonPos + 1)
        Else
            m_ws_host = LCase$(authority)
        End If

        If m_ws_host <> "localhost" And m_ws_host <> "127.0.0.1" Then Exit Function
    End If

    If Len(portText) = 0 Then
        m_ws_port = 80
    Else
        If Not IsNumeric(portText) Then Exit Function
        m_ws_port = CLng(portText)
        If m_ws_port < 1 Or m_ws_port > 65535 Then Exit Function
    End If

    If m_ws_host = "[::1]" Then
        m_ws_hostHeader = "[::1]:" & CStr(m_ws_port)
    Else
        m_ws_hostHeader = m_ws_host & ":" & CStr(m_ws_port)
    End If

    WsParseLoopbackUrl = True
End Function

Private Function WsHostToNetworkShort(ByVal value As Long) As Integer
    Dim swapped As Long
    swapped = ((value And &HFF&) * &H100&) Or ((value \ &H100&) And &HFF&)

    If swapped > 32767 Then swapped = swapped - 65536
    WsHostToNetworkShort = CInt(swapped)
End Function

Private Function WsEncodeAscii(ByVal text As String) As Byte()
    Dim result() As Byte
    Dim count As Long
    count = Len(text)

    If count = 0 Then
        WsEncodeAscii = result
        Exit Function
    End If

    ReDim result(0 To count - 1)

    Dim i As Long
    For i = 1 To count
        result(i - 1) = CByte(AscW(Mid$(text, i, 1)) And &HFF&)
    Next

    WsEncodeAscii = result
End Function

Private Function WsEncodeUtf8(ByVal text As String) As Byte()
    Dim stream As ADODB.stream
    Set stream = New ADODB.stream

    stream.Type = adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.position = 0
    stream.Type = adTypeBinary

    If stream.size <= 3 Then
        Dim emptyResult() As Byte
        stream.Close
        WsEncodeUtf8 = emptyResult
        Exit Function
    End If

    stream.position = 3
    WsEncodeUtf8 = stream.Read
    stream.Close
End Function

Private Function WsDecodeUtf8Chunks(ByVal chunks As Collection, _
                                  ByVal totalBytes As Long) As String
    If totalBytes <= 0 Then
        WsDecodeUtf8Chunks = vbNullString
        Exit Function
    End If

    Dim stream As ADODB.stream
    Set stream = New ADODB.stream

    stream.Type = adTypeBinary
    stream.Open

    Dim i As Long
    Dim chunkValue As Variant

    For i = 1 To chunks.count
        chunkValue = chunks(i)
        stream.Write chunkValue
    Next

    stream.position = 0
    stream.Type = adTypeText
    stream.Charset = "utf-8"
    WsDecodeUtf8Chunks = stream.ReadText

    stream.Close
    Set stream = Nothing
End Function

Private Function WsByteArrayLength(ByRef Bytes() As Byte) As Long
    On Error GoTo EmptyArray
    WsByteArrayLength = UBound(Bytes) - LBound(Bytes) + 1
    Exit Function

EmptyArray:
    WsByteArrayLength = 0
End Function

Private Function WsByteArrayCapacity(ByRef Bytes() As Byte) As Long
    On Error GoTo EmptyArray
    WsByteArrayCapacity = UBound(Bytes) - LBound(Bytes) + 1
    Exit Function

EmptyArray:
    WsByteArrayCapacity = 0
End Function
