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
' MIT License / Copyright (c) hanamichi77777
Private Const BIDI_SOCKET_VERSION As String = "4.4"
' BiDiSocketCommunicator owns loopback WebSocket transport, RFC 6455 framing,
' request/system IDs, event FIFO and monotonic timing. Retry classification,
' context recovery and SPA synchronization belong to BiDiCommandWrapper.
'
' Preserve these invariants:
'   - Accept only local unencrypted ws:// endpoints.
'   - Never send an Origin header. Firefox RemoteAgent rejects a non-empty
'     Origin with HTTP 400; Chromium drivers ignore it either way.
'   - One SendAndReceive call owns response correlation. Guarded command-correlation/connect-wait
'     paths pump Office messages only when AllowUiPump is explicitly enabled (default False).
'   - m_isReading and m_isRecovering are released only by structured cleanup.
'   - IDs [1..99] are system traffic; normal commands use [100..ID_MAX].
'   - A nonmatching root response ID is dropped as a late/unknown reply, never queued as an event.
'   - Event FIFO retention is bounded by both count and VBA string bytes.
'   - At most one application-level session.status heartbeat is outstanding.
'   - Recovery may heal the socket but never replay the same payload/request ID.
'   - Every native Pump is non-blocking and work-bounded. One size-bounded native UTF-8
'     codec operation may run past the nominal Pump time budget; the hard-overrun guard remains.
'   - Strict UTF-8 validation belongs here; invalid peer text closes with status 1007.
'   - Do not call Trace from a cleanup-critical error handler.
' Outer communicator diagnostics: request/response summaries, state changes and Trace output.
Private Const COMMUNICATOR_DEBUG_ENABLED As Boolean = False
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
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 100
Private Const MIN_POLLING_INTERVAL_MS As Long = 10
Private Const MAX_POLLING_INTERVAL_MS As Long = 5000 ' Must not exceed the application-heartbeat timeout
Private Const EVENT_BUFFER_MAX_COUNT As Long = 2000
Private Const EVENT_BUFFER_MAX_BYTES As Long = 33554432 ' 32 MiB of VBA string storage measured by LenB
Private Const DEFERRED_TEXT_QUEUE_MAX_COUNT As Long = 2000
Private Const DEFERRED_TEXT_QUEUE_TARGET_BYTES As Long = 33554432 ' 32 MiB normal backlog target measured by LenB(String)
Private Const PARSED_ASYNC_EVENT_MAX_CHARS As Long = 524288 ' decoded VBA-character ceiling used by TryGetParsedEvent/Wrapper
Private Const PARSED_ASYNC_EVENT_MAX_UTF8_BYTES As Long = 524288 ' pre-decode UTF-8 budget for a positively identified async event
Private Const COMMAND_UI_PUMP_INTERVAL_MS As Long = 50 ' bounded correlation read slice before a guarded Office yield
Private Const APPLICATION_HEARTBEAT_INTERVAL_MS As Long = 1000
Private Const APPLICATION_HEARTBEAT_TIMEOUT_MS As Long = 5000
Private Const MAX_CONSECUTIVE_APPLICATION_HEARTBEAT_TIMEOUTS As Long = 3
Private Const MAX_COMMAND_CORRELATION_WAIT_MS As Long = 605000 ' 600s operation ceiling + 5s response-correlation margin
Private Const COMMAND_CORRELATION_HARD_OVERRUN_MS As Long = 5000 ' secondary loop-overrun fail-safe only

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
Private m_recoveryConnectAuthorized As Boolean ' Allows only AttemptRecovery to enter public Connect
Private m_recoveryCancelRequested As Boolean ' Lifecycle teardown cancels interactive recovery backoff
Private m_isRaisingStateChanged As Boolean ' Suppresses lifecycle re-entry from event handlers
Private m_isCommandUiYielding As Boolean ' True only while SendAndReceive temporarily yields to Office
Private m_isConnectUiYielding As Boolean ' True only while a Connect wait temporarily yields to Office
Private m_isTerminating As Boolean ' True only during Class_Terminate; state updates continue but events are suppressed

Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum

Private Enum IncomingMessageClass
    IncomingClass_Event = 0
    IncomingClass_MatchingCommandResponse = 1
    IncomingClass_ReservedSystemResponse = 2
    IncomingClass_LateOrUnknownResponse = 3
End Enum

Private m_state As BiDiSocketState
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal StateName As String)

Private m_EventBuffer As Collection
Private m_EventBufferBytes As Long

Private m_DefaultTimeout As Long
Private m_PollingInterval As Long
Private m_DroppedEventsCount As Long
Private m_DroppedLateResponsesCount As Long
Private m_DroppedParsedOversizeCount As Long
Private m_DroppedParsedMalformedCount As Long
Private m_DroppedParsedFilteredCount As Long
Private m_LifecycleEventLossCount As Long
Private m_DownloadEventLossCount As Long
Private m_FileDialogEventLossCount As Long
Private m_UnknownCriticalEventLossCount As Long

' Receive-cap oversize diagnostics are transport-level evidence, separate from parsed-event caps.
' Event-classified oversize messages also increment the existing async event-loss counters.
Private m_TransportOversizedMessageCount As Long
Private m_TransportOversizedEventCount As Long
Private m_TransportOversizedCommandResponseCount As Long
Private m_TransportOversizedUnclassifiedCount As Long
Private m_LastTransportOversizeDiagnostic As String
Private m_activeCommandId As String
Private m_activeCommandOversizedResponseSeen As Boolean
Private m_activeCommandOversizedUnclassifiedSeen As Boolean
Private m_activeCommandOversizeDiagnostic As String

' Unexpected transport continuity is conservative loss evidence for exact event correlation.
' Complete WebSocket text messages are salvaged before transport cleanup; partial receive state
' cannot be reconstructed and is counted separately for diagnostics.
Private m_TransportDiscontinuityCount As Long
Private m_TransportPartialDiscardCount As Long
Private m_SalvagedDeferredMessageCount As Long
Private m_transportDiscontinuityLatched As Boolean
Private m_LastParsedEventDropReason As String
Private m_lastPollTime As Double
Private m_lastApplicationHeartbeatAt As Double
Private m_RegExp As RegExp
Private m_AllowUiPump As Boolean
Private m_qpcFreq As Currency ' QueryPerformanceFrequency cache (set in Class_Initialize)
Private m_lastPumpTicks As Currency ' Throttle state for PumpMessages

Private m_nextRequestId As Long
Private m_nextSystemPingId As Long ' heartbeat id rotation cursor [3..99]
Private m_applicationHeartbeatOutstanding As Boolean
Private m_applicationHeartbeatId As Long
Private m_consecutiveApplicationHeartbeatTimeouts As Long
Private m_LastCommandRequestedWaitMs As Long ' caller-requested correlation wait before communicator cap
Private m_LastCommandEffectiveWaitMs As Long ' effective response-correlation ownership wait for the last command
Private m_LastCommandCorrelationTimedOut As Boolean ' last command exited at the correlation deadline

Private Const MAX_MSG_SIZE As Long = 300000 ' Threshold to prevent VBA hanging (mirrors Wrapper constant)
Private Const ERR_INCOMING_MESSAGE_OVERSIZED As Long = 608
Private Const DEFAULT_MAX_INCOMING_SIZE As Long = 5242880 ' 5 MiB: C7-validated headroom for a 4 MiB payload plus BiDi/JSON envelope
Private Const MAX_CONFIGURED_INCOMING_SIZE As Long = 33554432 ' 32 MiB hard ceiling for VBA memory safety
Private m_maxIncomingSize As Long ' Receive-side cap; oversized messages are drained with bounded identity diagnostics

' Ws*/WS_*/m_ws_* members form the RFC 6455 engine: socket, handshake,
' framing, masking and WebSocket ping/pong. Each Pump is work-bounded; BiDi IDs,
' heartbeats, recovery and the event FIFO remain in the outer communicator layer.

' Lowest-level integrated WinSock transport diagnostics only (OPEN/FAILED).
Private Const WINSOCK_TRANSPORT_DEBUG_ENABLED As Boolean = False

Private Declare PtrSafe Function WsApi_WSAStartup Lib "ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Integer, ByRef lpWSAData As Byte) As Long
Private Declare PtrSafe Function WsApi_WSACleanup Lib "ws2_32.dll" Alias "WSACleanup" () As Long
Private Declare PtrSafe Function WsApi_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socketType As Long, ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function WsApi_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
Private Declare PtrSafe Function WsApi_connectSocketIPv4 Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef socketAddress As WsSOCKADDR_IN, ByVal nameLength As Long) As Long
Private Declare PtrSafe Function WsApi_connectSocketIPv6 Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef socketAddress As WsSOCKADDR_IN6, ByVal nameLength As Long) As Long
Private Declare PtrSafe Function WsApi_sendSocket Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef Buffer As Byte, ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_recvSocket Lib "ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByRef Buffer As Byte, ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As LongPtr, ByVal command As Long, ByRef argument As Long) As Long
Private Declare PtrSafe Function WsApi_selectSocket Lib "ws2_32.dll" Alias "select" (ByVal ignored As Long, ByRef readSet As WsFD_SET, ByRef writeSet As WsFD_SET, ByRef exceptSet As WsFD_SET, ByRef TIMEOUT As WsTIMEVAL) As Long
Private Declare PtrSafe Function WsApi_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optionName As Long, ByRef optionValue As Long, ByRef optionLength As Long) As Long
Private Declare PtrSafe Function WsApi_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optionName As Long, ByRef optionValue As Long, ByVal optionLength As Long) As Long
Private Declare PtrSafe Function WsApi_shutdownSocket Lib "ws2_32.dll" Alias "shutdown" (ByVal s As LongPtr, ByVal how As Long) As Long
Private Declare PtrSafe Function WsApi_WSAGetLastError Lib "ws2_32.dll" Alias "WSAGetLastError" () As Long
Private Declare PtrSafe Function WsApi_InetPtonW_IPv4 Lib "ws2_32.dll" Alias "InetPtonW" (ByVal addressFamily As Long, ByVal addressText As LongPtr, ByRef addressBuffer As Long) As Long
Private Declare PtrSafe Function WsApi_InetPtonW_IPv6 Lib "ws2_32.dll" Alias "InetPtonW" (ByVal addressFamily As Long, ByVal addressText As LongPtr, ByRef addressBuffer As Byte) As Long

Private Declare PtrSafe Sub WsApi_SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function WsApi_QueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef performanceCount As Currency) As Long
Private Declare PtrSafe Function WsApi_QueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef frequency As Currency) As Long
Private Declare PtrSafe Function WsApi_WideCharToMultiByteSize Lib "kernel32" Alias "WideCharToMultiByte" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal wideText As LongPtr, _
    ByVal wideCharCount As Long, ByVal utf8Buffer As LongPtr, ByVal utf8BufferSize As Long, _
    ByVal defaultChar As LongPtr, ByVal usedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function WsApi_WideCharToMultiByteBytes Lib "kernel32" Alias "WideCharToMultiByte" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal wideText As LongPtr, _
    ByVal wideCharCount As Long, ByRef utf8Buffer As Byte, ByVal utf8BufferSize As Long, _
    ByVal defaultChar As LongPtr, ByVal usedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function WsApi_MultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" ( _
    ByVal codePage As Long, ByVal flags As Long, ByRef utf8Text As Byte, _
    ByVal utf8ByteCount As Long, ByVal wideBuffer As LongPtr, ByVal wideBufferSize As Long) As Long

Private Declare PtrSafe Function WsApi_BCryptGenRandom Lib "bcrypt.dll" Alias "BCryptGenRandom" (ByVal hAlgorithm As LongPtr, ByRef randomBuffer As Byte, ByVal bufferSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptOpenAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptOpenAlgorithmProvider" (ByRef algorithmHandle As LongPtr, ByVal algorithmId As LongPtr, ByVal implementation As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptCloseAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptCloseAlgorithmProvider" (ByVal algorithmHandle As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptGetPropertyLong Lib "bcrypt.dll" Alias "BCryptGetProperty" (ByVal objectHandle As LongPtr, ByVal propertyName As LongPtr, ByRef outputBuffer As Long, ByVal outputSize As Long, ByRef resultSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptCreateHash Lib "bcrypt.dll" Alias "BCryptCreateHash" (ByVal algorithmHandle As LongPtr, ByRef hashHandle As LongPtr, ByRef hashObject As Byte, ByVal hashObjectSize As Long, ByVal secret As LongPtr, ByVal secretSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptHashData Lib "bcrypt.dll" Alias "BCryptHashData" (ByVal hashHandle As LongPtr, ByRef inputBuffer As Byte, ByVal inputSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptFinishHash Lib "bcrypt.dll" Alias "BCryptFinishHash" (ByVal hashHandle As LongPtr, ByRef outputBuffer As Byte, ByVal outputSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function WsApi_BCryptDestroyHash Lib "bcrypt.dll" Alias "BCryptDestroyHash" (ByVal hashHandle As LongPtr) As Long


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
Private Const WS_DEFAULT_MAX_INCOMING_SIZE As Long = 5242880 ' Keep native transport default aligned with outer communicator
Private Const OVERSIZED_DIAGNOSTIC_PREFIX_BYTES As Long = 512
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
Private Const WS_CLIENT_PONG_TIMEOUT_MS As Long = 10000
Private Const WS_CLIENT_PING_PAYLOAD_SIZE As Long = 4
Private Const WS_CLIENT_PING_BUSY_RETRY_MS As Long = 1000

Private Const WS_HANDSHAKE_NONCE_SIZE As Long = 16
Private Const WS_WEBSOCKET_GUID As String = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"

' WebSocket text is UTF-8 without a BOM. Strict flags prevent malformed peer data
' or isolated UTF-16 surrogates from being silently replaced.
Private Const WS_CP_UTF8 As Long = 65001
Private Const WS_MB_ERR_INVALID_CHARS As Long = &H8
Private Const WS_WC_ERR_INVALID_CHARS As Long = &H80
Private Const WS_ERROR_NO_UNICODE_TRANSLATION As Long = 1113
Private Const WS_CLOSE_INVALID_PAYLOAD_DATA As Long = 1007
Private Const WS_ERR_INVALID_UTF8 As Long = 9250
Private Const WS_ERR_INVALID_UTF16 As Long = 9251
Private Const WS_ERR_UTF8_CODEC As Long = 9252
Private Const WS_BCRYPT_SHA1_ALGORITHM As String = "SHA1"
Private Const WS_BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const WS_BCRYPT_HASH_LENGTH As String = "HashDigestLength"

Private Const WS_HANDSHAKE_USER_AGENT As String = "SeleniumVBA-BiDi"
Private Const WS_HANDSHAKE_STATUS_ECHO_MAX As Long = 120


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

Private m_ws_host As String
Private m_ws_hostHeader As String
Private m_ws_path As String
Private m_ws_port As Long
Private m_ws_phaseDeadlineMs As Double
Private m_ws_handshakeTimeoutMs As Long

Private m_ws_handshakeClientKey As String
Private m_ws_expectedHandshakeAccept As String
Private m_ws_handshakeUserAgent As String
Private m_ws_handshakeValidationError As String
Private m_ws_urlValidationError As String
Private m_ws_LastErrorCode As Long
Private m_ws_lastProtocolCloseStatus As Long ' RFC 6455 close status; separate from VBA/WSA/internal error codes
Private m_ws_lastErrorText As String

Private m_ws_sendQueue As Collection
Private m_ws_sendCurrent() As Byte
Private m_ws_sendCurrentCount As Long
Private m_ws_sendOffset As Long

Private m_ws_rawBuffer() As Byte
Private m_ws_rawCount As Long
Private m_ws_rawReadPos As Long

Private m_ws_deferredTextQueue As Collection
Private m_ws_deferredTextQueueBytes As Long ' decoded VBA String storage measured by LenB
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
Private m_ws_messageBuffer() As Byte ' used only for fragmented in-limit text messages
Private m_ws_messageCount As Long
Private m_ws_messageTextQueued As Boolean
Private m_ws_directDecodeCandidate As Boolean
Private m_ws_discardMessage As Boolean
Private m_ws_messageObservedBytes As Double
Private m_ws_messageOversized As Boolean
Private m_ws_messageDiagnosticPrefix() As Byte
Private m_ws_messageDiagnosticPrefixCount As Long
Private m_ws_messageIdentityResolved As Boolean
Private m_ws_messageKnownAsyncEvent As Boolean
Private m_ws_messageEarlyParsedOversizeDrop As Boolean
Private m_ws_fragmentReservationDeferred As Boolean
Private m_ws_pendingUtf8Rejection As Boolean
Private m_ws_pendingUtf8RejectionText As String

Private m_ws_closeSent As Boolean

Private m_ws_currentPumpCompletedMessages As Long

Private m_ws_clientPingEnabled As Boolean
Private m_ws_nextClientPingAtMs As Double
Private m_ws_clientPingOutstanding As Boolean
Private m_ws_clientPingSequence As Long
Private m_ws_clientPingOutstandingSequence As Long
Private m_ws_clientPingSentAtMs As Double

Private Sub Class_Initialize()
    If QueryPerformanceFrequency(m_qpcFreq) = 0 Or m_qpcFreq <= 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", _
                  "QueryPerformanceFrequency is unavailable on this system. QPC support is required."
    End If

    m_DefaultTimeout = DEFAULT_TIMEOUT_MS
    m_PollingInterval = DEFAULT_POLLING_INTERVAL_MS
    m_AllowUiPump = False
    m_maxIncomingSize = DEFAULT_MAX_INCOMING_SIZE
    ReDim m_ws_messageDiagnosticPrefix(0 To OVERSIZED_DIAGNOSTIC_PREFIX_BYTES - 1)

    ResetEventBuffer

    Set m_RegExp = New RegExp
    m_RegExp.Global = False
    m_RegExp.IgnoreCase = True

    m_MaxRetryAttempts = 5
    m_AutoReconnect = True

    m_nextRequestId = CMD_ID_FLOOR - 1
    m_nextSystemPingId = SYSTEM_PING_ID_MIN - 1
    ResetApplicationHeartbeatState

    SetState StateDisconnected
End Sub

Private Sub Class_Terminate()
    ' Teardown is non-raising. Suppress outbound StateChanged callbacks before any lifecycle
    ' transition, then clear communicator-owned guards before native handles.
    On Error Resume Next
    m_isTerminating = True
    m_isSendReceiving = False
    m_isReading = False
    m_isCommandUiYielding = False
    m_isConnectUiYielding = False

    Disconnect
    If m_wsTransportInitialized Then
        WsTransportTerminate
        m_wsTransportInitialized = False
    End If
    Set m_RegExp = Nothing
    Set m_EventBuffer = Nothing
    m_EventBufferBytes = 0
    ResetApplicationHeartbeatState
    On Error GoTo 0
End Sub

' Mirrors BiDiCommandWrapper.Version. A mismatch between the two classes means the
' pair was imported from different releases.
Public Property Get Version() As String
    Version = BIDI_SOCKET_VERSION
End Property

' Logical connectivity requires agreement between the communicator state machine and
' the native WebSocket. StateError must never surface as connected merely because a stale
' native handle still reports OPEN (notably when AutoReconnect is disabled).
Public Property Get IsConnected() As Boolean
    ' VBA And does not short-circuit; guard native access explicitly.
    If m_state <> StateConnected Then Exit Property
    If Not m_wsTransportInitialized Then Exit Property
    IsConnected = WsIsOpen
End Property

' Diagnostic-only native state. This may remain True briefly while logical state is
' StateError; callers deciding whether commands are usable must use IsConnected instead.
Public Property Get IsTransportOpen() As Boolean
    If m_wsTransportInitialized Then IsTransportOpen = WsIsOpen
End Property

' Compatibility property retained for older callers. The integrated WinSock
' transport is always enabled and can only be confirmed while disconnected.
Public Property Let UseWinsockTransport(ByVal v As Boolean)
    If Not v Then
        Err.Raise vbObjectError + 9105, "BiDiSocketCommunicator", _
                  "The integrated WinSock transport is always enabled and cannot be disabled."
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

' Polling drives both event receive work and the application-heartbeat opportunity.
' Clamp it away from a hot loop and below the heartbeat timeout.
Public Property Let PollingInterval(ByVal v As Long)
    If v < MIN_POLLING_INTERVAL_MS Then v = MIN_POLLING_INTERVAL_MS
    If v > MAX_POLLING_INTERVAL_MS Then v = MAX_POLLING_INTERVAL_MS
    m_PollingInterval = v
End Property
Public Property Get PollingInterval() As Long: PollingInterval = m_PollingInterval: End Property
' Office message pumping during guarded waits is opt-in; the default is False.
Public Property Let AllowUiPump(ByVal v As Boolean): m_AllowUiPump = v: End Property
Public Property Get AllowUiPump() As Boolean: AllowUiPump = m_AllowUiPump: End Property
Public Property Get state() As BiDiSocketState: state = m_state: End Property
Public Property Get MaxCommandCorrelationWaitMs() As Long: MaxCommandCorrelationWaitMs = MAX_COMMAND_CORRELATION_WAIT_MS: End Property
Public Property Get LastCommandRequestedWaitMs() As Long: LastCommandRequestedWaitMs = m_LastCommandRequestedWaitMs: End Property
Public Property Get LastCommandEffectiveWaitMs() As Long: LastCommandEffectiveWaitMs = m_LastCommandEffectiveWaitMs: End Property
Public Property Get LastCommandCorrelationTimedOut() As Boolean: LastCommandCorrelationTimedOut = m_LastCommandCorrelationTimedOut: End Property
Public Property Let AutoReconnect(ByVal v As Boolean): m_AutoReconnect = v: End Property
Public Property Get AutoReconnect() As Boolean: AutoReconnect = m_AutoReconnect: End Property
Public Property Get ConsecutiveInternalErrors() As Long: ConsecutiveInternalErrors = m_consecutiveInternalErrors: End Property
' MaxIncomingSize is a receive-assembly byte cap. Keep the existing command-size floor,
' and impose a hard upper ceiling so a public setting cannot disable memory protection.
Public Property Let MaxIncomingSize(ByVal v As Long)
    If m_isConnectUiYielding Then Exit Property
    If m_isCommandUiYielding Then Exit Property
    If v < MAX_MSG_SIZE Then v = MAX_MSG_SIZE
    If v > MAX_CONFIGURED_INCOMING_SIZE Then v = MAX_CONFIGURED_INCOMING_SIZE
    m_maxIncomingSize = v

    If m_wsTransportInitialized Then
        WsMaxIncomingSize = v
    End If

End Property
Public Property Get MaxIncomingSize() As Long: MaxIncomingSize = m_maxIncomingSize: End Property
Public Property Get ParsedAsyncEventMaxChars() As Long: ParsedAsyncEventMaxChars = PARSED_ASYNC_EVENT_MAX_CHARS: End Property
Public Property Get ParsedAsyncEventMaxUtf8Bytes() As Long: ParsedAsyncEventMaxUtf8Bytes = PARSED_ASYNC_EVENT_MAX_UTF8_BYTES: End Property
Public Property Get DroppedEventsCount() As Long: DroppedEventsCount = m_DroppedEventsCount: End Property
Public Property Get TotalDroppedEventsCount() As Long
    Dim total As Double
    total = CDbl(m_DroppedEventsCount) + CDbl(m_DroppedParsedOversizeCount) + _
            CDbl(m_DroppedParsedMalformedCount) + CDbl(m_DroppedParsedFilteredCount)
    If total >= 2147483647# Then
        TotalDroppedEventsCount = 2147483647
    Else
        TotalDroppedEventsCount = CLng(total)
    End If
End Property
Public Property Get TotalAsyncEventLossCount() As Long
    Dim total As Double
    total = CDbl(m_DroppedEventsCount) + CDbl(m_DroppedParsedOversizeCount) + _
            CDbl(m_DroppedParsedMalformedCount) + CDbl(m_TransportDiscontinuityCount) + _
            CDbl(m_TransportOversizedUnclassifiedCount)
    If total >= 2147483647# Then
        TotalAsyncEventLossCount = 2147483647
    Else
        TotalAsyncEventLossCount = CLng(total)
    End If
End Property
' Current depth of the bounded asynchronous-event FIFO. Diagnostic only; it does not classify
' observer health and does not include messages that have not yet reached this FIFO.
Public Property Get EventQueueDepth() As Long
    If m_EventBuffer Is Nothing Then Exit Property
    EventQueueDepth = m_EventBuffer.count
End Property
Public Property Get DroppedLateResponsesCount() As Long: DroppedLateResponsesCount = m_DroppedLateResponsesCount: End Property
Public Property Get DroppedParsedOversizeCount() As Long: DroppedParsedOversizeCount = m_DroppedParsedOversizeCount: End Property
Public Property Get DroppedParsedMalformedCount() As Long: DroppedParsedMalformedCount = m_DroppedParsedMalformedCount: End Property
Public Property Get DroppedParsedFilteredCount() As Long: DroppedParsedFilteredCount = m_DroppedParsedFilteredCount: End Property
Public Property Get LifecycleEventLossCount() As Long: LifecycleEventLossCount = m_LifecycleEventLossCount: End Property
Public Property Get DownloadEventLossCount() As Long: DownloadEventLossCount = m_DownloadEventLossCount: End Property
Public Property Get FileDialogEventLossCount() As Long: FileDialogEventLossCount = m_FileDialogEventLossCount: End Property
Public Property Get UnknownCriticalEventLossCount() As Long: UnknownCriticalEventLossCount = m_UnknownCriticalEventLossCount: End Property
Public Property Get TransportOversizedMessageCount() As Long: TransportOversizedMessageCount = m_TransportOversizedMessageCount: End Property
Public Property Get TransportOversizedEventCount() As Long: TransportOversizedEventCount = m_TransportOversizedEventCount: End Property
Public Property Get TransportOversizedCommandResponseCount() As Long: TransportOversizedCommandResponseCount = m_TransportOversizedCommandResponseCount: End Property
Public Property Get TransportOversizedUnclassifiedCount() As Long: TransportOversizedUnclassifiedCount = m_TransportOversizedUnclassifiedCount: End Property
Public Property Get LastTransportOversizeDiagnostic() As String: LastTransportOversizeDiagnostic = m_LastTransportOversizeDiagnostic: End Property
Public Property Get TransportDiscontinuityCount() As Long: TransportDiscontinuityCount = m_TransportDiscontinuityCount: End Property
Public Property Get TransportPartialDiscardCount() As Long: TransportPartialDiscardCount = m_TransportPartialDiscardCount: End Property
Public Property Get SalvagedDeferredMessageCount() As Long: SalvagedDeferredMessageCount = m_SalvagedDeferredMessageCount: End Property
Public Property Get LastParsedEventDropReason() As String: LastParsedEventDropReason = m_LastParsedEventDropReason: End Property

Private Sub SetState(ByVal NewState As BiDiSocketState)
    If m_state = NewState Then Exit Sub

    ' Only an established session can suffer an observable continuity break. Connection-attempt
    ' failures occur in StateConnecting and therefore do not poison fresh-session loss counters.
    If NewState = StateError And m_state = StateConnected Then
        MarkUnexpectedTransportDiscontinuity "StateConnected -> StateError"
    End If

    m_state = NewState

    Dim StateName As String
    Select Case m_state
        Case StateDisconnected: StateName = "DISCONNECTED"
        Case StateConnecting:   StateName = "CONNECTING"
        Case StateConnected:    StateName = "CONNECTED"
        Case StateDisconnecting: StateName = "DISCONNECTING"
        Case StateError:         StateName = "ERROR"
    End Select

    If COMMUNICATOR_DEBUG_ENABLED Then Debug.Print "BiDi State Changed: " & StateName

    ' Class_Terminate may still need the internal state transition for deterministic cleanup,
    ' but object destruction must never synchronously call back into an external subscriber.
    If m_isTerminating Then Exit Sub

    ' StateChanged is advisory. Subscriber errors must not interrupt cleanup, and transport
    ' or command entry points are rejected until the synchronous notification returns.
    If m_isRaisingStateChanged Then Exit Sub

    m_isRaisingStateChanged = True
    On Error Resume Next
    Err.Clear
    RaiseEvent StateChanged(m_state, StateName)
    If Err.Number <> 0 Then
        Debug.Print "BiDi: StateChanged handler error ignored: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    m_isRaisingStateChanged = False
End Sub

Public Function Connect(ByVal wsUrl As String, ByVal localPort As Long) As Boolean
    If m_isConnectUiYielding Then
        Trace ">>> [Lifecycle] Connect rejected during connection-wait UI yield."
        Exit Function
    End If

    If m_isCommandUiYielding Then
        Trace ">>> [Lifecycle] Connect rejected during command-correlation UI yield."
        Exit Function
    End If

    If m_isRaisingStateChanged Then
        Trace ">>> [Lifecycle] Connect rejected during StateChanged notification. " & _
              "Schedule the connection attempt after the event handler returns.", True
        Exit Function
    End If

    If m_isRecovering And Not m_recoveryConnectAuthorized Then
        Trace ">>> [Lifecycle] Connect rejected while recovery owns the connection lifecycle.", True
        Exit Function
    End If

    If m_state = StateConnecting Or m_state = StateConnected Then
        If COMMUNICATOR_DEBUG_ENABLED Then Debug.Print "BiDi: Connect aborted. Current state is " & m_state
        Exit Function
    End If

    If LenB(wsUrl) = 0 Or localPort < 1 Or localPort > 65535 Then
        Trace "[Security] Blocked: WebSocket URL is empty or localPort is outside 1..65535.", True
        SetState StateError
        Exit Function
    End If

    ' A direct Connect starts a fresh diagnostic session. Recovery reconnects keep
    ' the original transport failure and retry history in the same trace file.
    If Not m_isRecovering Then ResetSocketTraceLog

    ' Reject anything except loopback ws:// before entering the RFC 6455 parser.
    m_RegExp.pattern = "^ws://(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?(?:/|$)"
    m_RegExp.IgnoreCase = True

    If Not m_RegExp.Test(wsUrl) Then
        Trace "[Security] Blocked: endpoint must be a ws:// loopback URL " & _
              "using localhost, 127.0.0.1, or [::1].", True
        SetState StateError
        Exit Function
    End If

    Dim effectiveWsUrl As String
    Dim normalizeError As String
    If Not NormalizeLoopbackWsUrlPort(wsUrl, localPort, effectiveWsUrl, normalizeError) Then
        Trace "[Security] Blocked: " & normalizeError, True
        SetState StateError
        Exit Function
    End If

    m_LastWsUrl = effectiveWsUrl
    m_LastPort = localPort

    On Error GoTo ErrorHandler

    SetState StateConnecting
    InternalCleanup

    ' An explicit new connection starts a new event/diagnostic session. Recovery reconnects
    ' preserve both the FIFO and drop counters accumulated before the transport fault.
    If Not m_isRecovering Then
        ResetEventBuffer
        m_DroppedEventsCount = 0
        m_DroppedLateResponsesCount = 0
        m_DroppedParsedOversizeCount = 0
        m_DroppedParsedMalformedCount = 0
        m_DroppedParsedFilteredCount = 0
        m_LifecycleEventLossCount = 0
        m_DownloadEventLossCount = 0
        m_FileDialogEventLossCount = 0
        m_UnknownCriticalEventLossCount = 0
        m_TransportOversizedMessageCount = 0
        m_TransportOversizedEventCount = 0
        m_TransportOversizedCommandResponseCount = 0
        m_TransportOversizedUnclassifiedCount = 0
        m_LastTransportOversizeDiagnostic = vbNullString
        m_TransportDiscontinuityCount = 0
        m_TransportPartialDiscardCount = 0
        m_SalvagedDeferredMessageCount = 0
        m_transportDiscontinuityLatched = False
        m_LastParsedEventDropReason = vbNullString
        m_LastCommandRequestedWaitMs = 0
        m_LastCommandEffectiveWaitMs = 0
        m_LastCommandCorrelationTimedOut = False
        m_consecutiveInternalErrors = 0
    End If

    If ConnectWithWinsock(effectiveWsUrl) Then
        Connect = True
        Exit Function
    End If

    ' A guarded Office yield may let an explicit Disconnect/ForceClose cancel an in-flight
    ' connection attempt. Preserve that caller-owned lifecycle result instead of converting the
    ' cancellation into a synthetic connection error when control returns here.
    If m_state <> StateConnecting Then
        Connect = False
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

    If m_state <> StateConnecting Then Exit Function

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

        ' Connection waits stay responsive without turning the non-blocking socket poll into a
        ' hot Sleep(0) loop. The dedicated yield guard blocks synchronous command/I/O re-entry.
        If m_AllowUiPump Then
            YieldConnectWaitUi
            If m_state <> StateConnecting Then Exit Function
        End If
        SleepWinAPI 1
    Loop

ErrorHandler:
    m_ws_LastErrorCode = Err.Number
    m_ws_lastErrorText = "TryConnectWithWinsockUrl VBA error at " & diagnosticStage & ": " & Err.description
    TryConnectWithWinsockUrl = False
End Function

Private Function NormalizeLoopbackWsUrlPort(ByVal wsUrl As String, _
                                                ByVal localPort As Long, _
                                                ByRef normalizedUrl As String, _
                                                ByRef errorText As String) As Boolean
    normalizedUrl = vbNullString
    errorText = vbNullString

    On Error GoTo InvalidUrl

    If localPort < 1 Or localPort > 65535 Then
        errorText = "localPort must be in 1..65535."
        Exit Function
    End If

    If LCase$(Left$(wsUrl, 5)) <> "ws://" Then
        errorText = "endpoint must use ws://."
        Exit Function
    End If

    Dim authorityStart As Long
    authorityStart = 6

    Dim pathStart As Long
    pathStart = InStr(authorityStart, wsUrl, "/", vbBinaryCompare)
    If pathStart = 0 Then pathStart = Len(wsUrl) + 1

    Dim authority As String
    authority = Mid$(wsUrl, authorityStart, pathStart - authorityStart)
    If Len(authority) = 0 Then
        errorText = "WebSocket authority is empty."
        Exit Function
    End If

    Dim explicitPortText As String
    If Left$(authority, 1) = "[" Then
        Dim closeBracket As Long
        closeBracket = InStr(1, authority, "]", vbBinaryCompare)
        If closeBracket = 0 Then
            errorText = "IPv6 loopback authority is missing ']'."
            Exit Function
        End If

        If Len(authority) > closeBracket Then
            If Mid$(authority, closeBracket + 1, 1) <> ":" Then
                errorText = "Unexpected data follows the IPv6 loopback host."
                Exit Function
            End If
            explicitPortText = Mid$(authority, closeBracket + 2)
        End If
    Else
        Dim colonPos As Long
        colonPos = InStrRev(authority, ":", -1, vbBinaryCompare)
        If colonPos > 0 Then explicitPortText = Mid$(authority, colonPos + 1)
    End If

    If Len(explicitPortText) > 0 Then
        If Not IsNumeric(explicitPortText) Then
            errorText = "WebSocket URL contains a non-numeric port."
            Exit Function
        End If

        Dim explicitPort As Long
        explicitPort = CLng(explicitPortText)
        If explicitPort < 1 Or explicitPort > 65535 Then
            errorText = "WebSocket URL port is outside 1..65535."
            Exit Function
        End If

        If explicitPort <> localPort Then
            errorText = "WebSocket URL port " & CStr(explicitPort) & _
                        " does not match localPort " & CStr(localPort) & "."
            Exit Function
        End If

        normalizedUrl = wsUrl
    Else
        normalizedUrl = Left$(wsUrl, pathStart - 1) & ":" & CStr(localPort) & Mid$(wsUrl, pathStart)
    End If

    NormalizeLoopbackWsUrlPort = True
    Exit Function

InvalidUrl:
    errorText = "WebSocket URL port normalization failed: " & Err.description
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
    m_lastPollTime = 0
    m_nextSystemPingId = SYSTEM_PING_ID_MIN - 1
    ResetApplicationHeartbeatState
    m_RetryCount = 0
    ' A successful OPEN begins a new transport epoch. Recovery preserves cumulative counters,
    ' but a later independent discontinuity must be counted once again.
    m_transportDiscontinuityLatched = False
    SetState StateConnected
End Sub

Private Function GetWinsockFailureText() As String
    Dim details As String

    If m_ws_LastErrorCode <> 0 Or Len(m_ws_lastErrorText) > 0 Then
        details = CStr(m_ws_LastErrorCode) & " / " & m_ws_lastErrorText
    ElseIf Not m_wsTransportInitialized Then
        details = "integrated transport was not initialized"
    Else
        details = CStr(WsLastErrorCode) & " / " & WsLastErrorText
    End If

    If m_ws_lastProtocolCloseStatus <> 0 Then
        details = details & " / WebSocket close=" & CStr(m_ws_lastProtocolCloseStatus)
    End If

    GetWinsockFailureText = details
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

' allowRecovery=False is PostScript's best-effort send-only path. Recovery may repair a
' failed socket, but ordinary state/transport send failures return False and never replay the
' payload. Deterministic invalid UTF-16 caller input is the one raising case: encoding fails
' before any frame is built or queued. PostScript intentionally suppresses that error.
Public Function SendMessage(ByVal msg As String, Optional ByVal allowRecovery As Boolean = True) As Boolean
    If m_isConnectUiYielding Then Exit Function
    If m_isCommandUiYielding Then Exit Function
    If m_isRaisingStateChanged Or m_isRecovering Then Exit Function
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

' Fire-and-forget contract: invalid or oversized payloads are dropped.
' This path must never raise to the caller.
Public Sub PostScript(ByVal msg As String)
    On Error Resume Next

    If Len(msg) > MAX_MSG_SIZE Then
        Trace ">>> [PostScript] Outgoing payload dropped. Size=" & _
              CStr(Len(msg)) & " chars; Max=" & CStr(MAX_MSG_SIZE) & ".", True
        Exit Sub
    End If

    SendMessage msg, False
    On Error GoTo 0
End Sub

Public Function GetNextRequestId() As Long
    m_nextRequestId = m_nextRequestId + 1
    If m_nextRequestId > ID_MAX Then m_nextRequestId = CMD_ID_FLOOR
    GetNextRequestId = m_nextRequestId
End Function

' Heartbeat IDs rotate inside [3..99], separate from normal command IDs. Only one
' application heartbeat is tracked; a timed-out ID is retired before allocating the next.
Private Function GetNextSystemPingId() As Long
    m_nextSystemPingId = m_nextSystemPingId + 1
    If m_nextSystemPingId > SYSTEM_PING_ID_MAX Then m_nextSystemPingId = SYSTEM_PING_ID_MIN
    GetNextSystemPingId = m_nextSystemPingId
End Function

' Only integer numeric root IDs in [1..CMD_ID_FLOOR) are reserved replies. An id-less
' BiDi event is asynchronous application traffic. Keep integer validation aligned with
' IsNormalCommandId so malformed fractional IDs cannot enter the reserved band.
Private Function IsReservedSystemId(ByVal idStr As String) As Boolean
    On Error GoTo InvalidId

    If Len(idStr) = 0 Then Exit Function
    If Not IsNumeric(idStr) Then Exit Function

    Dim v As Double
    v = CDbl(idStr)
    IsReservedSystemId = (v >= 1 And v < CMD_ID_FLOOR And v = Fix(v))
    Exit Function

InvalidId:
    IsReservedSystemId = False
End Function


' SendAndReceive is a normal-command API. Requiring its own root ID prevents an id-less
' event from being mistaken for the response and keeps the reserved system band private.
Private Function IsNormalCommandId(ByVal idStr As String) As Boolean
    On Error GoTo InvalidId

    If Len(idStr) = 0 Then Exit Function
    If Not IsNumeric(idStr) Then Exit Function

    Dim v As Double
    v = CDbl(idStr)
    IsNormalCommandId = (v >= CMD_ID_FLOOR And v <= ID_MAX And v = Fix(v))
    Exit Function

InvalidId:
    IsNormalCommandId = False
End Function

' All receive sites use the same root-ID classification. expectedCommandId is supplied
' only while SendAndReceive owns a command; outside correlation, every normal-band root ID
' is necessarily a late or otherwise unknown response.
Private Function ClassifyIncomingMessage(ByVal inId As String, _
                                         Optional ByVal expectedCommandId As String = "") As IncomingMessageClass
    If Len(expectedCommandId) > 0 And inId = expectedCommandId Then
        ClassifyIncomingMessage = IncomingClass_MatchingCommandResponse
    ElseIf IsReservedSystemId(inId) Then
        ClassifyIncomingMessage = IncomingClass_ReservedSystemResponse
    ElseIf Len(inId) > 0 Then
        ClassifyIncomingMessage = IncomingClass_LateOrUnknownResponse
    Else
        ClassifyIncomingMessage = IncomingClass_Event
    End If
End Function

' Root IDs other than the currently owned command or reserved system IDs are stale or
' unknown replies. They are counted separately and never consume asynchronous FIFO space.
Private Sub DropLateOrUnknownResponse(ByVal idStr As String)
    ' Any root response proves peer progress even when its command owner already timed out.
    m_consecutiveApplicationHeartbeatTimeouts = 0

    If m_DroppedLateResponsesCount < 2147483647 Then
        m_DroppedLateResponsesCount = m_DroppedLateResponsesCount + 1
    End If

    If COMMUNICATOR_DEBUG_ENABLED Then
        Debug.Print "BiDi: Dropped late/unknown response ID=" & idStr
    End If
End Sub

Private Sub TraceSessionDropSummary()
    If m_DroppedEventsCount = 0 And _
       m_DroppedLateResponsesCount = 0 And _
       m_DroppedParsedOversizeCount = 0 And _
       m_DroppedParsedMalformedCount = 0 And _
       m_DroppedParsedFilteredCount = 0 And _
       m_TransportOversizedMessageCount = 0 And _
       m_TransportDiscontinuityCount = 0 And _
       m_SalvagedDeferredMessageCount = 0 Then Exit Sub

    Trace ">>> [Session Summary] Connection closing. Dropped asynchronous events=" & _
          m_DroppedEventsCount & "; dropped late/unknown command responses=" & _
          m_DroppedLateResponsesCount & "; parsed-event oversize drops=" & _
          m_DroppedParsedOversizeCount & "; malformed drops=" & _
          m_DroppedParsedMalformedCount & "; filtered drops=" & _
          m_DroppedParsedFilteredCount & "; lifecycle losses=" & _
          m_LifecycleEventLossCount & "; download losses=" & _
          m_DownloadEventLossCount & "; file-dialog losses=" & _
          m_FileDialogEventLossCount & "; unknown critical losses=" & _
          m_UnknownCriticalEventLossCount & "; transport oversized messages=" & _
          m_TransportOversizedMessageCount & " (events=" & m_TransportOversizedEventCount & _
          ", commandResponses=" & m_TransportOversizedCommandResponseCount & _
          ", unclassified=" & m_TransportOversizedUnclassifiedCount & "); transport discontinuities=" & _
          m_TransportDiscontinuityCount & "; partial receive discards=" & _
          m_TransportPartialDiscardCount & "; salvaged deferred messages=" & _
          m_SalvagedDeferredMessageCount, True
End Sub

Private Sub IncrementLossCounter(ByRef counter As Long)
    If counter < 2147483647 Then counter = counter + 1
End Sub

Private Function HasPartialTransportReceiveState() As Boolean
    ' A complete message already in m_ws_deferredTextQueue is NOT partial. This function only
    ' diagnoses parser/raw bytes that cannot survive WsForceClose.
    If m_ws_messageActive Then
        HasPartialTransportReceiveState = True
        Exit Function
    End If

    If m_ws_frameState <> WsFrameReadBaseHeader Then
        HasPartialTransportReceiveState = True
        Exit Function
    End If

    If m_ws_rawCount > m_ws_rawReadPos Then
        HasPartialTransportReceiveState = True
    End If
End Function

Private Sub MarkUnexpectedTransportDiscontinuity(ByVal reason As String)
    If m_transportDiscontinuityLatched Then Exit Sub
    m_transportDiscontinuityLatched = True

    IncrementLossCounter m_TransportDiscontinuityCount

    ' Existing wrapper capture guards already include UnknownCriticalEventLossCount in
    ' lifecycle/download/file-dialog loss baselines. Reusing that channel makes every
    ' exact-correlation arm fail safe without changing the wrapper API.
    IncrementLossCounter m_UnknownCriticalEventLossCount

    Dim hadPartial As Boolean
    hadPartial = HasPartialTransportReceiveState()
    If hadPartial Then IncrementLossCounter m_TransportPartialDiscardCount

    Trace ">>> [Transport Continuity] Unexpected stream discontinuity recorded. " & _
          "Reason=" & reason & "; partialReceive=" & CStr(hadPartial) & _
          "; deferredDepth=" & CStr(WsDeferredTextQueueDepth()) & _
          "; deferredBytes=" & CStr(m_ws_deferredTextQueueBytes) & ".", True
End Sub

Private Function WsDeferredTextQueueDepth() As Long
    If m_ws_deferredTextQueue Is Nothing Then Exit Function
    WsDeferredTextQueueDepth = m_ws_deferredTextQueue.count
End Function
Private Sub SalvageDeferredMessagesAfterTransportFailure(ByVal reason As String)
    ' Only COMPLETE text messages reach the deferred queue. Move them into the upper event FIFO
    ' before WsForceClose resets transport-owned buffers. Numeric replies have no active command
    ' owner at the call sites that use this helper and are therefore classified normally.
    Dim candidate As String
    Dim inId As String
    Dim salvaged As Long

    On Error GoTo SalvageFailed

    Do While WsTryDequeueText(candidate)
        salvaged = salvaged + 1
        IncrementLossCounter m_SalvagedDeferredMessageCount

        inId = ExtractJsonId(candidate)
        Select Case ClassifyIncomingMessage(inId)
            Case IncomingClass_ReservedSystemResponse
                ProcessReservedSystemResponse inId, candidate, False

            Case IncomingClass_LateOrUnknownResponse
                DropLateOrUnknownResponse inId

            Case IncomingClass_Event
                ObserveWinsockMessage candidate
                EnqueueEventBounded candidate
        End Select

        candidate = vbNullString
        inId = vbNullString
    Loop

    If salvaged > 0 Then
        Trace ">>> [Transport Continuity] Salvaged " & CStr(salvaged) & _
              " complete deferred message(s) before cleanup. Reason=" & reason & ".", True
    End If
    Exit Sub

SalvageFailed:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & _
                "] [Transport Continuity] Deferred salvage error: " & Err.description
    Err.Clear
End Sub

' Extracts only a top-level BiDi method from a bounded prefix in one pass.
' Escaped method keys/values are treated conservatively as unclassified; standard BiDi
' method names are unescaped ASCII tokens.
Private Function TryExtractTopLevelEventMethod(ByVal rawMsg As String, _
                                               ByRef methodName As String) As Boolean
    Const METHOD_SCAN_LIMIT As Long = 4096

    methodName = vbNullString

    Dim scanText As String
    scanText = rawMsg
    If Len(scanText) > METHOD_SCAN_LIMIT Then scanText = Left$(scanText, METHOD_SCAN_LIMIT)

    Dim textLength As Long
    textLength = Len(scanText)

    Dim i As Long
    i = 1
    Do While i <= textLength
        Dim leading As String
        leading = Mid$(scanText, i, 1)
        If leading <> " " And leading <> vbTab And leading <> vbCr And leading <> vbLf Then Exit Do
        i = i + 1
    Loop

    If i > textLength Then Exit Function
    If Mid$(scanText, i, 1) <> "{" Then Exit Function

    Dim depth As Long
    depth = 1
    i = i + 1

    Do While i <= textLength And depth > 0
        Dim ch As String
        ch = Mid$(scanText, i, 1)

        If ch = """" Then
            Dim tokenStart As Long
            Dim tokenEnd As Long
            Dim tokenEscaped As Boolean
            Dim p As Long

            tokenStart = i + 1
            tokenEnd = 0
            tokenEscaped = False
            p = tokenStart

            Do While p <= textLength
                ch = Mid$(scanText, p, 1)
                If ch = "\" Then
                    tokenEscaped = True
                    p = p + 2
                ElseIf ch = """" Then
                    tokenEnd = p
                    Exit Do
                Else
                    p = p + 1
                End If
            Loop

            If tokenEnd = 0 Then Exit Function

            If depth = 1 And Not tokenEscaped Then
                Dim afterKey As Long
                afterKey = tokenEnd + 1
                Do While afterKey <= textLength
                    ch = Mid$(scanText, afterKey, 1)
                    If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
                    afterKey = afterKey + 1
                Loop

                If afterKey <= textLength And Mid$(scanText, afterKey, 1) = ":" Then
                    If Mid$(scanText, tokenStart, tokenEnd - tokenStart) = "method" Then
                        Dim valuePos As Long
                        valuePos = afterKey + 1
                        Do While valuePos <= textLength
                            ch = Mid$(scanText, valuePos, 1)
                            If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
                            valuePos = valuePos + 1
                        Loop

                        If valuePos > textLength Or Mid$(scanText, valuePos, 1) <> """" Then Exit Function

                        Dim valueStart As Long
                        valueStart = valuePos + 1
                        p = valueStart
                        Do While p <= textLength
                            ch = Mid$(scanText, p, 1)
                            If ch = "\" Then
                                Exit Function
                            ElseIf ch = """" Then
                                methodName = Mid$(scanText, valueStart, p - valueStart)
                                TryExtractTopLevelEventMethod = (Len(methodName) > 0)
                                Exit Function
                            End If
                            p = p + 1
                        Loop
                        Exit Function
                    End If
                End If
            End If

            i = tokenEnd + 1

        Else
            Select Case ch
                Case "{", "["
                    depth = depth + 1
                Case "}", "]"
                    depth = depth - 1
                    If depth < 0 Then Exit Function
            End Select
            i = i + 1
        End If
    Loop
End Function

Private Sub RecordAsyncMessageLoss(ByVal rawMsg As String, _
                                   ByRef detectedMethod As String)
    detectedMethod = vbNullString

    If Not TryExtractTopLevelEventMethod(rawMsg, detectedMethod) Then
        IncrementLossCounter m_UnknownCriticalEventLossCount
        Exit Sub
    End If

    Select Case detectedMethod
        Case "browsingContext.contextCreated", _
             "browsingContext.contextDestroyed", _
             "browsingContext.navigationStarted"
            IncrementLossCounter m_LifecycleEventLossCount

        Case "browsingContext.downloadWillBegin", _
             "browsingContext.downloadEnd"
            IncrementLossCounter m_DownloadEventLossCount

        Case "input.fileDialogOpened"
            IncrementLossCounter m_FileDialogEventLossCount
    End Select
End Sub

Private Sub ResetEventBuffer()
    Set m_EventBuffer = New Collection
    m_EventBufferBytes = 0
End Sub

Private Sub DropOldestBufferedEvent()
    If m_EventBuffer Is Nothing Then Exit Sub
    If m_EventBuffer.count = 0 Then Exit Sub

    Dim dropped As String
    dropped = CStr(m_EventBuffer(1))
    m_EventBuffer.Remove 1

    m_EventBufferBytes = m_EventBufferBytes - LenB(dropped)
    If m_EventBufferBytes < 0 Then m_EventBufferBytes = 0

    If m_DroppedEventsCount < 2147483647 Then _
        m_DroppedEventsCount = m_DroppedEventsCount + 1

    Dim droppedMethod As String
    RecordAsyncMessageLoss dropped, droppedMethod
End Sub

' Preserve FIFO order while bounding both retained event count and VBA string bytes.
' A single event larger than the entire byte budget is dropped without evicting the queue.
Private Sub EnqueueEventBounded(ByVal msg As String)
    Dim msgBytes As Long
    msgBytes = LenB(msg)

    If msgBytes > EVENT_BUFFER_MAX_BYTES Then
        If m_DroppedEventsCount < 2147483647 Then _
            m_DroppedEventsCount = m_DroppedEventsCount + 1

        Dim oversizedMethod As String
        RecordAsyncMessageLoss msg, oversizedMethod
        Exit Sub
    End If

    If m_EventBuffer Is Nothing Then ResetEventBuffer

    Do While m_EventBuffer.count >= EVENT_BUFFER_MAX_COUNT Or _
             m_EventBufferBytes > (EVENT_BUFFER_MAX_BYTES - msgBytes)
        If m_EventBuffer.count = 0 Then
            m_EventBufferBytes = 0
            Exit Do
        End If
        DropOldestBufferedEvent
    Loop

    m_EventBuffer.Add msg
    m_EventBufferBytes = m_EventBufferBytes + msgBytes
End Sub

Private Function TryDequeueBufferedEvent(ByRef msg As String) As Boolean
    If m_EventBuffer Is Nothing Then Exit Function
    If m_EventBuffer.count = 0 Then
        m_EventBufferBytes = 0
        Exit Function
    End If

    msg = CStr(m_EventBuffer(1))
    m_EventBuffer.Remove 1

    m_EventBufferBytes = m_EventBufferBytes - LenB(msg)
    If m_EventBufferBytes < 0 Then m_EventBufferBytes = 0

    TryDequeueBufferedEvent = True
End Function

Private Sub ResetApplicationHeartbeatState()
    m_applicationHeartbeatOutstanding = False
    m_applicationHeartbeatId = 0
    m_lastApplicationHeartbeatAt = 0
    m_consecutiveApplicationHeartbeatTimeouts = 0
End Sub

Public Function SendCommand(ByVal method As String, _
                            ByRef params As Object, _
                            Optional ByVal Silent As Boolean = False, _
                            Optional ByVal waitTimeMs As Long = 0) As String
    If m_isConnectUiYielding Then
        Err.Raise 606, "BiDiSocket", _
            "SendCommand blocked during connection-wait UI yield."
    End If

    If m_isCommandUiYielding Then
        Err.Raise 606, "BiDiSocket", _
            "SendCommand blocked during command-correlation UI yield."
    End If
    
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
    
    If COMMUNICATOR_DEBUG_ENABLED And Not Silent Then
        Dim displayMsg As String: displayMsg = sendMsg
        If Len(displayMsg) > 500 Then displayMsg = Left$(displayMsg, 497) & "..."
        LogEvent method & " (ID: " & requestId & ")", displayMsg, "[REQ]"
    End If
    
    Dim strRes As String
    strRes = SendAndReceive(sendMsg, waitTimeMs)
    
    If COMMUNICATOR_DEBUG_ENABLED And Not Silent Then
        Dim displayRes As String: displayRes = strRes
        If Len(displayRes) > 500 Then displayRes = Left$(displayRes, 497) & "..."
        LogEvent method & " (ID: " & requestId & ")", displayRes, "[RES]"
    End If
    
    SendCommand = strRes
End Function

Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    If m_isCommandUiYielding Then
        Trace ">>> [Lifecycle] Disconnect rejected during command-correlation UI yield."
        Exit Sub
    End If

    If m_isRaisingStateChanged Then
        Trace ">>> [Lifecycle] Disconnect rejected during StateChanged notification. " & _
              "Schedule teardown after the event handler returns.", True
        Exit Sub
    End If

    If m_isRecovering Then
        m_recoveryCancelRequested = True
        Trace ">>> [Recovery] Disconnect requested; cancelling the active recovery attempt.", True
    End If

    If m_state = StateDisconnecting Or m_state = StateDisconnected Then Exit Sub

    ' keepErrorState=True is transient transport cleanup before possible recovery; defer the
    ' session summary until an actual session close so counters survive reconnects.
    If Not keepErrorState Then TraceSessionDropSummary

    If Not keepErrorState Then SetState StateDisconnecting

    ' Transient error cleanup must not erase complete messages already assembled by the old
    ' transport. During command correlation, SendAndReceive owns response classification.
    If keepErrorState And Not m_isSendReceiving Then
        SalvageDeferredMessagesAfterTransportFailure "Disconnect(keepErrorState=True)"
    End If
    
    Call InternalCleanup
    
    ' A transient recovery disconnect preserves complete queued events. Partial WebSocket
    ' frames are owned and cleared by InternalCleanup/WsForceClose, not by this FIFO.
    If Not keepErrorState Then ResetEventBuffer
    
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

    ResetApplicationHeartbeatState

    ' Do not release m_isRecovering here. InternalCleanup is also called from Connect
    ' inside AttemptRecovery; only AttemptRecovery.CleanExit owns that guard.
    On Error GoTo 0
End Sub

' SendAndReceive correlates one ID around bounded non-blocking Pump calls. During a long
' correlation wait it uses short read slices and guarded Office yields; public transport/lifecycle
' entry points reject synchronous re-entry while that yield is active.
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    If m_isConnectUiYielding Then
        Err.Raise 606, "BiDiSocket", _
            "SendAndReceive blocked during connection-wait UI yield."
    End If

    If m_isRaisingStateChanged Then
        Err.Raise 606, "BiDiSocket", _
            "SendAndReceive blocked during synchronous StateChanged notification."
    End If

    If m_isRecovering Then
        Err.Raise 606, "BiDiSocket", _
            "SendAndReceive blocked while connection recovery is in progress."
    End If

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

    ' Validate before recovery or any transport work: an invalid public call must fail
    ' without changing connection state.
    cmdId = ExtractJsonId(jsonMessage)
    If Not IsNormalCommandId(cmdId) Then
        Err.Raise 607, "BiDiSocket", _
            "SendAndReceive requires a numeric root command ID in [" & _
            CMD_ID_FLOOR & ".." & ID_MAX & "]."
    End If

    m_activeCommandId = cmdId
    m_activeCommandOversizedResponseSeen = False
    m_activeCommandOversizedUnclassifiedSeen = False
    m_activeCommandOversizeDiagnostic = vbNullString

    If m_state <> StateConnected Then
        If Not AttemptRecovery() Then
            Err.Raise 600, "BiDiSocket", "Not connected and recovery failed."
        End If
    End If
    
    If Not SendMessage(jsonMessage) Then
        If m_state <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost during send operation."
        Else
            Err.Raise 601, "BiDiSocket", "Send failed due to an internal Winsock transport error."
        End If
    End If
    
    ' Correlation timeout is deliberately separate from the browser-side operation timeout.
    ' The wrapper normally supplies an already-capped wait. Short receive slices plus a guarded
    ' Office yield keep the UI responsive without allowing re-entrant transport/lifecycle work.
    '
    ' Reaching the correlation deadline is NOT proof of transport failure: the browser operation
    ' may still be running and its late response will be discarded by ID classification later.
    ' Therefore Err 604 preserves the live socket and is never escalated merely because another
    ' command also reaches its correlation deadline. WebSocket/ping/send state owns transport loss.
    Dim requestedWaitMs As Long
    If waitTimeMs > 0 Then
        requestedWaitMs = waitTimeMs
    Else
        requestedWaitMs = m_DefaultTimeout
    End If
    If requestedWaitMs <= 0 Then requestedWaitMs = DEFAULT_TIMEOUT_MS

    Dim correlationWaitMs As Long
    correlationWaitMs = requestedWaitMs
    If correlationWaitMs > MAX_COMMAND_CORRELATION_WAIT_MS Then _
        correlationWaitMs = MAX_COMMAND_CORRELATION_WAIT_MS

    m_LastCommandRequestedWaitMs = requestedWaitMs
    m_LastCommandEffectiveWaitMs = correlationWaitMs
    m_LastCommandCorrelationTimedOut = False

    Dim correlationDeadline As Double
    correlationDeadline = nowMs() + correlationWaitMs

    Dim hardDeadline As Double
    hardDeadline = correlationDeadline + COMMAND_CORRELATION_HARD_OVERRUN_MS
    Dim readDeadline As Double
    Dim uiSliceDeadline As Double
    
    ' The command-time event FIFO is bounded by both count and retained VBA string bytes.
    ' Oldest asynchronous events are dropped first while preserving response correlation.
    
    Do
        If m_state <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        readDeadline = correlationDeadline
        If m_AllowUiPump Then
            uiSliceDeadline = nowMs() + COMMAND_UI_PUMP_INTERVAL_MS
            If uiSliceDeadline < readDeadline Then readDeadline = uiSliceDeadline
        End If

        msg = ReadFromWire(readDeadline)

        If Len(msg) = 0 And m_AllowUiPump And nowMs() < correlationDeadline Then
            Dim uiYieldStartedAt As Double
            Dim uiYieldElapsedMs As Double
            uiYieldStartedAt = nowMs()
            YieldCommandCorrelationUi
            uiYieldElapsedMs = nowMs() - uiYieldStartedAt

            ' Correlation timeout remains a wall-clock contract, but time spent executing Office/VBA
            ' callbacks is not a receive-loop invariant failure. Exclude it from the secondary
            ' hard-overrun fail-safe so a slow UI handler cannot poison a healthy transport.
            If uiYieldElapsedMs > 0 Then hardDeadline = hardDeadline + uiYieldElapsedMs
        End If

        ' Receive-cap loss is not socket silence. A matching oversized response was already
        ' dispatched by the browser but could not be retained, so surface a distinct ambiguous
        ' outcome immediately. An unclassified oversized text/binary message during correlation
        ' also fails closed rather than waiting for an unrelated 604 timeout.
        If m_activeCommandOversizedResponseSeen Then
            Err.Raise ERR_INCOMING_MESSAGE_OVERSIZED, "BiDiSocket", _
                      "IncomingMessageOversized: matching command response for ID " & cmdId & _
                      " exceeded MaxIncomingSize and was discarded. " & m_activeCommandOversizeDiagnostic
        End If

        If m_activeCommandOversizedUnclassifiedSeen Then
            Err.Raise ERR_INCOMING_MESSAGE_OVERSIZED, "BiDiSocket", _
                      "IncomingMessageOversized: an oversized incoming message could not be " & _
                      "classified while waiting for ID " & cmdId & ". Response correlation is " & _
                      "not reliable. " & m_activeCommandOversizeDiagnostic
        End If
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            Select Case ClassifyIncomingMessage(inId, cmdId)
                Case IncomingClass_MatchingCommandResponse
                    SendAndReceive = msg
                    succeeded = True
                    GoTo Cleanup

                Case IncomingClass_ReservedSystemResponse
                    ProcessReservedSystemResponse inId, msg, False

                Case IncomingClass_LateOrUnknownResponse
                    DropLateOrUnknownResponse inId

                Case IncomingClass_Event
                    EnqueueEventBounded msg
            End Select
        End If
        
        Dim nowT As Double: nowT = nowMs()

        ' Receive/parser work is work-bounded except for one size-bounded, non-preemptible native
        ' UTF-8 codec operation. A substantial overrun still indicates that the correlation loop
        ' exceeded its fail-safe budget and remains a transport/state error.
        If nowT > hardDeadline Then
            SetState StateError
            Err.Raise 604, "BiDiSocket", _
                      "CommandCorrelationOverrun: receive loop exceeded the correlation deadline by more than " & _
                      CStr(COMMAND_CORRELATION_HARD_OVERRUN_MS) & "ms while waiting for ID " & cmdId & "."
        End If

        If nowT > correlationDeadline Then
            m_LastCommandCorrelationTimedOut = True
            Err.Raise 604, "BiDiSocket", _
                      "CommandCorrelationTimeout: no matching response within " & _
                      CStr(correlationWaitMs) & "ms while waiting for ID " & cmdId & _
                      ". RequestedWaitMs=" & CStr(requestedWaitMs) & _
                      "; EffectiveWaitMs=" & CStr(correlationWaitMs) & _
                      "; MaxCorrelationWaitMs=" & CStr(MAX_COMMAND_CORRELATION_WAIT_MS) & _
                      ". Socket preserved; the browser operation may still be running."
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
    
        ' The matching response may have been fully assembled immediately before the peer closed.
        ' Preserve any complete messages that followed it, then mark the transport unusable for
        ' the next operation. The returned command result itself remains valid.
        If Wsstate = WsTransportFailed Then
            SalvageDeferredMessagesAfterTransportFailure "command response completed before transport failure"
            SetState StateError
        End If
    End If
    
    m_isSendReceiving = False
    m_isCommandUiYielding = False
    m_activeCommandId = vbNullString
    m_activeCommandOversizedResponseSeen = False
    m_activeCommandOversizedUnclassifiedSeen = False
    m_activeCommandOversizeDiagnostic = vbNullString
    
    ' A normal command-correlation timeout preserves StateConnected. Only an actual StateError
    ' (transport failure or the secondary hard-overrun invariant) releases stale native handles.
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
    m_recoveryConnectAuthorized = False
    m_recoveryCancelRequested = False
    
    On Error GoTo ErrorHandler

    If Not m_AutoReconnect Then GoTo CleanExit
    If m_RetryCount >= m_MaxRetryAttempts Then
        Trace ">>> [Recovery] Max retry attempts reached.", True
        GoTo CleanExit
    End If

    m_RetryCount = m_RetryCount + 1
    
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    Trace ">>> [Recovery] Attempt #" & m_RetryCount & " (Waiting " & waitTime & "ms)", True
    
    ' Recovery backoff stays interactive. Public I/O entries reject re-entry while
    ' lifecycle teardown requests cancel the connection rebuild.
    YieldExecution waitTime, True

    If m_recoveryCancelRequested Then
        Trace ">>> [Recovery] Attempt cancelled by a lifecycle teardown request.", True
        GoTo CleanExit
    End If

    m_recoveryConnectAuthorized = True
    Dim recoveryConnected As Boolean
    recoveryConnected = Connect(m_LastWsUrl, m_LastPort)
    m_recoveryConnectAuthorized = False

    If recoveryConnected Then
        Trace ">>> [Recovery] Success.", True
        m_RetryCount = 0
        AttemptRecovery = True
    Else
        Trace ">>> [Recovery] Failed to reconnect.", True
        AttemptRecovery = False
    End If
    
CleanExit:
    m_recoveryConnectAuthorized = False
    m_recoveryCancelRequested = False
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
    ' Any assembled peer message proves transport progress and clears heartbeat suspicion.
    ' Do not scan every payload for an "error" token: response semantics belong to callers,
    ' and the hot receive path must not turn ordinary BiDi errors into synchronous file I/O.
    m_consecutiveApplicationHeartbeatTimeouts = 0
    Trace ">>> [Winsock] Message assembled. Length: " & Len(rawMsg)
End Sub

Private Sub ProcessReservedSystemResponse(ByVal inId As String, _
                                          ByVal rawMsg As String, _
                                          ByVal isPostHeartbeat As Boolean)
    If m_applicationHeartbeatOutstanding Then
        If inId = CStr(m_applicationHeartbeatId) Then
            m_applicationHeartbeatOutstanding = False
            m_applicationHeartbeatId = 0
        End If
    End If

    If inId <> CStr(POST_SCRIPT_REPLY_ID) Then
        ' A reserved reply, including a late heartbeat response, proves the peer is alive.
        m_consecutiveApplicationHeartbeatTimeouts = 0
        If m_RetryCount > 0 Then m_RetryCount = 0
    End If
End Sub

' Drain completed messages before an optional bounded Pump; a real message suppresses
' heartbeat generation for that GetMessage call.
Private Function TryReadWinsockBeforeHeartbeat(ByRef realMsg As String, _
                                               ByVal allowPump As Boolean) As Boolean
    If Not m_wsTransportInitialized Then Exit Function

    ' m_isReading is a same-thread guard around non-yielding Pump/dequeue work and is
    ' released only by structured cleanup or emergency teardown.
    If m_isReading Then Exit Function

    On Error GoTo ErrorHandler

    m_isReading = True

    Dim didPump As Boolean
    Dim candidate As String
    Dim inId As String

    ' If a prior transport step already failed, completed deferred messages still belong to
    ' the old stream and must be migrated before any reconnect can clear transport buffers.
    If Not WsIsOpen Then
        If Wsstate = WsTransportFailed Then
            SalvageDeferredMessagesAfterTransportFailure "event polling entered with failed transport"
            SetState StateError
            If TryDequeueBufferedEvent(realMsg) Then
                TryReadWinsockBeforeHeartbeat = True
            End If
        End If
        GoTo FinalCleanup
    End If

DrainCompleted:
    Do While WsTryDequeueText(candidate)
        inId = ExtractJsonId(candidate)

        Select Case ClassifyIncomingMessage(inId)
            Case IncomingClass_ReservedSystemResponse
                ProcessReservedSystemResponse inId, candidate, False

            Case IncomingClass_LateOrUnknownResponse
                DropLateOrUnknownResponse inId

            Case IncomingClass_Event
                ObserveWinsockMessage candidate
                realMsg = candidate
                TryReadWinsockBeforeHeartbeat = True
                GoTo FinalCleanup
        End Select
    Loop

    If Not allowPump Or didPump Then GoTo FinalCleanup

    didPump = True
    WsPump WINSOCK_PUMP_MAX_MS, _
           WINSOCK_PUMP_MAX_BYTES, _
           WINSOCK_PUMP_MAX_CALLS

    If Wsstate = WsTransportFailed Then
        ' IMPORTANT: a single bounded Pump may enqueue a complete text message and then observe
        ' EOF/reset on the next recv. Drain every completed message BEFORE StateError/recovery.
        SalvageDeferredMessagesAfterTransportFailure "event polling pump observed transport failure"
        Trace ">>> [Socket] Winsock transport failed during event polling. " & _
              GetWinsockFailureText(), True
        SetState StateError

        If TryDequeueBufferedEvent(realMsg) Then
            TryReadWinsockBeforeHeartbeat = True
        End If
        GoTo FinalCleanup
    End If

    If Not WsIsOpen Then
        SalvageDeferredMessagesAfterTransportFailure "event polling pump closed transport"
        Trace ">>> [Socket] Winsock transport is no longer open during event polling. " & _
              GetWinsockFailureText(), True
        SetState StateError

        If TryDequeueBufferedEvent(realMsg) Then
            TryReadWinsockBeforeHeartbeat = True
        End If
        GoTo FinalCleanup
    End If

    GoTo DrainCompleted

FinalCleanup:
    m_isReading = False
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & _
                "] [Winsock] Pre-heartbeat read error: " & Err.description
    SalvageDeferredMessagesAfterTransportFailure "event polling error handler"
    SetState StateError
    Resume FinalCleanup
End Function

' Non-yielding event read for wrapper-side lifecycle reconciliation. It never sends a
' heartbeat, attempts recovery, sleeps, or pumps Office messages.
Private Function GetMessageNonBlocking() As String
    If m_isRaisingStateChanged Or m_isRecovering Then Exit Function
    If m_isSendReceiving Then Exit Function

    Dim bufferedEvent As String
    If TryDequeueBufferedEvent(bufferedEvent) Then
        GetMessageNonBlocking = bufferedEvent
        Exit Function
    End If

    If m_state <> StateConnected Then Exit Function

    Dim transportEvent As String
    If TryReadWinsockBeforeHeartbeat(transportEvent, True) Then
        GetMessageNonBlocking = transportEvent
    End If
End Function

' GetMessage serves the event FIFO, performs bounded receive work, then manages the
' reserved-band session.status heartbeat.
Public Function GetMessage() As String
    If m_isConnectUiYielding Then Exit Function
    If m_isRaisingStateChanged Or m_isRecovering Then Exit Function
    If m_isSendReceiving Then Exit Function

    Dim bufferedEvent As String
    If TryDequeueBufferedEvent(bufferedEvent) Then
        GetMessage = bufferedEvent
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

    ' Poll incoming traffic at PollingInterval, but pace the application heartbeat
    ' independently so receive latency is not coupled to session.status frequency.
    nowT = nowMs()
    If m_applicationHeartbeatOutstanding Then
        If (nowT - m_lastApplicationHeartbeatAt) < APPLICATION_HEARTBEAT_TIMEOUT_MS Then Exit Function

        If m_consecutiveApplicationHeartbeatTimeouts < 2147483647 Then
            m_consecutiveApplicationHeartbeatTimeouts = m_consecutiveApplicationHeartbeatTimeouts + 1
        End If

        Trace ">>> [Heartbeat] session.status reply timeout. Consecutive=" & _
              m_consecutiveApplicationHeartbeatTimeouts & _
              "/" & MAX_CONSECUTIVE_APPLICATION_HEARTBEAT_TIMEOUTS & _
              "; retiring ID=" & m_applicationHeartbeatId & ".", True

        m_applicationHeartbeatOutstanding = False
        m_applicationHeartbeatId = 0

        If m_consecutiveApplicationHeartbeatTimeouts >= _
           MAX_CONSECUTIVE_APPLICATION_HEARTBEAT_TIMEOUTS Then
            Trace ">>> [Heartbeat] Application liveness failure. " & _
                  "Consecutive session.status timeouts reached the limit.", True
            SetState StateError
            Exit Function
        End If
    End If

    If (nowT - m_lastApplicationHeartbeatAt) < APPLICATION_HEARTBEAT_INTERVAL_MS Then Exit Function

    Dim pingId As Long
    pingId = GetNextSystemPingId()

    Dim pingCmd As String
    pingCmd = "{""id"":" & pingId & ", ""method"":""session.status"", ""params"":{}}"
    If Not SendMessage(pingCmd) Then Exit Function

    m_applicationHeartbeatOutstanding = True
    m_applicationHeartbeatId = pingId
    m_lastApplicationHeartbeatAt = nowMs()

    Dim rawMsg As String
    rawMsg = ReadFromWire(nowMs() + 10)

    If LenB(rawMsg) = 0 Then
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If

    Dim inId As String
    inId = ExtractJsonId(rawMsg)

    Select Case ClassifyIncomingMessage(inId)
        Case IncomingClass_ReservedSystemResponse
            ProcessReservedSystemResponse inId, rawMsg, True

        Case IncomingClass_LateOrUnknownResponse
            DropLateOrUnknownResponse inId

        Case IncomingClass_Event
            GetMessage = rawMsg
    End Select
End Function

' Three-state parsed-event reader:
'   False                     = no message was available.
'   True + parsed Is Object   = one valid message was consumed and parsed.
'   True + parsed Is Nothing  = one invalid/filtered message was consumed and dropped.
' This distinction lets queue drains skip a bad event and continue with following events.
Public Function TryGetParsedEvent(ByRef parsed As Object, _
                                  ByRef dropReason As String, _
                                  Optional ByVal minMsgSize As Long = 10, _
                                  Optional ByVal maxMsgSize As Long = PARSED_ASYNC_EVENT_MAX_CHARS, _
                                  Optional ByVal nonBlocking As Boolean = False) As Boolean
    Set parsed = Nothing
    dropReason = vbNullString

    If maxMsgSize > PARSED_ASYNC_EVENT_MAX_CHARS Then maxMsgSize = PARSED_ASYNC_EVENT_MAX_CHARS

    Dim rawMsg As String
    If nonBlocking Then
        rawMsg = GetMessageNonBlocking()
    Else
        rawMsg = GetMessage()
    End If
    If Len(rawMsg) = 0 Then Exit Function

    ' One event message was consumed even when it is rejected below.
    TryGetParsedEvent = True

    Dim rawLength As Long
    rawLength = Len(rawMsg)

    If rawLength < minMsgSize Then
        dropReason = "too-short event (chars=" & rawLength & _
                     "; minimum=" & minMsgSize & ")"
        If m_DroppedParsedFilteredCount < 2147483647 Then _
            m_DroppedParsedFilteredCount = m_DroppedParsedFilteredCount + 1
        m_LastParsedEventDropReason = dropReason
        Exit Function
    End If

    If rawLength > maxMsgSize Then
        Dim oversizeMethod As String
        RecordAsyncMessageLoss rawMsg, oversizeMethod

        dropReason = "oversized parsed event (chars=" & rawLength & _
                     "; maximum=" & maxMsgSize & "; method="

        If Len(oversizeMethod) > 0 Then
            dropReason = dropReason & oversizeMethod & ")"
        Else
            dropReason = dropReason & "unknown)"
        End If

        If m_DroppedParsedOversizeCount < 2147483647 Then _
            m_DroppedParsedOversizeCount = m_DroppedParsedOversizeCount + 1
        m_LastParsedEventDropReason = dropReason
        Exit Function
    End If

    If InStr(rawMsg, """method"":") = 0 And _
       InStr(rawMsg, """id"":") = 0 Then
        dropReason = "message has neither a root method nor id"
        If m_DroppedParsedFilteredCount < 2147483647 Then _
            m_DroppedParsedFilteredCount = m_DroppedParsedFilteredCount + 1
        m_LastParsedEventDropReason = dropReason
        Exit Function
    End If

    Dim parseErr As Long
    Dim candidate As Object
    On Error Resume Next
    Set candidate = WebJsonConverter.ParseJson(rawMsg)
    parseErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If parseErr <> 0 Or candidate Is Nothing Then
        Dim malformedMethod As String
        RecordAsyncMessageLoss rawMsg, malformedMethod

        dropReason = "malformed JSON event"
        If parseErr <> 0 Then dropReason = dropReason & " (Err " & parseErr & ")"

        If Len(malformedMethod) > 0 Then
            dropReason = dropReason & "; method=" & malformedMethod
        Else
            dropReason = dropReason & "; method=unknown"
        End If

        If m_DroppedParsedMalformedCount < 2147483647 Then _
            m_DroppedParsedMalformedCount = m_DroppedParsedMalformedCount + 1
        m_LastParsedEventDropReason = dropReason
        Exit Function
    End If

    m_LastParsedEventDropReason = vbNullString
    Set parsed = candidate
End Function

' Compatibility API: callers that only understand Object/Nothing retain the old contract.
Public Function GetParsedEvent(Optional ByVal minMsgSize As Long = 10, _
                               Optional ByVal maxMsgSize As Long = PARSED_ASYNC_EVENT_MAX_CHARS) As Object
    Dim parsed As Object
    Dim dropReason As String

    If TryGetParsedEvent(parsed, dropReason, minMsgSize, maxMsgSize) Then
        If Not parsed Is Nothing Then Set GetParsedEvent = parsed
    End If
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

        ' A completed oversized command response/unclassified message is a known ambiguous
        ' correlation outcome. Return immediately so SendAndReceive can raise 608 now rather
        ' than waiting for the correlation deadline.
        If m_isSendReceiving And _
           (m_activeCommandOversizedResponseSeen Or m_activeCommandOversizedUnclassifiedSeen) Then
            Exit Do
        End If

        If Wsstate = WsTransportFailed Then
            Trace ">>> [Socket] Winsock transport failed while awaiting a response. " & _
                  GetWinsockFailureText(), True
            SetState StateError
            Exit Do
        End If

        If Not WsIsOpen Then
            Trace ">>> [Socket] Winsock transport is no longer open while awaiting a response. " & _
                  GetWinsockFailureText(), True
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

        ' SendAndReceive must classify its own matching ID before StateError is raised. Other
        ' readers have no command owner, so they can safely migrate trailing complete messages now.
        If Wsstate = WsTransportFailed And Not m_isSendReceiving Then
            SalvageDeferredMessagesAfterTransportFailure "non-command read completed before transport failure"
            SetState StateError
        End If
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

' PumpMessages may run DoEvents only outside SendAndReceive and recovery; throttleMs
' limits how often the Office queue is pumped.
Public Sub PumpMessages(Optional ByVal throttleMs As Long = 0)
    If m_isRaisingStateChanged Or m_isRecovering Then Exit Sub
    If m_isSendReceiving Then Exit Sub
    If throttleMs > 0 And m_lastPumpTicks <> 0 Then
        If elapsedMs(m_lastPumpTicks) < throttleMs Then Exit Sub
    End If
    DoEvents
    m_lastPumpTicks = NowTicks()
End Sub

' YieldExecution pumps Office messages only when explicitly requested. Command correlation uses
' YieldCommandCorrelationUi instead so the re-entry guard is active around DoEvents.
' Guarded connect waits and recovery backoff are the other intentional Office-pump paths.
Public Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    Dim vMs As Long: vMs = ms
    If vMs < 0 Then vMs = 1

    If vMs > 50 Then
        Call InterruptibleSleep(vMs, allowDoEvents)
    Else
        If allowDoEvents And (Not m_isSendReceiving Or m_isRecovering) Then
            DoEvents
        End If

        If vMs > 0 Then SleepWinAPI vMs
    End If
End Sub

Private Sub YieldConnectWaitUi()
    If Not m_AllowUiPump Then Exit Sub
    If m_isConnectUiYielding Then Exit Sub

    m_isConnectUiYielding = True
    On Error Resume Next
    DoEvents
    Err.Clear
    On Error GoTo 0
    m_isConnectUiYielding = False
End Sub

Private Sub YieldCommandCorrelationUi()
    If Not m_AllowUiPump Then Exit Sub
    If m_isCommandUiYielding Then Exit Sub

    m_isCommandUiYielding = True
    On Error Resume Next
    DoEvents
    Err.Clear
    On Error GoTo 0
    m_isCommandUiYielding = False
End Sub

' COMMUNICATOR_DEBUG_ENABLED controls outer-communicator Immediate output and
' keyword-based file detection. In production, file logging is driven only by
' the caller's explicit isCritical classification.
Public Sub Trace(ByVal msg As String, Optional ByVal isCritical As Boolean = False)
    If Not COMMUNICATOR_DEBUG_ENABLED And Not isCritical Then Exit Sub
    
    Dim logMsg As String: logMsg = "[" & Format(Now, "hh:mm:ss") & "] " & msg
    
    If COMMUNICATOR_DEBUG_ENABLED Then Debug.Print logMsg
    
    Dim shouldLogToFile As Boolean
    shouldLogToFile = isCritical
    If COMMUNICATOR_DEBUG_ENABLED Then
        shouldLogToFile = shouldLogToFile Or _
                          (InStr(1, msg, "Return Code", vbTextCompare) > 0) Or _
                          (InStr(1, msg, "failed", vbTextCompare) > 0) Or _
                          (InStr(1, msg, "error", vbTextCompare) > 0)
    End If

    If shouldLogToFile Then
        Dim localDir As String: localDir = ResolveLogDirectory()
        
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
    If Not COMMUNICATOR_DEBUG_ENABLED And Not LogToFile Then Exit Sub

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

    If COMMUNICATOR_DEBUG_ENABLED Then
        Debug.Print headerLine

        If Len(cleanText) > 0 Then
            For i = 1 To Len(cleanText) Step CHUNK_SIZE
                Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
            Next i
            Debug.Print ""
        End If
    End If
    
    If LogToFile Then
        WriteLogToFile headerLine & vbCrLf & cleanText
    End If
End Sub

Private Function ResolveLogDirectory() As String
    Dim localDir As String

    ' WebShared.thisLibFolderPath is host-neutral:
    ' Excel -> host workbook folder, Access -> CodeProject.Path.
    On Error Resume Next
    localDir = WebShared.thisLibFolderPath
    If Len(localDir) > 0 Then
        localDir = WebShared.getFullLocalPath(localDir)
    End If
    On Error GoTo 0

    ' An unsaved host document or an unavailable project path must not prevent logging.
    If Len(localDir) = 0 Then localDir = Environ$("TEMP")
    If Len(localDir) = 0 Then localDir = CurDir$
    If Len(localDir) = 0 Then localDir = Environ$("USERPROFILE")
    If Len(localDir) = 0 Then localDir = "."
    If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"

    ResolveLogDirectory = localDir
End Function

Private Sub ResetSocketTraceLog()
    Dim localDir As String
    Dim logPath As String
    Dim fNum As Integer

    On Error Resume Next
    localDir = ResolveLogDirectory()
    logPath = localDir & "bidi_socket_trace.log"
    fNum = FreeFile

    Err.Clear
    Open logPath For Output As #fNum
    If Err.Number = 0 Then Close #fNum
    On Error GoTo 0
End Sub

Private Sub WriteLogToFile(ByVal msg As String)
    Dim localDir As String
    On Error Resume Next
    localDir = ResolveLogDirectory()
    
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

' InterruptibleSleep normally suppresses Office messages during SendAndReceive. Guarded
' recovery backoff may pump them because public I/O entries reject re-entry and
' lifecycle teardown requests cancel the attempt.
Private Sub InterruptibleSleep(ByVal totalMs As Long, Optional ByVal allowDoEvents As Boolean = True)
    Dim startTs As Double: startTs = nowMs()
    Const CHUNK_MS As Long = 50 ' Small granularity for better responsiveness
    
    Do While (nowMs() - startTs) < totalMs
        If m_isRecovering And m_recoveryCancelRequested Then Exit Do

        If allowDoEvents And (Not m_isSendReceiving Or m_isRecovering) Then
            DoEvents
        End If
        
        SleepWinAPI CHUNK_MS
    Loop
End Sub

' ForceClose skips the close handshake and performs non-raising emergency teardown.
Public Sub ForceClose()
    If m_isCommandUiYielding Then
        Trace ">>> [Lifecycle] ForceClose rejected during command-correlation UI yield."
        Exit Sub
    End If

    If m_isRaisingStateChanged Then
        Trace ">>> [Lifecycle] ForceClose rejected during StateChanged notification. " & _
              "Schedule emergency teardown after the event handler returns.", True
        Exit Sub
    End If

    If m_isRecovering Then
        m_recoveryCancelRequested = True
        Trace ">>> [Recovery] ForceClose requested; cancelling the active recovery attempt.", True
    End If

    ' Clear I/O guards before closing native transport so interrupted operations do not
    ' poison object reuse.
    On Error Resume Next
    If m_state <> StateDisconnected Then TraceSessionDropSummary
    m_isSendReceiving = False
    m_isReading = False
    m_isCommandUiYielding = False

    If m_wsTransportInitialized Then
        WsForceClose
    End If

    ResetEventBuffer
    ResetApplicationHeartbeatState
    SetState StateDisconnected
    On Error GoTo 0
End Sub

Private Sub WsTransportInitialize()
    Dim frequency As Currency
    If WsApi_QueryPerformanceFrequency(frequency) = 0 Or frequency = 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", "QPC is unavailable."
    End If
    m_ws_qpcFrequency = frequency

    Dim wsaDataBuffer(0 To 511) As Byte
    Dim result As Long
    result = WsApi_WSAStartup(&H202, wsaDataBuffer(0))
    If result <> 0 Then
        Err.Raise vbObjectError + 9101, "BiDiSocketCommunicator", "WSAStartup failed: " & result
    End If
    m_ws_wsaStarted = True

    Set m_ws_sendQueue = New Collection
    WsResetDeferredTextQueue
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
    If value < MAX_MSG_SIZE Then value = MAX_MSG_SIZE
    If value > MAX_CONFIGURED_INCOMING_SIZE Then value = MAX_CONFIGURED_INCOMING_SIZE
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
        If Len(m_ws_urlValidationError) > 0 Then
            WsSetFailure 9201, m_ws_urlValidationError
        Else
            WsSetFailure 9201, "Only ws:// loopback URLs are supported."
        End If
        Exit Function
    End If

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

        If WsApi_InetPtonW_IPv6(WS_AF_INET6, StrPtr("::1"), address6.address(0)) <> 1 Then
            WsSetFailure WsApi_WSAGetLastError(), "IPv6 loopback conversion failed."
            Exit Function
        End If

        connectResult = WsApi_connectSocketIPv6(m_ws_socket, address6, LenB(address6))
    Else
        Dim address4 As WsSOCKADDR_IN
        address4.family = WS_AF_INET
        address4.port = WsHostToNetworkShort(m_ws_port)

        If WsApi_InetPtonW_IPv4(WS_AF_INET, StrPtr("127.0.0.1"), address4.address) <> 1 Then
            WsSetFailure WsApi_WSAGetLastError(), "IPv4 loopback conversion failed."
            Exit Function
        End If

        connectResult = WsApi_connectSocketIPv4(m_ws_socket, address4, LenB(address4))
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

    ' Reserve the final parser millisecond for bounded VBA copying on fragmented text only.
    ' Non-fragmented text uses raw-buffer direct decode after the complete payload is present.
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

            ' After our Close frame is sent, accept an already-available peer Close/EOF within
            ' the same bounded Pump. We still force-close if no peer response is immediately ready;
            ' higher-level cleanup remains deterministic and does not gain a new wait loop.
            If m_ws_state = WsTransportClosing And WsMonotonicMs() < pumpDeadline Then
                WsPumpReceiveInternal pumpDeadline, copyGuardDeadline, maxBytes, maxCalls, True
            End If

            If m_ws_state = WsTransportClosing And Not WsHasPendingSend Then WsForceClose
    End Select

    ' UTF-8 rejection is finalized only after parser/receive loops have unwound. The finalizer
    ' performs a send-only bounded close flush and never re-enters WsPump.
    If m_ws_pendingUtf8Rejection Then
        WsFinalizePendingUtf8Rejection pumpDeadline, maxBytes, maxCalls
    End If

    WsPump = (m_ws_state <> WsTransportFailed)
End Function
Private Function WsQueueText(ByVal message As String) As Boolean
    If m_ws_state <> WsTransportOpen Then Exit Function
    On Error GoTo ErrorHandler

    Dim payload() As Byte
    payload = WsEncodeUtf8(message)

    Dim frame() As Byte
    frame = WsBuildClientFrame(WS_OPCODE_TEXT, payload, True)

    If m_ws_state = WsTransportFailed Then Exit Function
    If WsByteArrayLength(frame) = 0 Then Exit Function

    WsEnqueueSendBytes frame
    WsQueueText = True
    Exit Function

ErrorHandler:
    Dim savedNumber As Long
    Dim savedDescription As String
    savedNumber = Err.Number
    savedDescription = Err.description

    If savedNumber = WS_ERR_INVALID_UTF16 Then Resume RejectCallerText
    Resume FailTransport

RejectCallerText:
    ' Encoding failed before a frame was built or queued. The peer saw nothing, so the
    ' socket remains healthy and the deterministic caller-input error is propagated.
    On Error GoTo 0
    Err.Raise savedNumber, "BiDiSocketCommunicator.WsQueueText", savedDescription

FailTransport:
    On Error GoTo 0
    WsSetFailure savedNumber, _
                 "Outgoing WebSocket frame construction failed: " & savedDescription
End Function

Private Sub WsResetDeferredTextQueue()
    Set m_ws_deferredTextQueue = New Collection
    m_ws_deferredTextQueueBytes = 0
End Sub
' This queue contains both events and command responses, so overflow is fatal.
' Silent eviction could break BiDi response-ID correlation.
Private Function WsEnqueueDeferredText(ByVal message As String) As Boolean
    On Error GoTo ErrorHandler

    If m_ws_deferredTextQueue Is Nothing Then
        WsSetFailure 9222, "Deferred text queue is unavailable."
        Exit Function
    End If

    Dim storageBytes As Long
    storageBytes = LenB(message)

    ' MaxIncomingSize is a UTF-8 byte cap. ASCII can expand to 2x storage after decode to a
    ' VBA UTF-16 String, so single-message admission is separate from the normal backlog target.
    Dim singleMessageStorageLimit As Double
    singleMessageStorageLimit = CDbl(m_ws_maxIncomingSize) * 2#
    If CDbl(storageBytes) > singleMessageStorageLimit Then
        WsSetFailure 9222, _
            "Deferred text message exceeds the decoded single-message storage limit. IncomingLenB=" & _
            CStr(storageBytes) & "; MaxLenB=" & Format$(singleMessageStorageLimit, "0") & "."
        Exit Function
    End If

    If m_ws_deferredTextQueue.count >= DEFERRED_TEXT_QUEUE_MAX_COUNT Then
        WsSetFailure 9222, _
            "Deferred text queue count limit exceeded. Count=" & _
            CStr(m_ws_deferredTextQueue.count) & "; MaxCount=" & _
            CStr(DEFERRED_TEXT_QUEUE_MAX_COUNT) & "."
        Exit Function
    End If

    ' Keep 32 MiB as the normal decoded backlog target. A single configured-max message may
    ' legitimately exceed that target after UTF-16 expansion; allow it only when the queue is empty.
    If storageBytes > DEFERRED_TEXT_QUEUE_TARGET_BYTES Then
        If m_ws_deferredTextQueue.count <> 0 Or m_ws_deferredTextQueueBytes <> 0 Then
            WsSetFailure 9222, _
                "Deferred text queue retention target exceeded by a large single message while backlog exists. " & _
                "RetainedLenB=" & CStr(m_ws_deferredTextQueueBytes) & _
                "; IncomingLenB=" & CStr(storageBytes) & _
                "; TargetLenB=" & CStr(DEFERRED_TEXT_QUEUE_TARGET_BYTES) & "."
            Exit Function
        End If
    ElseIf m_ws_deferredTextQueueBytes > (DEFERRED_TEXT_QUEUE_TARGET_BYTES - storageBytes) Then
        WsSetFailure 9222, _
            "Deferred text queue retention limit exceeded. Count=" & _
            CStr(m_ws_deferredTextQueue.count) & "; RetainedLenB=" & _
            CStr(m_ws_deferredTextQueueBytes) & "; IncomingLenB=" & CStr(storageBytes) & _
            "; TargetLenB=" & CStr(DEFERRED_TEXT_QUEUE_TARGET_BYTES) & "."
        Exit Function
    End If

    m_ws_deferredTextQueue.Add item:=message
    m_ws_deferredTextQueueBytes = m_ws_deferredTextQueueBytes + storageBytes

    WsEnqueueDeferredText = True
    Exit Function

ErrorHandler:
    WsSetFailure 9222, "Deferred text queue enqueue failed: " & Err.description
End Function
Private Function WsTryDequeueText(ByRef message As String) As Boolean
    On Error GoTo ErrorHandler

    If m_ws_deferredTextQueue Is Nothing Then
        WsSetFailure 9218, "Deferred text queue is unavailable."
        Exit Function
    End If

    If m_ws_deferredTextQueue.count = 0 Then
        If m_ws_deferredTextQueueBytes <> 0 Then
            WsSetFailure 9218, "Deferred text queue byte accounting is inconsistent."
        End If
        Exit Function
    End If

    message = CStr(m_ws_deferredTextQueue(1))

    Dim storageBytes As Long
    storageBytes = LenB(message)
    If storageBytes < 0 Or storageBytes > m_ws_deferredTextQueueBytes Then
        WsSetFailure 9218, "Deferred text queue byte accounting is invalid."
        Exit Function
    End If

    m_ws_deferredTextQueue.Remove 1
    m_ws_deferredTextQueueBytes = m_ws_deferredTextQueueBytes - storageBytes

    ' A complete decoded message is immutable application data. A later TCP/WebSocket failure
    ' does not invalidate text that was already fully assembled, validated and queued.
    WsTryDequeueText = True
    Exit Function

ErrorHandler:
    WsSetFailure 9218, "Deferred text queue dequeue failed: " & Err.description
End Function
Private Sub WsPendInvalidPeerUtf8(ByVal description As String)
    ' First malformed text wins. Do not change transport state here: parser/receive loops use
    ' the pending flag to unwind without re-entering WsPump or consuming subsequent frames.
    If m_ws_pendingUtf8Rejection Then Exit Sub
    m_ws_pendingUtf8Rejection = True
    m_ws_pendingUtf8RejectionText = description
End Sub

Private Sub WsFinalizePendingUtf8Rejection(ByVal pumpDeadline As Double, _
                                           ByVal maxBytes As Long, _
                                           ByVal maxCalls As Long)
    If Not m_ws_pendingUtf8Rejection Then Exit Sub

    Dim rejectionText As String
    rejectionText = m_ws_pendingUtf8RejectionText
    If Len(rejectionText) = 0 Then rejectionText = "Peer text message contains malformed UTF-8."

    ' Clear the pending marker before teardown so normal reset paths cannot recursively finalize it.
    m_ws_pendingUtf8Rejection = False
    m_ws_pendingUtf8RejectionText = vbNullString

    On Error Resume Next

    If m_ws_state = WsTransportOpen Then
        If WsBeginClose(WS_CLOSE_INVALID_PAYLOAD_DATA) Then
            m_ws_lastProtocolCloseStatus = WS_CLOSE_INVALID_PAYLOAD_DATA

            ' Send-only bounded flush: never call WsPump from inside the outer Pump.
            If WsHasPendingSend And WsMonotonicMs() < pumpDeadline Then
                WsPumpSendInternal pumpDeadline, maxBytes, maxCalls
            End If
        End If
    End If

    On Error GoTo 0

    ' Ordering is intentional: ForceClose releases native/parser state while preserving complete
    ' decoded messages; WsSetFailure then restores FAILED after ForceClose's DISCONNECTED state.
    WsForceClose preserveDeferredText:=True
    WsSetFailure WS_ERR_INVALID_UTF8, rejectionText
End Sub
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

Private Sub WsForceClose(Optional ByVal preserveDeferredText As Boolean = False)
    On Error Resume Next

    If m_ws_socket <> -1 Then
        WsApi_shutdownSocket m_ws_socket, WS_SD_BOTH
        WsApi_closesocket m_ws_socket
        m_ws_socket = -1
    End If

    Set m_ws_sendQueue = New Collection
    If Not preserveDeferredText Then WsResetDeferredTextQueue

    Erase m_ws_sendCurrent
    Erase m_ws_rawBuffer
    Erase m_ws_controlPayload
    Erase m_ws_messageBuffer

    m_ws_sendCurrentCount = 0
    m_ws_sendOffset = 0
    m_ws_rawCount = 0
    m_ws_rawReadPos = 0
    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_framePayloadLength = 0
    m_ws_framePayloadRemaining = 0
    m_ws_frameDiscardPayload = False
    m_ws_controlCount = 0
    m_ws_pendingUtf8Rejection = False
    m_ws_pendingUtf8RejectionText = vbNullString
    WsResetMessageAccumulator
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

    If WINSOCK_TRANSPORT_DEBUG_ENABLED Then Debug.Print "BiDiSocketCommunicator[IntegratedWinsock]: OPEN"

    WsParseAvailableFrames pumpDeadline, copyGuardDeadline
End Sub

Private Function WsIsHttp101StatusLine(ByVal firstLine As String) As Boolean
    ' Parse the status-code token rather than searching for " 101 " anywhere in the line.
    ' This rejects malformed lines such as "X 101 Y" while tolerating repeated SP before 101.
    Dim lineText As String
    lineText = Trim$(firstLine)
    If Len(lineText) < 9 Then Exit Function
    If StrComp(Left$(lineText, 5), "HTTP/", vbTextCompare) <> 0 Then Exit Function

    Dim firstSpace As Long
    firstSpace = InStr(1, lineText, " ", vbBinaryCompare)
    If firstSpace = 0 Then Exit Function

    Dim remainder As String
    remainder = LTrim$(Mid$(lineText, firstSpace + 1))
    If Len(remainder) = 0 Then Exit Function

    Dim secondSpace As Long
    secondSpace = InStr(1, remainder, " ", vbBinaryCompare)

    Dim statusCode As String
    If secondSpace = 0 Then
        statusCode = remainder
    Else
        statusCode = Left$(remainder, secondSpace - 1)
    End If

    WsIsHttp101StatusLine = (StrComp(statusCode, "101", vbBinaryCompare) = 0)
End Function

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
    statusValid = WsIsHttp101StatusLine(firstLine)

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
                ' Multiple field lines are semantically cumulative here. Once the required
                ' token is observed, a later duplicate header must not erase that evidence.
                If WsHeaderContainsToken(headerValue, "websocket") Then hasUpgrade = True
            ElseIf headerName = "connection" Then
                If WsHeaderContainsToken(headerValue, "upgrade") Then hasConnectionUpgrade = True
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
        m_ws_handshakeValidationError = _
            "HTTP status is not 101. Peer status line: " & _
            Left$(firstLine, WS_HANDSHAKE_STATUS_ECHO_MAX)
    ElseIf Not hasUpgrade Then
        m_ws_handshakeValidationError = _
            "Upgrade header has no websocket token"
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
                   Right$("00000000" & Hex$(status), 8)
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
            Right$("00000000" & Hex$(status), 8)
        GoTo Cleanup
    End If

    Dim objectProperty As String
    objectProperty = WS_BCRYPT_OBJECT_LENGTH

    Dim objectLength As Long
    Dim resultSize As Long

    status = WsApi_BCryptGetPropertyLong( _
        algorithmHandle, StrPtr(objectProperty), _
        objectLength, 4, resultSize, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Or objectLength <= 0 Then
        failureText = _
            "ObjectLength query NTSTATUS=0x" & _
            Right$("00000000" & Hex$(status), 8)
        GoTo Cleanup
    End If

    Dim hashProperty As String
    hashProperty = WS_BCRYPT_HASH_LENGTH

    Dim hashLength As Long
    resultSize = 0

    status = WsApi_BCryptGetPropertyLong( _
        algorithmHandle, StrPtr(hashProperty), _
        hashLength, 4, resultSize, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Or hashLength <= 0 Then
        failureText = _
            "HashDigestLength query NTSTATUS=0x" & _
            Right$("00000000" & Hex$(status), 8)
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
            Right$("00000000" & Hex$(status), 8)
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
            Right$("00000000" & Hex$(status), 8)
        GoTo Cleanup
    End If

    status = WsApi_BCryptFinishHash( _
        hashHandle, digest(0), hashLength, 0)

    If status <> WS_BCRYPT_STATUS_SUCCESS Then
        failureText = _
            "BCryptFinishHash NTSTATUS=0x" & _
            Right$("00000000" & Hex$(status), 8)
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
        If m_ws_pendingUtf8Rejection Then Exit Sub
        If m_ws_state = WsTransportFailed Then Exit Sub
        If m_ws_currentPumpCompletedMessages >= _
           WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Sub
        If WsMonotonicMs() >= pumpDeadline Then Exit Sub
    End If

    Do While callCount < maxCalls And receivedBytes < maxBytes
        If m_ws_pendingUtf8Rejection Then Exit Do
        If m_ws_currentPumpCompletedMessages >= _
           WS_MAX_COMPLETED_MESSAGES_PER_PUMP Then Exit Do
        If WsMonotonicMs() >= pumpDeadline Then Exit Do

        ' Direct-decode payloads were reserved with their current raw offset included. Avoid
        ' byte-wise compaction while waiting, otherwise the zero-copy path would reintroduce a
        ' potentially large VBA copy of unread payload bytes.
        If Not (parseFrames And m_ws_directDecodeCandidate) Then WsCompactRawIfNeeded

        Dim allowed As Long
        allowed = WS_RECEIVE_CHUNK_SIZE
        If allowed > (maxBytes - receivedBytes) Then allowed = maxBytes - receivedBytes

        ' A direct-decode frame reserves exactly the current payload span. Do not read beyond the
        ' remaining payload and accidentally force a second buffer growth for following frames.
        If parseFrames And m_ws_directDecodeCandidate Then
            Dim payloadBytesNeeded As Long
            payloadBytesNeeded = m_ws_framePayloadLength - WsRawAvailable
            If payloadBytesNeeded <= 0 Then
                WsParseAvailableFrames pumpDeadline, copyGuardDeadline
                If m_ws_pendingUtf8Rejection Then Exit Do
                If m_ws_state = WsTransportFailed Then Exit Do
                GoTo ContinueReceiveLoop
            End If
            If allowed > payloadBytesNeeded Then allowed = payloadBytesNeeded

            ' For an oversized direct-decode candidate, retain only enough raw payload to
            ' classify the bounded root prefix. Once identity is resolved, the parser either
            ' discards a known async event or reserves the complete response/unclassified span.
            If m_ws_framePayloadLength > PARSED_ASYNC_EVENT_MAX_UTF8_BYTES And _
               Not m_ws_messageIdentityResolved Then
                Dim identityBytesNeeded As Long
                identityBytesNeeded = OVERSIZED_DIAGNOSTIC_PREFIX_BYTES - WsRawAvailable
                If identityBytesNeeded > 0 And allowed > identityBytesNeeded Then _
                    allowed = identityBytesNeeded
            End If
        End If

        If allowed <= 0 Then Exit Do
        WsEnsureRawCapacity m_ws_rawCount + allowed
        If m_ws_state = WsTransportFailed Then Exit Do

        Dim result As Long
        result = WsApi_recvSocket(m_ws_socket, m_ws_rawBuffer(m_ws_rawCount), allowed, 0)
        callCount = callCount + 1

        If result > 0 Then
            m_ws_rawCount = m_ws_rawCount + result
            receivedBytes = receivedBytes + result

            If parseFrames Then
                WsParseAvailableFrames pumpDeadline, copyGuardDeadline
                If m_ws_pendingUtf8Rejection Then Exit Do
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

ContinueReceiveLoop:
    Loop
End Sub
Private Sub WsParseAvailableFrames(ByVal pumpDeadline As Double, _
                                 ByVal copyGuardDeadline As Double)
    Do While m_ws_state = WsTransportOpen Or m_ws_state = WsTransportClosing
        If m_ws_pendingUtf8Rejection Then Exit Do
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

                    Dim length16High As Long
                    Dim length16Low As Long
                    Dim length16 As Long
                    length16High = WsReadRawByte()
                    length16Low = WsReadRawByte()
                    length16 = length16High * &H100& + length16Low
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
                    If m_ws_directDecodeCandidate Then
                        WsEvaluateDirectTextRetention
                        If m_ws_state = WsTransportFailed Then Exit Do
                    End If

                    If m_ws_directDecodeCandidate Then
                    ' Zero-VBA-copy fast path: leave payload untouched until the whole frame is
                    ' contiguous in raw storage, then pass its first Byte directly to Windows UTF-8.
                    If WsRawAvailable < m_ws_framePayloadLength Then Exit Do

                    Dim directText As String
                    If m_ws_framePayloadLength > 0 Then
                        If Not WsTryDecodeUtf8Span( _
                            m_ws_rawBuffer(m_ws_rawReadPos), _
                            m_ws_framePayloadLength, directText) Then
                            Exit Do
                        End If
                    Else
                        directText = vbNullString
                    End If

                    If m_ws_pendingUtf8Rejection Or m_ws_state = WsTransportFailed Then Exit Do
                    If Not WsEnqueueDeferredText(directText) Then Exit Do

                    m_ws_messageTextQueued = True
                    m_ws_messageCount = m_ws_framePayloadLength
                    m_ws_messageObservedBytes = CDbl(m_ws_framePayloadLength)
                    WsConsumeRaw m_ws_framePayloadLength
                    m_ws_framePayloadRemaining = 0
                    m_ws_frameState = WsFrameFinalize

                    Else
                    Dim available As Long
                    available = WsRawAvailable
                    If available <= 0 Then Exit Do

                    Dim takeCount As Long
                    takeCount = available
                    If takeCount > m_ws_framePayloadRemaining Then takeCount = m_ws_framePayloadRemaining

                    If m_ws_frameOpcode >= WS_OPCODE_CLOSE Then
                        WsConsumeControlPayload takeCount

                    ElseIf m_ws_messageOversized Then
                        ' Preserve only the bounded identity prefix; all remaining oversize bytes
                        ' are discarded directly from raw storage without message-buffer copies.
                        m_ws_messageObservedBytes = m_ws_messageObservedBytes + CDbl(takeCount)
                        WsCaptureMessageDiagnosticPrefix takeCount
                        WsConsumeRaw takeCount

                    ElseIf m_ws_discardMessage Or m_ws_frameDiscardPayload Then
                        ' Binary/non-retained data needs no VBA-level payload copy. For a parsed-event
                        ' early drop, continue wire-byte accounting so a later transport-cap crossing
                        ' remains visible even though message retention already stopped.
                        If m_ws_messageEarlyParsedOversizeDrop Then _
                            m_ws_messageObservedBytes = m_ws_messageObservedBytes + CDbl(takeCount)
                        WsConsumeRaw takeCount

                    Else
                        ' Only fragmented in-limit text uses a VBA Byte() copy. Keep that copy chunk
                        ' bounded so the parser can respect its nominal Pump work budget.
                        If WsMonotonicMs() >= copyGuardDeadline Then Exit Do
                        If takeCount > WS_PARSER_COPY_CHUNK_SIZE Then _
                            takeCount = WS_PARSER_COPY_CHUNK_SIZE

                        If m_ws_fragmentReservationDeferred Then
                            Dim deferredRoom As Long
                            deferredRoom = WsByteArrayCapacity(m_ws_messageBuffer) - m_ws_messageCount
                            If deferredRoom <= 0 Then
                                WsSetFailure 9219, "Deferred fragmented-text classification buffer is full before identity resolution."
                                Exit Do
                            End If
                            If takeCount > deferredRoom Then takeCount = deferredRoom
                        End If

                        WsAppendFragmentPayload takeCount
                        If m_ws_state = WsTransportFailed Then Exit Do
                        WsConsumeRaw takeCount
                    End If

                    m_ws_framePayloadRemaining = m_ws_framePayloadRemaining - takeCount
                    If m_ws_framePayloadRemaining = 0 Then
                        m_ws_frameState = WsFrameFinalize
                    End If
                    End If
                End If

            Case WsFrameFinalize
                If WsMonotonicMs() >= pumpDeadline Then Exit Do

                WsCompleteCurrentFrame
                If m_ws_pendingUtf8Rejection Then Exit Do
                If m_ws_state = WsTransportFailed Then Exit Do
        End Select
    Loop
End Sub
Private Sub WsPrepareFramePayload(ByVal payloadLength As Long)
    m_ws_directDecodeCandidate = False
    m_ws_fragmentReservationDeferred = False

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
        m_ws_messageTextQueued = False
        m_ws_discardMessage = (m_ws_frameOpcode = WS_OPCODE_BINARY)
        m_ws_messageIdentityResolved = False
        m_ws_messageKnownAsyncEvent = False
        m_ws_messageEarlyParsedOversizeDrop = False
        Erase m_ws_messageBuffer
        WsBeginMessageDiagnostics
    End If

    m_ws_framePayloadLength = payloadLength
    m_ws_framePayloadRemaining = payloadLength
    m_ws_frameDiscardPayload = False
    m_ws_controlCount = 0
    Erase m_ws_controlPayload

    ' Reject by transport message cap before any large raw/message-buffer reservation.
    If m_ws_frameOpcode <= WS_OPCODE_BINARY Then
        If payloadLength > m_ws_maxIncomingSize Then
            ' A later oversized continuation can be the first point where a fragmented text
            ' message crosses the transport cap. Preserve the already-buffered message head
            ' before the fragment buffer is erased so root id/method classification remains
            ' available to WsRecordOversizedMessage.
            If m_ws_frameOpcode = WS_OPCODE_CONTINUATION And _
               m_ws_messageOpcode = WS_OPCODE_TEXT And _
               Not m_ws_messageOversized And _
               Not m_ws_messageEarlyParsedOversizeDrop And _
               m_ws_messageCount > 0 Then
                WsSeedDiagnosticPrefixFromMessageBuffer
                m_ws_messageObservedBytes = CDbl(m_ws_messageCount)
            End If

            m_ws_messageOversized = True
            m_ws_discardMessage = True
            m_ws_frameDiscardPayload = True

        ElseIf m_ws_messageEarlyParsedOversizeDrop Then
            ' Parsed-event early drop stopped growing m_ws_messageCount. Continue tracking
            ' actual wire bytes so a later fragmented message crossing MaxIncomingSize still
            ' records the transport-level oversize diagnostic without duplicating async loss.
            If m_ws_messageObservedBytes + CDbl(payloadLength) > CDbl(m_ws_maxIncomingSize) Then
                m_ws_messageOversized = True
                m_ws_frameDiscardPayload = True
            End If

        ElseIf Not m_ws_discardMessage Then
            If CDbl(m_ws_messageCount) + CDbl(payloadLength) > CDbl(m_ws_maxIncomingSize) Then
                ' If fragmentation crossed the cap, preserve the original message head from the
                ' bytes already accumulated before discarding the fragment buffer.
                WsSeedDiagnosticPrefixFromMessageBuffer
                m_ws_messageObservedBytes = CDbl(m_ws_messageCount)
                m_ws_messageOversized = True
                m_ws_discardMessage = True
                m_ws_frameDiscardPayload = True
            End If
        End If
    End If

    If m_ws_messageOversized Then
        Erase m_ws_messageBuffer

    ElseIf m_ws_frameOpcode = WS_OPCODE_TEXT And _
           m_ws_frameFin And _
           Not m_ws_discardMessage Then
        ' Non-fragmented text is the dominant path. For messages above the async-event parse
        ' budget, defer the full raw reservation until a bounded root id/method prefix decides
        ' whether this is an event (drop early) or response/unclassified traffic (retain).
        m_ws_directDecodeCandidate = True
        If payloadLength <= PARSED_ASYNC_EVENT_MAX_UTF8_BYTES Then
            WsEnsureRawCapacityForFrame payloadLength
        Else
            WsEnsureRawCapacityForFrame OVERSIZED_DIAGNOSTIC_PREFIX_BYTES
        End If
        If m_ws_state = WsTransportFailed Then Exit Sub

    ElseIf (m_ws_frameOpcode = WS_OPCODE_TEXT Or _
            m_ws_frameOpcode = WS_OPCODE_CONTINUATION) And _
           m_ws_messageOpcode = WS_OPCODE_TEXT And _
           Not m_ws_discardMessage Then
        Dim plannedMessageBytes As Double
        plannedMessageBytes = CDbl(m_ws_messageCount) + CDbl(payloadLength)

        If plannedMessageBytes > PARSED_ASYNC_EVENT_MAX_UTF8_BYTES Then
            WsExtendDiagnosticPrefixFromMessageBuffer
            WsResolveMessageIdentityFromPrefix

            If m_ws_messageIdentityResolved And m_ws_messageKnownAsyncEvent Then
                WsMarkEarlyParsedEventDrop
            ElseIf m_ws_messageIdentityResolved Then
                WsEnsureMessageCapacity CLng(plannedMessageBytes)
            Else
                ' First oversized fragment: retain only one bounded parser chunk until the first
                ' 512 bytes can classify the message, then expand once only if it is not an event.
                m_ws_fragmentReservationDeferred = True
                Dim initialFragmentReserve As Long
                initialFragmentReserve = payloadLength
                If initialFragmentReserve > WS_PARSER_COPY_CHUNK_SIZE Then _
                    initialFragmentReserve = WS_PARSER_COPY_CHUNK_SIZE
                WsEnsureMessageCapacity m_ws_messageCount + initialFragmentReserve
            End If
        Else
            WsEnsureMessageCapacity CLng(plannedMessageBytes)
        End If

        If m_ws_state = WsTransportFailed Then Exit Sub
    End If

    m_ws_frameState = WsFrameReadPayload
End Sub
Private Sub WsConsumeControlPayload(ByVal count As Long)
    If count <= 0 Then Exit Sub

    WsEnsureControlCapacity m_ws_controlCount + count

    Dim i As Long
    For i = 0 To count - 1
        m_ws_controlPayload(m_ws_controlCount + i) = CByte(WsPeekRawByte(i))
    Next

    m_ws_controlCount = m_ws_controlCount + count
    WsConsumeRaw count
End Sub

Private Sub WsAppendFragmentPayload(ByVal count As Long)
    If count <= 0 Then Exit Sub

    If WsByteArrayCapacity(m_ws_messageBuffer) < (m_ws_messageCount + count) Then
        WsSetFailure 9219, "Fragmented text accumulator capacity is insufficient."
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To count - 1
        m_ws_messageBuffer(m_ws_messageCount + i) = CByte(WsPeekRawByte(i))
    Next

    m_ws_messageCount = m_ws_messageCount + count

    If m_ws_fragmentReservationDeferred Then
        WsExtendDiagnosticPrefixFromMessageBuffer
        WsResolveMessageIdentityFromPrefix

        If m_ws_messageIdentityResolved Then
            Dim remainingAfterThisCopy As Long
            remainingAfterThisCopy = m_ws_framePayloadRemaining - count
            If remainingAfterThisCopy < 0 Then remainingAfterThisCopy = 0

            Dim plannedBytes As Double
            plannedBytes = CDbl(m_ws_messageCount) + CDbl(remainingAfterThisCopy)

            If m_ws_messageKnownAsyncEvent And plannedBytes > PARSED_ASYNC_EVENT_MAX_UTF8_BYTES Then
                WsMarkEarlyParsedEventDrop
            Else
                WsEnsureMessageCapacity CLng(plannedBytes)
            End If

            m_ws_fragmentReservationDeferred = False
        End If
    End If
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
                If Not m_ws_messageTextQueued Then
                    Dim fragmentedText As String
                    If m_ws_messageCount > 0 Then
                        If Not WsTryDecodeUtf8Span( _
                            m_ws_messageBuffer(0), m_ws_messageCount, fragmentedText) Then
                            Exit Sub
                        End If
                    Else
                        fragmentedText = vbNullString
                    End If

                    If m_ws_pendingUtf8Rejection Or m_ws_state = WsTransportFailed Then Exit Sub
                    If Not WsEnqueueDeferredText(fragmentedText) Then Exit Sub
                    m_ws_messageTextQueued = True
                End If
            End If

            If m_ws_messageOversized Then
                WsRecordOversizedMessage
                WsTraceOversizedMessageDiagnostic
            End If

            m_ws_currentPumpCompletedMessages = m_ws_currentPumpCompletedMessages + 1
            WsResetMessageAccumulator
        End If
    End If

    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_framePayloadLength = 0
    m_ws_framePayloadRemaining = 0
    m_ws_frameDiscardPayload = False
    m_ws_directDecodeCandidate = False
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

    Dim currentMs As Double
    currentMs = WsMonotonicMs()

    If m_ws_clientPingOutstanding Then
        If m_ws_clientPingSentAtMs > 0 Then
            If (currentMs - m_ws_clientPingSentAtMs) >= WS_CLIENT_PONG_TIMEOUT_MS Then
                WsSetFailure 9221, _
                    "Client Ping timed out waiting for matching Pong. Sequence=" & _
                    CStr(m_ws_clientPingOutstandingSequence) & "."
            End If
        End If
        Exit Sub
    End If

    If m_ws_nextClientPingAtMs <= 0 Then
        m_ws_nextClientPingAtMs = currentMs + WS_CLIENT_PING_INTERVAL_MS
        Exit Sub
    End If

    If currentMs < m_ws_nextClientPingAtMs Then Exit Sub

    If WsHasPendingSend Then
        m_ws_nextClientPingAtMs = currentMs + WS_CLIENT_PING_BUSY_RETRY_MS
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
    m_ws_clientPingSentAtMs = currentMs
    m_ws_nextClientPingAtMs = currentMs + WS_CLIENT_PING_INTERVAL_MS
End Sub

Private Sub WsHandleClientPong()
    If m_ws_controlCount <> WS_CLIENT_PING_PAYLOAD_SIZE Then Exit Sub

    Dim sequence As Long
    If Not WsTryDecodePingSequence(m_ws_controlPayload, sequence) Then Exit Sub

    If Not m_ws_clientPingOutstanding Or _
       sequence <> m_ws_clientPingOutstandingSequence Then
        Exit Sub
    End If

    Dim currentMs As Double
    currentMs = WsMonotonicMs()

    m_ws_clientPingOutstanding = False
    m_ws_clientPingOutstandingSequence = 0
    m_ws_clientPingSentAtMs = 0
    m_ws_nextClientPingAtMs = currentMs + WS_CLIENT_PING_INTERVAL_MS
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
                   Right$("00000000" & Hex$(status), 8)
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

Private Sub WsEnsureRawCapacityForFrame(ByVal payloadLength As Long)
    If payloadLength < 0 Then
        WsSetFailure 9219, "Negative direct-decode payload length."
        Exit Sub
    End If

    Dim requiredDouble As Double
    requiredDouble = CDbl(m_ws_rawReadPos) + CDbl(payloadLength)
    If requiredDouble > 2147483647# Then
        WsSetFailure 9219, "Direct-decode raw buffer requirement exceeds VBA Long range."
        Exit Sub
    End If

    Dim requiredCount As Long
    requiredCount = CLng(requiredDouble)
    If requiredCount <= 0 Then Exit Sub
    If WsByteArrayCapacity(m_ws_rawBuffer) >= requiredCount Then Exit Sub

    If WsByteArrayCapacity(m_ws_rawBuffer) = 0 Then
        ReDim m_ws_rawBuffer(0 To requiredCount - 1)
    Else
        ReDim Preserve m_ws_rawBuffer(0 To requiredCount - 1)
    End If
End Sub

Private Sub WsEnsureMessageCapacity(ByVal requiredCount As Long)
    If requiredCount <= 0 Then Exit Sub
    If requiredCount > m_ws_maxIncomingSize Then
        WsSetFailure 9219, "Fragmented text buffer request exceeds MaxIncomingSize."
        Exit Sub
    End If

    Dim capacity As Long
    capacity = WsByteArrayCapacity(m_ws_messageBuffer)
    If capacity >= requiredCount Then Exit Sub

    Dim newCapacity As Long
    newCapacity = capacity
    If newCapacity < WS_RECEIVE_CHUNK_SIZE Then newCapacity = WS_RECEIVE_CHUNK_SIZE

    Do While newCapacity < requiredCount
        If newCapacity > (m_ws_maxIncomingSize \ 2) Then
            newCapacity = requiredCount
            Exit Do
        End If
        newCapacity = newCapacity * 2
    Loop

    If newCapacity > m_ws_maxIncomingSize Then newCapacity = requiredCount

    If capacity = 0 Then
        ReDim m_ws_messageBuffer(0 To newCapacity - 1)
    Else
        ReDim Preserve m_ws_messageBuffer(0 To newCapacity - 1)
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
        Mid$(text, i + 1, 1) = ChrW$(m_ws_rawBuffer(m_ws_rawReadPos + i))
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

Private Sub WsBeginMessageDiagnostics()
    m_ws_messageObservedBytes = 0
    m_ws_messageOversized = False
    m_ws_messageDiagnosticPrefixCount = 0
End Sub

Private Sub WsCaptureMessageDiagnosticPrefix(ByVal count As Long)
    If count <= 0 Then Exit Sub

    Dim available As Long
    available = OVERSIZED_DIAGNOSTIC_PREFIX_BYTES - m_ws_messageDiagnosticPrefixCount
    If available <= 0 Then Exit Sub
    If count < available Then available = count

    Dim i As Long
    For i = 0 To available - 1
        m_ws_messageDiagnosticPrefix(m_ws_messageDiagnosticPrefixCount + i) = _
            CByte(WsPeekRawByte(i))
    Next

    m_ws_messageDiagnosticPrefixCount = m_ws_messageDiagnosticPrefixCount + available
End Sub

Private Sub WsSeedDiagnosticPrefixFromMessageBuffer()
    ' Transport-oversize paths and parsed-event early-drop paths share the same bounded prefix.
    WsExtendDiagnosticPrefixFromMessageBuffer
End Sub

Private Sub WsCaptureDirectDiagnosticPrefix()
    If m_ws_messageDiagnosticPrefixCount >= OVERSIZED_DIAGNOSTIC_PREFIX_BYTES Then Exit Sub

    Dim visibleBytes As Long
    visibleBytes = WsRawAvailable
    If visibleBytes <= m_ws_messageDiagnosticPrefixCount Then Exit Sub

    Dim targetCount As Long
    targetCount = visibleBytes
    If targetCount > OVERSIZED_DIAGNOSTIC_PREFIX_BYTES Then _
        targetCount = OVERSIZED_DIAGNOSTIC_PREFIX_BYTES

    Dim i As Long
    For i = m_ws_messageDiagnosticPrefixCount To targetCount - 1
        m_ws_messageDiagnosticPrefix(i) = _
            m_ws_rawBuffer(m_ws_rawReadPos + i)
    Next

    m_ws_messageDiagnosticPrefixCount = targetCount
End Sub

Private Sub WsExtendDiagnosticPrefixFromMessageBuffer()
    If m_ws_messageDiagnosticPrefixCount >= OVERSIZED_DIAGNOSTIC_PREFIX_BYTES Then Exit Sub
    If m_ws_messageCount <= m_ws_messageDiagnosticPrefixCount Then Exit Sub

    Dim targetCount As Long
    targetCount = m_ws_messageCount
    If targetCount > OVERSIZED_DIAGNOSTIC_PREFIX_BYTES Then _
        targetCount = OVERSIZED_DIAGNOSTIC_PREFIX_BYTES

    If WsByteArrayCapacity(m_ws_messageBuffer) < targetCount Then Exit Sub

    Dim i As Long
    For i = m_ws_messageDiagnosticPrefixCount To targetCount - 1
        m_ws_messageDiagnosticPrefix(i) = m_ws_messageBuffer(i)
    Next

    m_ws_messageDiagnosticPrefixCount = targetCount
End Sub

Private Sub WsResolveMessageIdentityFromPrefix()
    If m_ws_messageIdentityResolved Then Exit Sub
    If m_ws_messageDiagnosticPrefixCount <= 0 Then Exit Sub

    Dim prefixJson As String
    prefixJson = WsDiagnosticPrefixJsonText()

    Dim rootId As String
    rootId = ExtractJsonId(prefixJson)

    Dim eventMethod As String
    Dim isEvent As Boolean
    isEvent = TryExtractTopLevelEventMethod(prefixJson, eventMethod)

    If isEvent And Len(rootId) = 0 Then
        m_ws_messageIdentityResolved = True
        m_ws_messageKnownAsyncEvent = True
    ElseIf Len(rootId) > 0 Then
        m_ws_messageIdentityResolved = True
        m_ws_messageKnownAsyncEvent = False
    ElseIf m_ws_messageDiagnosticPrefixCount >= OVERSIZED_DIAGNOSTIC_PREFIX_BYTES Then
        ' Bounded prefix could not prove event identity. Preserve the message through the normal
        ' transport cap rather than risking a command-response drop.
        m_ws_messageIdentityResolved = True
        m_ws_messageKnownAsyncEvent = False
    End If
End Sub

Private Sub WsMarkEarlyParsedEventDrop()
    If m_ws_messageEarlyParsedOversizeDrop Then Exit Sub

    Dim prefixJson As String
    prefixJson = WsDiagnosticPrefixJsonText()

    Dim detectedMethod As String
    RecordAsyncMessageLoss prefixJson, detectedMethod
    IncrementLossCounter m_DroppedParsedOversizeCount

    m_ws_messageEarlyParsedOversizeDrop = True
    ' For fragmented text, m_ws_messageCount is the byte count already accumulated in the
    ' message buffer (including the current copied chunk). Subsequent discarded bytes continue
    ' accumulating in m_ws_messageObservedBytes. Direct text starts from zero here.
    m_ws_messageObservedBytes = CDbl(m_ws_messageCount)
    m_ws_discardMessage = True
    m_ws_frameDiscardPayload = True
    m_ws_directDecodeCandidate = False
    m_ws_fragmentReservationDeferred = False

    m_LastParsedEventDropReason = _
        "oversized async event dropped before UTF-8 decode (transportBytes>" & _
        CStr(PARSED_ASYNC_EVENT_MAX_UTF8_BYTES) & "; method=" & _
        IIf(Len(detectedMethod) > 0, detectedMethod, "unknown") & ")"

    Erase m_ws_messageBuffer
End Sub

Private Sub WsEvaluateDirectTextRetention()
    If Not m_ws_directDecodeCandidate Then Exit Sub
    If m_ws_framePayloadLength <= PARSED_ASYNC_EVENT_MAX_UTF8_BYTES Then Exit Sub

    WsCaptureDirectDiagnosticPrefix
    WsResolveMessageIdentityFromPrefix

    If Not m_ws_messageIdentityResolved Then Exit Sub

    If m_ws_messageKnownAsyncEvent Then
        WsMarkEarlyParsedEventDrop
        Exit Sub
    End If

    ' Response/reserved/unclassified text retains the transport-level MaxIncomingSize contract.
    WsEnsureRawCapacityForFrame m_ws_framePayloadLength
End Sub

Private Function WsDiagnosticPrefixText() As String
    Dim result As String
    Dim i As Long
    Dim b As Long

    For i = 0 To m_ws_messageDiagnosticPrefixCount - 1
        b = CLng(m_ws_messageDiagnosticPrefix(i))

        Select Case b
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 13
                result = result & "\r"
            Case 32 To 126
                result = result & Chr$(b)
            Case Else
                result = result & "."
        End Select
    Next

    WsDiagnosticPrefixText = result
End Function

Private Function WsDiagnosticFusionMarkerCount() As Long
    Dim i As Long

    For i = 0 To m_ws_messageDiagnosticPrefixCount - 2
        If m_ws_messageDiagnosticPrefix(i) = 125 And _
           m_ws_messageDiagnosticPrefix(i + 1) = 123 Then
            WsDiagnosticFusionMarkerCount = WsDiagnosticFusionMarkerCount + 1
        End If
    Next
End Function

Private Function WsDiagnosticPrefixJsonText() As String
    ' Preserve ASCII JSON syntax and whitespace from the retained prefix. Non-ASCII payload
    ' bytes are replaced with spaces; root id/method keys are ASCII and normally occur near
    ' the start of a BiDi message, so bounded classification does not require full decoding.
    Dim result As String
    Dim i As Long
    Dim b As Long

    For i = 0 To m_ws_messageDiagnosticPrefixCount - 1
        b = CLng(m_ws_messageDiagnosticPrefix(i))
        Select Case b
            Case 9
                result = result & vbTab
            Case 10
                result = result & vbLf
            Case 13
                result = result & vbCr
            Case 32 To 126
                result = result & Chr$(b)
            Case Else
                result = result & " "
        End Select
    Next

    WsDiagnosticPrefixJsonText = result
End Function

Private Sub WsRecordOversizedMessage()
    IncrementLossCounter m_TransportOversizedMessageCount

    Dim prefixJson As String
    prefixJson = WsDiagnosticPrefixJsonText()

    Dim rootId As String
    rootId = ExtractJsonId(prefixJson)

    Dim eventMethod As String
    Dim isEvent As Boolean
    isEvent = TryExtractTopLevelEventMethod(prefixJson, eventMethod)

    Dim classification As String
    classification = "unclassified"

    If m_ws_messageOpcode = WS_OPCODE_TEXT And isEvent And Len(rootId) = 0 Then
        classification = "event"
        IncrementLossCounter m_TransportOversizedEventCount

        ' This event never reaches the upper FIFO/parsed-event path. Count it as a dropped
        ' async event unless the same message was already recorded by the earlier parsed-event cap.
        If Not m_ws_messageEarlyParsedOversizeDrop Then
            IncrementLossCounter m_DroppedEventsCount
            Dim detectedMethod As String
            RecordAsyncMessageLoss prefixJson, detectedMethod
        End If

    ElseIf m_ws_messageOpcode = WS_OPCODE_TEXT And IsNormalCommandId(rootId) Then
        classification = "command-response"
        IncrementLossCounter m_TransportOversizedCommandResponseCount

        If m_isSendReceiving And _
           StrComp(rootId, m_activeCommandId, vbBinaryCompare) = 0 Then
            m_activeCommandOversizedResponseSeen = True
        End If

    Else
        IncrementLossCounter m_TransportOversizedUnclassifiedCount
        ' Identity could not be proven from the bounded prefix. Treat it as conservative
        ' critical async-loss evidence so exact wrapper captures fail closed.
        IncrementLossCounter m_UnknownCriticalEventLossCount
        If m_isSendReceiving Then m_activeCommandOversizedUnclassifiedSeen = True
    End If

    m_LastTransportOversizeDiagnostic = _
        "class=" & classification & _
        " | observedBytes=" & Format$(m_ws_messageObservedBytes, "0") & _
        " | limitBytes=" & CStr(m_ws_maxIncomingSize) & _
        " | id=" & IIf(Len(rootId) > 0, rootId, "(none)") & _
        " | method=" & IIf(Len(eventMethod) > 0, eventMethod, "(none)")

    If m_isSendReceiving And _
       (m_activeCommandOversizedResponseSeen Or m_activeCommandOversizedUnclassifiedSeen) Then
        m_activeCommandOversizeDiagnostic = m_LastTransportOversizeDiagnostic
    End If
End Sub

Private Sub WsTraceOversizedMessageDiagnostic()
    Trace ">>> [Winsock] OVERSIZED MESSAGE DISCARDED. observedBytes=" & _
          Format$(m_ws_messageObservedBytes, "0") & _
          " limitBytes=" & CStr(m_ws_maxIncomingSize) & _
          " retainedBytes=" & CStr(m_ws_messageDiagnosticPrefixCount) & _
          " fusionMarkers(}{)=" & CStr(WsDiagnosticFusionMarkerCount()) & _
          " HEAD" & CStr(m_ws_messageDiagnosticPrefixCount) & "=[" & _
          WsDiagnosticPrefixText() & "]", True
End Sub

Private Sub WsResetMessageAccumulator()
    m_ws_messageActive = False
    m_ws_messageOpcode = 0
    m_ws_messageCount = 0
    m_ws_messageTextQueued = False
    m_ws_directDecodeCandidate = False
    m_ws_discardMessage = False
    m_ws_messageObservedBytes = 0
    m_ws_messageOversized = False
    m_ws_messageDiagnosticPrefixCount = 0
    m_ws_messageIdentityResolved = False
    m_ws_messageKnownAsyncEvent = False
    m_ws_messageEarlyParsedOversizeDrop = False
    m_ws_fragmentReservationDeferred = False
    Erase m_ws_messageBuffer
End Sub
Private Sub WsResetRuntimeState()
    m_ws_LastErrorCode = 0
    m_ws_lastProtocolCloseStatus = 0
    m_ws_lastErrorText = vbNullString

    Set m_ws_sendQueue = New Collection
    WsResetDeferredTextQueue

    Erase m_ws_sendCurrent
    Erase m_ws_rawBuffer
    Erase m_ws_controlPayload
    Erase m_ws_messageBuffer

    m_ws_sendCurrentCount = 0
    m_ws_sendOffset = 0
    m_ws_rawCount = 0
    m_ws_rawReadPos = 0
    m_ws_frameState = WsFrameReadBaseHeader
    m_ws_framePayloadLength = 0
    m_ws_framePayloadRemaining = 0
    m_ws_frameDiscardPayload = False
    m_ws_directDecodeCandidate = False
    m_ws_controlCount = 0
    m_ws_closeSent = False
    m_ws_pendingUtf8Rejection = False
    m_ws_pendingUtf8RejectionText = vbNullString
    m_ws_nextClientPingAtMs = 0
    m_ws_clientPingOutstanding = False
    m_ws_clientPingOutstandingSequence = 0
    m_ws_clientPingSentAtMs = 0
    m_ws_handshakeClientKey = vbNullString
    m_ws_expectedHandshakeAccept = vbNullString
    m_ws_handshakeUserAgent = vbNullString
    m_ws_handshakeValidationError = vbNullString
    m_ws_urlValidationError = vbNullString

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

    If WINSOCK_TRANSPORT_DEBUG_ENABLED Then
        Debug.Print "BiDiSocketCommunicator[IntegratedWinsock] FAILED " & errorCode & ": " & description
    End If
End Sub

Private Function WsParseLoopbackUrl(ByVal wsUrl As String) As Boolean
    m_ws_urlValidationError = vbNullString
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

    If Not WsIsSafeAsciiRequestTarget(m_ws_path) Then
        m_ws_urlValidationError = "WebSocket request target must start with '/' and contain only visible ASCII characters; CR/LF, spaces and raw non-ASCII are rejected."
        Exit Function
    End If

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

Private Function WsIsSafeAsciiRequestTarget(ByVal target As String) As Boolean
    If Len(target) = 0 Then Exit Function
    If Left$(target, 1) <> "/" Then Exit Function

    Dim i As Long
    Dim codePoint As Long
    For i = 1 To Len(target)
        codePoint = AscW(Mid$(target, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536

        ' HTTP request-target is emitted verbatim into the request line. Require visible ASCII;
        ' callers must percent-encode non-ASCII data before constructing the WebSocket URL.
        If codePoint < 33 Or codePoint > 126 Then Exit Function
    Next

    WsIsSafeAsciiRequestTarget = True
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
    Dim codePoint As Long
    For i = 1 To count
        codePoint = AscW(Mid$(text, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        If codePoint > 127 Then
            Err.Raise 9232, "BiDiSocketCommunicator.ASCII", _
                      "Non-ASCII character rejected while building WebSocket HTTP data."
        End If
        result(i - 1) = CByte(codePoint)
    Next

    WsEncodeAscii = result
End Function

Private Function WsEncodeUtf8(ByVal text As String) As Byte()
    Dim result() As Byte
    Dim wideCharCount As Long
    wideCharCount = Len(text)

    If wideCharCount = 0 Then
        WsEncodeUtf8 = result
        Exit Function
    End If

    Dim requiredBytes As Long
    Dim lastError As Long

    requiredBytes = WsApi_WideCharToMultiByteSize( _
        WS_CP_UTF8, WS_WC_ERR_INVALID_CHARS, StrPtr(text), wideCharCount, _
        0, 0, 0, 0)
    lastError = Err.LastDllError

    If requiredBytes = 0 Then WsRaiseEncodeError lastError

    ReDim result(0 To requiredBytes - 1)

    Dim convertedBytes As Long
    convertedBytes = WsApi_WideCharToMultiByteBytes( _
        WS_CP_UTF8, WS_WC_ERR_INVALID_CHARS, StrPtr(text), wideCharCount, _
        result(0), requiredBytes, 0, 0)
    lastError = Err.LastDllError

    If convertedBytes = 0 Then
        WsRaiseEncodeError lastError
    ElseIf convertedBytes <> requiredBytes Then
        Err.Raise WS_ERR_UTF8_CODEC, "BiDiSocketCommunicator.UTF8", _
                  "WideCharToMultiByte returned an unexpected byte count."
    End If

    WsEncodeUtf8 = result
End Function

Private Function WsTryDecodeUtf8Span(ByRef firstByte As Byte, _
                                     ByVal byteCount As Long, _
                                     ByRef decodedText As String) As Boolean
    On Error GoTo DecodeError

    If byteCount <= 0 Then
        decodedText = vbNullString
        WsTryDecodeUtf8Span = True
        Exit Function
    End If

    decodedText = WsDecodeUtf8Span(firstByte, byteCount)
    WsTryDecodeUtf8Span = True
    Exit Function

DecodeError:
    Dim savedNumber As Long
    Dim savedDescription As String
    savedNumber = Err.Number
    savedDescription = Err.description
    Err.Clear
    On Error GoTo 0

    If savedNumber = WS_ERR_INVALID_UTF8 Then
        WsPendInvalidPeerUtf8 savedDescription
    Else
        WsSetFailure savedNumber, "UTF-8 decode failed: " & savedDescription
    End If
End Function

Private Function WsDecodeUtf8Span(ByRef firstByte As Byte, _
                                  ByVal byteCount As Long) As String
    If byteCount <= 0 Then Exit Function

    Dim requiredWideChars As Long
    Dim lastError As Long

    requiredWideChars = WsApi_MultiByteToWideChar( _
        WS_CP_UTF8, WS_MB_ERR_INVALID_CHARS, _
        firstByte, byteCount, 0, 0)
    lastError = Err.LastDllError

    If requiredWideChars = 0 Then WsRaiseDecodeError lastError

    Dim result As String
    result = String$(requiredWideChars, vbNullChar)

    Dim convertedWideChars As Long
    convertedWideChars = WsApi_MultiByteToWideChar( _
        WS_CP_UTF8, WS_MB_ERR_INVALID_CHARS, _
        firstByte, byteCount, StrPtr(result), requiredWideChars)
    lastError = Err.LastDllError

    If convertedWideChars = 0 Then
        WsRaiseDecodeError lastError
    ElseIf convertedWideChars <> requiredWideChars Then
        Err.Raise WS_ERR_UTF8_CODEC, "BiDiSocketCommunicator.UTF8", _
                  "MultiByteToWideChar returned an unexpected character count."
    End If

    WsDecodeUtf8Span = result
End Function

Private Sub WsRaiseEncodeError(ByVal windowsError As Long)
    If windowsError = WS_ERROR_NO_UNICODE_TRANSLATION Then
        Err.Raise WS_ERR_INVALID_UTF16, "BiDiSocketCommunicator.UTF8", _
                  "Outgoing text contains an invalid UTF-16 surrogate sequence."
    End If

    Err.Raise WS_ERR_UTF8_CODEC, "BiDiSocketCommunicator.UTF8", _
              "WideCharToMultiByte failed. Windows error: " & windowsError
End Sub

Private Sub WsRaiseDecodeError(ByVal windowsError As Long)
    If windowsError = WS_ERROR_NO_UNICODE_TRANSLATION Then
        Err.Raise WS_ERR_INVALID_UTF8, "BiDiSocketCommunicator.UTF8", _
                  "Peer text message contains malformed UTF-8."
    End If

    Err.Raise WS_ERR_UTF8_CODEC, "BiDiSocketCommunicator.UTF8", _
              "MultiByteToWideChar failed. Windows error: " & windowsError
End Sub

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
