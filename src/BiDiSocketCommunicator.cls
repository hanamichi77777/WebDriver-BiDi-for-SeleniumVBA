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
' ==========================================================================
' WebDriver BiDi for SeleniumVBA(https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA)
' Version: 2.4
' MIT License Copyright (c) hanamichi77777
' ==========================================================================
' Class Name : BiDiSocketCommunicator
' Purpose    : Professional WebSocket Client for WebDriver BiDi (W3C Standard).
'
' [SECURITY & COMPLIANCE - ANALYST REFERENCE]
' 1. SANDBOXED IPC SCOPE: Strictly restricted to the Local Loopback interface (127.0.0.1).
'    Engineered exclusively for Inter-Process Communication (IPC) between the VBA host
'    and the local WebDriver/Browser instance.
' 2. ANTI-RAT DESIGN: This implementation enforces a strict URL validation policy that
'    blocks non-localhost routing, preventing its use as a Remote Access Tool (RAT).
' 3. NO DATA EXFILTRATION: By architectural design, this module cannot initiate connections
'    to remote C2 (Command & Control) servers, and is designed to prevent outbound exfiltration.
' 4. NATIVE OS INTEGRITY: Utilizes 'WinHttp.dll' (a signed, core Windows System Component)
'    for socket handshaking, complying with enterprise-level EDR auditing and OS policies.
' 5. OPEN-SOURCE TRANSPARENCY: Part of the SeleniumVBA project (GCuser99).
'    The full source code and logic are available for independent security audits:
'    Audit URL: https://github.com/GCuser99/SeleniumVBA
' 6. ARCHITECTURAL NECESSITY (W3C BiDi): Required for compliance with the W3C WebDriver BiDi
'    standard. Unlike legacy one-way REST APIs, BiDi mandates a persistent WebSocket for
'    bidirectional events (e.g., real-time network interception, console logs).
'    The communication is inherently safe as the WebDriver process only exposes its
'    BiDi endpoint to the local loopback interface.
' ==========================================================================
' AI NAVIGATION MAP  (transport layer; the Wrapper is the upper layer)
' Read before editing any single method. The risk is not "where is X" but
' "which wire-level invariant does X participate in". This class owns the
' socket; business logic, retries and SPA waits live in BiDiCommandWrapper.
' [FUNNELS]  Route through these; do not open ad-hoc paths.

'   Connect        : 7-stage WinHttp WebSocket handshake; the ONLY place handles are acquired.
'   SendCommand    : unified id-gen + envelope + size-check + SendAndReceive (callers pass method+params).
'   SendAndReceive : core synchronous I/O, held under the m_isSendReceiving lock.
'   SendMessage    : raw frame send; never re-sends the same id (see INVARIANTS).
'                    allowRecovery:=False = best-effort path (PostScript), never recovers.
'   GetMessage     : 100ms poll; swallows reserved-band replies (id < CMD_ID_FLOOR),
'                    returns all other BiDi events.
'
' [STATE OWNERSHIP]  This class is the owner. Do not duplicate these upstream.
'   WinHttp/WebSocket handles : hSession_, hConnection_, hRequest_, hWebSocket_
'   Request id generation     : GetNextRequestId (seed = CMD_ID_FLOOR)
'   QPC monotonic clock       : NowTicks / NowMs / ElapsedMs (single source of time)
'   I/O locks                 : m_isSendReceiving, m_isReading
'
' [INVARIANTS]  Breaking these = double-free, deadlock, or protocol violation.
'   HANDLE OWNERSHIP : on a successful upgrade the request handle is closed and ownership
'                      passes to hWebSocket_; on a FAILED CompleteUpgrade, hRequest_ is closed
'                      defensively. Never retain or double-close it.
'   RE-ENTRANCY      : no DoEvents / PumpMessages while m_isSendReceiving is True (WinHttp
'                      state corrupts under re-entrant receive).
'   ID BANDS         : [1..99] reserved (PostScript reply=2 fixed; heartbeat pings rotate
'                      in [3..99]); commands [CMD_ID_FLOOR=100 .. ID_MAX]. Recognition is
'                      band-based: any root id below CMD_ID_FLOOR is system traffic.
'   NO RESEND        : on send failure SendMessage recovers the socket but returns False; the same
'                      id is never re-sent (W3C requires unique ids). Retry is the Wrapper's job, new id.
'   SILENT PEER      : sync WebSocket Receive has NO timeout; the timeout tiers work only
'                      while frames arrive. A TCP-alive, frame-silent peer blocks
'                      indefinitely (see the LIMITATION note in SendAndReceive).
'   ERROR HANDLER    : no file I/O (Trace) inside the error handler.
'   QPC FAIL-FAST    : Class_Initialize raises vbObjectError+9001 if QPC is unavailable.
' ==========================================================================
' --- Configuration ---
Private Const DEBUG_MODE As Boolean = False
Private Const AGENT_NAME As String = "SeleniumVBA/BiDi-Client (Open-Source; Local-IPC-Only)"

' --- WinHttp API Declarations ---
Private Declare PtrSafe Function WinHttpOpen Lib "WinHttp" (ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpConnect Lib "WinHttp" (ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr
Private Declare PtrSafe Function WinHttpOpenRequest Lib "WinHttp" (ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpSendRequest Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long) As Long
Private Declare PtrSafe Function WinHttpReceiveResponse Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long
Private Declare PtrSafe Function WinHttpQueryHeaders Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, ByVal pwszName As LongPtr, ByRef lpBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare PtrSafe Function WinHttpSetOption Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpSetTimeouts Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal nResolveTimeout As Long, ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, ByVal nReceiveTimeout As Long) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHttp" (ByVal hInternet As LongPtr) As Long

' --- WebSocket API Declarations ---
Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal pContext As LongPtr) As LongPtr
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal eBufferType As Long, ByRef pvBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByRef pvBuffer As Any, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peBufferType As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal usStatus As Long, ByRef pvReason As Any, ByVal dwReasonLength As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketShutdown Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal usStatus As Long, ByRef pvReason As Any, ByVal dwReasonLength As Long) As Long

' --- Utility API Declarations ---
Private Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- WinHttp Constants ---
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY As Long = 1
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL As Long = 116
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX As Long = 0
Private Const WINHTTP_NO_HEADER_INDEX As Long = 0

' --- WebSocket Buffer Types ---
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE  As Long = 0
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE As Long = 1
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE    As Long = 2
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE   As Long = 3
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE           As Long = 4

' --- Library Settings ---
Private Const NO_ERROR As Long = 0
' [BiDi spec compliance] Reserved POSITIVE id band for system traffic (js-uint conformant).
' Commands are issued only from [CMD_ID_FLOOR .. ID_MAX]; ids [1..99] never collide with them.
' Heartbeat pings ROTATE within [3..99] so no duplicate id is ever outstanding (W3C unique-id
' compliance, same rationale as the NO RESEND invariant). Recognition is band-based:
' IsReservedSystemId discards every root id below CMD_ID_FLOOR at both receive sites.
Private Const CMD_ID_FLOOR As Long = 100         ' command ids: [100 .. ID_MAX]
Private Const ID_MAX As Long = 2000000000        ' wrap ceiling (within VBA Long range)
Private Const POST_SCRIPT_REPLY_ID As Long = 2   ' reserved fixed id: PostScript fire-and-forget reply (discarded)
Private Const SYSTEM_PING_ID_MIN As Long = 3     ' heartbeat rotating id band floor
Private Const SYSTEM_PING_ID_MAX As Long = 99    ' heartbeat rotating id band ceiling
Private Const MAX_CONSECUTIVE_SOFT_TIMEOUTS As Long = 2 ' soft timeouts on one connection before StateError
Private Const DEFAULT_BUFFER_SIZE As Long = 8192
Private Const MAX_BUFFER_SIZE As Long = 10485760 ' 10MB
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 100
Private Const KEEPALIVE_INTERVAL_MS As Long = 15000

' --- State & Handle Variables ---
Private hSession_ As LongPtr
Private hConnection_ As LongPtr
Private hRequest_ As LongPtr
Private hWebSocket_ As LongPtr
Private m_LastWsUrl As String
Private m_LastPort As Long
Private m_RetryCount As Long
Private m_MaxRetryAttempts As Long
Private m_AutoReconnect As Boolean
Private m_consecutiveInternalErrors As Long

' --- Robustness Semaphores & Timestamps ---
Private m_isReading As Boolean
Private m_isReadingSince As Double ' Tracks the timestamp of the last read lock
Private m_discardActive As Boolean ' TRUE iff m_InStream holds a partially-discarded oversized message (persists across ReadFromWire calls; cleared only where m_InStream is truncated)
Private m_isSendReceiving As Boolean ' True when a command sequence is active
Private m_isRecovering As Boolean ' Re-entrancy lock for AttemptRecovery

Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum
Private m_State As BiDiSocketState
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal stateName As String)

' --- Internal Buffers & Objects ---
Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long
Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_ReceiveBuffer() As Byte
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long
Private m_InStream As ADODB.Stream  ' Dedicated Read Stream
Private m_OutStream As ADODB.Stream ' Dedicated Write Stream
Private m_DroppedEventsCount As Long
Private m_lastPollTime As Double
Private m_RegExp As RegExp
Private m_AllowUiPump As Boolean
Private m_qpcFreq As Currency ' QueryPerformanceFrequency cache (set in Class_Initialize)
Private m_lastPumpTicks As Currency ' Throttle state for PumpMessages

' --- Request ID Management (owner; see AI NAVIGATION MAP) ---
Private m_nextRequestId As Long
Private m_nextSystemPingId As Long ' heartbeat id rotation cursor [3..99]
Private m_consecutiveSoftTimeouts As Long ' consecutive soft timeouts on the current connection
Private Const MAX_MSG_SIZE As Long = 300000 ' Threshold to prevent VBA hanging (mirrors Wrapper constant)
Private Const DEFAULT_MAX_INCOMING_SIZE As Long = 2097152 ' 2MB: receive-side assembly cap (runtime-tunable via MaxIncomingSize)
Private m_MaxIncomingSize As Long ' Receive-side cap; oversized messages are discarded with identity diagnostics

' Initialization & Termination
Private Sub Class_Initialize()
    ' Fail-fast QPC initialization. Required for all timer APIs (NowTicks/NowMs/ElapsedMs).
    If QueryPerformanceFrequency(m_qpcFreq) = 0 Or m_qpcFreq <= 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", _
                  "QueryPerformanceFrequency is unavailable on this system. QPC support is required."
    End If
    m_DefaultTimeout = DEFAULT_TIMEOUT_MS
    m_PollingInterval = DEFAULT_POLLING_INTERVAL_MS
    m_AllowUiPump = True
    
    m_BufferSize = DEFAULT_BUFFER_SIZE
    ReDim m_ReceiveBuffer(m_BufferSize - 1)
    m_MaxBufferSize = MAX_BUFFER_SIZE
    m_MaxIncomingSize = DEFAULT_MAX_INCOMING_SIZE

    Set m_EventBuffer = New Collection
    
    ' Maintain separate streams for Read and Write to prevent reentrancy collisions
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open

    Set m_OutStream = New ADODB.Stream
    m_OutStream.mode = adModeReadWrite
    m_OutStream.Open
    
    ' Initial pattern for ExtractJsonId (will be updated JIT if needed)
    Set m_RegExp = New RegExp
    m_RegExp.Global = False
    m_RegExp.IgnoreCase = True
    
    ' Enable controlled auto-reconnect by default
    m_MaxRetryAttempts = 5      ' up to 5 attempts with exponential backoff
    m_AutoReconnect = True      ' allow recovery

    ' Moderate buffer growth threshold (3 consecutive full reads)
    m_GrowThresholdCount = 3

    ' Request ID seed (first issued id = CMD_ID_FLOOR)
    m_nextRequestId = CMD_ID_FLOOR - 1
    ' Heartbeat id seed (first issued ping id = SYSTEM_PING_ID_MIN)
    m_nextSystemPingId = SYSTEM_PING_ID_MIN - 1

    SetState StateDisconnected
End Sub

Private Sub Class_Terminate()
' Ensure all handles, streams, and semaphores are cleaned up.
    ' If the macro was forcefully stopped (Stop button, End statement, Ctrl+Break)
    ' during an active SendAndReceive or ReadFromWire cycle, these flags remain True.
    ' Without resetting them here, any subsequent reuse of the object (or a new instance
    ' that inherits the module-level state in edge cases) would deadlock permanently.
    m_isSendReceiving = False
    m_isReading = False
    m_isReadingSince = 0
    
    Disconnect
    If Not m_InStream Is Nothing Then
        If m_InStream.State = adStateOpen Then m_InStream.Close
        Set m_InStream = Nothing
    End If
    If Not m_OutStream Is Nothing Then
        If m_OutStream.State = adStateOpen Then m_OutStream.Close
        Set m_OutStream = Nothing
    End If
    Set m_RegExp = Nothing
    Set m_EventBuffer = Nothing
End Sub

' Resets the target stream state and sets the required type without cross-talk
Private Sub PrepareStream(ByVal strm As ADODB.Stream, ByVal sType As StreamTypeEnum)
    If strm.State = adStateClosed Then strm.Open
    strm.Position = 0
    strm.SetEOS ' Truncate existing data
    strm.Type = sType
End Sub

' Helper to safely release stream resources
Private Sub CloseStream(ByRef strm As ADODB.Stream)
    If Not strm Is Nothing Then
        If strm.State = adStateOpen Then strm.Close
        Set strm = Nothing
    End If
End Sub

' Public Properties
Public Property Get IsConnected() As Boolean: IsConnected = (hWebSocket_ <> 0): End Property
Public Property Let PollingInterval(ByVal v As Long): m_PollingInterval = v: End Property
Public Property Let AllowUiPump(ByVal v As Boolean): m_AllowUiPump = v: End Property
Public Property Get AllowUiPump() As Boolean: AllowUiPump = m_AllowUiPump: End Property
Public Property Get State() As BiDiSocketState: State = m_State: End Property
Public Property Let AutoReconnect(ByVal v As Boolean): m_AutoReconnect = v: End Property
Public Property Get AutoReconnect() As Boolean: AutoReconnect = m_AutoReconnect: End Property
Public Property Get ConsecutiveInternalErrors() As Long: ConsecutiveInternalErrors = m_consecutiveInternalErrors: End Property
' MaxIncomingSize: receive-side assembly cap in bytes. Floor guard: values below
' MAX_MSG_SIZE would discard legitimate command responses, so they are clamped.
Public Property Let MaxIncomingSize(ByVal v As Long)
    If v < MAX_MSG_SIZE Then v = MAX_MSG_SIZE
    m_MaxIncomingSize = v
End Property
Public Property Get MaxIncomingSize() As Long: MaxIncomingSize = m_MaxIncomingSize: End Property
Public Property Get DroppedEventsCount() As Long: DroppedEventsCount = m_DroppedEventsCount: End Property

Private Sub SetState(ByVal NewState As BiDiSocketState)
    If m_State = NewState Then Exit Sub
    
    m_State = NewState
    
    Dim stateName As String
    Select Case m_State
        Case StateDisconnected: stateName = "DISCONNECTED"
        Case StateConnecting:   stateName = "CONNECTING"
        Case StateConnected:    stateName = "CONNECTED"
        Case StateDisconnecting: stateName = "DISCONNECTING"
        Case StateError:         stateName = "ERROR"
    End Select
    
    If DEBUG_MODE Then Debug.Print "BiDi State Changed: " & stateName
    RaiseEvent StateChanged(m_State, stateName)
End Sub

' Establishes a WebSocket connection to the WebDriver BiDi endpoint.
' Utilizes a state machine to prevent redundant connection attempts and manages lifecycle.
Public Function Connect(ByVal wsUrl As String, ByVal localPort As Long) As Boolean
    ' Guard: Prevent connection attempts if already connected or in progress
    If m_State = StateConnecting Or m_State = StateConnected Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect aborted. Current state is " & m_State
        Exit Function
    End If
    
    ' --- SECURITY ENFORCEMENT: STRICT LOOPBACK VALIDATION ---
    ' [INTENT] Ensure the WebSocket connection is strictly bound to the local machine.
    m_RegExp.pattern = "^wss?://(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?(?:/|$)"
    m_RegExp.IgnoreCase = True
    
    If Not m_RegExp.Test(wsUrl) Then
        Trace "[Security] Blocked: Connection attempt to non-local address. " & _
              "BiDi communication is restricted to localhost for safety.", True
        SetState StateError
        Exit Function
    End If
    
    ' Store connection parameters for potential recovery attempts
    m_LastWsUrl = wsUrl
    m_LastPort = localPort

    On Error GoTo ErrorHandler

    SetState StateConnecting
    
    ' Ensure any residual handles are cleaned up before a new attempt
    Call InternalCleanup
    
    If wsUrl = "" Or localPort = 0 Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect failed. Invalid URL or Port."
        GoTo ErrorHandler
    End If
    
    Dim Path As String: Path = ExtractPath(wsUrl)
    
    ' WHY: dynamically extract the host to match the actual binding
    Dim host As String: host = ExtractHost(wsUrl)

    hSession_ = WinHttpOpen(StrPtr(AGENT_NAME), WINHTTP_ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If hSession_ = 0 Then Err.Raise 513, "BiDiSocket", "WinHttpOpen failed"

    ' --- RATIONALE: UI RESPONSIVENESS vs CPU BALANCE ---
    ' Set connection-related timeouts (Resolve, Connect, Send, Receive)
    WinHttpSetTimeouts hSession_, 2000, 2000, 5000, 100
    
    hConnection_ = WinHttpConnect(hSession_, StrPtr(host), localPort, 0)
    If hConnection_ = 0 Then Err.Raise 514, "BiDiSocket", "WinHttpConnect failed"

    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(Path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    ' Configure the request for WebSocket upgrade
    Dim nullPtr As LongPtr: nullPtr = 0
    If WinHttpSetOption(hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, ByVal nullPtr, 0) = 0 Then
        Err.Raise 516, "BiDiSocket", "Failed to set WebSocket upgrade option"
    End If

    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    ' Verify Status 101 (Switching Protocols)
    Dim status As Long, Size As Long: Size = 4
    Dim idx As Long: idx = WINHTTP_NO_HEADER_INDEX
    Dim headerRet As Long
    
    headerRet = WinHttpQueryHeaders(hRequest_, _
                        WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, _
                        WINHTTP_HEADER_NAME_BY_INDEX, _
                        status, Size, idx)
    
    If headerRet <> 0 And status = 101 Then
        hWebSocket_ = WinHttpWebSocketCompleteUpgrade(hRequest_, 0)

        ' INVARIANT (HANDLE OWNERSHIP): the request handle and the WebSocket handle are independent.
        ' If CompleteUpgrade fails, close the request handle and abort.
        If hWebSocket_ = 0 Then
            On Error Resume Next
            WinHttpCloseHandle hRequest_
            On Error GoTo ErrorHandler
            hRequest_ = 0
            Err.Raise 520, "BiDiSocket", "WebSocket upgrade failed"
        End If

        ' INVARIANT (HANDLE OWNERSHIP): on success the request handle is no longer needed and must be
        ' closed (per WinHTTP docs/official sample); the WebSocket handle stays valid.
        ' Not closing it would leak one request handle per connection.
        On Error Resume Next
        WinHttpCloseHandle hRequest_
        On Error GoTo ErrorHandler
        hRequest_ = 0

        ' Configure Keep-Alive to maintain connection
        Dim ka As Long: ka = KEEPALIVE_INTERVAL_MS
        WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, ka, 4
        
        m_DroppedEventsCount = 0
        
        m_lastPollTime = 0
        m_RetryCount = 0
        m_consecutiveSoftTimeouts = 0 ' fresh connection: reset silence tracking
        SetState StateConnected
        Connect = True
    Else
        Err.Raise 519, "BiDiSocket", "Upgrade rejected. API Result: " & headerRet & " Status: " & status
    End If
    Exit Function

ErrorHandler:
    ' On any failure, transition to Error state and perform handle cleanup
    SetState StateError
    Call InternalCleanup
    Connect = False
End Function

' Sends a raw UTF-8 string message through the WebSocket.
' allowRecovery:=False makes this a best-effort send: on StateError or send failure
' it exits/returns False WITHOUT invoking AttemptRecovery (no backoff wait). Used by
' PostScript so fire-and-forget can never block teardown or HUD updates.
Public Function SendMessage(ByVal Msg As String, Optional ByVal allowRecovery As Boolean = True) As Boolean
    If LenB(Msg) = 0 Then Exit Function
    
    ' NEW: Log every send attempt before it happens
    Trace ">>> [Socket] Attempting to send message..."

    If m_State = StateError Then
        If Not allowRecovery Then Exit Function
        If Not AttemptRecovery() Then Exit Function
    End If

    If m_State <> StateConnected Then
        Trace ">>> [Socket] Send aborted. Not connected.", True
        Exit Function
    End If

    Dim data() As Byte: data = ToUtf8Bytes(Msg)
    
    Dim dataSize As Long
    On Error Resume Next
    dataSize = UBound(data) - LBound(data) + 1
    If Err.Number <> 0 Then dataSize = 0
    Err.Clear
    On Error GoTo 0
    
    If dataSize <= 0 Then
        Trace ">>> [Socket] Empty payload after UTF-8 conversion. Aborting.", True
        Exit Function
    End If
    
    Dim ret As Long
    ret = WinHttpWebSocketSend(hWebSocket_, _
                               WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                               data(0), _
                               dataSize)
    
    If ret = NO_ERROR Then
        SendMessage = True
        Exit Function
    End If

    Trace ">>> [Socket] SendMessage failed. Error: " & ret, True
    SetState StateError
    Disconnect keepErrorState:=True
    
    ' INVARIANT (NO RESEND): protocol compliance
    ' W3C WebDriver BiDi request IDs must remain unique within the session.
    ' Resending an identical JSON payload after a socket drop may reuse an ID
    ' that the peer is still tracking, so we never replay it on the recovered
    ' connection.
    ' We heal the connection here, but strictly return False so the upper layer
    ' handles the failure (e.g., re-generating the ID or raising an exception).
    ' Best-effort sends (allowRecovery:=False) skip healing entirely.
    If allowRecovery Then Call AttemptRecovery
    SendMessage = False
End Function

' PostScript: Fire & Forget transport. Sends a message without expecting a response.
' Best-effort: never triggers recovery/backoff (see SendMessage allowRecovery).
Public Sub PostScript(ByVal Msg As String)
    SendMessage Msg, False
End Sub

' FUNCTION: GetNextRequestId
' SUMMARY: Generates a unique sequential Request ID for BiDi commands.
'          Migrated from Wrapper to consolidate protocol-level concerns.
'          VBA is single-threaded and SendAndReceive is synchronous, so the sequential
'          counter alone guarantees uniqueness within any realistic session lifetime.
Public Function GetNextRequestId() As Long
    m_nextRequestId = m_nextRequestId + 1
    ' Wrap within the command band [CMD_ID_FLOOR .. ID_MAX]; never re-enter reserved 1..99
    If m_nextRequestId > ID_MAX Then m_nextRequestId = CMD_ID_FLOOR
    GetNextRequestId = m_nextRequestId
End Function

' Next heartbeat id, rotating within [SYSTEM_PING_ID_MIN .. SYSTEM_PING_ID_MAX].
' Rotation keeps every outstanding ping id unique (W3C unique-id compliance).
Private Function GetNextSystemPingId() As Long
    m_nextSystemPingId = m_nextSystemPingId + 1
    If m_nextSystemPingId > SYSTEM_PING_ID_MAX Then m_nextSystemPingId = SYSTEM_PING_ID_MIN
    GetNextSystemPingId = m_nextSystemPingId
End Function

' True if idStr is a numeric root id inside the reserved system band [1 .. CMD_ID_FLOOR).
' BiDi events carry no id (idStr = "") and are NOT system traffic.
Private Function IsReservedSystemId(ByVal idStr As String) As Boolean
    If Len(idStr) = 0 Then Exit Function
    If Not IsNumeric(idStr) Then Exit Function
    Dim v As Double: v = val(idStr)
    IsReservedSystemId = (v >= 1 And v < CMD_ID_FLOOR)
End Function

' FUNCTION: SendCommand
' SUMMARY: High-level protocol command sender. Constructs the BiDi JSON envelope
'          (id + method + params), validates payload size, serializes, and delegates
'          to SendAndReceive for synchronous response retrieval.
'
' NOTE ON RESPONSIBILITY BOUNDARY:
'   This method owns: ID generation, payload construction, size validation, serialization.
'   It does NOT own: retry logic, context recovery, or error classification.
'   Those remain in Wrapper's ExecuteBiDiCommand to preserve the existing recovery paths.
'
' PARAMETERS:
'   method   - BiDi method name (e.g. "browsingContext.navigate")
'   params   - Dictionary of method parameters
'   Silent   - Suppresses debug logging when True
' RETURNS: Raw JSON response string from the browser.
' RAISES:  ERR_ACTION_FAILED (vbObjectError+1003) if payload exceeds MAX_MSG_SIZE.
'          Propagates errors from SendAndReceive (600-604).
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
    
    ' Release the dictionary immediately (memory hygiene)
    Set requestDic = Nothing
    
    ' Payload Size Guard (prevents VBA hang from oversized payloads)
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

' Properly closes the WebSocket and all associated WinHttp handles.
Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    ' If we are already in the middle of a connection attempt,
    ' we should be cautious about forced disconnection unless it's a hard reset.
    If m_State = StateDisconnecting Or m_State = StateDisconnected Then Exit Sub

    ' [STEP 1] Log the final count of dropped events to bidi_socket_trace.log.
    ' This acts as a "Post-Mortem" report for debugging high-load scenarios.
    ' We do not reset the counter here so it remains accessible via property after disconnect.
    If m_DroppedEventsCount > 0 Then
        Trace ">>> [Session Summary] Connection closing. Total events dropped during this session: " & m_DroppedEventsCount, True
    End If

    If Not keepErrorState Then SetState StateDisconnecting
    
    ' Perform the actual handle cleanup
    Call InternalCleanup
    
    ' Clear the event buffer
    Set m_EventBuffer = New Collection
    
    ' Finalize state only if we are not preserving an Error state for recovery
    If Not keepErrorState Then
        SetState StateDisconnected
    End If
End Sub

' Private helper to release WinHttp handles.
' STRICT RULE: This method MUST NOT modify m_State.
Private Sub InternalCleanup()
    On Error Resume Next
    
    If hWebSocket_ <> 0 Then
        ' Send a close frame (Status 1000) WITHOUT waiting for the peer's reply.
        ' DANGER: WinHttpWebSocketClose performs the FULL close handshake (send +
        ' wait for the peer's close frame) and blocks indefinitely against a
        ' half-dead peer (e.g. suspended driver: TCP still ACKs, no frames ever
        ' arrive). WinHttpWebSocketShutdown is the send-only half; the handle
        ' release below tears the connection down regardless of the peer.
        ' (ForceClose already avoided Close for the same reason.)
        WinHttpWebSocketShutdown hWebSocket_, 1000, ByVal 0&, 0
        WinHttpCloseHandle hWebSocket_
        hWebSocket_ = 0
    End If
    
    If hRequest_ <> 0 Then
        ' INVARIANT (HANDLE OWNERSHIP): the request handle is independent of the WebSocket handle.
        ' On the success path it was already closed in Connect (so this is skipped);
        ' here it only releases a handle left open by a failed/partial upgrade.
        WinHttpCloseHandle hRequest_
        hRequest_ = 0
    End If
    
    If hConnection_ <> 0 Then
        WinHttpCloseHandle hConnection_
        hConnection_ = 0
    End If
    
    If hSession_ <> 0 Then
        WinHttpCloseHandle hSession_
        hSession_ = 0
    End If
    
    ' Purge the input stream to prevent cross-session data corruption.
    ' If a connection drops mid-fragment, the residual bytes must be truncated
    ' before the next socket connection is established.
    If Not m_InStream Is Nothing Then
        If m_InStream.State = adStateOpen Then
            m_InStream.Position = 0
            m_InStream.SetEOS
            m_InStream.Type = adTypeBinary
        End If
    End If
    
    ' Reset module-scoped semaphores and latches.
    m_isRecovering = False
    m_discardActive = False ' input stream truncated above: drop any partial-discard latch
    
    On Error GoTo 0
End Sub

' Sends a JSON message and waits for a specific response ID with a strict hard deadline.
' Refactored to establish a "High-Purity Loop" that prevents WinHttp state corruption
' by suppressing DoEvents during the active command-response cycle.
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    ' Re-entry Guard: Lock the communication channel.
    ' Raise (instead of returning "") so the caller cannot mistake a blocked
    ' re-entry for an empty, successful response. This Raise happens BEFORE
    ' "On Error GoTo Cleanup" is armed, so it propagates straight to the caller
    ' and does NOT run Cleanup -- the outer in-flight call keeps the lock.
    If m_isSendReceiving Then
        Err.Raise 606, "BiDiSocket", _
            "SendAndReceive re-entry blocked (a command was issued while another was still in flight)."
    End If
    m_isSendReceiving = True
    
    Dim succeeded As Boolean: succeeded = False
    Dim Msg As String, inId As String
    Dim cmdId As String
    
    On Error GoTo Cleanup

    If m_State <> StateConnected Then
        If Not AttemptRecovery() Then
            Err.Raise 600, "BiDiSocket", "Not connected and recovery failed."
        End If
    End If
    
    cmdId = ExtractJsonId(jsonMessage)
    
    If Not SendMessage(jsonMessage) Then
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost during send operation."
        Else
            Err.Raise 601, "BiDiSocket", "Send failed due to internal WinHttp error."
        End If
    End If
    
    ' WHY: two-tier timeout strategy distinguishes transient browser slowness from
    '       sustained socket silence without immediately destroying handles.
    '
    '       Soft Timeout (1x): no matching response within the normal window. Raises
    '           604 but PRESERVES the connection; the caller (ExecuteBiDiCommand)
    '           retries as a NEW SendAndReceive call with a new id and fresh deadlines.
    '           Each soft timeout on the same connection increments
    '           m_consecutiveSoftTimeouts; reaching MAX_CONSECUTIVE_SOFT_TIMEOUTS is
    '           judged sustained silence (half-dead socket that still accepts sends)
    '           and escalates to StateError. The counter resets on any successfully
    '           matched response and on a fresh Connect.
    '       Hard Timeout (2x): reachable only WITHIN a single call, when ReadFromWire
    '           (whose internal deadline is hardDeadline) stays in fragment assembly
    '           past twice the window. Transitions to StateError immediately.
    '       On StateError, handles are cleaned up on exit ONLY when AutoReconnect
    '           is enabled (see Cleanup); otherwise they persist until Disconnect.
    '       LIMITATION (empirically verified): both tiers require ReadFromWire to
    '           regain control, i.e. frames must still arrive. A peer that is
    '           TCP-alive but sends nothing (suspended or deadlocked driver: the
    '           kernel keeps ACKing) blocks WinHttpWebSocketReceive indefinitely;
    '           NO WinHTTP timeout option applies to synchronous WebSocket
    '           receives. A KILLED peer does unblock it (TCP reset -> receive
    '           error). Escaping a silent-but-alive peer would require async
    '           WinHTTP or a second thread, neither available to VBA.
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim softDeadline As Double: softDeadline = NowMs() + timeout
    Dim hardDeadline As Double: hardDeadline = NowMs() + (timeout * 2)
    
    ' Memory Protection Constants
    ' [PATCH] Raised from 1000. Because the drain fix below keeps the matching
    ' response from being buried, a deeper buffer can now safely retain events
    ' that arrive during a long-held command (instead of dropping them, which
    ' would lose recorder log entries). ~20k small log/network events is a few
    ' MB at most; lower it if memory is tight, raise it for very chatty pages.
    Const MAX_EVENT_BUFFER_SIZE As Long = 20000
    
    ' Receive Loop (High-Purity)
    Do
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        Msg = ReadFromWire(hardDeadline)
        
        If Len(Msg) > 0 Then
            inId = ExtractJsonId(Msg)
            If inId = cmdId Then
                SendAndReceive = Msg
                succeeded = True
                GoTo Cleanup
            Else
                ' Async events or pongs: Queue for GetMessage processing
                ' BiDi events (network.beforeRequestSent etc.) have no "id" field,
                ' so the old condition (inId <> "") silently discarded them.
                ' Losing these one-shot events degraded recorder/[REQ] completeness.
                ' Discard reserved-band system replies (rotating heartbeat pongs,
                ' PostScript reply) via a single band check.
                If Not IsReservedSystemId(inId) Then
                    m_EventBuffer.Add Msg
                    
                    ' FIFO Protection to prevent memory exhaustion
                    If m_EventBuffer.count > MAX_EVENT_BUFFER_SIZE Then
                        m_EventBuffer.Remove 1
                        m_DroppedEventsCount = m_DroppedEventsCount + 1
                    End If
                End If
            End If
        End If
        
        Dim nowT As Double: nowT = NowMs()
        
        ' HARD Timeout: Sustained silence beyond 2x the normal timeout.
        ' This indicates a genuine connection problem, not just browser slowness.
        If nowT > hardDeadline Then
            SetState StateError
            Err.Raise 604, "BiDiSocket", "Socket Silence: Hard timeout (connection presumed lost) waiting for ID: " & cmdId
        End If
        
        ' SOFT Timeout: Response not received within the normal timeout window.
        ' First occurrence preserves the connection for a caller retry; consecutive
        ' occurrences on the same connection indicate a half-dead socket (sends
        ' still succeed, nothing ever matches) and escalate to StateError.
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

        ' [PATCH] Drain at full speed while messages are flowing; only sleep when
        ' the wire produced nothing this iteration. ReadFromWire blocks inside
        ' WinHttpWebSocketReceive when idle, so this never busy-spins. This keeps
        ' the m_EventBuffer from overflowing under event floods and prevents the
        ' matching response from being buried behind queued events.
        If Len(Msg) = 0 Then YieldExecution 5, False
    Loop

Cleanup:
    ' WHY: save Err info immediately to prevent overwrite during InternalCleanup.
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrDesc As String: savedErrDesc = Err.description
    Dim savedErrSrc As String: savedErrSrc = Err.source
    
    If succeeded Then
        savedErrNum = 0
        Err.Clear
        ' A matched response proves the connection is alive: reset silence tracking.
        m_consecutiveSoftTimeouts = 0
    End If
    
    m_isSendReceiving = False
    
    ' WHY: handle cleanup is conditional on the two-tier timeout strategy:
    '   - Soft timeout (604 without StateError): Connection is preserved.
    '     The caller (ExecuteBiDiCommand) can retry with the same socket.
    '   - Hard timeout (604 with StateError): Connection is presumed dead.
    '     Clear stale handles so the next attempt triggers a full reconnect.
    If m_State = StateError And m_AutoReconnect Then
        Call InternalCleanup
    End If
    
    ' Explicit cleanup of potentially large strings
    Msg = ""
    inId = ""
    
    ' Error Propagation
    If Not succeeded And savedErrNum <> 0 Then
        On Error GoTo 0
        Err.Raise savedErrNum, savedErrSrc, savedErrDesc
    End If
End Function

' Internal recovery logic with re-entrancy protection.
Private Function AttemptRecovery() As Boolean
    ' Re-entrancy guard: Acquire the lock IMMEDIATELY.
    ' Moving this to the very top ensures atomic-like locking before any other
    ' evaluations or error handler setups occur.
    If m_isRecovering Then Exit Function
    m_isRecovering = True
    
    ' WHY: route all runtime errors to the ErrorHandler to guarantee flag reset
    On Error GoTo ErrorHandler

    ' Pre-condition checks now route to CleanExit to properly release the lock
    If Not m_AutoReconnect Then GoTo CleanExit
    If m_RetryCount >= m_MaxRetryAttempts Then
        Trace ">>> [Recovery] Max retry attempts reached.", True
        GoTo CleanExit
    End If

    m_RetryCount = m_RetryCount + 1
    
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    Trace ">>> [Recovery] Attempt #" & m_RetryCount & " (Waiting " & waitTime & "ms)", True
    
    ' During YieldExecution, DoEvents is pumped (allowDoEvents = True).
    ' External COM events or Timers could attempt re-entrancy here.
    ' Because m_isRecovering is already True, they will be safely rejected.
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
    ' [CRITICAL] Always release the recovery lock regardless of success, failure, or runtime errors
    m_isRecovering = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    ' INVARIANT (ERROR HANDLER): prohibit Trace (File I/O) within the error handler.
    ' If File I/O fails (e.g., file lock, permissions), a new exception occurs here,
    ' preventing Resume CleanExit from executing and leaving m_isRecovering=True forever.
    ' Debug.Print is non-throwing and safe in this context.
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [Recovery] Critical Error during attempt: " & Err.description
    
    AttemptRecovery = False
    ' Clear the error state and jump to the cleanup block to release the lock
    Resume CleanExit
End Function

' Polling-based BiDi message pump; processes queued events and manages heartbeat.
' Optimized to prevent Excel UI freezing by using a micro-deadline for socket I/O.
Public Function GetMessage() As String
    ' Block heartbeat/ping while a command response is pending in SendAndReceive
    If m_isSendReceiving Then Exit Function
    
    ' Serve from local event buffer first
    If m_EventBuffer.count > 0 Then
        GetMessage = m_EventBuffer(1): m_EventBuffer.Remove 1: Exit Function
    End If

    ' Autonomous recovery and connectivity guard
    If m_State = StateError Then If Not AttemptRecovery() Then Exit Function
    If m_State <> StateConnected Then Exit Function

    ' Enforce polling cadence (100ms) to prevent CPU spikes
    Dim nowT As Double: nowT = NowMs()
    If (nowT - m_lastPollTime) < m_PollingInterval Then Exit Function
    m_lastPollTime = nowT

    ' [SMART HEARTBEAT]
    ' Send a silent keep-alive ping every poll cycle to hold the pipe open at
    ' 100ms cadence. The pong is swallowed below; no console/file logging.
    Dim pingId As Long: pingId = GetNextSystemPingId()
    Dim pingCmd As String: pingCmd = "{""id"":" & pingId & ", ""method"":""session.status"", ""params"":{}}"

    If Not SendMessage(pingCmd) Then Exit Function

    Dim rawMsg As String: rawMsg = ReadFromWire(nowT + 10)

    If LenB(rawMsg) = 0 Then
        YieldExecution 1, m_AllowUiPump: Exit Function
    End If

    Dim inId As String: inId = ExtractJsonId(rawMsg)
    If IsReservedSystemId(inId) Then
        ' Reserved band: heartbeat pong (rotating [3..99]) or PostScript reply (2).
        ' All are swallowed so they never reach the recorder/event stream.
        If inId <> CStr(POST_SCRIPT_REPLY_ID) Then
            ' Heartbeat pong: connection proven healthy -> clear any retry backoff.
            If m_RetryCount > 0 Then m_RetryCount = 0
        End If
        GetMessage = ""
    Else
        ' Return the actual BiDi message (Event/Response)
        GetMessage = rawMsg
    End If
End Function

' FUNCTION: GetParsedEvent
' SUMMARY: Higher-level event retrieval that combines GetMessage + validation + JSON parse.
'          Returns a parsed Dictionary object ready for dispatch, or Nothing if no
'          valid event is available.
'
' This consolidates the "dequeue -> validate -> parse" pipeline that was previously
' duplicated in Wrapper's ProcessRecorderQueue, reducing the Wrapper's responsibility
' to pure dispatch logic.
'
' PARAMETERS:
'   minMsgSize  - Minimum message length to accept (filters noise). Default: 10
'   maxMsgSize  - Maximum message length to accept (prevents VBA hang). Default: 300000
' RETURNS: Parsed Dictionary object, or Nothing if buffer is empty / message invalid.
Public Function GetParsedEvent(Optional ByVal minMsgSize As Long = 10, _
                               Optional ByVal maxMsgSize As Long = 300000) As Object
    Dim rawMsg As String: rawMsg = GetMessage()
    If Len(rawMsg) = 0 Then Exit Function
    
    If Len(rawMsg) < minMsgSize Or Len(rawMsg) > maxMsgSize Then Exit Function
    
    ' Structural pre-check (must contain at least a method or id key)
    If InStr(rawMsg, """method"":") = 0 And InStr(rawMsg, """id"":") = 0 Then Exit Function
    
    ' JSON Parse (under OERN; returns Nothing on malformed input)
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

' Reads a WebSocket frame.
' Refactored to include a Liveness Check (deadlock detection) and strict memory management.
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function
    
    ' Liveness Check (Deadlock Detection)
    If m_isReading Then
        If (NowMs() - m_isReadingSince) > 5000 Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] BiDi Warning: Read lock stuck. Force-resetting."
            m_isReading = False
        Else
            Exit Function
        End If
    End If

    On Error GoTo ErrorHandler

    m_isReading = True
    m_isReadingSince = NowMs()

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim isMessageFinished As Boolean
    Dim hasValidData As Boolean
    ' Counter to track fragment throughput
    Dim fragmentCount As Long: fragmentCount = 0
    ' Receive-side cap bookkeeping. Discard state is held in the module-level
    ' m_discardActive latch (NOT call-local): a 12002/deadline exit can interrupt
    ' an oversized message mid-discard, and the latch keeps discard engaged across
    ' calls until the MESSAGE boundary truncates m_InStream -- mirroring how a normal
    ' partial message persists via the retained stream + non-zero Position. The latch
    ' is cleared ONLY where m_InStream is truncated (CleanExit finished block,
    ' StreamErrorHandler, InternalCleanup, ForceClose), so 'empty stream yet still
    ' discarding' cannot occur. totalBytes stays call-local (current call only).
    Dim totalBytes As Long

    If m_InStream.Position = 0 Then
        PrepareStream m_InStream, adTypeBinary
    End If
    
    Do
        ' Receive data from the WebSocket
        ret = WinHttpWebSocketReceive(hWebSocket_, _
                                      m_ReceiveBuffer(0), _
                                      m_BufferSize, _
                                      bytesRead, _
                                      bufType)
        
        If ret <> NO_ERROR Then
            If ret <> 12002 Then
                Trace ">>> [WinHttp] Receive Failed. Code: " & ret, True
                SetState StateError
                Disconnect keepErrorState:=True
            End If
            GoTo CleanExit
        End If

        If bytesRead > 0 Then
            hasValidData = True
            totalBytes = totalBytes + bytesRead

            If Not m_discardActive Then
                If m_InStream.Size + bytesRead > m_MaxIncomingSize Then
                    ' Cap exceeded: stop writing, keep draining fragments until the
                    ' MESSAGE frame closes the boundary. The retained head is used
                    ' for identity diagnostics in CleanExit.
                    m_discardActive = True
                    Trace ">>> [WinHttp] Incoming message exceeded MaxIncomingSize (" & _
                          m_MaxIncomingSize & "). Discard mode engaged.", True
                Else
                    If bytesRead = m_BufferSize Then
                        m_InStream.Write m_ReceiveBuffer
                    Else
                        Dim chunk() As Byte
                        chunk = m_ReceiveBuffer
                        ReDim Preserve chunk(bytesRead - 1)
                        m_InStream.Write chunk
                        Erase chunk
                    End If
                End If
            End If

            MaybeGrowBuffer bytesRead
        End If

        ' Frame Assembly Logic
        Select Case bufType
            Case WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE
                isMessageFinished = True
                GoTo CleanExit

            Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
                ' Server initiated close. Handle is now invalid.
                Debug.Print "[BiDi] WebSocket Close frame received from server."
                SetState StateError
                Disconnect keepErrorState:=True
                GoTo CleanExit
                
            Case WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE
                ' WHY: prevent Excel freezing during high-volume fragment assembly
                fragmentCount = fragmentCount + 1
                If fragmentCount Mod 20 = 0 Then
                    ' Yield to the OS message pump only if allowed by class settings
                    If m_AllowUiPump And Not m_isSendReceiving Then DoEvents
                    ' WHY: do not reset the counter here. The Mod operator cycles every 20
                    ' iterations on its own; resetting to 0 would fire the next DoEvents
                    ' after 21 fragments instead of 20 (off-by-one).
                End If
                
            Case Else
                GoTo CleanExit
        End Select
        
        ' Overall Deadline Check
        If deadline > 0 And NowMs() > deadline Then
            Trace ">>> [WinHttp] Read loop hit internal deadline.", True
            GoTo CleanExit
        End If

        ' NO SLEEP HERE (deliberate). This is a synchronous WinHttp session
        ' (WinHttpOpen dwFlags=0): WinHttpWebSocketReceive blocks internally until
        ' data arrives or the 100ms receive timeout fires (12002 -> exit above),
        ' so Receive itself paces the loop. A former "YieldExecution 1, False"
        ' here cost ~15.6ms per fragment (kernel32 Sleep(1) rounds up to the
        ' system timer quantum; no timeBeginPeriod in this process), turning a
        ' 12.6MB / 8KB = ~1,540-fragment assembly into a ~24s UI freeze.
        ' UI responsiveness is handled by the Mod-20 DoEvents guard above.
    Loop

CleanExit:
    On Error GoTo StreamErrorHandler

    If isMessageFinished And hasValidData Then
        m_InStream.Position = 0
        m_InStream.Type = adTypeText
        m_InStream.Charset = "UTF-8"

        If m_discardActive Then
            ' IDENTITY DIAGNOSTICS for oversized messages:
            '   HEAD300           -> shows the leading JSON ("method"/"id") of the giant
            '   fusionMarkers(}{) -> 0  = genuine single message
            '                        >0 = multiple messages fused into one frame
            Dim headText As String: headText = m_InStream.ReadText
            Dim fusionCount As Long
            fusionCount = (Len(headText) - Len(Replace(headText, "}{", ""))) \ 2
            Trace ">>> [WinHttp] OVERSIZED MESSAGE DISCARDED. thisCallBytes=" & totalBytes & _
                  " retainedChars=" & Len(headText) & _
                  " fusionMarkers(}{)=" & fusionCount & _
                  " HEAD300=[" & Left$(headText, 300) & "]", True
            ReadFromWire = ""
        Else
            ReadFromWire = m_InStream.ReadText

            Dim isSuspicious As Boolean
            isSuspicious = (InStr(1, ReadFromWire, """error"":", vbTextCompare) > 0) Or _
                           (InStr(1, ReadFromWire, "chrome-error", vbTextCompare) > 0)

            If isSuspicious Then m_consecutiveInternalErrors = m_consecutiveInternalErrors + 1
            Trace ">>> [WinHttp] Message assembled. Length: " & Len(ReadFromWire), isSuspicious
        End If

        PrepareStream m_InStream, adTypeBinary
        m_discardActive = False ' MESSAGE boundary reached: stream truncated, latch cleared
    End If
    
    GoTo FinalCleanup

StreamErrorHandler:
    On Error Resume Next
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [WinHttp] Stream Error: " & Err.description
    CloseStream m_InStream
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open
    PrepareStream m_InStream, adTypeBinary
    m_discardActive = False ' stream rebuilt empty: no partial/discard data survives
    Resume FinalCleanup

FinalCleanup:
    m_isReading = False
    m_isReadingSince = 0
    On Error GoTo 0
    Exit Function

ErrorHandler:
    m_isReading = False
    m_isReadingSince = 0
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [WinHttp] Critical Read Error: " & Err.description
    Resume CleanExit
End Function

' Helpers & Utils
' Unified Timer API (QPC-based, monotonic)
' NowTicks   : Opaque Currency snapshot (raw QPC counter). Use as a token for ElapsedMs.
' ElapsedMs  : Difference between a NowTicks snapshot and the current counter, in ms.
' NowMs      : Absolute monotonic timestamp in ms (since QPC origin).
' Note: m_qpcFreq is initialized in Class_Initialize with fail-fast on missing QPC support.
Public Function NowTicks() As Currency
    QueryPerformanceCounter NowTicks
End Function

Public Function ElapsedMs(ByVal startTicks As Currency) As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    ElapsedMs = (CDbl(cur - startTicks) / CDbl(m_qpcFreq)) * 1000#
End Function

Public Function NowMs() As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    NowMs = (CDbl(cur) / CDbl(m_qpcFreq)) * 1000#
End Function

' Peak-holding buffer growth logic
Private Sub MaybeGrowBuffer(ByVal bytesRead As Long)
    If bytesRead >= m_BufferSize Then
        m_RecentFullReads = m_RecentFullReads + 1
        If m_RecentFullReads >= m_GrowThresholdCount Then
            Dim newSize As Long: newSize = m_BufferSize * 2
            If newSize > m_MaxBufferSize Then newSize = m_MaxBufferSize
            
            If newSize > m_BufferSize Then
                m_BufferSize = newSize
                m_RecentFullReads = 0
                
                ' Sync the actual byte array with the new buffer size
                ReDim m_ReceiveBuffer(m_BufferSize - 1)
                
                If DEBUG_MODE Then Debug.Print "BiDi: Buffer grown to " & m_BufferSize
            Else
                ' WHY: reset consecutive counter even when MAX_BUFFER_SIZE is reached
                ' to prevent infinite accumulation.
                m_RecentFullReads = 0
            End If
        End If
    Else
        ' WHY: reset consecutive counter when a partial read occurs
        m_RecentFullReads = 0
    End If
End Sub

' Standardized port retrieval with type-safety
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

' Extracts the root-level "id" value from a JSON string.
' Designed to be robust against variations in key order and large payloads.
Private Function ExtractJsonId(ByVal json As String) As String
    Dim s As String: s = LTrim$(json)
    If Len(s) < 10 Or Left$(s, 1) <> "{" Then Exit Function

    ' Fast pre-check: Avoid running the scanner on payloads without "id"
    If InStr(1, s, """id""", vbTextCompare) = 0 Then Exit Function

    Dim depth As Long: depth = 0
    Dim inString As Boolean: inString = False
    Dim isEscaped As Boolean: isEscaped = False
    Dim i As Long
    
    ' Map the string directly to a Byte array (UTF-16 LE) for blazing fast memory access
    Dim b() As Byte: b = s
    Dim bLen As Long: bLen = UBound(b)
    
    ' Constants for ASCII character codes
    Const CHAR_QUOTE As Byte = 34  ' "
    Const CHAR_LBRACE As Byte = 123 ' {
    Const CHAR_RBRACE As Byte = 125 ' }
    Const CHAR_SLASH As Byte = 92  ' \
    
    ' In VBA, strings are UTF-16, so each character is 2 bytes. We step by 2.
    For i = 0 To bLen Step 2
        ' [CRITICAL ASCII GUARD]
        ' VBA strings are UTF-16 LE. Each char is 2 bytes: low byte (b(i)) + high byte (b(i+1)).
        ' We MUST verify the high byte is 0 (=ASCII range) before evaluating b(i).
        ' Without this guard, any non-ASCII char whose low byte happens to match
        ' '{' (0x7B), '}' (0x7D), '"' (0x22), or '\' (0x5C) would corrupt depth/inString.
        ' Example: Many Japanese chars have low bytes that collide with these ASCII codes.
        ' Bounds check (i+1 <= bLen) is included to guarantee no array overrun on
        ' an oddly-sized buffer (defensive; VBA strings should always be even).
        If i + 1 > bLen Then Exit For
        If b(i + 1) <> 0 Then GoTo ContinueLoop

        ' Check if we are currently escaping a character
        If isEscaped Then
            isEscaped = False
            GoTo ContinueLoop
        End If

        Select Case b(i)
            Case CHAR_SLASH
                ' If inside a string, the next character is escaped
                If inString Then isEscaped = True
                
            Case CHAR_QUOTE
                inString = Not inString
                
                ' If we just OPENED a string at the root level, check if it's "id"
                If inString And depth = 1 Then
                    ' Ensure we don't go out of bounds (need 4 chars: i, d, ", :)
                    If i + 8 <= bLen Then
                        ' Check for 'i' (105) and 'd' (100) and '"' (34)
                        ' Also verify each char's high byte is 0 (ASCII range)
                        If b(i + 2) = 105 And b(i + 3) = 0 _
                           And b(i + 4) = 100 And b(i + 5) = 0 _
                           And b(i + 6) = 34 And b(i + 7) = 0 Then
                            
                            Dim j As Long: j = i + 8 ' Start right after "id"
                            Dim idVal As String
                            
                            ' Bounds safety: explicitly check both j and j+1 in every iteration.
                            ' This ensures no reliance on implicit mathematical safety (j <= bLen-1).
                            Do While j + 1 <= bLen
                                ' Validate that the character is strictly within ASCII range
                                ' (Upper byte must be 0 in UTF-16 LE).
                                If b(j + 1) = 0 Then
                                    Select Case b(j)
                                        Case 48 To 57, 45 ' '0'-'9', '-'
                                            idVal = idVal & Chr$(b(j))
                                        Case 32, 58, 9, 10, 13 ' spaces, colon, tabs, CRLF
                                            ' Stop parsing if we hit a delimiter AFTER numbers have started
                                            If Len(idVal) > 0 Then Exit Do
                                        Case Else
                                            ' Hit an invalid ASCII char (e.g., letters, quote, bracket)
                                            Exit Do
                                    End Select
                                Else
                                    ' Non-ASCII character encountered; terminate extraction immediately.
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
                    ' Optimization: If depth drops to 0, we have parsed the root object
                    If depth = 0 Then Exit Function
                End If
        End Select
        
ContinueLoop:
    Next i
End Function

Private Function ExtractPath(ByVal url As String) As String
    ' Explicit branch: IIf would evaluate both arms, so a path-less URL
    ' (e.g. "ws://localhost:12345", no "/" after the host) would still run
    ' Mid$(url, 0) and raise runtime error 5. Default to "/" instead.
    Dim p As Long
    p = InStr(1, url, "://")
    If p > 0 Then p = p + 3 Else p = 1

    p = InStr(p, url, "/")
    If p > 0 Then
        ExtractPath = Mid$(url, p)
    Else
        ExtractPath = "/"
    End If
End Function

' UTF-8 Conversion Utility
Private Function ToUtf8Bytes(ByVal txt As String) As Byte()
    ' Use the dedicated output stream to avoid corrupting m_InStream
    PrepareStream m_OutStream, adTypeText
    m_OutStream.Charset = "UTF-8"
    m_OutStream.WriteText txt
    
    m_OutStream.Position = 0
    m_OutStream.Type = adTypeBinary
    
    ' Dynamic BOM Detection (EF BB BF)
    If m_OutStream.Size >= 3 Then
        Dim Header() As Byte: Header = m_OutStream.Read(3)
        If Header(0) = &HEF And Header(1) = &HBB And Header(2) = &HBF Then
            m_OutStream.Position = 3 ' Skip BOM
        Else
            m_OutStream.Position = 0
        End If
    Else
        m_OutStream.Position = 0
    End If
    
    If m_OutStream.Size > m_OutStream.Position Then
        ToUtf8Bytes = m_OutStream.Read
    Else
        ToUtf8Bytes = StrConv("", vbFromUnicode)
    End If
End Function

' PumpMessages: Drains the Windows message queue (DoEvents) without sleeping.
'   throttleMs = 0 : pump unconditionally
'   throttleMs > 0 : pump only if at least throttleMs has elapsed since the last pump
' Skipped while m_isSendReceiving is True to prevent WinHttp re-entrancy.
' On the very first call (m_lastPumpTicks = 0), DoEvents always runs.
Public Sub PumpMessages(Optional ByVal throttleMs As Long = 0)
    If m_isSendReceiving Then Exit Sub
    If throttleMs > 0 And m_lastPumpTicks <> 0 Then
        If ElapsedMs(m_lastPumpTicks) < throttleMs Then Exit Sub
    End If
    DoEvents
    m_lastPumpTicks = NowTicks()
End Sub

' YieldExecution: Sleeps for the given duration, optionally pumping the Windows
' message queue. Companion to PumpMessages (which only pumps, no sleep).
'   ms <= 50  : direct SleepWinAPI for precision
'   ms > 50   : delegates to InterruptibleSleep for chunked, responsive waiting
' DoEvents fires only when allowDoEvents=True AND m_isSendReceiving=False
' (re-entrancy guard against WinHttp corruption during active I/O).
Public Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    ' Floor Guard: Prevent negative values from causing undefined SleepWinAPI behavior.
    ' This is the only addition; the optimized 50ms branching logic below is preserved as-is.
    Dim vMs As Long: vMs = ms
    If vMs < 0 Then vMs = 1

    ' 1. Use the robust chunked sleep for any duration longer than 50ms.
    ' This ensures that long waits (like in AttemptRecovery) do not freeze Excel.
    If vMs > 50 Then
        Call InterruptibleSleep(vMs, allowDoEvents)
    Else
        ' 2. For very short durations, perform a direct sleep to maintain precision.
        ' Even for short sleeps, we apply the m_isSendReceiving guard to prevent
        ' unintended re-entrancy during active command cycles.
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If

        If vMs > 0 Then SleepWinAPI vMs
    End If
End Sub

' Leaves breadcrumbs in both the external file and the Immediate Window.
' CRITICAL PASS-THROUGH: isCritical=True bypasses the DEBUG_MODE gate for the
' FILE log only (production incident monitoring, e.g. oversized-message
' diagnostics whose root cause is still unconfirmed). Immediate Window output
' and the substring-based detection (Return Code/failed/error) remain
' DEBUG_MODE-only, so production log volume is limited to explicitly
' flagged exceptional events.
Public Sub Trace(ByVal Msg As String, Optional ByVal isCritical As Boolean = False)
    If Not DEBUG_MODE And Not isCritical Then Exit Sub
    
    Dim logMsg As String: logMsg = "[" & Format(Now, "hh:mm:ss") & "] " & Msg
    
    ' Only print the message itself to the Immediate Window (debug builds only)
    If DEBUG_MODE Then Debug.Print logMsg
    
    ' Detection Logic
    ' WHY: the substring probes are gated on DEBUG_MODE so that production
    ' file logging is driven solely by the caller's explicit isCritical flag,
    ' not by incidental keywords in high-volume messages.
    Dim shouldLogToFile As Boolean
    shouldLogToFile = isCritical
    If DEBUG_MODE Then
        shouldLogToFile = shouldLogToFile Or _
                          (InStr(1, Msg, "Return Code", vbTextCompare) > 0) Or _
                          (InStr(1, Msg, "failed", vbTextCompare) > 0) Or _
                          (InStr(1, Msg, "error", vbTextCompare) > 0)
    End If

    If shouldLogToFile Then
        Dim localDir As String: localDir = WebShared.getFullLocalPath(ThisWorkbook.Path)
        If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
        If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
        
        Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
        Dim fNum As Integer: fNum = FreeFile
        
        On Error Resume Next
        Open logPath For Append As #fNum
        Print #fNum, logMsg & " (LOG_EVENT)"
        Close #fNum
        On Error GoTo 0
        ' [SILENT] Removed LOG_SUCCESS print for clean operation
    End If
End Sub

' LogEvent: Unified logging with QPC timestamps and chunked output for long strings.
Public Sub LogEvent(ByVal Header As String, ByVal Body As String, _
                   Optional ByVal tag As String = "[INFO]", _
                   Optional ByVal LogToFile As Boolean = False)
    Const CHUNK_SIZE As Long = 250
    Dim i As Long
    Dim cleanText As String
    Dim timePrefix As String
    Dim currentTick As Currency
    
    ' High-resolution timestamp for synchronization and performance analysis
    currentTick = Me.NowTicks()
    timePrefix = "[" & Format(Now, "hh:mm:ss") & " | " & Format(currentTick, "0") & "] "

    ' Normalize text: remove null characters and trailing whitespace/newlines
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

    ' Output the header with the structured tag.
    ' Format: [Time | Ticks] [TAG] Header (Length: N)
    Dim sfx As String: sfx = "": If Len(cleanText) > 0 Then sfx = " (Length: " & Len(cleanText) & ")"
    Dim headerLine As String: headerLine = timePrefix & tag & " " & Header & sfx
    Debug.Print headerLine

    ' Chunked output to bypass VBA's Immediate Window line length limits
    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            ' Print parts of the string without an automatic newline
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
        Debug.Print ""
    End If
    
    ' Optional file logging
    If LogToFile Then
        WriteLogToFile headerLine & vbCrLf & cleanText
    End If
End Sub

' Internal helper for LogToFile
Private Sub WriteLogToFile(ByVal Msg As String)
    Dim localDir As String
    On Error Resume Next
    localDir = WebShared.getFullLocalPath(ThisWorkbook.Path)
    If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
    If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
    
    Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
    Dim fNum As Integer: fNum = FreeFile
    
    Open logPath For Append As #fNum
    Print #fNum, Msg
    Close #fNum
    On Error GoTo 0
End Sub

' Explicitly resets the consecutive error counter.
' Called by the Wrapper only when a truly healthy response is confirmed.
Public Sub ClearInternalErrorCounter()
    m_consecutiveInternalErrors = 0
End Sub

' Extracts the host name/IP from a URL.
' Designed to be future-proof against IPv6 loopbacks (e.g., [::1]) and custom ports.
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

    ' Check for "://" (standard), "//" (protocol-relative), or none
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
    
    ' Check for Query string indicator (?)
    qPos = InStr(startPos, url, "?")
    If qPos > 0 And (authEnd = 0 Or qPos < authEnd) Then authEnd = qPos
    
    ' Check for Fragment indicator (#)
    fPos = InStr(startPos, url, "#")
    If fPos > 0 And (authEnd = 0 Or fPos < authEnd) Then authEnd = fPos
    
    If authEnd = 0 Then authEnd = Len(url) + 1
    
    authority = Mid$(url, startPos, authEnd - startPos)
    If authority = "" Then Exit Function
    
    ' Strip user credentials; InStrRev handles '@' inside passwords
    atPos = InStrRev(authority, "@")
    If atPos > 0 Then authority = Mid$(authority, atPos + 1)
    
    If Left$(authority, 1) = "[" Then
        ' Case: IPv6 address [::1]:port
        closeBracket = InStr(1, authority, "]")
        If closeBracket > 0 Then
            hostPart = Mid$(authority, 2, closeBracket - 2)
        Else
            ' Invalid IPv6 format (missing closing bracket)
            Exit Function
        End If
    Else
        ' Case: Standard domain or IPv4 address
        ' Locate the last colon to identify the port number
        portPos = InStrRev(authority, ":")
        If portPos > 0 Then
            hostPart = Left$(authority, portPos - 1)
        Else
            hostPart = authority
        End If
    End If
    
    ' Convert host to lowercase for consistency
    If hostPart <> "" Then ExtractHost = LCase$(hostPart)
End Function

' Extracts the origin (scheme + host + port) from a URL string.
' Example: "https://example.com:8080/path" -> "https://example.com:8080"
' Returns empty string for non-parseable or empty URLs.
Public Function ExtractOrigin(ByVal url As String) As String
    If Len(url) = 0 Then Exit Function
    Dim schemeEnd As Long: schemeEnd = InStr(1, url, "://", vbTextCompare)
    If schemeEnd = 0 Then Exit Function
    ' Find the first "/" after "://" which terminates the authority part
    Dim authStart As Long: authStart = schemeEnd + 3
    Dim pathStart As Long: pathStart = InStr(authStart, url, "/")
    If pathStart = 0 Then
        ExtractOrigin = LCase$(url)
    Else
        ExtractOrigin = LCase$(Left$(url, pathStart - 1))
    End If
End Function

' Internal Helper: Performs an interruptible sleep in small chunks.
' It intelligently yields to the OS (DoEvents) only when it's safe to do so,
' ensuring that the WebSocket state machine is not corrupted by re-entrancy.
Private Sub InterruptibleSleep(ByVal totalMs As Long, Optional ByVal allowDoEvents As Boolean = True)
    Dim startTs As Double: startTs = NowMs()
    Const CHUNK_MS As Long = 50 ' Small granularity for better responsiveness
    
    Do While (NowMs() - startTs) < totalMs
        ' CRITICAL GUARD:
        ' We only call DoEvents if:
        ' 1. The caller explicitly allows it (allowDoEvents).
        ' 2. We are NOT in the middle of a SendAndReceive operation (m_isSendReceiving).
        ' corrupting the WinHttp handle state during active I/O.
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If
        
        ' Physical sleep to prevent high CPU usage during the wait period.
        SleepWinAPI CHUNK_MS
    Loop
End Sub

' Releases all WinHttp handles without sending a WebSocket close frame.
Public Sub ForceClose()
    ' OERN SCOPE = ENTIRE PROCEDURE (no matching "On Error GoTo 0"): emergency
    ' teardown must never raise. Every handle release below is best-effort.
    On Error Resume Next
    ' INVARIANT (TEARDOWN ORDER): clear communication locks first, before any handle operations.
    ' ForceClose is the emergency exit path (called from Wrapper.Shutdown and
    ' Class_Terminate). If the caller was interrupted mid-SendAndReceive,
    ' these flags must be reset to prevent deadlock on reconnection.
    m_isSendReceiving = False
    m_isReading = False
    m_isReadingSince = 0
    m_discardActive = False
    
    ' Skip WinHttpWebSocketClose (no close frame, no timeout).
    ' Directly release all WinHttp handles.
    '
    ' INVARIANT (HANDLE OWNERSHIP): the request handle is independent of the WebSocket handle.
    ' On the success path it was already closed in Connect; here we simply close
    ' whatever handles are still open (best-effort emergency teardown).
    If hWebSocket_ <> 0 Then
        WinHttpCloseHandle hWebSocket_
        hWebSocket_ = 0
    End If
    If hRequest_ <> 0 Then
        WinHttpCloseHandle hRequest_
        hRequest_ = 0
    End If
    If hConnection_ <> 0 Then
        WinHttpCloseHandle hConnection_
        hConnection_ = 0
    End If
    If hSession_ <> 0 Then
        WinHttpCloseHandle hSession_
        hSession_ = 0
    End If

    ' Purge the input stream so a reused instance never carries stale fragments into the
    ' next connection's first message. Connect -> InternalCleanup already truncates on
    ' reconnect, so this is defense-in-depth; it also keeps ForceClose consistent with
    ' InternalCleanup and with the m_discardActive invariant cleared above.
    If Not m_InStream Is Nothing Then
        If m_InStream.State = adStateOpen Then
            m_InStream.Position = 0
            m_InStream.SetEOS
            m_InStream.Type = adTypeBinary
        End If
    End If

    Set m_EventBuffer = New Collection
    SetState StateDisconnected
    On Error GoTo 0
End Sub
