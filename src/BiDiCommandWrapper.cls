VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BiDiCommandWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' WebDriver BiDi for SeleniumVBA
' https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA
' MIT License / Copyright (c) hanamichi77777
Private Const BIDI_WRAPPER_VERSION As String = "4.0"
'
' BiDiCommandWrapper owns command policy, recovery and SPA synchronization;
' BiDiSocketCommunicator owns transport and request IDs.
'
' Main funnels:
'   Result JS: CallScript -> WrapInSafetyShell -> ExecuteBiDiCommand.
'   DOM action: ExecuteBaseAction -> Prepare -> Execute -> Sync.
'   SPA wait: MasterWaitEngine -> EvaluateWaitCondition -> EvaluateSpaIdleCondition.
'
' Preserve these invariants:
'   - m_waitStackDepth tracks nested active wait loops.
'   - Do not pump Office messages during SendAndReceive correlation.
'   - Ambiguous outcomes are never replayed when they may duplicate state or overlap.
'   - ERR_CONSENSUS_HALT propagates; ordinary post-action sync failure is best-effort.
'   - Completion arms are one-shot and consumed by the next SPA-idle wait.
'   - File-dialog composite actions preserve completion arms until input.setFiles completes.
'   - Polling retries only classified transient protocol errors; fatal errors surface immediately.
'   - PostScript id 2 is reserved, send-only, and its response is discarded.
'
' OERN is limited to captured or deliberately best-effort scopes; it must not hide
' the primary action result or structural response errors.
Private socket_ As BiDiSocketCommunicator
' True when no live session still requires teardown.
Private m_shutdownCompleted As Boolean
Private m_useTrustedClick As Boolean        ' default False -> synthetic JS clicks (proven path); True = trusted pointer clicks (opt-in)
Private m_debugMode As Boolean              ' Backing field for the DebugMode property. Default On; set in Class_Initialize.

' Session-scoped preload registration; it is not retroactive for existing closed roots.
Private m_shadowUnlockerRegistered As Boolean
Private m_shadowUnlockerScriptId As String

' Last post-action sync; Attempted=False means no sync was requested.
Private m_lastSyncAttempted As Boolean
Private m_lastSyncSucceeded As Boolean
Private m_lastSyncErrorNumber As Long
Private m_lastSyncErrorSource As String
Private m_lastSyncErrorDescription As String

Private m_isRecovering As Boolean
Private m_lastTelemetryTick As Currency
Private Const MAX_CONTEXT_RECOVERY_RETRIES As Long = 5
Private m_consecutiveCdpFailures As Long
Private Const MAX_CONSECUTIVE_CDP_FAILURES As Long = 5

Private m_navigationDetected As Boolean   ' Flag: navigationStarted event received
Private m_navigationContextId As String   ' Context ID where navigation was detected

Private m_waitStackDepth As Long ' Tracks recursion depth of active wait loops

Private m_lastHudKey As String        ' Fingerprint of last HUD payload (msg+color)
Private m_lastHudCtx As String        ' Context the last HUD was rendered into
Private m_lastHudRetryTick As Currency ' Throttle timestamp for HUD send retries

Private m_mainContextId As String
Private m_lostMainContextId As String
Private m_mainContextState As MainContextState
Private m_contextLifecycleDroppedEventsBaseline As Long
Private m_contextLifecycleNeedsReconcile As Boolean
Private m_cdpSessionId As String
Private m_cachedRealmIds As Dictionary ' Cached Execution Realm ID
Private m_validatedContexts As Dictionary ' Context IDs confirmed present in the tree (skips repeat getTree)
Private m_activeSubscriptions As Dictionary ' Event Management (Aggregation)
Private m_isProcessingQueue As Boolean

' Chromium BiDi may leave fetchError.errorText empty for blocked requests;
' classify [BLOCK] by matching the URL against the installed block patterns.
Private m_blockingPatterns As Collection

Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_COMMAND_TIMEOUT_MS As Long = 30000
Private Const NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS As Long = 60000
Private Const TRANSPORT_TIMEOUT_MARGIN_MS As Long = 5000
Private Const MAX_TRANSPORT_TIMEOUT_MS As Long = 605000
Private Const DEFAULT_STALE_RETRIES As Long = 3
Private Const SHORT_WAIT_MS As Long = 500
' Discovery Log ring-buffer limit.
Private Const MAX_LOG_SIZE As Long = 5000
Private Const MAX_SUPPRESSED_KEYS As Long = 2000
Private Const MAX_SUPPRESSED_SUMMARY_LINES As Long = 25
Private Const MAX_CONSOLE_LOG_TEXT_SIZE As Long = 2000
Private Const MAX_PARSED_EVENT_SIZE As Long = 524288 ' 512 KiB cap for recorder event parsing
Private Const MIN_MSG_SIZE As Long = 10     ' Ignore extremely short messages
' Reserved send-only ID; communicator discards its response.
Private Const POST_SCRIPT_ID As Long = 2

Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore, #vba-bidi-hud, [id^='sizzle']"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"
Private Const SPA_PROBE_VERSION As Long = 38 ' Internal probe generation/quiet/filter-diagnostic contract
Private m_userIgnoreSelectors As Collection
Private m_userIgnoreNetworkPatterns As Collection
Private m_idleFilterRevision As Long          ' increments whenever VBA-side ignore filters change
Private m_probeFilterRevisions As Dictionary ' context ID -> last filter revision pushed into that probe
Private m_reportedInvalidSelectors As Dictionary ' selector -> one-shot diagnostic emitted this session

Private m_recorderLogs As Collection
Private m_isRecording As Boolean
Private m_recorderStartTime As Double
Private m_excludeResources As Boolean
' beforeRequestSent is recorder-only; no latency pairing is retained.
Private m_recSuppByKey As Dictionary       ' normalized key -> suppressed-event count
Private m_recSuppOverflowCount As Long      ' suppressed events not keyed after the key cap
Private m_droppedHeadCount As Long          ' count of oldest timeline lines evicted by the MAX_LOG_SIZE ring buffer (0 = none)

' Completion arms gate SPA idle; a hit restarts the full stable window.
Private m_sigDomArmed As Boolean            ' ArmContentSignal active (existing-element rewrite)
Private m_sigDomSeen As Boolean             ' latched: target-subtree mutation observed
Private m_sigDomXPath As String             ' armed target (diagnostics only)
Private m_sigNetArmed As Boolean            ' ArmNetworkSignal active (API response arrival)
Private m_sigNetSeen As Boolean             ' latched: matching responseCompleted observed
Private m_sigNetPattern As String           ' armed URL pattern (UrlMatchesPattern notation)
Private m_signalOwnerContextId As String    ' one browsing context owns every active completion arm
Private m_sigVisArmed As Boolean            ' ArmVisibilitySignal active (new-element appearance)
Private m_sigVisSeen As Boolean             ' latched: target became visible
Private m_sigVisXPath As String             ' armed target (diagnostics only)
Private m_sigLostWarned As Boolean          ' one-shot [SPA-SIGNAL-LOST] warning guard
Private m_sigNetContextMissWarned As Boolean ' one-shot foreign/empty-context match diagnostic
Private m_lastProbeQuietMs As Double        ' last polled domQuiet; quietAtHit for net hits

' Short-lived file-input capture. input.fileDialogOpened is subscribed persistently;
' this local arm is one-shot and owned by exactly one browsing context.
Private m_fileDlgArmed As Boolean
Private m_fileDlgSeen As Boolean
Private m_fileDlgOwnerContextId As String
Private m_fileDlgHasElement As Boolean
Private m_fileDlgSharedId As String
Private m_fileDlgMultipleKnown As Boolean
Private m_fileDlgMultiple As Boolean
Private m_fileDlgContextMissWarned As Boolean
Private m_fileDlgDroppedEventsAtArm As Long

' New top-level context capture: baseline snapshot plus opener correlation.
Private Const NEW_CONTEXT_CORRELATION_OPENER As String = "originalOpener"
Private Const NEW_CONTEXT_CORRELATION_BASELINE As String = "baseline-diff"
Private Const NEW_CONTEXT_SETTLE_MS As Long = 150
Private m_newContextArmed As Boolean
Private m_newContextSeen As Boolean
Private m_newContextAmbiguous As Boolean
Private m_newContextCandidateDestroyed As Boolean
Private m_newContextOwnerDestroyed As Boolean
Private m_newContextOwnerContextId As String
Private m_newContextOwnerClientWindow As String
Private m_newContextId As String
Private m_newContextUrl As String
Private m_newContextOriginalOpener As String
Private m_newContextClientWindow As String
Private m_newContextCorrelationMode As String
Private m_newContextCandidateCount As Long
Private m_newContextForeignIgnoredCount As Long
Private m_newContextDroppedEventsAtArm As Long
Private m_newContextBaseline As Dictionary

' Download transaction capture: one-shot owner-context arm, begin/end correlation,
' and exactly-one start detection. Correlation prefers the BiDi download id; context+navigation
' is the measured Edge fallback. Multiple owner-context starts are latched as ambiguity and
' are not followed as multiple transactions.
Private Const DOWNLOAD_CORRELATION_UUID As String = "download"
Private Const DOWNLOAD_CORRELATION_CONTEXT_NAV As String = "context+navigation"
Private Const DOWNLOAD_CORRELATION_UNCORRELATABLE As String = "uncorrelatable"
Private Const DEFAULT_DOWNLOAD_TIMEOUT_MS As Long = 30000
Private m_downloadArmed As Boolean
Private m_downloadStartedSeen As Boolean
Private m_downloadEndSeen As Boolean
Private m_downloadOwnerContextId As String
Private m_downloadAcceptedContextId As String
Private m_downloadCorrelationMode As String
Private m_downloadCorrelationDownloadId As String
Private m_downloadCorrelationNavigationId As String
Private m_downloadEndContextId As String
Private m_downloadEndStatusDiagnostic As String
Private m_downloadEndFilePathDiagnostic As String
Private m_downloadContextMissWarned As Boolean
Private m_downloadDroppedEventsAtArm As Long
Private m_downloadStartedCount As Long
Private m_downloadMultipleStarted As Boolean
Private m_downloadSecondCorrelationMode As String
Private m_downloadSecondDownloadId As String
Private m_downloadSecondNavigationId As String
Private m_downloadSuggestedFilename As String
Private m_downloadUrl As String
Private m_downloadTerminalEligible As Boolean
Private m_downloadUncorrelatableCount As Long
Private m_downloadFirstUncorrelatableContext As String

' browser.setDownloadBehavior is browser-side state that may outlive this wrapper/WebSocket.
' Once this wrapper attempts to change it, Shutdown must best-effort clear the global behavior
' before transport teardown. v1 intentionally does not expose userContexts scoping.
Private m_downloadBehaviorCleanupRequired As Boolean
Private m_downloadBehaviorStateUnknown As Boolean

' Native file-chooser interception is target/session state outside the VBA object.
' "MayBeActive" is intentionally conservative: it is set BEFORE the enable command is sent
' so an ambiguous transport failure cannot make us forget a browser-side interception state.
Private m_fileChooserInterceptMayBeActive As Boolean
Private m_fileChooserInterceptSessionId As String
Private m_fileChooserInterceptContextId As String


' Carryover is a one-wait diagnostic hint and never changes consensus.
Private m_unresolvedJsonThisWait As Boolean       ' this wait saw a tracked json/API, no content render yet
Private m_unresolvedJsonFirstKeyThisWait As String ' first such json's key (SUSPECT line only; NOT used by NOTE)
Private m_unresolvedJsonFirstSecThisWait As Double ' first such json's arrival relTime (SUSPECT line only)
Private m_hasCarryoverHintForNextWait As Boolean  ' reservation: hand a one-shot NOTE arm to the NEXT wait
Private m_carryoverNoteArmedThisWait As Boolean   ' this wait may emit ONE carryover note (consumed from the reservation)
Private m_carryoverNoteEmittedThisWait As Boolean ' this wait already emitted its one note (1-per-wait cap)

Private Const ERR_BIDI_PROTOCOL As Long = vbObjectError + 1001
Private Const ERR_JS_EXCEPTION As Long = vbObjectError + 1002
Private Const ERR_ACTION_FAILED As Long = vbObjectError + 1003
Private Const ERR_CONSENSUS_HALT As Long = vbObjectError + 1010  ' Consensus failure: process halted for review
Private Const ERR_CONNECT_FAILED As Long = vbObjectError + 1011 ' WebSocket connection could not be established
Private Const ERR_DOWNLOAD_START_TIMEOUT As Long = vbObjectError + 1020
Private Const ERR_DOWNLOAD_COMPLETION_TIMEOUT As Long = vbObjectError + 1021
Private Const ERR_DOWNLOAD_MULTIPLE_STARTED As Long = vbObjectError + 1022
Private Const ERR_MAIN_CONTEXT_LOST As Long = vbObjectError + 1023
Private Const ERR_NEW_CONTEXT_TIMEOUT As Long = vbObjectError + 1024
Private Const ERR_NEW_CONTEXT_MULTIPLE As Long = vbObjectError + 1025
Private Const ERR_NEW_CONTEXT_DESTROYED As Long = vbObjectError + 1026

Private Enum MainContextState
    MainContextUnbound = 0
    MainContextPinned = 1
    MainContextLost = 2
End Enum

Public Enum BiDiMatchBy
    MatchByUrl = 0
    MatchByTitle = 1
End Enum

Public Enum BiDiContextType
    ContextTypeTab = 0
    ContextTypeWindow = 1
End Enum

Private Enum waitMode
    WaitMode_UrlContains
    WaitMode_SpaIdle
End Enum
' Extension contracts: result JS uses the Safety Shell; long loops yield and call hb();
' public timeouts use shared validation; wrapper waits yield outside correlation only.

Private Sub Class_Initialize()
    Set socket_ = New BiDiSocketCommunicator
    Set m_activeSubscriptions = New Dictionary
    Set m_cachedRealmIds = New Dictionary
    Set m_validatedContexts = New Dictionary
    Set m_probeFilterRevisions = New Dictionary
    Set m_reportedInvalidSelectors = New Dictionary
    Set m_blockingPatterns = New Collection
    Set m_newContextBaseline = New Dictionary
    m_mainContextState = MainContextUnbound

    ResetSignalState
    ResetLastSyncState
    m_shutdownCompleted = True
    m_debugMode = True
End Sub

Public Property Get Version() As String
    Version = BIDI_WRAPPER_VERSION
End Property

Public Sub ConnectTo(ByVal webSocketUrl As String, _
                     Optional ByVal requireShadowUnlocker As Boolean = False)
    ' Validate before ClearState; otherwise a rejected reconnect can erase wrapper state
    ' while the existing browser/socket session remains live.
    
    If Len(Trim$(webSocketUrl)) = 0 Then
        Err.Raise ERR_CONNECT_FAILED, "BiDiCommandWrapper:ConnectTo", _
            "webSocketUrl is empty. Call caps.EnableBiDiMode BEFORE driver.OpenBrowser, " & _
            "then pass driver.GetWebSocketUrl to ConnectTo."
    End If
    
    If socket_.state = StateConnected Or socket_.state = StateConnecting Then
        Err.Raise ERR_CONNECT_FAILED, "BiDiCommandWrapper:ConnectTo", _
            "The BiDi socket is already connected. Call Shutdown before reconnecting."
    End If
    
    ' Connection plus mandatory subscription is atomic. Shadow Unlocker registration is
    ' best-effort by default, or part of the atomic initialization when explicitly required.
    On Error GoTo InitializationFailed

    ResetSessionBoundaryState
    Me.ClearState

    Dim attempt As Long
    Dim connected As Boolean
    For attempt = 1 To 3
        connected = socket_.AttemptAutoConnect(webSocketUrl)
        If connected Then Exit For
        If attempt < 3 Then socket_.YieldExecution 250 * attempt, False
    Next attempt

    If Not connected Then
        Err.Raise ERR_CONNECT_FAILED, "BiDiCommandWrapper:ConnectTo", _
            "Failed to establish the BiDi WebSocket connection to: " & webSocketUrl & vbCrLf & _
            "Checks: (1) the URL must use ws:// and point to localhost/127.0.0.1/[::1], " & _
            "(2) it must contain ""/session"" (i.e. a webSocketUrl issued by the driver), " & _
            "(3) the browser/driver session must still be alive."
    End If

    m_shutdownCompleted = False
    
    Me.ExecuteSessionSubscribe Array( _
        "log.entryAdded", _
        "browsingContext.navigationStarted", _
        "browsingContext.contextCreated", _
        "browsingContext.contextDestroyed"), , True
    m_contextLifecycleDroppedEventsBaseline = CurrentLifecycleRelevantLossCount()

    ' Hard-reset insurance: a previous VBA project may have died after enabling Chromium's
    ' file-chooser interceptor. Sweep current Chromium page sessions back to disabled.
    ' Unsupported browsers are ignored without failing ConnectTo.
    RecoverResidualFileChooserInterceptionOnConnectBestEffort

    ' Hard-reset insurance for browser.setDownloadBehavior. Unsupported implementations are
    ' accepted; an indeterminate clear leaves download execution blocked until a setter succeeds.
    RecoverResidualDownloadBehaviorOnConnectBestEffort

    Dim shadowErrNum As Long
    Dim shadowErrSrc As String
    Dim shadowErrDesc As String
    If Not TryEnableGlobalShadowUnlocker(shadowErrNum, shadowErrSrc, shadowErrDesc) Then
        If requireShadowUnlocker Then
            If shadowErrNum = 0 Then shadowErrNum = ERR_BIDI_PROTOCOL
            If Len(shadowErrSrc) = 0 Then shadowErrSrc = "BiDi:Protocol:script.addPreloadScript"
            If Len(shadowErrDesc) = 0 Then shadowErrDesc = "Shadow Unlocker registration failed."
            Err.Raise shadowErrNum, shadowErrSrc, shadowErrDesc
        End If

        Debug.Print "BiDi Warning: Shadow Unlocker unavailable; connection continues in best-effort mode. " & _
                    "Source=" & shadowErrSrc & " | " & shadowErrDesc
    End If
    Exit Sub

InitializationFailed:
    Dim initErrNum As Long: initErrNum = Err.Number
    Dim initErrDesc As String: initErrDesc = Err.description
    Dim initErrSrc As String: initErrSrc = Err.source

    On Error Resume Next
    socket_.ForceClose
    ResetSessionBoundaryState
    Me.ClearState
    m_shutdownCompleted = True
    On Error GoTo 0

    If Len(initErrSrc) = 0 Then initErrSrc = "BiDiCommandWrapper:ConnectTo"
    Err.Raise initErrNum, initErrSrc, initErrDesc
End Sub

Private Sub Class_Terminate()
    ' Destructor cleanup is intentionally non-raising; the OERN scope covers the whole
    ' procedure because teardown must continue after any individual failure.
    On Error Resume Next

    If Not m_shutdownCompleted Then
        If m_mainContextId <> "" Then
            Call Me.PostScript(GetHudJs("remove"), Nothing, m_mainContextId)
        End If

        ' Default destructor shutdown force-closes the socket. The CDP tunnel is session-bound,
        ' so no acknowledged detach is required on this emergency path.
        Call Shutdown
    End If
    
    If m_debugMode Then
        Debug.Print "BiDiCommandWrapper: Terminated."
    End If
End Sub
Public Property Get GetSocket() As BiDiSocketCommunicator
    Set GetSocket = socket_
End Property
Public Property Set SetSocket(ByVal socket As BiDiSocketCommunicator)
    If socket Is Nothing Then
        Err.Raise ERR_ACTION_FAILED, "BiDiCommandWrapper:SetSocket", _
            "The replacement BiDiSocketCommunicator cannot be Nothing."
    End If

    If Not socket_ Is Nothing Then
        If socket_ Is socket Then Exit Property

        Select Case socket_.state
            Case StateConnecting, StateConnected, StateDisconnecting, StateError
                Err.Raise ERR_ACTION_FAILED, "BiDiCommandWrapper:SetSocket", _
                    "Cannot replace the communicator while the current socket may still own " & _
                    "transport resources. Call Shutdown before SetSocket."
        End Select
    End If

    Select Case socket.state
        Case StateConnecting, StateConnected, StateDisconnecting, StateError
            Err.Raise ERR_ACTION_FAILED, "BiDiCommandWrapper:SetSocket", _
                "The replacement communicator must be fully disconnected and free of " & _
                "an error-state transport."
    End Select

    ResetSessionBoundaryState
    ClearState
    Set socket_ = socket
    m_shutdownCompleted = True
End Property

' Compatibility shim: WinSock is the only transport; False is rejected.
Public Property Let UseWinsockTransport(ByVal v As Boolean)
    socket_.UseWinsockTransport = v
End Property

Public Property Get UseWinsockTransport() As Boolean
    UseWinsockTransport = socket_.UseWinsockTransport
End Property

Private Sub DispatchGlobalEvents(ByVal json As Object)
    If Not json.Exists("method") Then Exit Sub
    Dim method As String: method = json("method")

    ' Blocked requests surface as fetchError. They remain visible to diagnostics but do
    ' not participate in SPA-idle consensus because blocking is browser-side.

    Select Case method
        Case "log.entryAdded"
            Call HandleConsoleLog(json)

        Case "input.fileDialogOpened"
            Call HandleFileDialogOpenedEvent(json)

        Case "browsingContext.downloadWillBegin"
            Call HandleDownloadWillBeginEvent(json)

        Case "browsingContext.downloadEnd"
            Call HandleDownloadEndEvent(json)

        Case "browsingContext.contextCreated"
            Call HandleContextCreatedEvent(json)

        Case "browsingContext.contextDestroyed"
            Call HandleContextDestroyedEvent(json)

        Case "browsingContext.navigationStarted"
            m_navigationDetected = True
            m_navigationContextId = GetPathString(json, "params.context")

            If m_cachedRealmIds.Exists(m_navigationContextId) Then
                m_cachedRealmIds.Remove m_navigationContextId
            End If

            ' Navigation invalidates the cached realm and validated-context entry for that context.
            If Not m_validatedContexts Is Nothing Then
                If m_validatedContexts.Exists(m_navigationContextId) Then m_validatedContexts.Remove m_navigationContextId
            End If

            ' A new document has no applied VBA-side filter revision until its probe is injected.
            If Not m_probeFilterRevisions Is Nothing Then
                If m_probeFilterRevisions.Exists(m_navigationContextId) Then m_probeFilterRevisions.Remove m_navigationContextId
            End If

            If m_debugMode Then
                If m_navigationContextId = m_mainContextId Then
                    Debug.Print "BiDi: Navigation detected in main context " & _
                                Left$(m_navigationContextId, 16) & "... (Probe pre-invalidated)"
                End If
            End If

        Case "network.responseCompleted", "network.fetchError"
            ' Network-signal matching runs on every responseCompleted before noise filtering;
            ' an ignored endpoint may still be an explicitly armed completion signal.
            If method = "network.responseCompleted" And m_sigNetArmed And Not m_sigNetSeen Then
                Dim sigUrl As String: sigUrl = GetPathString(json, "params.response.url")
                Dim sigContext As String: sigContext = GetPathString(json, "params.context")
                If Len(sigUrl) > 0 And UrlMatchesPattern(sigUrl, m_sigNetPattern) Then
                    If StrComp(sigContext, m_signalOwnerContextId, vbBinaryCompare) = 0 Then
                        m_sigNetSeen = True
                        Dim sigStatus As String: sigStatus = GetPathString(json, "params.response.status")
                        LogSignalLine "[SPA-SIGNAL-HIT] type=net | key=" & NormalizeEndpointKey(sigUrl) & _
                                      " | status=" & sigStatus & _
                                      " | ctx=" & sigContext & _
                                      " | quietAtHit=" & Format(m_lastProbeQuietMs, "0") & "ms"
                    ElseIf Not m_sigNetContextMissWarned Then
                        m_sigNetContextMissWarned = True
                        Dim actualSigContext As String: actualSigContext = sigContext
                        If Len(actualSigContext) = 0 Then actualSigContext = "(empty)"
                        LogSignalLine "[SPA-SIGNAL-CTX-MISS] type=net | key=" & _
                                      NormalizeEndpointKey(sigUrl) & _
                                      " | expectedCtx=" & m_signalOwnerContextId & _
                                      " | actualCtx=" & actualSigContext
                    End If
                End If
            End If

            ' Because blocked fetchError often lacks errorText, URL-pattern matching distinguishes
            ' [BLOCK] from ordinary network failures.
            If m_debugMode And method = "network.fetchError" Then
                Dim blockedUrl As String
                blockedUrl = GetPathString(json, "params.request.url")
                If Len(blockedUrl) > 0 And m_blockingPatterns.count > 0 Then
                    Dim tok As Variant
                    For Each tok In m_blockingPatterns
                        If UrlMatchesPattern(blockedUrl, CStr(tok)) Then
                            Debug.Print "CDP ResourceBlocking [BLOCK]: " & Left$(blockedUrl, 120)
                            Exit For
                        End If
                    Next
                End If
            End If

    End Select

    If m_isRecording Then
        FormatAndStoreLog json
    End If
End Sub

Private Sub HandleFileDialogOpenedEvent(ByVal json As Object)
    If Not m_fileDlgArmed Or m_fileDlgSeen Then Exit Sub

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")

    If Len(eventCtx) = 0 Then
        LogSignalLine "[FILE-DIALOG-IGNORED] reason=missing-context"
        Exit Sub
    End If

    If StrComp(eventCtx, m_fileDlgOwnerContextId, vbBinaryCompare) <> 0 Then
        If Not m_fileDlgContextMissWarned Then
            m_fileDlgContextMissWarned = True
            LogSignalLine "[FILE-DIALOG-CTX-MISS] expectedCtx=" & m_fileDlgOwnerContextId & _
                          " | actualCtx=" & eventCtx
        End If
        Exit Sub
    End If

    Dim p As Object
    Set p = GetPathObject(json, "params")
    If p Is Nothing Then
        LogSignalLine "[FILE-DIALOG-IGNORED] reason=missing-params | ctx=" & eventCtx
        Exit Sub
    End If

    m_fileDlgSeen = True
    m_fileDlgArmed = False

    m_fileDlgMultipleKnown = p.Exists("multiple")
    If m_fileDlgMultipleKnown Then m_fileDlgMultiple = CBool(p("multiple"))

    m_fileDlgHasElement = False
    m_fileDlgSharedId = vbNullString

    If p.Exists("element") Then
        If IsObject(p("element")) Then
            Dim elementRef As Object
            Set elementRef = p("element")
            If Not elementRef Is Nothing Then
                If elementRef.Exists("sharedId") Then
                    m_fileDlgSharedId = CStr(elementRef("sharedId"))
                    m_fileDlgHasElement = (Len(m_fileDlgSharedId) > 0)
                End If
            End If
        End If
    End If

    LogSignalLine "[FILE-DIALOG-HIT] ctx=" & eventCtx & _
                  " | hasElement=" & LCase$(CStr(m_fileDlgHasElement)) & _
                  " | multiple=" & IIf(m_fileDlgMultipleKnown, LCase$(CStr(m_fileDlgMultiple)), "unknown")
End Sub

Private Sub HandleContextCreatedEvent(ByVal json As Object)
    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then Exit Sub

    If m_validatedContexts Is Nothing Then Set m_validatedContexts = New Dictionary
    m_validatedContexts(eventCtx) = True

    If Not m_newContextArmed Then Exit Sub

    Dim parentId As String
    parentId = GetPathString(json, "params.parent")
    If Len(parentId) > 0 Then Exit Sub

    If Not m_newContextBaseline Is Nothing Then
        If m_newContextBaseline.Exists(eventCtx) Then Exit Sub
    End If

    Dim openerId As String
    openerId = GetPathString(json, "params.originalOpener")

    Dim correlationMode As String
    If Len(openerId) > 0 Then
        If StrComp(openerId, m_newContextOwnerContextId, vbBinaryCompare) <> 0 Then
            m_newContextForeignIgnoredCount = m_newContextForeignIgnoredCount + 1
            If m_debugMode Then
                Debug.Print "BiDi: New context ignored; opener does not match owner. ctx=" & eventCtx
            End If
            Exit Sub
        End If
        correlationMode = NEW_CONTEXT_CORRELATION_OPENER
    Else
        correlationMode = NEW_CONTEXT_CORRELATION_BASELINE
    End If

    If Not m_newContextSeen Then
        m_newContextSeen = True
        m_newContextCandidateCount = 1
        m_newContextId = eventCtx
        m_newContextUrl = GetPathString(json, "params.url")
        m_newContextOriginalOpener = openerId
        m_newContextClientWindow = GetPathString(json, "params.clientWindow")
        m_newContextCorrelationMode = correlationMode

        If m_debugMode Then
            Debug.Print "BiDi: New context captured. ctx=" & eventCtx & _
                        " | correlation=" & correlationMode
        End If
        Exit Sub
    End If

    If StrComp(eventCtx, m_newContextId, vbBinaryCompare) = 0 Then Exit Sub

    m_newContextCandidateCount = m_newContextCandidateCount + 1
    m_newContextAmbiguous = True
    m_newContextArmed = False
End Sub

Private Sub HandleContextDestroyedEvent(ByVal json As Object)
    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then Exit Sub

    InvalidateContextState eventCtx, True

    If m_newContextArmed Then
        If StrComp(eventCtx, m_newContextOwnerContextId, vbBinaryCompare) = 0 Then
            m_newContextOwnerDestroyed = True
        End If
    End If

    If m_newContextSeen Then
        If StrComp(eventCtx, m_newContextId, vbBinaryCompare) = 0 Then
            m_newContextCandidateDestroyed = True
            m_newContextArmed = False
        End If
    End If
End Sub

Private Sub InvalidateContextState(ByVal contextId As String, _
                                   Optional ByVal markMainLost As Boolean = True)
    If Len(contextId) = 0 Then Exit Sub

    If Not m_cachedRealmIds Is Nothing Then
        If m_cachedRealmIds.Exists(contextId) Then m_cachedRealmIds.Remove contextId
    End If

    If Not m_validatedContexts Is Nothing Then
        If m_validatedContexts.Exists(contextId) Then m_validatedContexts.Remove contextId
    End If

    If Not m_probeFilterRevisions Is Nothing Then
        If m_probeFilterRevisions.Exists(contextId) Then m_probeFilterRevisions.Remove contextId
    End If

    If StrComp(m_lastHudCtx, contextId, vbBinaryCompare) = 0 Then
        m_lastHudCtx = ""
        m_lastHudKey = ""
        m_lastHudRetryTick = 0
    End If

    RemoveContextScopedSubscriptionRegistry contextId

    If markMainLost Then MarkMainContextLost contextId
End Sub

Private Sub RemoveContextScopedSubscriptionRegistry(ByVal contextId As String)
    ' Local registry only; context-scoped remote unsubscribe is not used by current callers.
    If m_activeSubscriptions Is Nothing Then Exit Sub
    If m_activeSubscriptions.count = 0 Then Exit Sub

    Dim suffix As String
    suffix = "|" & contextId

    Dim keys As Variant
    keys = m_activeSubscriptions.keys

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim key As String
        key = CStr(keys(i))
        If Len(key) >= Len(suffix) Then
            If StrComp(Right$(key, Len(suffix)), suffix, vbBinaryCompare) = 0 Then
                m_activeSubscriptions.Remove key
            End If
        End If
    Next i
End Sub

Private Sub ResetNewContextCaptureState()
    m_newContextArmed = False
    m_newContextSeen = False
    m_newContextAmbiguous = False
    m_newContextCandidateDestroyed = False
    m_newContextOwnerDestroyed = False
    m_newContextOwnerContextId = ""
    m_newContextOwnerClientWindow = ""
    m_newContextId = ""
    m_newContextUrl = ""
    m_newContextOriginalOpener = ""
    m_newContextClientWindow = ""
    m_newContextCorrelationMode = ""
    m_newContextCandidateCount = 0
    m_newContextForeignIgnoredCount = 0
    m_newContextDroppedEventsAtArm = 0
    Set m_newContextBaseline = New Dictionary
End Sub

Private Sub ArmNewContextCapture(ByVal ownerContextId As String)
    ResetNewContextCaptureState

    Dim rawTree As String
    rawTree = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, True)

    Dim tree As Object
    Set tree = ValidateProtocolResponse(rawTree, "browsingContext.getTree")

    Dim contexts As Collection
    Set contexts = GetPathObject(tree, "result.contexts")
    If contexts Is Nothing Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ArmNewContextCapture", _
                  "browsingContext.getTree returned no top-level context list."
    End If

    ReconcileContextTreeCache contexts

    Dim ownerInfo As Dictionary
    Set ownerInfo = FindContextInfoExactId(contexts, ownerContextId)
    If ownerInfo Is Nothing Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ArmNewContextCapture", _
                  "Owner context is no longer present: " & ownerContextId
    End If

    Dim ctx As Dictionary
    For Each ctx In contexts
        Dim existingId As String
        existingId = GetPathString(ctx, "context")
        If Len(existingId) > 0 Then m_newContextBaseline(existingId) = True
    Next ctx

    ' Drain events accumulated during getTree while the local capture arm is still down.
    ' This narrows the stale baseline-diff window without adding a queue watermark.
    ProcessRecorderQueue

    m_newContextOwnerContextId = ownerContextId
    m_newContextOwnerClientWindow = GetPathString(ownerInfo, "clientWindow")
    m_newContextDroppedEventsAtArm = CurrentLifecycleRelevantLossCount()
    m_newContextArmed = True
End Sub

Private Function NewContextDroppedEventsSinceArm() As Long
    Dim delta As Long
    delta = CurrentLifecycleRelevantLossCount() - m_newContextDroppedEventsAtArm
    If delta < 0 Then delta = 0
    NewContextDroppedEventsSinceArm = delta
End Function

Private Sub WaitForNewContextCapture(ByVal timeoutMs As Long, ByVal methodName As String)
    Dim startTick As Currency
    startTick = socket_.NowTicks()

    Dim settleTick As Currency
    Dim settleStarted As Boolean

    Do
        ProcessRecorderQueue

        If NewContextDroppedEventsSinceArm() > 0 Then
            m_newContextArmed = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                      "Async event loss occurred after the new-context arm boundary; " & _
                      "exact context capture cannot be guaranteed."
        End If

        If m_newContextAmbiguous Then
            Err.Raise ERR_NEW_CONTEXT_MULTIPLE, "BiDi:" & methodName, _
                      "Multiple new top-level browsing contexts matched one trigger. " & _
                      "The trigger is not replayed. CandidateCount=" & _
                      CStr(m_newContextCandidateCount) & "."
        End If

        If m_newContextCandidateDestroyed Then
            Err.Raise ERR_NEW_CONTEXT_DESTROYED, "BiDi:" & methodName, _
                      "The captured browsing context was destroyed before capture completed."
        End If

        If m_newContextSeen Then
            If Not settleStarted Then
                settleTick = socket_.NowTicks()
                settleStarted = True
            ElseIf socket_.elapsedMs(settleTick) >= NEW_CONTEXT_SETTLE_MS Then
                m_newContextArmed = False
                Exit Sub
            End If
        ElseIf IsTimeout(startTick, CDbl(timeoutMs)) Then
            m_newContextArmed = False

            Dim ownerNote As String
            If m_newContextOwnerDestroyed Then
                ownerNote = " Owner context was destroyed during the wait."
            End If

            Err.Raise ERR_NEW_CONTEXT_TIMEOUT, "BiDi:" & methodName, _
                      "No new top-level browsing context was captured within " & _
                      CStr(timeoutMs) & "ms." & ownerNote & _
                      " ForeignContextsIgnored=" & CStr(m_newContextForeignIgnoredCount) & "."
        End If

        WaitHeartbeat 50
    Loop
End Sub

Private Function BuildNewContextResult() As Dictionary
    Dim result As New Dictionary
    result.Add "context", m_newContextId
    result.Add "url", m_newContextUrl
    result.Add "originalOpener", m_newContextOriginalOpener
    result.Add "clientWindow", m_newContextClientWindow
    result.Add "ownerContext", m_newContextOwnerContextId
    result.Add "ownerClientWindow", m_newContextOwnerClientWindow
    result.Add "correlation", m_newContextCorrelationMode
    result.Add "candidateCount", m_newContextCandidateCount
    result.Add "foreignIgnored", m_newContextForeignIgnoredCount

    Dim kind As String
    kind = "unknown"
    If Len(m_newContextOwnerClientWindow) > 0 And Len(m_newContextClientWindow) > 0 Then
        If StrComp(m_newContextOwnerClientWindow, m_newContextClientWindow, vbBinaryCompare) = 0 Then
            kind = "tab"
        Else
            kind = "window"
        End If
    End If
    result.Add "kind", kind

    Set BuildNewContextResult = result
End Function

' Diagnostic formatter preserves missing and explicit JSON null separately.
Private Function GetDownloadDiagnosticField(ByVal params As Object, ByVal key As String) As String
    If params Is Nothing Then
        GetDownloadDiagnosticField = "(missing)"
        Exit Function
    End If

    If Not params.Exists(key) Then
        GetDownloadDiagnosticField = "(missing)"
        Exit Function
    End If

    If IsObject(params(key)) Then
        GetDownloadDiagnosticField = "(object)"
    ElseIf IsNull(params(key)) Then
        GetDownloadDiagnosticField = "(null)"
    Else
        GetDownloadDiagnosticField = CStr(params(key))
    End If
End Function

Private Function TryGetDownloadCorrelationValue(ByVal p As Object, _
                                                ByVal key As String, _
                                                ByRef value As String) As Boolean
    value = vbNullString
    If p Is Nothing Then Exit Function
    If Not p.Exists(key) Then Exit Function
    If IsObject(p(key)) Or IsNull(p(key)) Then Exit Function

    value = CStr(p(key))
    TryGetDownloadCorrelationValue = (Len(value) > 0)
End Function

Private Sub HandleDownloadWillBeginEvent(ByVal json As Object)
    Dim p As Object
    Set p = GetPathObject(json, "params")
    If p Is Nothing Then
        LogSignalLine "[DOWNLOAD-WILL-BEGIN-IGNORED] reason=missing-params"
        Exit Sub
    End If

    LogSignalLine "[DOWNLOAD-WILL-BEGIN] ctx=" & GetDownloadDiagnosticField(p, "context") & _
                  " | download=" & GetDownloadDiagnosticField(p, "download") & _
                  " | navigation=" & GetDownloadDiagnosticField(p, "navigation") & _
                  " | filename=" & GetDownloadDiagnosticField(p, "suggestedFilename") & _
                  " | url=" & GetDownloadDiagnosticField(p, "url")

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")

    If Len(eventCtx) = 0 Then
        If m_downloadArmed Or m_downloadStartedSeen Then
            LogSignalLine "[DOWNLOAD-WILL-BEGIN-IGNORED] reason=missing-context"
        End If
        Exit Sub
    End If

    ' Once the first owner-context start has been accepted, another owner-context start means
    ' the exactly-one operation is ambiguous. Foreign contexts remain diagnostic only and never
    ' multiply the owner transaction. Do not attempt to track both downloads after this point.
    If m_downloadStartedSeen Then
        If StrComp(eventCtx, m_downloadAcceptedContextId, vbBinaryCompare) <> 0 Then
            If Not m_downloadContextMissWarned Then
                m_downloadContextMissWarned = True
                LogSignalLine "[DOWNLOAD-CTX-MISS] expectedCtx=" & m_downloadAcceptedContextId & _
                              " | actualCtx=" & eventCtx
            End If
            Exit Sub
        End If

        m_downloadStartedCount = m_downloadStartedCount + 1
        m_downloadMultipleStarted = True
        m_downloadTerminalEligible = False
        m_downloadArmed = False

        Dim secondDownloadId As String
        Dim secondNavigationId As String
        m_downloadSecondCorrelationMode = DOWNLOAD_CORRELATION_UNCORRELATABLE
        m_downloadSecondDownloadId = vbNullString
        m_downloadSecondNavigationId = vbNullString

        If TryGetDownloadCorrelationValue(p, "download", secondDownloadId) Then
            m_downloadSecondCorrelationMode = DOWNLOAD_CORRELATION_UUID
            m_downloadSecondDownloadId = secondDownloadId
        ElseIf TryGetDownloadCorrelationValue(p, "navigation", secondNavigationId) Then
            m_downloadSecondCorrelationMode = DOWNLOAD_CORRELATION_CONTEXT_NAV
            m_downloadSecondNavigationId = secondNavigationId
        End If

        LogSignalLine "[DOWNLOAD-MULTIPLE-STARTED] ctx=" & eventCtx & _
                      " | count=" & CStr(m_downloadStartedCount) & _
                      " | secondCorrelation=" & m_downloadSecondCorrelationMode & _
                      " | secondDownload=" & IIf(Len(m_downloadSecondDownloadId) > 0, m_downloadSecondDownloadId, "(none)") & _
                      " | secondNavigation=" & IIf(Len(m_downloadSecondNavigationId) > 0, m_downloadSecondNavigationId, "(none)")
        Exit Sub
    End If

    If Not m_downloadArmed Then Exit Sub

    If StrComp(eventCtx, m_downloadOwnerContextId, vbBinaryCompare) <> 0 Then
        If Not m_downloadContextMissWarned Then
            m_downloadContextMissWarned = True
            LogSignalLine "[DOWNLOAD-CTX-MISS] expectedCtx=" & m_downloadOwnerContextId & _
                          " | actualCtx=" & eventCtx
        End If
        Exit Sub
    End If

    Dim downloadId As String
    Dim navigationId As String

    m_downloadCorrelationMode = DOWNLOAD_CORRELATION_UNCORRELATABLE
    m_downloadCorrelationDownloadId = vbNullString
    m_downloadCorrelationNavigationId = vbNullString

    ' Spec-preferred identity. Tested Edge currently omits this field, so the second branch is
    ' operationally important. No inference is made when both identities are unavailable.
    If TryGetDownloadCorrelationValue(p, "download", downloadId) Then
        m_downloadCorrelationMode = DOWNLOAD_CORRELATION_UUID
        m_downloadCorrelationDownloadId = downloadId
    ElseIf TryGetDownloadCorrelationValue(p, "navigation", navigationId) Then
        m_downloadCorrelationMode = DOWNLOAD_CORRELATION_CONTEXT_NAV
        m_downloadCorrelationNavigationId = navigationId
    Else
        ' v1 cannot safely pair downloadEnd when both protocol identities are absent. Keep the
        ' one-shot arm available for a later correlatable owner-context start; the raw event remains
        ' visible in diagnostics and no browser-side cause is inferred.
        m_downloadUncorrelatableCount = m_downloadUncorrelatableCount + 1
        If Len(m_downloadFirstUncorrelatableContext) = 0 Then
            m_downloadFirstUncorrelatableContext = eventCtx
        End If
        LogSignalLine "[DOWNLOAD-WILL-BEGIN-IGNORED] reason=uncorrelatable | ctx=" & eventCtx & _
                      " | count=" & CStr(m_downloadUncorrelatableCount)
        Exit Sub
    End If

    m_downloadAcceptedContextId = eventCtx
    m_downloadSuggestedFilename = GetPathString(json, "params.suggestedFilename")
    m_downloadUrl = GetPathString(json, "params.url")
    m_downloadStartedSeen = True
    m_downloadStartedCount = 1
    m_downloadTerminalEligible = True
    m_downloadArmed = False

    LogSignalLine "[DOWNLOAD-START-ACCEPTED] ctx=" & eventCtx & _
                  " | correlation=" & m_downloadCorrelationMode & _
                  " | download=" & IIf(Len(m_downloadCorrelationDownloadId) > 0, m_downloadCorrelationDownloadId, "(none)") & _
                  " | navigation=" & IIf(Len(m_downloadCorrelationNavigationId) > 0, m_downloadCorrelationNavigationId, "(none)")
End Sub

Private Sub HandleDownloadEndEvent(ByVal json As Object)
    Dim p As Object
    Set p = GetPathObject(json, "params")
    If p Is Nothing Then
        LogSignalLine "[DOWNLOAD-END-IGNORED] reason=missing-params"
        Exit Sub
    End If

    LogSignalLine "[DOWNLOAD-END] ctx=" & GetDownloadDiagnosticField(p, "context") & _
                  " | download=" & GetDownloadDiagnosticField(p, "download") & _
                  " | navigation=" & GetDownloadDiagnosticField(p, "navigation") & _
                  " | status=" & GetDownloadDiagnosticField(p, "status") & _
                  " | filepath=" & GetDownloadDiagnosticField(p, "filepath")

    ' MultipleDownloadsStarted is an operation-level ambiguity. The measured lifecycle after
    ' multiplicity is intentionally not modeled; raw end events remain visible but cannot restore
    ' a single successful transaction.
    If m_downloadMultipleStarted Then
        LogSignalLine "[DOWNLOAD-END-IGNORED] reason=multiple-starts"
        Exit Sub
    End If

    ' Raw end events remain diagnostic even when no accepted start exists. They must not satisfy
    ' a future transaction merely because they were delivered after the local arm boundary.
    If Not m_downloadStartedSeen Or Not m_downloadTerminalEligible Or m_downloadEndSeen Then Exit Sub

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then
        LogSignalLine "[DOWNLOAD-END-IGNORED] reason=missing-context"
        Exit Sub
    End If

    If StrComp(eventCtx, m_downloadAcceptedContextId, vbBinaryCompare) <> 0 Then
        LogSignalLine "[DOWNLOAD-END-CTX-MISS] expectedCtx=" & m_downloadAcceptedContextId & _
                      " | actualCtx=" & eventCtx
        Exit Sub
    End If

    Dim actualId As String
    Select Case m_downloadCorrelationMode
        Case DOWNLOAD_CORRELATION_UUID
            If Not TryGetDownloadCorrelationValue(p, "download", actualId) Then
                LogSignalLine "[DOWNLOAD-END-CORRELATION-MISS] mode=download | reason=missing-download-id"
                Exit Sub
            End If
            If StrComp(actualId, m_downloadCorrelationDownloadId, vbBinaryCompare) <> 0 Then
                LogSignalLine "[DOWNLOAD-END-CORRELATION-MISS] mode=download | expected=" & _
                              m_downloadCorrelationDownloadId & " | actual=" & actualId
                Exit Sub
            End If

        Case DOWNLOAD_CORRELATION_CONTEXT_NAV
            If Not TryGetDownloadCorrelationValue(p, "navigation", actualId) Then
                LogSignalLine "[DOWNLOAD-END-CORRELATION-MISS] mode=context+navigation | reason=missing-navigation"
                Exit Sub
            End If
            If StrComp(actualId, m_downloadCorrelationNavigationId, vbBinaryCompare) <> 0 Then
                LogSignalLine "[DOWNLOAD-END-CORRELATION-MISS] mode=context+navigation | expected=" & _
                              m_downloadCorrelationNavigationId & " | actual=" & actualId
                Exit Sub
            End If

        Case Else
            LogSignalLine "[DOWNLOAD-END-IGNORED] reason=uncorrelatable-start"
            Exit Sub
    End Select

    Dim terminalStatus As String
    If Not TryGetDownloadCorrelationValue(p, "status", terminalStatus) Then
        LogSignalLine "[DOWNLOAD-END-IGNORED] reason=missing-status"
        Exit Sub
    End If
    Select Case LCase$(terminalStatus)
        Case "complete", "canceled"
        Case Else
            LogSignalLine "[DOWNLOAD-END-IGNORED] reason=unknown-status | status=" & terminalStatus
            Exit Sub
    End Select

    m_downloadEndSeen = True
    m_downloadTerminalEligible = False
    m_downloadEndContextId = eventCtx
    m_downloadEndStatusDiagnostic = terminalStatus
    m_downloadEndFilePathDiagnostic = vbNullString
    Call TryGetDownloadCorrelationValue(p, "filepath", m_downloadEndFilePathDiagnostic)

    LogSignalLine "[DOWNLOAD-END-MATCHED] ctx=" & eventCtx & _
                  " | correlation=" & m_downloadCorrelationMode & _
                  " | status=" & m_downloadEndStatusDiagnostic & _
                  " | filepath=" & m_downloadEndFilePathDiagnostic
End Sub

Private Sub ResetDownloadCaptureState()
    m_downloadArmed = False
    m_downloadStartedSeen = False
    m_downloadEndSeen = False
    m_downloadOwnerContextId = vbNullString
    m_downloadAcceptedContextId = vbNullString
    m_downloadCorrelationMode = vbNullString
    m_downloadCorrelationDownloadId = vbNullString
    m_downloadCorrelationNavigationId = vbNullString
    m_downloadEndContextId = vbNullString
    m_downloadEndStatusDiagnostic = vbNullString
    m_downloadEndFilePathDiagnostic = vbNullString
    m_downloadContextMissWarned = False
    m_downloadDroppedEventsAtArm = 0
    m_downloadStartedCount = 0
    m_downloadMultipleStarted = False
    m_downloadSecondCorrelationMode = vbNullString
    m_downloadSecondDownloadId = vbNullString
    m_downloadSecondNavigationId = vbNullString
    m_downloadSuggestedFilename = vbNullString
    m_downloadUrl = vbNullString
    m_downloadTerminalEligible = False
    m_downloadUncorrelatableCount = 0
    m_downloadFirstUncorrelatableContext = vbNullString
End Sub

Private Sub ArmDownloadCapture(ByVal contextId As String)
    ResetDownloadCaptureState
    m_downloadOwnerContextId = contextId
    m_downloadDroppedEventsAtArm = CurrentDownloadRelevantLossCount()
    m_downloadArmed = True
    LogSignalLine "[DOWNLOAD-ARM] ctx=" & contextId & _
                  " | droppedBase=" & CStr(m_downloadDroppedEventsAtArm)
End Sub

Private Function DownloadDroppedEventsSinceArm() As Long
    Dim delta As Long
    delta = CurrentDownloadRelevantLossCount() - m_downloadDroppedEventsAtArm
    If delta < 0 Then delta = 0
    DownloadDroppedEventsSinceArm = delta
End Function

Private Function DownloadEventQueueDepth() As Long
    ' Mechanical diagnostic only. A nonzero value means events remain in the bounded FIFO at
    ' the observation point; zero does not prove the observer was fully caught up.
    DownloadEventQueueDepth = socket_.EventQueueDepth
End Function

Private Sub RaiseDownloadMultipleStarted(ByVal methodName As String)
    m_downloadTerminalEligible = False
    m_downloadArmed = False
    Err.Raise ERR_DOWNLOAD_MULTIPLE_STARTED, "BiDi:" & methodName, _
              "MultipleDownloadsStarted: more than one download started in the owner context. " & _
              "The trigger is not replayed. FirstCorrelation=" & m_downloadCorrelationMode & _
              "; SecondCorrelation=" & m_downloadSecondCorrelationMode & "."
End Sub

Private Sub WaitForDownloadTransaction(ByVal timeoutMs As Long, ByVal methodName As String)
    Dim stageTick As Currency
    stageTick = socket_.NowTicks()

    ' Stage 1: wait for one correlatable owner-context start. The same public timeout value is
    ' deliberately reused for both stages so the API has one timeout knob while diagnostics retain
    ' distinct StartTimeout and CompletionTimeout outcomes.
    Do
        ProcessRecorderQueue

        If DownloadDroppedEventsSinceArm() > 0 Then
            m_downloadArmed = False
            m_downloadTerminalEligible = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                      "Async event loss occurred while waiting for download start; reliable " & _
                      "download transaction capture cannot be determined. The trigger click has " & _
                      "already been dispatched and is not replayed automatically. Do not blindly " & _
                      "retry the whole download operation because a download may already have " & _
                      "started. AsyncEventsDroppedSinceArm=" & _
                      CStr(DownloadDroppedEventsSinceArm()) & "."
        End If

        If m_downloadMultipleStarted Then RaiseDownloadMultipleStarted methodName
        If m_downloadStartedSeen Then Exit Do

        If IsTimeout(stageTick, CDbl(timeoutMs)) Then
            m_downloadArmed = False
            Err.Raise ERR_DOWNLOAD_START_TIMEOUT, "BiDi:" & methodName, _
                      "DownloadStartTimeout: no correlatable browsingContext.downloadWillBegin " & _
                      "was accepted within " & CStr(timeoutMs) & "ms. The trigger is not replayed. " & _
                      "UncorrelatableStarts=" & CStr(m_downloadUncorrelatableCount) & _
                      "; FirstUncorrelatableContext=" & _
                      IIf(Len(m_downloadFirstUncorrelatableContext) > 0, m_downloadFirstUncorrelatableContext, "(none)") & _
                      "; AsyncQueueDepthAtTimeout=" & CStr(DownloadEventQueueDepth()) & "."
        End If
        WaitHeartbeat 50
    Loop

    stageTick = socket_.NowTicks()

    ' Stage 2: wait only for the terminal event correlated to the accepted begin. A second owner
    ' start is an ambiguity and terminates the operation immediately rather than tracking both.
    Do
        ProcessRecorderQueue

        If DownloadDroppedEventsSinceArm() > 0 Then
            m_downloadArmed = False
            m_downloadTerminalEligible = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                      "Async event loss occurred while waiting for download completion; reliable " & _
                      "download transaction capture cannot be determined. The trigger click has " & _
                      "already been dispatched, and the download may already have completed or " & _
                      "left a file on disk. Do not blindly retry the whole download operation. " & _
                      "AsyncEventsDroppedSinceArm=" & CStr(DownloadDroppedEventsSinceArm()) & "."
        End If

        If m_downloadMultipleStarted Then RaiseDownloadMultipleStarted methodName
        If m_downloadEndSeen Then Exit Sub

        If IsTimeout(stageTick, CDbl(timeoutMs)) Then
            m_downloadTerminalEligible = False
            Err.Raise ERR_DOWNLOAD_COMPLETION_TIMEOUT, "BiDi:" & methodName, _
                      "DownloadCompletionTimeout: no matching browsingContext.downloadEnd was " & _
                      "accepted within " & CStr(timeoutMs) & "ms after the download start. " & _
                      "The trigger is not replayed. AsyncQueueDepthAtTimeout=" & _
                      CStr(DownloadEventQueueDepth()) & "."
        End If
        WaitHeartbeat 50
    Loop
End Sub

Private Function BuildDownloadResult() As Dictionary
    Dim result As New Dictionary
    result.Add "status", m_downloadEndStatusDiagnostic
    result.Add "filePath", m_downloadEndFilePathDiagnostic
    result.Add "suggestedFilename", m_downloadSuggestedFilename
    result.Add "url", m_downloadUrl
    result.Add "context", m_downloadAcceptedContextId
    result.Add "correlationMode", m_downloadCorrelationMode
    result.Add "downloadId", m_downloadCorrelationDownloadId
    result.Add "navigationId", m_downloadCorrelationNavigationId
    result.Add "asyncEventsDroppedSinceArm", DownloadDroppedEventsSinceArm()
    Set BuildDownloadResult = result
End Function

' Sends the global browser.setDownloadBehavior command. This v1 helper intentionally does not
' expose userContexts because that scope remains unmeasured in the current Download evidence set.
Private Sub SetGlobalDownloadBehavior(ByVal behaviorType As String, _
                                      Optional ByVal destinationFolder As String = vbNullString, _
                                      Optional ByVal clearBehavior As Boolean = False)
    Dim params As New Dictionary

    If clearBehavior Then
        params.Add "downloadBehavior", Null
    Else
        Dim behavior As New Dictionary
        behavior.Add "type", behaviorType
        If StrComp(behaviorType, "allowed", vbBinaryCompare) = 0 Then
            behavior.Add "destinationFolder", destinationFolder
        End If
        params.Add "downloadBehavior", behavior
    End If

    ' Set conservatively before dispatch. A definite protocol rejection restores the prior
    ' state; only an ambiguous outcome keeps the browser-side state marked unknown.
    Dim previousCleanupRequired As Boolean
    Dim previousStateUnknown As Boolean
    previousCleanupRequired = m_downloadBehaviorCleanupRequired
    previousStateUnknown = m_downloadBehaviorStateUnknown

    m_downloadBehaviorCleanupRequired = True
    m_downloadBehaviorStateUnknown = True

    On Error GoTo Failed

    Dim strRes As String
    strRes = ExecuteBiDiCommand("browser.setDownloadBehavior", params, True, _
                                allowAmbiguousRetry:=False, allowContextRetarget:=False)

    Dim parsed As Object
    Set parsed = ValidateProtocolResponse(strRes, "browser.setDownloadBehavior")

    m_downloadBehaviorStateUnknown = False
    If clearBehavior Then m_downloadBehaviorCleanupRequired = False
    Exit Sub

Failed:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrSrc As String: savedErrSrc = Err.source
    Dim savedErrDesc As String: savedErrDesc = Err.description

    If IsUnsupportedDownloadBehaviorError(savedErrNum, savedErrSrc, savedErrDesc) Then
        m_downloadBehaviorCleanupRequired = False
        m_downloadBehaviorStateUnknown = False
    ElseIf Not IsAmbiguousDispatchFailure(savedErrNum, savedErrSrc, savedErrDesc) Then
        m_downloadBehaviorCleanupRequired = previousCleanupRequired
        m_downloadBehaviorStateUnknown = previousStateUnknown
    End If

    Err.Raise savedErrNum, savedErrSrc, savedErrDesc
End Sub

' Sets the global BiDi download destination for the current browser session. The folder is resolved
' through WebShared so OneDrive/SharePoint-backed VBA projects use their local synchronized path.
Public Sub SetDownloadFolder(ByVal folderPath As String)
    Dim resolvedFolder As String
    resolvedFolder = WebShared.getFullLocalPath(folderPath, targetExists:=True)

    Dim fso As New FileSystemObject
    If Not fso.FolderExists(resolvedFolder) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:SetDownloadFolder", _
                  "Download destination folder does not exist: " & resolvedFolder
    End If

    SetGlobalDownloadBehavior "allowed", resolvedFolder
    If m_debugMode Then Debug.Print "BiDi: Download behavior set to allowed. Destination=" & resolvedFolder
End Sub

' Denies downloads globally for the current browser session. A resulting canceled download is still
' reported by ExecuteDownloadByXPath as the browser-authoritative terminal status "canceled".
Public Sub SetDownloadDenied()
    SetGlobalDownloadBehavior "denied"
    If m_debugMode Then Debug.Print "BiDi: Download behavior set to denied."
End Sub

' Clears this wrapper's global BiDi download behavior override and restores browser/default policy.
Public Sub ClearDownloadBehavior()
    SetGlobalDownloadBehavior vbNullString, clearBehavior:=True
    If m_debugMode Then Debug.Print "BiDi: Download behavior cleared."
End Sub

Private Function ClearDownloadBehaviorBestEffort(Optional ByVal emitWarning As Boolean = True) As Boolean
    If Not m_downloadBehaviorCleanupRequired Then
        ClearDownloadBehaviorBestEffort = True
        Exit Function
    End If

    Dim localErrNum As Long
    Dim localErrDesc As String

    On Error Resume Next
    Err.Clear
    ClearDownloadBehavior
    localErrNum = Err.Number
    localErrDesc = Err.description
    Err.Clear
    On Error GoTo 0

    If localErrNum = 0 Then
        ClearDownloadBehaviorBestEffort = True
    ElseIf emitWarning And m_debugMode Then
        Debug.Print "BiDi CRITICAL: browser.setDownloadBehavior cleanup could not be confirmed. " & _
                    "Reconnect to the same browser session and clear the behavior before further downloads. " & _
                    "Error #" & CStr(localErrNum) & ": " & localErrDesc
    End If
End Function

' Executes exactly one trigger click and observes exactly one correlatable owner-context download.
' complete/canceled are browser-reported terminal outcomes returned in result("status"). filePath is
' browser-reported only; this method does not assert filesystem existence, size, hash or rename stability.
Public Function ExecuteDownloadByXPath(ByVal triggerXPath As String, _
                                       Optional ByVal searchTimeoutMs As Long = 10000, _
                                       Optional ByVal timeoutMs As Long = DEFAULT_DOWNLOAD_TIMEOUT_MS, _
                                       Optional ByVal contextId As String = "") As Dictionary
    Const METHOD_NAME As String = "ExecuteDownloadByXPath"

    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeoutMs, "timeoutMs", DEFAULT_DOWNLOAD_TIMEOUT_MS, 1)

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    If m_downloadBehaviorStateUnknown Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & METHOD_NAME, _
                  "DownloadBehaviorStateUnknown: browser.setDownloadBehavior cleanup could not " & _
                  "be confirmed. Call SetDownloadFolder, SetDownloadDenied, or " & _
                  "ClearDownloadBehavior before starting another download."
    End If

    Dim carryoverDiagnosticActive As Boolean
    Dim carryDomSeenBefore As Boolean
    Dim carryNetSeenBefore As Boolean
    Dim carryVisSeenBefore As Boolean
    carryoverDiagnosticActive = BeginSignalCarryoverDiagnostic( _
        METHOD_NAME, targetCtx, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore)

    ' Lazy subscription keeps download traffic out of sessions that never use the feature. The
    ' registry makes repeated calls delta-only and therefore cheap.
    Me.ExecuteSessionSubscribe Array( _
        "browsingContext.downloadWillBegin", _
        "browsingContext.downloadEnd"), , True

    ' Resolve the trigger before the local arm so element lookup cannot let an unrelated buffered
    ' owner-context event satisfy this operation.
    Dim triggerSharedId As String
    triggerSharedId = ResolveXPathOrThrow(triggerXPath, targetCtx, searchTimeoutMs, METHOD_NAME)

    ' Best-effort stale boundary shared with the hardened file-dialog path: drain first, snapshot
    ' cumulative drops in ArmDownloadCapture, then issue the side-effecting trigger exactly once.
    ProcessRecorderQueue
    ArmDownloadCapture targetCtx

    Dim clickArgs As Collection
    Set clickArgs = NewElementArgs(triggerSharedId)

    Dim clickResult As String
    Dim clickErrNum As Long
    Dim clickErrSrc As String
    Dim clickErrDesc As String

    On Error Resume Next
    clickResult = PerformTrustedClick(targetCtx, clickArgs, vTimeout)
    clickErrNum = Err.Number
    clickErrSrc = Err.source
    clickErrDesc = Err.description
    On Error GoTo 0

    If clickErrNum <> 0 Then
        ' The dispatch outcome may be ambiguous, so never re-click. Close the local arm before
        ' propagating the original transport/protocol/action error.
        m_downloadArmed = False
        m_downloadTerminalEligible = False
        EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
        Err.Raise clickErrNum, clickErrSrc, clickErrDesc
    End If

    On Error GoTo DownloadWaitFailed
    WaitForDownloadTransaction vTimeout, METHOD_NAME
    On Error GoTo 0
    EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    Set ExecuteDownloadByXPath = BuildDownloadResult()
    Exit Function

DownloadWaitFailed:
    Dim waitErrNum As Long: waitErrNum = Err.Number
    Dim waitErrSrc As String: waitErrSrc = Err.source
    Dim waitErrDesc As String: waitErrDesc = Err.description
    EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    Err.Raise waitErrNum, waitErrSrc, waitErrDesc
End Function

Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeoutMs As Long = 5000) As String
    Dim vTimeout As Long
    If timeoutMs = 0 Then
        vTimeout = 0
    Else
        vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    End If

    Dim startTick As Currency: startTick = socket_.NowTicks()
    Dim ctxId As String, jsonTree As Object, contexts As Collection
    Dim pendingTransient As String, lastLoggedTransient As String

    Do
        Dim strRes As String
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)

        If TryGetPollingResponse(strRes, "browsingContext.getTree", jsonTree, _
                                 pendingTransient, lastLoggedTransient) Then
            Set contexts = GetPathObject(jsonTree, "result.contexts")
            If Not contexts Is Nothing Then
                ctxId = RecursiveFrameSearch(contexts, partialUrl)
                If ctxId <> "" Then
                    GetIframeContextIdByUrl = ctxId
                    Exit Function
                End If
            End If
        End If

        If vTimeout <= 0 Then Exit Do
        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(vTimeout)) Then
            RaiseIfPollingEndedOnProtocolError "browsingContext.getTree", pendingTransient
            Err.Raise ERR_ACTION_FAILED, "BiDi:GetIframeContextIdByUrl", _
                      "Iframe context not found within " & vTimeout & _
                      "ms (Target URL pattern: " & partialUrl & ")"
        End If
    Loop

    RaiseIfPollingEndedOnProtocolError "browsingContext.getTree", pendingTransient
End Function

Private Function RecursiveFrameSearch(ByVal contextList As Collection, _
                                      ByVal partialUrl As String, _
                                      Optional ByVal currentDepth As Long = 0) As String
    
    Const MAX_RECURSION_DEPTH As Long = 20
    
    On Error GoTo StackGuard
    
    If currentDepth > MAX_RECURSION_DEPTH Then
        Debug.Print "BiDi Warning: Max recursion depth reached (" & MAX_RECURSION_DEPTH & ")"
        RecursiveFrameSearch = ""
        Exit Function
    End If
    
    Dim ctx As Dictionary
    Dim childRes As String

    For Each ctx In contextList
        If InStr(ctx("url"), partialUrl) > 0 Then
            RecursiveFrameSearch = ctx("context")
            Exit Function
        End If

        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    childRes = RecursiveFrameSearch(ctx("children"), partialUrl, currentDepth + 1)
                    If childRes <> "" Then
                        RecursiveFrameSearch = childRes
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ctx

    RecursiveFrameSearch = ""
    Exit Function

StackGuard:
    If Err.Number = 28 Then ' Out of stack space
        Debug.Print "BiDi Critical: Stack overflow detected in RecursiveFrameSearch"
        RecursiveFrameSearch = ""
    Else
        Debug.Print "BiDi Error in RecursiveFrameSearch: " & Err.description
        RecursiveFrameSearch = ""
    End If
End Function

Public Function ExecuteFindWindowContextId(ByVal matchType As BiDiMatchBy, _
                                           ByVal matchValue As String, _
                                           Optional ByVal timeoutMs As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    If matchType <> MatchByUrl And matchType <> MatchByTitle Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteFindWindowContextId", _
                  "matchType must be MatchByUrl or MatchByTitle."
    End If

    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    Dim startTick As Currency: startTick = socket_.NowTicks()
    Dim strRes As String
    Dim jsonTree As Object, ctxList As Collection, ctx As Dictionary
    Dim pendingTransient As String, lastLoggedTransient As String
    Dim candidateCtx As String
    Dim candidateTitle As String
    Dim readyState As String
    Dim isMatch As Boolean

    Do
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)

        If TryGetPollingResponse(strRes, "browsingContext.getTree", jsonTree, _
                                 pendingTransient, lastLoggedTransient) Then
            Set ctxList = GetPathObject(jsonTree, "result.contexts")
            If Not ctxList Is Nothing Then
                For Each ctx In ctxList
                    isMatch = False
                    candidateCtx = vbNullString

                    If Not ctx.Exists("context") Then GoTo NextWindowContext
                    candidateCtx = CStr(ctx("context"))
                    If Len(candidateCtx) = 0 Then GoTo NextWindowContext

                    If matchType = MatchByUrl Then
                        If Not ctx.Exists("url") Then GoTo NextWindowContext
                        isMatch = (InStr(CStr(ctx("url")), matchValue) > 0)
                    Else
                        candidateTitle = vbNullString
                        If Not TryGetWindowContextTitle(candidateCtx, candidateTitle) Then
                            GoTo NextWindowContext
                        End If
                        isMatch = (InStr(candidateTitle, matchValue) > 0)
                    End If

                    If isMatch Then
                        If Not waitForComplete Then
                            ExecuteFindWindowContextId = candidateCtx
                            Exit Function
                        End If

                        readyState = vbNullString
                        If Not TryGetWindowContextReadyState(candidateCtx, readyState) Then
                            GoTo NextWindowContext
                        End If
                        If StrComp(readyState, "complete", vbBinaryCompare) = 0 Then
                            ExecuteFindWindowContextId = candidateCtx
                            Exit Function
                        End If
                    End If
NextWindowContext:
                Next ctx
            End If
        End If

        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(vTimeout)) Then
            RaiseIfPollingEndedOnProtocolError "browsingContext.getTree", pendingTransient
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteFindWindowContextId", _
                      "Target window not found within " & vTimeout & _
                      "ms (Match: " & matchValue & ")"
        End If
    Loop
End Function

Private Function TryGetWindowContextTitle(ByVal contextId As String, _
                                          ByRef titleText As String) As Boolean
    On Error GoTo Failed

    titleText = ExecuteGetTitleByContextId(contextId)
    TryGetWindowContextTitle = True
    Exit Function

Failed:
    Dim errNum As Long: errNum = Err.Number
    Dim errSrc As String: errSrc = Err.source
    Dim errDesc As String: errDesc = Err.description
    Err.Clear

    If IsTransientWindowContextReadError(errNum, errSrc, errDesc) Then
        If m_debugMode Then
            Debug.Print "BiDi: Transient window-context read skipped | phase=title | ctx=" & _
                        Left$(contextId, 16) & "..." & " | " & errDesc
        End If
        Exit Function
    End If

    Err.Raise errNum, errSrc, errDesc
End Function

Private Function TryGetWindowContextReadyState(ByVal contextId As String, _
                                               ByRef readyState As String) As Boolean
    On Error GoTo Failed

    readyState = CStr(GetCleanResult(CallScript( _
                    "return document.readyState;", , contextId, _
                    allowAmbiguousRetry:=True)))
    TryGetWindowContextReadyState = True
    Exit Function

Failed:
    Dim errNum As Long: errNum = Err.Number
    Dim errSrc As String: errSrc = Err.source
    Dim errDesc As String: errDesc = Err.description
    Err.Clear

    If IsTransientWindowContextReadError(errNum, errSrc, errDesc) Then
        If m_debugMode Then
            Debug.Print "BiDi: Transient window-context read skipped | phase=readyState | ctx=" & _
                        Left$(contextId, 16) & "..." & " | " & errDesc
        End If
        Exit Function
    End If

    Err.Raise errNum, errSrc, errDesc
End Function

Private Function IsTransientWindowContextReadError(ByVal errNum As Long, _
                                                   ByVal errSrc As String, _
                                                   ByVal errDesc As String) As Boolean
    If errNum <> ERR_BIDI_PROTOCOL And _
       errNum <> ERR_ACTION_FAILED And _
       errNum <> ERR_JS_EXCEPTION Then Exit Function

    Dim d As String: d = LCase$(errDesc)

    If InStr(d, "inspected target navigated or closed") > 0 Or _
       InStr(d, "execution context was destroyed") > 0 Or _
       InStr(d, "document was unloaded") > 0 Or _
       InStr(d, "target navigated") > 0 Or _
       InStr(d, "[no such frame]") > 0 Or _
       InStr(d, "[no such window]") > 0 Or _
       InStr(d, "cannot find context") > 0 Then
        IsTransientWindowContextReadError = True
        Exit Function
    End If

    IsTransientWindowContextReadError = _
        (InStr(d, "context") > 0) And _
        ((InStr(d, "destroyed") > 0) Or _
         (InStr(d, "gone") > 0) Or _
         (InStr(d, "closed") > 0) Or _
         (InStr(d, "not found") > 0))
End Function

Public Function ExecuteOpenNewContextByXPath(ByVal triggerXPath As String, _
                                              Optional ByVal searchTimeoutMs As Long = 10000, _
                                              Optional ByVal timeoutMs As Long = 10000, _
                                              Optional ByVal ownerContextId As String = "") As Dictionary
    Const METHOD_NAME As String = "ExecuteOpenNewContextByXPath"

    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 10000, 1)

    Dim ownerCtx As String
    ownerCtx = ResolveTargetContext(ownerContextId)

    Dim carryoverDiagnosticActive As Boolean
    Dim carryDomSeenBefore As Boolean
    Dim carryNetSeenBefore As Boolean
    Dim carryVisSeenBefore As Boolean
    carryoverDiagnosticActive = BeginSignalCarryoverDiagnostic( _
        METHOD_NAME, ownerCtx, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore)

    ' Resolve before arming so lookup cannot consume the capture boundary.
    Dim triggerSharedId As String
    triggerSharedId = ResolveXPathOrThrow(triggerXPath, ownerCtx, searchTimeoutMs, METHOD_NAME)

    ProcessRecorderQueue
    ArmNewContextCapture ownerCtx

    Dim clickArgs As Collection
    Set clickArgs = NewElementArgs(triggerSharedId)

    Dim clickResult As String
    Dim clickErrNum As Long
    Dim clickErrSrc As String
    Dim clickErrDesc As String

    On Error Resume Next
    clickResult = PerformTrustedClick(ownerCtx, clickArgs, vTimeout, False)
    clickErrNum = Err.Number
    clickErrSrc = Err.source
    clickErrDesc = Err.description
    On Error GoTo 0

    If clickErrNum <> 0 Then
        ResetNewContextCaptureState
        EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
        Err.Raise clickErrNum, clickErrSrc, clickErrDesc
    End If

    On Error GoTo CaptureFailed
    WaitForNewContextCapture vTimeout, METHOD_NAME
    On Error GoTo 0

    Dim result As Dictionary
    Set result = BuildNewContextResult()
    ResetNewContextCaptureState
    EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    Set ExecuteOpenNewContextByXPath = result
    Exit Function

CaptureFailed:
    Dim captureErrNum As Long: captureErrNum = Err.Number
    Dim captureErrSrc As String: captureErrSrc = Err.source
    Dim captureErrDesc As String: captureErrDesc = Err.description
    ResetNewContextCaptureState
    EndSignalCarryoverDiagnostic METHOD_NAME, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    Err.Raise captureErrNum, captureErrSrc, captureErrDesc
End Function

Public Function ExecuteCreateContext(Optional ByVal createType As BiDiContextType = ContextTypeTab) As String
    Dim typeStr As String
    If createType = ContextTypeWindow Then
        typeStr = "window"
    Else
        typeStr = "tab"
    End If

    Dim params As New Dictionary
    params.Add "type", typeStr
    
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.create", params, allowAmbiguousRetry:=False)
    
    Dim json As Object
    Set json = ValidateProtocolResponse(strRes, "browsingContext.create")
    ExecuteCreateContext = GetPathString(json, "result.context")
End Function

Public Function ExecuteCloseContext(ByVal contextId As String) As String
    Dim wasMain As Boolean
    wasMain = (m_mainContextState = MainContextPinned And _
               StrComp(contextId, m_mainContextId, vbBinaryCompare) = 0)

    Dim params As New Dictionary
    params.Add "context", contextId
    Dim closeRes As String
    closeRes = ExecuteBiDiCommand("browsingContext.close", params, _
                                  allowAmbiguousRetry:=False, allowRecoveryRetry:=False)
    ValidateProtocolResponse closeRes, "browsingContext.close"

    InvalidateContextState contextId, Not wasMain
    If wasMain Then SetMainContextUnbound

    ExecuteCloseContext = closeRes
End Function

Public Function ExecuteGetTitleByContextId(ByVal contextId As String) As String
    Dim strRes As String: strRes = CallScript("return document.title || '';", , contextId, _
                                             allowAmbiguousRetry:=True)
    
    ExecuteGetTitleByContextId = CStr(GetCleanResult(strRes))
End Function

Public Function ExecuteHandleUserPrompt(ByVal Accept As Boolean, Optional ByVal userText As String = "", Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "accept", Accept
    If Len(userText) > 0 Then params.Add "userText", userText

    On Error Resume Next
    ExecuteHandleUserPrompt = ExecuteBiDiCommand("browsingContext.handleUserPrompt", params, _
                                allowAmbiguousRetry:=False, allowContextRetarget:=(targetCtx = m_mainContextId))
    ValidateProtocolResponse ExecuteHandleUserPrompt, "browsingContext.handleUserPrompt"
    Dim capturedErr As Long: capturedErr = Err.Number
    Dim capturedDesc As String: capturedDesc = Err.description
    Dim capturedSrc As String: capturedSrc = Err.source
    On Error GoTo 0
    
    ' Ignore only the expected 'no such alert'; propagate transport and fatal errors.
    If capturedErr <> 0 Then
        If InStr(1, capturedDesc, "no such alert", vbTextCompare) = 0 And _
           InStr(1, capturedDesc, "no user prompt", vbTextCompare) = 0 Then
            Err.Raise capturedErr, capturedSrc, capturedDesc
        End If
    End If
End Function

Public Function ExecuteIsUrlContains(ByVal partialUrl As String, _
                                     Optional ByVal waitNetworkIdle As Boolean = True, _
                                     Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                     Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                     Optional ByVal contextId As String = "") As Boolean
    
    If waitNetworkIdle Then
        If minStableMs = 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
        If maxTimeoutMs = 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    End If
    ' Exactly 0ms is a deliberate one-shot read only when SPA synchronization is disabled.
    If Not waitNetworkIdle And maxTimeoutMs = 0 Then
        Dim immediateCtx As String
        immediateCtx = ResolveTargetContext(contextId)

        ExecuteIsUrlContains = _
            (InStr(GetCurrentUrl(immediateCtx, True), partialUrl) > 0)
        Exit Function
    End If

    Dim vMinStable As Long: vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long: vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Me.ExecuteSessionSubscribe Array("network.responseCompleted", "network.fetchError", "browsingContext.load", "browsingContext.fragmentNavigated")
    
    Dim waitCtx As New Dictionary
    waitCtx.Add "contextId", targetCtx
    waitCtx.Add "partialUrl", partialUrl
    
    Dim isMatched As Boolean
    isMatched = MasterWaitEngine(WaitMode_UrlContains, vMaxTimeout, 200, waitCtx)
    
    If isMatched And waitNetworkIdle Then
        SyncAfterAction targetCtx, vMinStable, vMaxTimeout
    End If

    ExecuteIsUrlContains = isMatched
End Function

' The protocol transport uses a 60-second floor. Because the communicator hard deadline
' is twice the command timeout, a silent live connection may block one navigate call
' for about 120 seconds before the transport reports failure.
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String

    Dim vMinStable As Long:  vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long: vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim attempt As Long
    ' A committed navigation is never resent. Retries are limited to explicitly classified
    ' ordinary-navigation network failures and liveness verification.
    Const MAX_NAV_RETRIES As Long = 2
    
    Dim isNavigated As Boolean
    Dim modeStr As String
    If waitNetworkIdle Then
        modeStr = "complete"
    Else: modeStr = "none"
    End If

    Dim navigationAccepted As Boolean: navigationAccepted = False
    For attempt = 0 To MAX_NAV_RETRIES
        isNavigated = False
        
        If Not navigationAccepted Then
            On Error Resume Next
            Call ExecuteBrowsingContextNavigate(targetUrl, modeStr, vMinStable, vMaxTimeout, targetCtx)
            Dim navErr As Long: navErr = Err.Number
            Dim navDesc As String: navDesc = Err.description
            Dim navSrc As String: navSrc = Err.source
            On Error GoTo 0
            
            If navErr = ERR_CONSENSUS_HALT Then
                Err.Raise navErr, navSrc, navDesc
            End If
            
            ' Selected Chromium network failures are retryable only for ordinary page navigation.
            ' They do not prove the server saw no request; never rely on this policy for one-time,
            ' download, or otherwise side-effecting URLs.
            If navErr <> 0 Then
                If IsRetryableNavFailure(navDesc) Then
                    WaitHeartbeat 1000
                    GoTo NextNavAttempt   ' re-navigate next iteration (not accepted yet)
                End If
                ' All ambiguous or post-commit failures are terminal because resending may duplicate
                ' navigation or server-side work.
                ExecuteNavigateAndGetStatus = "500"
                Exit Function
            End If
            
            ' After commit, retry verification only; never issue another navigate command.
            navigationAccepted = True
        End If
        
        If Not waitNetworkIdle Then
            On Error Resume Next
            Dim curUrl As String: curUrl = GetCurrentUrl(targetCtx, True)
            On Error GoTo 0
            If Not (InStr(1, curUrl, "chrome-error", vbTextCompare) > 0 Or curUrl = "") Then
                isNavigated = True
            End If
        Else
            isNavigated = True
        End If
        
        If isNavigated Then
            ExecuteNavigateAndGetStatus = "200"
            Exit Function
        End If

        WaitHeartbeat 1000
NextNavAttempt:
    Next attempt
    
    ExecuteNavigateAndGetStatus = "500"
End Function

' browsingContext.traverseHistory returns before the specification guarantees that
' traversal has completed. The command result is therefore followed by the wrapper's
' bounded URL/SPA synchronization, and callers should still verify a destination-specific
' element when the restored page has a known completion condition.
Public Function ExecuteTraverseHistory(ByVal delta As Long, _
                                       Optional ByVal waitNetworkIdle As Boolean = True, _
                                       Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                       Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                       Optional ByVal contextId As String = "") As String

    ResetLastSyncState
    If delta = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteTraverseHistory", _
                  "delta must not be zero. Use -1 for Back or 1 for Forward."
    End If

    Dim vMinStable As Long
    Dim vMaxTimeout As Long
    Dim targetCtx As String
    Dim params As New Dictionary
    Dim strRes As String

    vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    targetCtx = ResolveTargetContext(contextId)

    params.Add "context", targetCtx
    params.Add "delta", delta

    ' History traversal is state-changing. A transport timeout, disconnect, or context-loss
    ' response can mean that the browser already accepted the traversal, so never resend it,
    ' recover it against another context, or retarget it to the current main context.
    strRes = ExecuteBiDiCommand( _
                "browsingContext.traverseHistory", _
                params, _
                allowAmbiguousRetry:=False, _
                allowRecoveryRetry:=False, _
                allowContextRetarget:=False)

    ValidateProtocolResponse strRes, "browsingContext.traverseHistory"

    ' Ordinary settle failure does not undo an accepted history traversal.
    ' TrySyncPhase propagates only ERR_CONSENSUS_HALT.
    Call TrySyncPhase(targetCtx, waitNetworkIdle, vMinStable, vMaxTimeout, True)

    ExecuteTraverseHistory = strRes
End Function

' Resource blocking deliberately uses CDP Network.setBlockedURLs.
' Fetch/BiDi interception would require per-request replies on the shared synchronous
' correlation path, starving SPA-probe polls and risking re-entry. Failing parser-blocking
' subresources can also stall load. Browser-side declarative blocking avoids both.
Private Function EnsureCdpBlockingInitialized() As Boolean
    Dim res As Object
    
    Set res = ExecuteCDPCommand("Network.enable", New Dictionary, True, True)
    EnsureCdpBlockingInitialized = res("Ok")
End Function

Public Function ExecuteEnableResourceBlocking(ByVal patterns As Variant) As String
    Dim params As New Dictionary
    Dim urlList As New Collection
    Dim item As Variant

    If IsArray(patterns) Then
        For Each item In patterns: urlList.Add CStr(item): Next
    Else
        urlList.Add CStr(patterns)
    End If
    params.Add "urls", urlList

    If Not EnsureCdpBlockingInitialized() Then
        ExecuteEnableResourceBlocking = "{""error"":""CDP_INIT_FAILURE"",""message"":""Failed to enable Network domain.""}"
        Exit Function
    End If

    Dim resObj As Object: Set resObj = ExecuteCDPCommand("Network.setBlockedURLs", params, True, True)
    
    If Not resObj("Ok") Then
        ExecuteEnableResourceBlocking = "{""error"":""CDP_FAILURE""}"
    Else
        ' Cache normalized block tokens because Chromium may omit ERR_BLOCKED_BY_CLIENT text
        ' from fetchError events.
        Set m_blockingPatterns = New Collection
        Dim pat As Variant
        For Each pat In urlList
            Dim token As String: token = CStr(pat)
            Do While Len(token) > 0 And (Left$(token, 1) = "*" Or Left$(token, 1) = " ")
                token = Mid$(token, 2)
            Loop
            Do While Len(token) > 0 And (Right$(token, 1) = "*" Or Right$(token, 1) = " ")
                token = Left$(token, Len(token) - 1)
            Loop
            If Len(token) > 0 Then m_blockingPatterns.Add token
        Next
        
        If m_debugMode Then
            Debug.Print "CDP ResourceBlocking: ENABLED. " & urlList.count & " URL pattern(s) installed."
        End If
        ExecuteEnableResourceBlocking = resObj("Json")
    End If
End Function

' Visibility means connected, style-visible, and non-zero geometry; it does not test
' viewport intersection or occlusion.
Public Function ExecuteIsElementVisible(ByVal xpath As String, _
                                        Optional ByVal searchTimeoutMs As Long = 10000, _
                                        Optional ByVal waitNetworkIdle As Boolean = False, _
                                        Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                        Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                        Optional ByVal contextId As String = "") As Boolean
    Dim vSearchTimeout As Long
    If searchTimeoutMs = 0 Then
        vSearchTimeout = 0
    Else
        vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    End If

    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = GetFindElementSharedIdByXPath(xpath, vSearchTimeout, targetCtx)
    If Len(sharedId) = 0 Then Exit Function

    Dim args As Collection: Set args = NewElementArgs(sharedId)
    Dim strRes As String
    strRes = ExecuteBaseAction(targetCtx, "visibility_check", args, waitNetworkIdle, _
                               SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS), _
                               True, DEFAULT_STALE_RETRIES, _
                               SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS))

    On Error Resume Next
    ExecuteIsElementVisible = CBool(GetCleanResult(strRes))
    On Error GoTo 0
End Function

' Reads normalized textContent; an absent element remains an action error.
Public Function ExecuteGetTextByXPath(ByVal xpath As String, _
                                      Optional ByVal searchTimeoutMs As Long = 10000, _
                                      Optional ByVal contextId As String = "") As String
    ExecuteGetTextByXPath = RequireStringResult( _
        ReadElementValueByXPath(xpath, "get_text", searchTimeoutMs, contextId, _
                                "ExecuteGetTextByXPath"), _
        "ExecuteGetTextByXPath", "text")
End Function

' Reads getAttribute(name); missing returns Null and empty remains empty.
Public Function ExecuteGetAttributeByXPath(ByVal xpath As String, _
                                           ByVal attributeName As String, _
                                           Optional ByVal searchTimeoutMs As Long = 10000, _
                                           Optional ByVal contextId As String = "") As Variant
    ValidateDomMemberName attributeName, "attributeName", "ExecuteGetAttributeByXPath"
    ExecuteGetAttributeByXPath = RequireScalarResult( _
        ReadElementValueByXPath(xpath, "get_attribute", searchTimeoutMs, contextId, _
                                "ExecuteGetAttributeByXPath", attributeName), _
        "ExecuteGetAttributeByXPath", _
        "The attribute result was not a scalar value or Null.")
End Function

' Reads a scalar DOM property (String, Boolean, numeric, or Null).
Public Function ExecuteGetPropertyByXPath(ByVal xpath As String, _
                                          ByVal propertyName As String, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal contextId As String = "") As Variant
    ValidateDomMemberName propertyName, "propertyName", "ExecuteGetPropertyByXPath"
    ExecuteGetPropertyByXPath = RequireScalarResult( _
        ReadElementValueByXPath(xpath, "get_property", searchTimeoutMs, contextId, _
                                "ExecuteGetPropertyByXPath", propertyName), _
        "ExecuteGetPropertyByXPath", _
        "The property result was not a supported scalar value or Null.")
End Function

' Returns outerHTML for the first element matching XPath.
Public Function ExecuteGetOuterHtmlByXPath(ByVal xpath As String, _
                                           Optional ByVal searchTimeoutMs As Long = 10000, _
                                           Optional ByVal contextId As String = "") As String
    ExecuteGetOuterHtmlByXPath = RequireStringResult( _
        ReadElementValueByXPath(xpath, "get_outer_html", searchTimeoutMs, contextId, _
                                "ExecuteGetOuterHtmlByXPath"), _
        "ExecuteGetOuterHtmlByXPath", "outerHTML")
End Function

Private Function NewLocalValue(ByVal valueType As String, ByVal value As Variant) As Dictionary
    Dim item As New Dictionary
    item.Add "type", valueType: item.Add "value", value
    Set NewLocalValue = item
End Function

Private Function NewElementReference(ByVal sharedId As String) As Dictionary
    Dim item As New Dictionary: item.Add "sharedId", sharedId
    Set NewElementReference = item
End Function

Private Function NewElementArgs(ByVal sharedId As String, _
                                Optional ByVal stringValue As Variant) As Collection
    Dim args As New Collection
    args.Add NewElementReference(sharedId)
    If Not IsMissing(stringValue) Then args.Add NewLocalValue("string", CStr(stringValue))
    Set NewElementArgs = args
End Function

' Rebuild element arguments with a fresh sharedId while preserving remaining values.
Private Function ReplaceFirstElementReference(ByVal sourceArgs As Collection, _
                                              ByVal sharedId As String) As Collection
    Dim rebuilt As New Collection
    Dim i As Long

    rebuilt.Add NewElementReference(sharedId)

    For i = 2 To sourceArgs.count
        If IsObject(sourceArgs(i)) Then
            Dim itemObject As Object
            Set itemObject = sourceArgs(i)
            rebuilt.Add itemObject
        Else
            rebuilt.Add sourceArgs(i)
        End If
    Next i

    Set ReplaceFirstElementReference = rebuilt
End Function

' Only controlled sentinels or protocol-scoped stale errors may trigger re-resolution.
Private Function IsStaleElementFailure(ByVal errNum As Long, _
                                       ByVal errSrc As String, _
                                       ByVal errDesc As String) As Boolean

    If errNum = ERR_ACTION_FAILED Then
        Dim isShellJsError As Boolean
        isShellJsError = _
            (Left$(errSrc, Len("BiDi:JSError:")) = "BiDi:JSError:") Or _
            (Left$(errSrc, Len("BiDi:SafetyTimeout:")) = "BiDi:SafetyTimeout:")

        If isShellJsError Then
            If errDesc = "STALE_ELEMENT" Or _
               errDesc = "STALE_ELEMENT_PRE_INPUT" Then
                IsStaleElementFailure = True
                Exit Function
            End If
        End If
    End If

    If errNum <> ERR_BIDI_PROTOCOL Then Exit Function
    If Left$(errSrc, Len("BiDi:Protocol:")) <> "BiDi:Protocol:" Then Exit Function

    Dim d As String
    d = LCase$(errDesc)

    If InStr(d, "stale element reference") > 0 Or _
       InStr(d, "no such node") > 0 Or _
       InStr(d, "node is detached from document") > 0 Or _
       InStr(d, "element is no longer connected") > 0 Then
        IsStaleElementFailure = True
    End If
End Function

Private Function ReadElementValueByXPath(ByVal xpath As String, _
                                         ByVal actionKey As String, _
                                         ByVal searchTimeoutMs As Long, _
                                         ByVal contextId As String, _
                                         ByVal callerName As String, _
                                         Optional ByVal memberValue As Variant) As Variant
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, callerName)

    Dim args As Collection
    If IsMissing(memberValue) Then
        Set args = NewElementArgs(sharedId)
    Else
        Set args = NewElementArgs(sharedId, CStr(memberValue))
    End If

    ReadElementValueByXPath = GetCleanResult(ExecuteBaseAction( _
        targetCtx, actionKey, args, False, DEFAULT_IDLE_STABLE_MS, True, _
        DEFAULT_STALE_RETRIES, DEFAULT_IDLE_TIMEOUT_MS, False, xpath, _
        searchTimeoutMs, callerName))
End Function

Private Sub ValidateDomMemberName(ByVal value As String, _
                                  ByVal parameterName As String, _
                                  ByVal callerName As String)
    If Len(value) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, parameterName & " must not be empty."
    End If
    If InStr(value, Chr$(0)) > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  parameterName & " contains a null byte (Chr$(0))."
    End If
End Sub

Private Function RequireStringResult(ByVal value As Variant, _
                                     ByVal callerName As String, _
                                     ByVal resultName As String) As String
    If IsEmpty(value) Or IsNull(value) Or IsObject(value) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "The " & resultName & " result was not a string value."
    End If
    RequireStringResult = CStr(value)
End Function

Private Function RequireScalarResult(ByVal value As Variant, _
                                     ByVal callerName As String, _
                                     ByVal invalidMessage As String) As Variant
    If IsEmpty(value) Or IsObject(value) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, invalidMessage
    End If
    RequireScalarResult = value
End Function

Public Function ExecuteSessionSubscribe(ByVal events As Variant, _
                                         Optional ByVal contexts As Collection = Nothing, _
                                         Optional ByVal Silent As Boolean = False) As String
    Dim eventArray() As String
    Dim i As Long, v As Variant
    Dim lb As Long, ub As Long
    If IsArray(events) Then
        ' Normalize to zero-based; callers may pass Dim x(1 To n) or a negative lower bound.
        lb = LBound(events): ub = UBound(events)
        ReDim eventArray(0 To ub - lb)
        For i = lb To ub
            eventArray(i - lb) = CStr(events(i))
        Next i
    Else
        ReDim eventArray(0)
        eventArray(0) = CStr(events)
    End If

    ' Subscription registry keys include both event and context; an empty context denotes
    ' a global subscription.
    Dim targetCtxs As Collection: Set targetCtxs = New Collection
    Dim c As Variant
    If contexts Is Nothing Then
        targetCtxs.Add ""
    ElseIf contexts.count = 0 Then
        targetCtxs.Add ""
    Else
        For Each c In contexts
            targetCtxs.Add CStr(c)
        Next c
    End If

    Dim eventsToRequest As Collection, oneCtx As Collection
    Dim params As Dictionary
    Dim finalArr() As String
    Dim strRes As String, subId As String, ctxId As String
    Dim tc As Variant
    Dim lastRes As String: lastRes = "{""status"":""already_subscribed""}"

    ' Delta-subscribe each event/context pair independently to avoid partial-overlap gaps.
    For Each tc In targetCtxs
        ctxId = IIf(CStr(tc) = "", "", "|" & CStr(tc))

        Set eventsToRequest = New Collection
        For i = LBound(eventArray) To UBound(eventArray)
            If Not m_activeSubscriptions.Exists(eventArray(i) & ctxId) Then
                eventsToRequest.Add eventArray(i)
            End If
        Next i

        If eventsToRequest.count = 0 Then
            If m_debugMode And Not Silent Then Debug.Print "BiDi: Subscription skipped (already active)."
        Else
            ReDim finalArr(eventsToRequest.count - 1)
            i = 0
            For Each v In eventsToRequest
                finalArr(i) = CStr(v): i = i + 1
            Next v

            Set params = New Dictionary
            params.Add "events", finalArr
            If CStr(tc) <> "" Then
                Set oneCtx = New Collection
                oneCtx.Add CStr(tc)
                params.Add "contexts", oneCtx
            End If

            strRes = ExecuteBiDiCommand("session.subscribe", params, Silent, allowAmbiguousRetry:=False)
            lastRes = strRes

            Dim subscribeJson As Object
            Set subscribeJson = ValidateProtocolResponse(strRes, "session.subscribe")
            subId = GetPathString(subscribeJson, "result.subscription")
            If Len(subId) = 0 Then
                Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:session.subscribe", _
                          "Successful session.subscribe response omitted result.subscription."
            End If

            For i = 0 To UBound(finalArr)
                m_activeSubscriptions(finalArr(i) & ctxId) = subId
            Next i
        End If
    Next tc

    ExecuteSessionSubscribe = lastRes
End Function

' browsingContext.navigate waits to interactive at the protocol layer. waitMode controls
' only the additional bounded wrapper-side settlement after the command returns. The transport
' wait has an independent 60-second floor and grows beyond it when maxTimeoutMs requires more.
Private Function ExecuteBrowsingContextNavigate(ByVal url As String, _
                                    Optional ByVal waitMode As String = "complete", _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String
    ResetLastSyncState
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim preNavOrigin As String: preNavOrigin = socket_.ExtractOrigin(url)

    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", url
    params.Add "wait", "interactive"

    ' Explicit child/popup/frame targets must never be retargeted to the main context.
    Dim allowRetarget As Boolean: allowRetarget = (targetCtx = m_mainContextId)
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params, _
                                allowAmbiguousRetry:=False, _
                                allowContextRetarget:=allowRetarget, _
                                commandTimeoutMs:=ResolveNavigationTransportTimeout(maxTimeoutMs))
    Dim navJson As Object
    Set navJson = ValidateProtocolResponse(strRes, "browsingContext.navigate")

    Dim postNavOrigin As String
    postNavOrigin = socket_.ExtractOrigin(GetPathString(navJson, "result.url"))
    If Len(preNavOrigin) > 0 And Len(postNavOrigin) > 0 Then
        If LCase$(preNavOrigin) <> LCase$(postNavOrigin) Then
            m_cdpSessionId = ""
            If m_debugMode Then
                Debug.Print "BiDi: CDP session invalidated (cross-origin: " & _
                            preNavOrigin & " -> " & postNavOrigin & ")"
            End If
        End If
    End If

    If LCase$(waitMode) = "complete" Then
        SettleAfterNavigation targetCtx, minStableMs, maxTimeoutMs
    End If
    ExecuteBrowsingContextNavigate = strRes
End Function

Private Sub SettleAfterNavigation(ByVal contextId As String, _
                                  ByVal minStableMs As Long, _
                                  ByVal maxTimeoutMs As Long)
    m_lastSyncAttempted = True
    On Error GoTo Failed

    Dim urlStable As Boolean
    urlStable = WaitForUrlStability(contextId, 1000, maxTimeoutMs)
    UpdateMainContextId True, True
    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, contextId

    m_lastSyncSucceeded = urlStable
    If Not urlStable Then
        RecordSyncFailure 504, "LogicalTimeout", _
            "URL and readyState did not remain stable within " & maxTimeoutMs & "ms."
        If m_debugMode Then Debug.Print "BiDi: navigate URL-stability timeout (best-effort)."
    End If
    Exit Sub

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    RecordSyncFailure errorNumber, errorSource, errorDescription
    On Error GoTo 0

    If errorNumber = ERR_CONSENSUS_HALT Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    If m_debugMode Then Debug.Print "BiDi: navigate settle best-effort skip: " & errorDescription
End Sub

Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal checkUrl As Boolean = False) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteClickByXPath")

    Dim args As Collection: Set args = NewElementArgs(sharedId)
    ExecuteClickByXPath = ExecuteBaseAction(targetCtx, "click", args, waitNetworkIdle, _
                                            minStableMs, True, DEFAULT_STALE_RETRIES, _
                                            maxTimeoutMs, checkUrl, xpath, searchTimeoutMs, _
                                            "ExecuteClickByXPath")
End Function

Public Function ExecuteInputValueByXPath(ByVal xpath As String, _
                                         ByVal valueToSet As String, _
                                         Optional ByVal searchTimeoutMs As Long = 10000, _
                                         Optional ByVal waitNetworkIdle As Boolean = True, _
                                         Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                         Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                         Optional ByVal contextId As String = "", _
                                         Optional ByVal useKeyEvents As Boolean = True, _
                                         Optional ByVal useNativeKeys As Boolean = False) As String
    ResetLastSyncState
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteInputValueByXPath")

    Dim args As Collection: Set args = NewElementArgs(sharedId, valueToSet)
    If useNativeKeys Then
        Dim focusArgs As Collection: Set focusArgs = NewElementArgs(sharedId)
        Call ExecuteBaseAction(targetCtx, "click", focusArgs, False, minStableMs, _
                               True, DEFAULT_STALE_RETRIES, maxTimeoutMs, False, _
                               xpath, searchTimeoutMs, "ExecuteInputValueByXPath")
        ExecuteInputValueByXPath = PerformNativeKeyInput( _
            targetCtx, valueToSet, maxTimeoutMs, waitNetworkIdle, minStableMs)
        If m_debugMode Then Debug.Print "BiDi: Action 'native_keys' completed."
    ElseIf Not useKeyEvents Then
        ExecuteInputValueByXPath = ExecuteBaseAction( _
            targetCtx, "input", args, waitNetworkIdle, minStableMs, True, _
            DEFAULT_STALE_RETRIES, maxTimeoutMs, False, xpath, searchTimeoutMs, _
            "ExecuteInputValueByXPath")
    Else
        Dim keySafetyMs As Long: keySafetyMs = maxTimeoutMs
        If keySafetyMs < 20000 Then keySafetyMs = 20000

        Dim keyResult As String
        keyResult = ExecuteBaseAction( _
            targetCtx, "input_keys", args, False, minStableMs, False, _
            DEFAULT_STALE_RETRIES, keySafetyMs, False, xpath, searchTimeoutMs, _
            "ExecuteInputValueByXPath")

        Call TrySyncPhase(targetCtx, waitNetworkIdle, _
                          SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS), _
                          SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS), False)
        ExecuteInputValueByXPath = keyResult
        If m_debugMode Then Debug.Print "BiDi: Action 'input_keys' completed."
    End If
End Function

Public Function ExecuteSetFilesByXPath(ByVal xpath As String, _
                                       ByVal filePaths As Variant, _
                                       Optional ByVal searchTimeoutMs As Long = 10000, _
                                       Optional ByVal waitNetworkIdle As Boolean = True, _
                                       Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                       Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                       Optional ByVal contextId As String = "") As String
    ResetLastSyncState
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteSetFilesByXPath")
    
    Dim paths As Variant
    paths = NormalizeFilePaths(filePaths)
    
    ExecuteSetFilesByXPath = PerformSetFiles(targetCtx, sharedId, paths, maxTimeoutMs, waitNetworkIdle, minStableMs)
    If m_debugMode Then Debug.Print "BiDi: Action 'set_files' completed."
End Function

' Short-lived input[type=file] path.
'
' PUBLIC CONTRACT
' ---------------
' Name: ExecuteSetFileSelectionViaDialog
'
' "FileSelection" is intentional terminology. input.setFiles changes the selection state of
' an HTML file input; it does not itself mean that an application upload has completed.
' The caller SHOULD arm ArmNetworkSignal, ArmContentSignal, and/or ArmVisibilitySignal first;
' the selected signal(s) describe the application-level completion condition after setFiles.
'
' Arming is deliberately NOT enforced. Without an armed signal the final wait falls back to
' idle-only consensus, which cannot distinguish "the application finished" from "the application
' ignored the change" -- both are quiet. That weakness is reported rather than pre-empted:
' [SPA-WAIT-BEGIN]/[SPA-WAIT-END] simply omit the signal= field for an ungated wait, and the
' [SPA-ACT] burst recorded after setFiles lists the mutation paths this site actually produces.
' An unarmed run is therefore the discovery run: it completes, and its log yields the xpath to
' arm next time (DOM: -> ArmContentSignal, ATTR:..[style] -> ArmVisibilitySignal). Refusing to
' run would withhold exactly the evidence the caller needs to write the signal.
'
' Native picker suppression is Chromium-specific and deliberately narrow:
'   CDP Page.setInterceptFileChooserDialog(enabled=True, cancel=False)
' The CDP session is resolved from the SAME BiDi browsing context as triggerXPath and is held
' for cleanup so iframe/popup contexts do not silently fall back to the main-context session.
'
' File-dialog observation and file selection remain WebDriver BiDi:
'   input.fileDialogOpened -> event.element.sharedId -> input.setFiles
'
' The trigger always uses a trusted BiDi pointer action; UseTrustedClick does not change this
' composite operation. No SPA wait occurs after the trigger click. Final synchronization starts
' only after input.setFiles succeeds and after the native-picker interceptor is released.
'
' RETRY CONTRACT
' --------------
' The trigger click and input.setFiles are state-changing and are never replayed internally after
' an ambiguous outcome. If the final completion wait fails after input.setFiles succeeds, the
' raised error is prefixed PostSelectionCompletionFailed and explicitly warns that selection may
' already have started application work; callers must not blindly retry the whole operation.
Public Function ExecuteSetFileSelectionViaDialog(ByVal triggerXPath As String, _
                                                  ByVal filePaths As Variant, _
                                                  Optional ByVal searchTimeoutMs As Long = 10000, _
                                                  Optional ByVal fileDialogTimeoutMs As Long = 10000, _
                                                  Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                                  Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                                  Optional ByVal contextId As String = "") As String
    ResetLastSyncState

    On Error GoTo Failed

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    ' Ownership is checked only when a signal is actually armed. With no arm, an empty owner is
    ' the normal ungated state, not a mismatch.
    If AnySignalArmed() Then
        If Len(m_signalOwnerContextId) = 0 Or _
           StrComp(targetCtx, m_signalOwnerContextId, vbBinaryCompare) <> 0 Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                      "CompletionSignalContextMismatch: The active completion signal does not belong " & _
                      "to the file-selection context." & vbCrLf & _
                      "Signal owner: " & m_signalOwnerContextId & vbCrLf & _
                      "Target: " & targetCtx
        End If
    End If

    ' Never begin another chooser operation while a previous interception may still be active.
    If m_fileChooserInterceptMayBeActive Then
        If Not ReleaseTrackedFileChooserInterceptionBestEffort(True) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                      "ResidualFileChooserInterception: a previous native file-chooser " & _
                      "interception could not be confirmed disabled. Call Shutdown or " & _
                      "reconnect before starting another file-selection operation."
        End If
    End If

    Dim paths As Variant
    paths = NormalizeFilePaths(filePaths)

    Dim pathCount As Long
    pathCount = VariantArrayCount(paths)
    If pathCount = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "NoFilesSupplied: No file paths were supplied for file selection."
    End If

    Dim vDialogTimeout As Long
    Dim vMinStable As Long
    Dim vMaxTimeout As Long
    vDialogTimeout = SanitizeMs(fileDialogTimeoutMs, "fileDialogTimeoutMs", 10000)
    vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    ' Subscribe before enabling interception and before the trigger click. Subscription remains
    ' session-scoped; the one-shot local arm below decides which event this operation consumes.
    Me.ExecuteSessionSubscribe Array("input.fileDialogOpened"), , True

    ' Resolve the trigger before the local capture arm. A slow XPath lookup must not create a
    ' window in which an unrelated chooser can satisfy this operation.
    Dim triggerSharedId As String
    triggerSharedId = ResolveXPathOrThrow(triggerXPath, targetCtx, searchTimeoutMs, _
                                          "ExecuteSetFileSelectionViaDialog")

    ' Resolve a fresh CDP session for the SAME BiDi context as the trigger, then hold that session
    ' id for cleanup. This matters for out-of-process frames and non-main top-level contexts.
    Dim fileChooserCdpSessionId As String
    fileChooserCdpSessionId = GetCdpSessionIdForContext(targetCtx)
    BeginNativeFileChooserInterception targetCtx, fileChooserCdpSessionId

    ' Dispatch already-buffered events immediately before the local arm. CDP resolution and
    ' interception above can perform command round-trips that preserve asynchronous events in
    ' the socket FIFO. Process them while the local arm is still down so ordinary stale
    ' fileDialogOpened events are rejected.
    ' This deliberately remains a best-effort drain rather than introducing a queue watermark.
    ProcessRecorderQueue

    ArmFileDialogCapture targetCtx

    Dim clickArgs As Collection
    Set clickArgs = NewElementArgs(triggerSharedId)

    ' File-picker triggers require real user activation. PerformTrustedClick uses BiDi
    ' input.performActions and does not replay an ambiguous click.
    Dim clickResult As String
    clickResult = PerformTrustedClick(targetCtx, clickArgs, vMaxTimeout, False)

    If Not WaitForFileDialogCapture(vDialogTimeout) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "FileDialogTimeout: input.fileDialogOpened was not observed within " & _
                  vDialogTimeout & "ms."
    End If

    If Not m_fileDlgMultipleKnown Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "FileDialogMissingMultiple: input.fileDialogOpened did not include the " & _
                  "required 'multiple' field."
    End If

    If Not m_fileDlgHasElement Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "FileDialogHasNoInputElement: input.fileDialogOpened had no HTML " & _
                  "file-input element. Elementless picker paths are outside this API contract."
    End If

    If pathCount > 1 And Not m_fileDlgMultiple Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "MultipleFilesNotAllowed: the opened file input does not allow multiple files."
    End If

    Dim capturedSharedId As String
    capturedSharedId = m_fileDlgSharedId

    ' Send input.setFiles first, but deliberately postpone the required application-completion wait.
    ' The interceptor is needed only through the chooser/setFiles hand-off.
    Dim setFilesResult As String
    setFilesResult = PerformSetFiles(targetCtx, capturedSharedId, paths, _
                                     vMaxTimeout, False, vMinStable)

    Dim selectionApplied As Boolean
    selectionApplied = True

    ResetFileDialogCaptureState

    ' input.setFiles is already state-changing, so cleanup failure is surfaced as NON-RETRYABLE.
    ' Keep the tracked residual state if both the held session and a fresh session fail to disable.
    If Not ReleaseTrackedFileChooserInterceptionBestEffort(True) Then
        Me.DisarmSignal
        Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "PostSelectionInterceptionCleanupFailed: input.setFiles succeeded, but native " & _
                  "file-chooser interception could not be confirmed disabled. The selection may " & _
                  "already have triggered application work; do not retry the whole operation. " & _
                  "Call Shutdown or reconnect before further file-chooser use."
    End If

    RunRequiredSpaSync targetCtx, vMinStable, vMaxTimeout
    ExecuteSetFileSelectionViaDialog = setFilesResult

    If m_debugMode Then Debug.Print "BiDi: Action 'set_file_selection_via_dialog' completed."
    Exit Function

Failed:
    Dim primaryErrNum As Long
    Dim primaryErrSrc As String
    Dim primaryErrDesc As String
    primaryErrNum = Err.Number
    primaryErrSrc = Err.source
    primaryErrDesc = Err.description

    ' Abnormal-flow cleanup must never replace the primary error. The tracked interceptor state
    ' survives a failed disable so Shutdown / reconnect can retry it.
    Dim cleanupConfirmed As Boolean
    On Error Resume Next
    ResetFileDialogCaptureState
    cleanupConfirmed = ReleaseTrackedFileChooserInterceptionBestEffort(True)
    Me.DisarmSignal
    On Error GoTo 0

    If primaryErrNum = 0 Then primaryErrNum = ERR_ACTION_FAILED

    If selectionApplied And InStr(1, primaryErrDesc, "PostSelectionInterceptionCleanupFailed:", vbBinaryCompare) = 0 Then
        Dim underlyingSrc As String
        underlyingSrc = primaryErrSrc
        If Len(underlyingSrc) = 0 Then underlyingSrc = "(unknown)"

        primaryErrSrc = "BiDi:ExecuteSetFileSelectionViaDialog"
        primaryErrDesc = _
            "PostSelectionCompletionFailed: input.setFiles succeeded before the completion " & _
            "wait failed. File selection may already have triggered application work; do not " & _
            "blindly retry the whole file-selection operation." & vbCrLf & _
            "Underlying source: " & underlyingSrc & vbCrLf & _
            "Underlying error: " & primaryErrDesc
    ElseIf Len(primaryErrSrc) = 0 Then
        primaryErrSrc = "BiDi:ExecuteSetFileSelectionViaDialog"
    End If

    If m_fileChooserInterceptMayBeActive And Not cleanupConfirmed Then
        primaryErrDesc = primaryErrDesc & vbCrLf & _
            "InterceptionCleanupPending: native file-chooser interception may still be active. " & _
            "Call Shutdown or reconnect before further file-chooser use."
    End If

    Err.Raise primaryErrNum, primaryErrSrc, primaryErrDesc
End Function

Private Function NormalizeFilePaths(ByVal value As Variant) As Variant
    Dim source As Variant
    If IsArray(value) Then
        source = value
    Else
        source = Split(Replace(CStr(value), vbCrLf, vbLf), vbLf)
    End If

    Dim srcLb As Long
    Dim srcUb As Long
    If Not TryGetVariantArrayBounds(source, srcLb, srcUb) Then
        NormalizeFilePaths = Array()
        Exit Function
    End If

    Dim out() As String, count As Long, i As Long
    ReDim out(0 To srcUb - srcLb)
    For i = srcLb To srcUb
        Dim item As String: item = Trim$(CStr(source(i)))
        If Len(item) > 0 Then out(count) = item: count = count + 1
    Next i

    If count = 0 Then
        NormalizeFilePaths = Array()
    Else
        ReDim Preserve out(0 To count - 1)
        NormalizeFilePaths = out
    End If
End Function

Private Function TryGetVariantArrayBounds(ByVal values As Variant, _
                                          ByRef lb As Long, _
                                          ByRef ub As Long) As Boolean
    If Not IsArray(values) Then Exit Function

    On Error GoTo NoBounds
    lb = LBound(values)
    ub = UBound(values)
    TryGetVariantArrayBounds = (ub >= lb)
    Exit Function

NoBounds:
    Err.Clear
End Function

Private Function VariantArrayCount(ByVal values As Variant) As Long
    Dim lb As Long
    Dim ub As Long
    If TryGetVariantArrayBounds(values, lb, ub) Then
        VariantArrayCount = ub - lb + 1
    End If
End Function

Private Sub ArmFileDialogCapture(ByVal contextId As String)
    ResetFileDialogCaptureState
    m_fileDlgOwnerContextId = contextId
    m_fileDlgDroppedEventsAtArm = CurrentFileDialogRelevantLossCount()
    m_fileDlgArmed = True
    LogSignalLine "[FILE-DIALOG-ARM] ctx=" & contextId & _
                  " | droppedBase=" & CStr(m_fileDlgDroppedEventsAtArm)
End Sub

Private Function FileDialogDroppedEventsSinceArm() As Long
    Dim delta As Long
    delta = CurrentFileDialogRelevantLossCount() - m_fileDlgDroppedEventsAtArm
    If delta < 0 Then delta = 0
    FileDialogDroppedEventsSinceArm = delta
End Function

Private Function WaitForFileDialogCapture(ByVal timeoutMs As Long) As Boolean
    Dim startTick As Currency
    startTick = socket_.NowTicks()

    Do
        ProcessRecorderQueue

        If FileDialogDroppedEventsSinceArm() > 0 Then
            m_fileDlgArmed = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                      "Async event loss occurred after the file-dialog arm boundary; reliable " & _
                      "file-input capture cannot be determined. AsyncEventsDroppedSinceArm=" & _
                      CStr(FileDialogDroppedEventsSinceArm()) & "."
        End If

        If m_fileDlgSeen Then
            WaitForFileDialogCapture = True
            Exit Function
        End If

        If IsTimeout(startTick, CDbl(timeoutMs)) Then Exit Do
        WaitHeartbeat 50
    Loop
End Function

Private Sub ResetFileDialogCaptureState()
    m_fileDlgArmed = False
    m_fileDlgSeen = False
    m_fileDlgOwnerContextId = vbNullString
    m_fileDlgHasElement = False
    m_fileDlgSharedId = vbNullString
    m_fileDlgMultipleKnown = False
    m_fileDlgMultiple = False
    m_fileDlgContextMissWarned = False
    m_fileDlgDroppedEventsAtArm = 0
End Sub

Private Function GetCdpSessionIdForContext(ByVal contextId As String) As String
    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    Dim params As New Dictionary
    params.Add "context", targetCtx

    ' Always resolve this operation's session fresh. A cached main-context CDP session may be
    ' wrong for an OOPIF/popup, or stale after navigation performed outside this wrapper.
    Dim raw As String
    raw = ExecuteBiDiCommand("goog:cdp.getSession", params, True, _
                             allowAmbiguousRetry:=True, _
                             allowRecoveryRetry:=True, _
                             allowContextRetarget:=False)

    Dim json As Object
    Set json = ValidateProtocolResponse(raw, "goog:cdp.getSession")

    Dim sessionId As String
    sessionId = GetPathString(json, "result.session")
    If Len(sessionId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "NativeFileChooserInterceptionUnavailable: goog:cdp.getSession returned no " & _
                  "CDP session for context " & targetCtx & "."
    End If

    ' Keep the existing main-context cache coherent for unrelated CDP features.
    If StrComp(targetCtx, m_mainContextId, vbBinaryCompare) = 0 Then
        m_cdpSessionId = sessionId
        m_consecutiveCdpFailures = 0
    End If

    GetCdpSessionIdForContext = sessionId
End Function

Private Sub BeginNativeFileChooserInterception(ByVal contextId As String, _
                                               ByVal cdpSessionId As String)
    ' Mark MAY-BE-ACTIVE before the state-changing CDP command. If the command reaches Chromium
    ' but its response is lost, cleanup still knows that it must send an explicit disable.
    m_fileChooserInterceptMayBeActive = True
    m_fileChooserInterceptSessionId = cdpSessionId
    m_fileChooserInterceptContextId = contextId

    SetNativeFileChooserInterceptionOnSession True, cdpSessionId
End Sub

Private Sub SetNativeFileChooserInterceptionOnSession(ByVal enabled As Boolean, _
                                                       ByVal cdpSessionId As String, _
                                                       Optional ByVal logSuccess As Boolean = True)
    If Len(cdpSessionId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "NativeFileChooserInterceptionUnavailable: CDP session id is empty."
    End If

    Dim cdpParams As New Dictionary
    cdpParams.Add "enabled", enabled

    ' cancel=False is the experimentally verified coexistence mode:
    ' suppress the native picker UI without synthesizing a site-visible cancel.
    cdpParams.Add "cancel", False

    Dim res As Object
    Set res = ExecuteCDPCommand("Page.setInterceptFileChooserDialog", cdpParams, _
                                True, True, cdpSessionId, enabled)

    If res Is Nothing Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "NativeFileChooserInterceptionUnavailable: CDP " & _
                  "Page.setInterceptFileChooserDialog returned no result object."
    End If

    If Not CBool(res("Ok")) Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                  "NativeFileChooserInterceptionUnavailable: CDP " & _
                  "Page.setInterceptFileChooserDialog failed: " & CStr(res("Error"))
    End If

    If m_debugMode And logSuccess Then
        Debug.Print "BiDi: Native file chooser interception " & _
                    IIf(enabled, "ENABLED", "DISABLED") & _
                    " (cancel=False, session=" & Left$(cdpSessionId, 16) & "...)."
    End If
End Sub

Private Sub ClearTrackedFileChooserInterception()
    m_fileChooserInterceptMayBeActive = False
    m_fileChooserInterceptSessionId = vbNullString
    m_fileChooserInterceptContextId = vbNullString
End Sub

Private Function TryDisableFileChooserInterceptionOnSession(ByVal cdpSessionId As String, _
                                                             Optional ByVal logSuccess As Boolean = True) As Boolean
    If Len(cdpSessionId) = 0 Then Exit Function

    On Error GoTo Failed
    SetNativeFileChooserInterceptionOnSession False, cdpSessionId, logSuccess
    TryDisableFileChooserInterceptionOnSession = True
    Exit Function

Failed:
    If m_debugMode Then
        Debug.Print "BiDi Warning: failed to disable native file chooser interception. " & _
                    Err.description
    End If
    Err.Clear
End Function

Private Function ReleaseTrackedFileChooserInterceptionBestEffort( _
    Optional ByVal tryFreshSession As Boolean = True) As Boolean

    If Not m_fileChooserInterceptMayBeActive Then
        ReleaseTrackedFileChooserInterceptionBestEffort = True
        Exit Function
    End If

    Dim oldSession As String
    oldSession = m_fileChooserInterceptSessionId

    If Len(oldSession) > 0 Then
        If TryDisableFileChooserInterceptionOnSession(oldSession, True) Then
            ClearTrackedFileChooserInterception
            ReleaseTrackedFileChooserInterceptionBestEffort = True
            Exit Function
        End If
    End If

    If tryFreshSession And Len(m_fileChooserInterceptContextId) > 0 Then
        On Error Resume Next
        Err.Clear
        Dim freshSession As String
        freshSession = GetCdpSessionIdForContext(m_fileChooserInterceptContextId)
        Dim freshErr As Long
        freshErr = Err.Number
        Err.Clear
        On Error GoTo 0

        If freshErr = 0 And Len(freshSession) > 0 Then
            If StrComp(freshSession, oldSession, vbBinaryCompare) <> 0 Or Len(oldSession) = 0 Then
                If TryDisableFileChooserInterceptionOnSession(freshSession, True) Then
                    ClearTrackedFileChooserInterception
                    ReleaseTrackedFileChooserInterceptionBestEffort = True
                    Exit Function
                End If
            End If
        End If
    End If

    ReleaseTrackedFileChooserInterceptionBestEffort = False
End Function

Private Sub RecoverResidualFileChooserInterceptionOnConnectBestEffort()
    ' This is the hard-reset insurance path. A fresh VBA object has no memory of an earlier
    ' interceptor, so on Chromium we explicitly send disabled=False-equivalent state (enabled=False)
    ' to every current browsing-context CDP session once at ConnectTo. Firefox/other BiDi
    ' implementations that do not support goog:cdp are silently left untouched.
    On Error GoTo Done

    Dim params As New Dictionary
    Dim raw As String
    raw = ExecuteBiDiCommand("browsingContext.getTree", params, True, _
                             allowAmbiguousRetry:=True, _
                             allowRecoveryRetry:=True, _
                             allowContextRetarget:=False)

    Dim tree As Object
    Set tree = ValidateProtocolResponse(raw, "browsingContext.getTree")

    Dim roots As Collection
    Set roots = GetPathObject(tree, "result.contexts")
    If roots Is Nothing Then GoTo Done
    If roots.count = 0 Then GoTo Done

    Dim contexts As New Collection
    CollectContextIdsForCleanup roots, contexts
    If contexts.count = 0 Then GoTo Done

    Dim seenSessions As New Dictionary
    Dim i As Long
    Dim ctx As String
    Dim sessionId As String
    Dim chromiumDetected As Boolean
    Dim allDisabled As Boolean
    allDisabled = True

    For i = 1 To contexts.count
        ctx = CStr(contexts(i))
        sessionId = TryGetCdpSessionIdForCleanup(ctx)

        If Len(sessionId) = 0 Then
            ' If the first probe cannot obtain goog:cdp, treat this as a non-Chromium browser.
            If Not chromiumDetected Then GoTo Done
        Else
            chromiumDetected = True
            If Not seenSessions.Exists(sessionId) Then
                seenSessions.Add sessionId, True
                If Not TryDisableFileChooserInterceptionOnSession(sessionId, False) Then
                    allDisabled = False
                    If Not m_fileChooserInterceptMayBeActive Then
                        m_fileChooserInterceptMayBeActive = True
                        m_fileChooserInterceptSessionId = sessionId
                        m_fileChooserInterceptContextId = ctx
                    End If
                    If m_debugMode Then
                        Debug.Print "BiDi Warning: ConnectTo could not confirm native file chooser " & _
                                    "interception cleanup for context " & Left$(ctx, 16) & "..."
                    End If
                End If
            End If
        End If
    Next i

    If chromiumDetected And allDisabled Then
        ClearTrackedFileChooserInterception
        If m_debugMode Then
            Debug.Print "BiDi: Native file chooser residual-state cleanup completed on ConnectTo."
        End If
    ElseIf chromiumDetected And Not allDisabled Then
        If m_debugMode Then
            Debug.Print "BiDi Warning: native file chooser residual-state cleanup remains pending."
        End If
    End If

Done:
    Err.Clear
End Sub

Private Sub RecoverResidualDownloadBehaviorOnConnectBestEffort()
    ' A fresh wrapper cannot know whether a previous VBA project left browser-global download
    ' behavior behind. Clear once after ConnectTo. Unsupported implementations are harmless.
    On Error GoTo Failed

    Dim params As New Dictionary
    params.Add "downloadBehavior", Null

    m_downloadBehaviorCleanupRequired = True
    m_downloadBehaviorStateUnknown = True

    Dim failureDetail As String
    Dim raw As String
    raw = ExecuteBiDiCommand("browser.setDownloadBehavior", params, True, _
                             allowAmbiguousRetry:=False, _
                             allowRecoveryRetry:=False, _
                             allowContextRetarget:=False)

    Dim json As Object
    Set json = WebJsonConverter.ParseJson(raw)
    If json Is Nothing Then GoTo Failed

    If json.Exists("error") Then
        Dim errType As String
        errType = LCase$(CStr(json("error")))
        If StrComp(errType, "unknown command", vbBinaryCompare) = 0 Or _
           StrComp(errType, "unsupported operation", vbBinaryCompare) = 0 Then
            m_downloadBehaviorCleanupRequired = False
            m_downloadBehaviorStateUnknown = False
            If m_debugMode Then Debug.Print "BiDi: browser.setDownloadBehavior is unsupported; residual sweep skipped."
            GoTo Done
        End If
        failureDetail = "[" & errType & "] " & GetPathString(json, "message")
        GoTo Failed
    End If

    m_downloadBehaviorCleanupRequired = False
    m_downloadBehaviorStateUnknown = False
    If m_debugMode Then Debug.Print "BiDi: Native download-behavior residual-state cleanup completed on ConnectTo."
    GoTo Done

Failed:
    m_downloadBehaviorStateUnknown = True
    If Len(failureDetail) = 0 Then
        failureDetail = "Error #" & CStr(Err.Number) & ": " & Err.description
    End If
    If m_debugMode Then
        Debug.Print "BiDi CRITICAL: browser.setDownloadBehavior residual cleanup could not be confirmed. " & _
                    "Download execution is blocked until an explicit setter or clear succeeds. " & _
                    failureDetail
    End If

Done:
    Err.Clear
End Sub

Private Function IsUnsupportedDownloadBehaviorError(ByVal errNum As Long, _
                                                    ByVal errSrc As String, _
                                                    ByVal errDesc As String) As Boolean
    If errNum <> ERR_BIDI_PROTOCOL Then Exit Function
    If InStr(1, errSrc, "browser.setDownloadBehavior", vbTextCompare) = 0 Then Exit Function
    IsUnsupportedDownloadBehaviorError = _
        (InStr(1, errDesc, "[unknown command]", vbTextCompare) > 0 Or _
         InStr(1, errDesc, "[unsupported operation]", vbTextCompare) > 0)
End Function


Private Function TryGetCdpSessionIdForCleanup(ByVal contextId As String) As String
    On Error GoTo Failed
    TryGetCdpSessionIdForCleanup = GetCdpSessionIdForContext(contextId)
    Exit Function
Failed:
    Err.Clear
    TryGetCdpSessionIdForCleanup = vbNullString
End Function

Private Sub CollectContextIdsForCleanup(ByVal nodes As Collection, ByRef output As Collection)
    Dim node As Variant
    For Each node In nodes
        If IsObject(node) Then
            If node.Exists("context") Then output.Add CStr(node("context"))
            If node.Exists("children") Then
                If IsObject(node("children")) Then
                    If Not node("children") Is Nothing Then
                        CollectContextIdsForCleanup node("children"), output
                    End If
                End If
            End If
        End If
    Next node
End Sub

Private Function PerformSetFiles(ByVal contextId As String, _
                                 ByVal sharedId As String, _
                                 ByVal paths As Variant, _
                                 Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                 Optional ByVal waitNetworkIdle As Boolean = True, _
                                 Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    If Not IsArray(paths) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:PerformSetFiles", "Internal: paths must be an array."
    End If
    If UBound(paths) < LBound(paths) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:PerformSetFiles", "No file paths supplied for file input."
    End If

    Dim files As New Collection, i As Long
    For i = LBound(paths) To UBound(paths)
        files.Add CStr(paths(i))
    Next i

    Dim params As New Dictionary
    params.Add "context", contextId
    params.Add "element", NewElementReference(sharedId)
    params.Add "files", files

    Dim strRes As String
    strRes = ExecuteBiDiCommand("input.setFiles", params, _
                                allowAmbiguousRetry:=False, _
                                allowContextRetarget:=(contextId = m_mainContextId), _
                                commandTimeoutMs:=ResolveTransportTimeout(maxTimeoutMs))
    ValidateProtocolResponse strRes, "input.setFiles"
    If waitNetworkIdle Then RunRequiredSpaSync contextId, minStableMs, maxTimeoutMs
    PerformSetFiles = strRes
End Function

Private Function PerformNativeKeyInput(ByVal contextId As String, _
                                       ByVal text As String, _
                                       Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                       Optional ByVal waitNetworkIdle As Boolean = True, _
                                       Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String

    Dim keyActions As New Collection
    Dim i As Long
    Dim ch As String
    Dim asciiCode As Long
    Dim skipNext As Boolean: skipNext = False

    For i = 1 To Len(text)
        If skipNext Then
            skipNext = False
        Else
            ch = Mid$(text, i, 1)
            asciiCode = AscW(ch)

            Select Case asciiCode
                Case 13 ' CR -> Enter (also swallow following LF if present: CRLF)
                    AppendKeyStroke keyActions, ChrW(&HE007&) ' U+E007 = Enter (W3C key code)
                    If i < Len(text) Then
                        If AscW(Mid$(text, i + 1, 1)) = 10 Then skipNext = True
                    End If
                Case 10 ' LF -> Enter
                    AppendKeyStroke keyActions, ChrW(&HE007&)
                Case 9  ' Tab
                    AppendKeyStroke keyActions, ChrW(&HE004&) ' U+E004 = Tab
                Case 8  ' Backspace
                    AppendKeyStroke keyActions, ChrW(&HE003&) ' U+E003 = Backspace
                Case Else
                    AppendKeyStroke keyActions, ch
            End Select
        End If
    Next i

    Dim source As New Dictionary
    source.Add "type", "key"
    source.Add "id", "kbd_main"
    source.Add "actions", keyActions

    Dim sources As New Collection: sources.Add source

    Dim params As New Dictionary
    params.Add "context", contextId
    params.Add "actions", sources

    Dim strRes As String
    strRes = ExecuteBiDiCommand("input.performActions", params, _
                                allowAmbiguousRetry:=False, _
                                allowContextRetarget:=(contextId = m_mainContextId), _
                                commandTimeoutMs:=ResolveTransportTimeout(maxTimeoutMs))
    ValidateProtocolResponse strRes, "input.performActions"
    If waitNetworkIdle Then RunRequiredSpaSync contextId, minStableMs, maxTimeoutMs
    PerformNativeKeyInput = strRes
End Function

Private Sub AppendKeyStroke(ByRef actions As Collection, ByVal value As String)
    Dim dn As New Dictionary
    dn.Add "type", "keyDown"
    dn.Add "value", value
    actions.Add dn

    Dim up As New Dictionary
    up.Add "type", "keyUp"
    up.Add "value", value
    actions.Add up
End Sub

' Trusted click is used only when the target is topmost at its center. Callers that require
' real user activation can disable the node-addressed synthetic fallback.
Private Function PerformTrustedClick(ByVal contextId As String, _
                                     ByVal args As Collection, _
                                     ByVal TIMEOUT As Long, _
                                     Optional ByVal allowSyntheticFallback As Boolean = True) As String

    Dim elemArg As Object: Set elemArg = args(1)
    Dim sharedId As String: sharedId = CStr(elemArg("sharedId"))

    Dim verdict As String
    On Error Resume Next
    verdict = CStr(GetCleanResult(CallScript(HitTestJs(), args, contextId, True, False, 5000, _
                                               Silent:=False, allowAmbiguousRetry:=True)))
    On Error GoTo 0
    Dim onTarget As Boolean: onTarget = (InStr(verdict, "hit=1") > 0)

    If onTarget Then
        Dim ptrActions As New Collection
        ptrActions.Add BuildPointerMove(sharedId, 0, 0)
        ptrActions.Add BuildPointerButton("pointerDown", 0)
        ptrActions.Add BuildPointerButton("pointerUp", 0)

        Dim ptrParams As New Dictionary: ptrParams.Add "pointerType", "mouse"
        Dim source As New Dictionary
        source.Add "type", "pointer": source.Add "id", "mouse_main"
        source.Add "parameters", ptrParams: source.Add "actions", ptrActions
        Dim sources As New Collection: sources.Add source
        Dim params As New Dictionary
        params.Add "context", contextId: params.Add "actions", sources

        Dim res As String, trustedOk As Boolean
        Dim tcErrNum As Long, tcErrDesc As String, tcErrSrc As String
        On Error Resume Next
        Err.Clear
        res = ExecuteBiDiCommand("input.performActions", params, _
                                 allowAmbiguousRetry:=False, _
                                 allowContextRetarget:=(contextId = m_mainContextId), _
                                 commandTimeoutMs:=ResolveTransportTimeout(TIMEOUT))
        tcErrNum = Err.Number: tcErrDesc = Err.description: tcErrSrc = Err.source
        trustedOk = (tcErrNum = 0)
        On Error GoTo 0
        If trustedOk Then
            ValidateProtocolResponse res, "input.performActions"
            If m_debugMode Then Debug.Print "BiDi: TRUSTED click (on-target)."
            PerformTrustedClick = res
            Exit Function
        End If
        ' Timeout or connection loss after dispatch is ambiguous: the pointer click may already
        ' have fired, so never fall back to a second JS click.
        If IsAmbiguousDispatchFailure(tcErrNum, tcErrSrc, tcErrDesc) Then Err.Raise tcErrNum, tcErrSrc, tcErrDesc
        If Not allowSyntheticFallback Then Err.Raise tcErrNum, tcErrSrc, tcErrDesc
        ' Fallback is safe only after a non-ambiguous pointer failure.
    Else
        ' A covered target must not receive a coordinate click because it would hit the front
        ' element; the default fallback addresses the intended node directly.
        If Not allowSyntheticFallback Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:PerformTrustedClick", _
                      "TrustedClickRequired: the target center is not topmost; synthetic fallback is disabled."
        End If
        If m_debugMode Then Debug.Print "BiDi: not on-target; node click fallback."
    End If

    PerformTrustedClick = CallScript(GetActionJs("click"), args, contextId, True, , TIMEOUT, _
                                     allowAmbiguousRetry:=False)
End Function

Private Function HitTestJs() As String
    Dim js As String
    js = "const e=arguments[0]; e.scrollIntoView({block:'center',inline:'center',behavior:'instant'}); "
    js = js & "if(window.__vbaIdleProbe) window.__vbaIdleProbe.lastMutationTs = performance.now(); "
    js = js & "const r=e.getBoundingClientRect(); if(!(r.width>0||r.height>0)) return 'hit=0'; "
    js = js & "const cx=Math.round(r.left+r.width/2), cy=Math.round(r.top+r.height/2); "
    js = js & "let top=document.elementFromPoint(cx,cy); "
    js = js & "let g=0; while(top&&top.shadowRoot&&g++<10){const d=top.shadowRoot.elementFromPoint(cx,cy); if(!d||d===top)break; top=d;} "
    js = js & "const within=(n)=>{let x=n,h=0;while(x&&h++<20){if(x===e||e.contains(x))return true;"
    js = js & "const r=x.getRootNode?x.getRootNode():null;if(r&&r.host){x=r.host;continue;}break;}return false;}; "
    js = js & "return 'hit='+((top&&within(top))?1:0);"
    HitTestJs = js
End Function

Private Function BuildPointerMove(ByVal sharedId As String, _
                                  ByVal x As Long, _
                                  ByVal y As Long) As Dictionary
    Dim elemRef As Dictionary: Set elemRef = NewElementReference(sharedId)

    Dim origin As New Dictionary
    origin.Add "type", "element"
    origin.Add "element", elemRef

    Dim mv As New Dictionary
    mv.Add "type", "pointerMove"
    mv.Add "x", x
    mv.Add "y", y
    mv.Add "duration", 0
    mv.Add "origin", origin

    Set BuildPointerMove = mv
End Function

Private Function BuildPointerButton(ByVal actionType As String, _
                                    ByVal button As Long) As Dictionary
    Dim a As New Dictionary
    a.Add "type", actionType
    a.Add "button", button
    Set BuildPointerButton = a
End Function

Public Function ExecuteSelectValueByXPath(ByVal xpath As String, _
                                          ByVal valueOrText As String, _
                                          Optional ByVal selectByText As Boolean = False, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                          Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteSelectValueByXPath")

    Dim elementArgs As Collection: Set elementArgs = NewElementArgs(sharedId)
    Dim isStandardSelect As Boolean
    isStandardSelect = CBool(GetCleanResult(CallScript( _
        "return arguments[0].tagName==='SELECT';", elementArgs, targetCtx, _
        False, False, 2000, Silent:=True, allowAmbiguousRetry:=True)))

    If isStandardSelect Then
        Dim jsKey As String
        If selectByText Then jsKey = "select_text" Else jsKey = "select"
        ExecuteSelectValueByXPath = ExecuteBaseAction( _
            targetCtx, jsKey, NewElementArgs(sharedId, valueOrText), waitNetworkIdle, _
            DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs, _
            False, xpath, searchTimeoutMs, "ExecuteSelectValueByXPath")
    Else
        Call ExecuteBaseAction(targetCtx, "click", elementArgs, False, _
                               DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, _
                               maxTimeoutMs, False, xpath, searchTimeoutMs, _
                               "ExecuteSelectValueByXPath")
        Dim optXPath As String
        optXPath = "//*[@role='option'][contains(., " & EscapeXPathString(valueOrText) & ")]"
        Dim optSharedId As String
        optSharedId = ResolveXPathOrThrow(optXPath, targetCtx, searchTimeoutMs, _
                                          "ExecuteSelectValueByXPath")
        ExecuteSelectValueByXPath = ExecuteBaseAction( _
            targetCtx, "click", NewElementArgs(optSharedId), waitNetworkIdle, _
            DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs, _
            False, optXPath, searchTimeoutMs, "ExecuteSelectValueByXPath")
    End If
End Function

Private Function ExecuteBrowsingContextLocateNodes(ByVal xpath As String, _
                                                  Optional ByVal contextId As String = "", _
                                                  Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim params As New Dictionary, locator As New Dictionary
    locator.Add "type", "xpath": locator.Add "value", xpath
    params.Add "context", targetCtx: params.Add "locator", locator
    
    ExecuteBrowsingContextLocateNodes = ExecuteBiDiCommand("browsingContext.locateNodes", params, Silent, _
                                allowContextRetarget:=(targetCtx = m_mainContextId))
End Function

Private Function GetFindElementSharedIdByXPath(ByVal xpath As String, _
                                               ByVal searchTimeoutMs As Long, _
                                               ByVal targetCtx As String) As String

    ' Test the original timeout before SanitizeMs; otherwise the exact-zero one-shot path
    ' would be converted to the minimum polling interval.
    Dim isImmediate As Boolean: isImmediate = (searchTimeoutMs <= 0)

    Dim vTimeoutMs As Long
    If isImmediate Then
        vTimeoutMs = 0
    Else
        vTimeoutMs = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    End If

    Dim startTick As Currency: startTick = socket_.NowTicks()
    Dim strRes As String, json As Object, nodes As Collection
    Dim sid As String, attemptCount As Long
    Dim pendingTransient As String, lastLoggedTransient As String

    If m_debugMode Then
        Debug.Print "BiDi: Searching for [" & xpath & "] (Timeout: " & _
                    IIf(isImmediate, "Immediate", vTimeoutMs & "ms") & ")"
    End If

    Do
        attemptCount = attemptCount + 1

        If attemptCount Mod 5 = 0 Then
            Dim currentUrl As String: currentUrl = GetCurrentUrl(targetCtx, True)
            If InStr(1, currentUrl, "about:blank", vbTextCompare) > 0 Or currentUrl = "" Then
                Debug.Print "BiDi: Aborting search. Target context navigated away."
                Exit Do
            End If
        End If

        strRes = ExecuteBrowsingContextLocateNodes(xpath, targetCtx, True)
        If TryGetPollingResponse(strRes, "browsingContext.locateNodes", json, _
                                 pendingTransient, lastLoggedTransient) Then
            Set nodes = GetPathObject(json, "result.nodes")
            If Not nodes Is Nothing Then
                If nodes.count > 0 Then
                    If nodes(1).Exists("sharedId") Then
                        sid = CStr(nodes(1)("sharedId"))
                        GetFindElementSharedIdByXPath = sid
                        Exit Function
                    End If
                End If
            End If
        End If

        If isImmediate Then Exit Do
        WaitHeartbeat 500
        If IsTimeout(startTick, CDbl(vTimeoutMs)) Then Exit Do
    Loop

    RaiseIfPollingEndedOnProtocolError "browsingContext.locateNodes", pendingTransient
    GetFindElementSharedIdByXPath = ""
End Function

Public Function GetCurrentUrl(Optional ByVal contextId As String = "", _
                              Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Dim strRes As String: strRes = CallScript("return window.location.href;", Nothing, targetCtx, False, False, 10000, _
                                             Silent:=Silent, allowAmbiguousRetry:=True)
    
    GetCurrentUrl = CStr(GetCleanResult(strRes))
End Function

Public Function ExecuteWebExtensionInstall(ByVal extensionPath As String) As String
    Dim params As New Dictionary, extData As New Dictionary
    extData.Add "type", "path"
    extData.Add "path", extensionPath
    params.Add "extensionData", extData
    Dim strRes As String
    strRes = ExecuteBiDiCommand("webExtension.install", params, allowAmbiguousRetry:=False)
    ValidateProtocolResponse strRes, "webExtension.install"
    ExecuteWebExtensionInstall = strRes
End Function

Public Function ExecuteShadowClick(ByVal selectorsArray As Variant, _
                                   Optional ByVal searchTimeoutMs As Long = 10000, _
                                   Optional ByVal waitNetworkIdle As Boolean = True, _
                                   Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                   Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                   Optional ByVal contextId As String = "", _
                                   Optional ByVal requireUnlocker As Boolean = False) As String
    
    If requireUnlocker And Not m_shadowUnlockerRegistered Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteShadowClick", _
            "Shadow Unlocker is not registered for this session. Reconnect with " & _
            "requireShadowUnlocker:=True, or verify script.addPreloadScript support."
    End If

    Dim vSearchTimeout As Long: vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 5000)
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Dim args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray)
    
    Dim argTime As New Dictionary
    argTime.Add "type", "number": argTime.Add "value", vSearchTimeout
    args.Add argTime
    
    On Error GoTo ShadowClickFailed
    ExecuteShadowClick = ExecuteBaseAction(targetCtx, "shadow_click", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
    Exit Function

ShadowClickFailed:
    Dim shadowErrNum As Long: shadowErrNum = Err.Number
    Dim shadowErrSrc As String: shadowErrSrc = Err.source
    Dim shadowErrDesc As String: shadowErrDesc = Err.description

    If Not m_shadowUnlockerRegistered And _
       InStr(1, shadowErrDesc, "ShadowSearchError", vbTextCompare) > 0 Then
        shadowErrDesc = shadowErrDesc & vbCrLf & _
            "Diagnostic: Shadow Unlocker is not registered. The target may be inside a " & _
            "closed shadow root. Open shadow roots remain accessible without the Unlocker."
    End If

    If Len(shadowErrSrc) = 0 Then shadowErrSrc = "BiDi:ExecuteShadowClick"
    Err.Raise shadowErrNum, shadowErrSrc, shadowErrDesc
End Function

' Lazy scrolling delegates only the scroll loop to JS; final quiescence remains the
' SPA-idle wait. The JS budget must stay below the Safety Shell timeout because a shell
' timeout stops waiting but does not cancel the browser task. Auto-reserve headroom to
' avoid an ambiguous still-running scroll.
Public Function ExecuteLazyLoadScroll(Optional ByVal scrollTimeoutMs As Long = 20000, _
                                      Optional ByVal waitNetworkIdle As Boolean = True, _
                                      Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                      Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                      Optional ByVal scrollPollMs As Long = 800, _
                                      Optional ByVal contextId As String = "") As String

    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)

    Const MIN_POLL_MS As Long = 100
    Const MAX_POLL_MS As Long = 5000
    If scrollPollMs < MIN_POLL_MS Or scrollPollMs > MAX_POLL_MS Then
        Err.Raise ERR_ACTION_FAILED, "ExecuteLazyLoadScroll", _
                  "scrollPollMs must be between " & MIN_POLL_MS & " and " & MAX_POLL_MS & "ms. Got: " & scrollPollMs
    End If

    ' Keep scrollTimeoutMs strictly below maxTimeoutMs so the browser task can return before
    ' the Safety Shell stops waiting.
    Const SCROLL_SAFETY_MARGIN_MS As Long = 5000
    Dim vScrollTimeout As Long: vScrollTimeout = SanitizeMs(scrollTimeoutMs, "scrollTimeoutMs", 20000)
    Dim vMaxTimeout As Long:    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    If vMaxTimeout <= vScrollTimeout + SCROLL_SAFETY_MARGIN_MS Then
        vMaxTimeout = vScrollTimeout + SCROLL_SAFETY_MARGIN_MS
        If m_debugMode Then _
            Debug.Print "[BiDi LazyScroll] maxTimeout auto-raised to " & vMaxTimeout & _
                        "ms to fit scroll budget (" & vScrollTimeout & "ms) + margin."
    End If

    Dim args As New Collection
    args.Add NewLocalValue("number", vScrollTimeout)
    args.Add NewLocalValue("number", scrollPollMs)

    ' awaitPromise must remain True because the scroll action is asynchronous.
    ExecuteLazyLoadScroll = ExecuteBaseAction(targetCtx, "lazy_scroll", args, _
                                              waitNetworkIdle, minStableMs, True, _
                                              DEFAULT_STALE_RETRIES, vMaxTimeout)
End Function

Private Function ExecuteBiDiCommand(ByVal method As String, _
                                     ByRef params As Dictionary, _
                                     Optional ByVal Silent As Boolean = False, _
                                     Optional ByVal allowAmbiguousRetry As Boolean = True, _
                                     Optional ByVal allowRecoveryRetry As Boolean = True, _
                                     Optional ByVal allowContextRetarget As Boolean = True, _
                                     Optional ByVal commandTimeoutMs As Long = 0) As String
    
    If socket_ Is Nothing Then Err.Raise 91, "ExecuteBiDiCommand", "Socket object is not set."

    Dim retryCount As Long
    Dim strRes As String
    Dim localErrNum As Long
    Dim localErrDesc As String
    
    Do While retryCount <= MAX_CONTEXT_RECOVERY_RETRIES
        
        On Error Resume Next
        strRes = socket_.SendCommand(method, params, Silent, commandTimeoutMs)
        localErrNum = Err.Number: localErrDesc = Err.description
        On Error GoTo 0
        
        If localErrNum = 0 Then
            If IsContextLostError(strRes) Then
                localErrNum = ERR_BIDI_PROTOCOL
                localErrDesc = "Environment desync (Context Lost) detected in response."
                Set m_validatedContexts = New Dictionary
            End If
            
            If localErrNum = 0 Then
                socket_.ClearInternalErrorCounter
                ExecuteBiDiCommand = strRes
                
                Exit Function
            End If
        End If

        If localErrNum = ERR_ACTION_FAILED Then Exit Do

        ' Ambiguous outcomes for state-changing commands include 604, 600, 601, 517, and
        ' context loss after dispatch. The browser may already have executed the action, so
        ' allowAmbiguousRetry=False must stop rather than risk a duplicate click/input/navigation.
        If Not allowAmbiguousRetry And _
           (IsAmbiguousOutcome(localErrNum) Or localErrNum = ERR_BIDI_PROTOCOL) Then Exit Do

        ' Commands bound to a specific disappearing context (notably close) disable recovery
        ' retry because the generic recovery path may retarget parameters to the surviving main
        ' context and act on the wrong tab.
        If Not allowRecoveryRetry Then Exit Do

        ' When main retargeting is forbidden, context loss is terminal: retrying the dead target
        ' and hijacking the main target are both incorrect.
        If Not allowContextRetarget And localErrNum = ERR_BIDI_PROTOCOL Then Exit Do

        If retryCount < MAX_CONTEXT_RECOVERY_RETRIES And Not m_isRecovering Then
            
            ' If auto-reconnect is enabled, recovery may continue for retry-safe commands even when
            ' the socket is currently down. Non-idempotent commands have already been stopped by the
            ' ambiguous-outcome guard.
            If Not socket_.IsConnected And Not socket_.AutoReconnect Then
                Debug.Print "BiDi [FATAL]: Socket down and auto-reconnect disabled. Aborting command retries."
                Exit Do
            End If
            
            m_isRecovering = True

            ' Own the recovery guard across both the backoff wait and context refresh.
            ' Any error is captured first; the guard is then released exactly once before
            ' deciding whether the command loop may continue.
            Dim recoveryErr As Long
            Dim recoveryErrDesc As String
            On Error Resume Next
            WaitHeartbeat 800 * (retryCount + 1)
            recoveryErr = Err.Number
            recoveryErrDesc = Err.description
            Err.Clear

            If recoveryErr = 0 Then
                If m_debugMode Then
                    Debug.Print "BiDi [RECOVERY]: Attempting to refresh context (" & (retryCount + 1) & "/" & MAX_CONTEXT_RECOVERY_RETRIES & ")..."
                End If

                Call UpdateMainContextId(True, True) ' forceRefresh:=True, Silent:=True
                recoveryErr = Err.Number
                recoveryErrDesc = Err.description
                Err.Clear
            End If
            On Error GoTo 0
            m_isRecovering = False

            If recoveryErr = 0 Then
                ' Retarget only when the caller explicitly permits main-context recovery.
                If m_mainContextId <> "" And allowContextRetarget Then Call UpdateContextAndRealm(params, m_mainContextId)
                retryCount = retryCount + 1
            Else
                Debug.Print "BiDi [RECOVERY FAIL]: " & recoveryErrDesc
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    Err.Raise IIf(localErrNum = 0, ERR_ACTION_FAILED, localErrNum), _
              "BiDi:ExecuteBiDiCommand", _
              "Command '" & method & "' failed after " & retryCount & " recovery attempt(s). " & _
              "Last Error: " & localErrDesc
End Function

Private Function IsContextLostError(ByVal response As String) As Boolean
    If InStr(1, response, """result"":", vbBinaryCompare) > 0 Then
        IsContextLostError = False
        Exit Function
    End If

    On Error Resume Next
    Dim json As Object: Set json = WebJsonConverter.ParseJson(response)
    On Error GoTo 0
    
    If json Is Nothing Then Exit Function
    
    If json.Exists("error") And json.Exists("message") Then
        Dim errType As String: errType = LCase(CStr(json("error")))
        Dim errMsg As String: errMsg = LCase(CStr(json("message")))
        
        Select Case errType
            Case "no such frame", "no such window", "unknown error"
                If InStr(errMsg, "context") > 0 Or InStr(errMsg, "gone") > 0 Or InStr(errMsg, "destroyed") > 0 Then
                    IsContextLostError = True
                    Exit Function
                End If
        End Select
    End If
End Function

Private Function UpdateContextAndRealmRecursive(ByRef node As Variant, _
                                                ByVal newCtx As String, _
                                                Optional ByVal newRealm As String = "") As Boolean
    On Error GoTo ErrH

    If IsObject(node) Then
        If TypeName(node) = "Dictionary" Then
            Dim k As Variant
            For Each k In node.keys
                If LCase(CStr(k)) = "context" Then
                    node(k) = newCtx
                ElseIf LCase(CStr(k)) = "realm" And Len(newRealm) > 0 Then
                    node(k) = newRealm
                End If
                Call UpdateContextAndRealmRecursive(node(k), newCtx, newRealm)
            Next k
            UpdateContextAndRealmRecursive = True
            
        ElseIf TypeName(node) = "Collection" Then
            Dim i As Long
            For i = 1 To node.count
                Call UpdateContextAndRealmRecursive(node(i), newCtx, newRealm)
            Next i
            UpdateContextAndRealmRecursive = True
        Else
            UpdateContextAndRealmRecursive = True
        End If
    Else
        UpdateContextAndRealmRecursive = True
    End If
    Exit Function

ErrH:
    UpdateContextAndRealmRecursive = False
End Function

Private Function UpdateContextAndRealm(ByRef params As Dictionary, ByVal newCtx As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim newRealm As String
    newRealm = GetScriptRealmId(newCtx, True, True) ' Silent:=True
    
    If newRealm = "" Then
        UpdateContextAndRealm = False
        Exit Function
    End If

    UpdateContextAndRealm = UpdateContextAndRealmRecursive(params, newCtx, newRealm)
    Exit Function

ErrorHandler:
    UpdateContextAndRealm = False
End Function

Private Sub ResetVolatileContextState()
    m_cdpSessionId = ""

    Set m_cachedRealmIds = New Dictionary
    Set m_validatedContexts = New Dictionary
    Set m_probeFilterRevisions = New Dictionary
    Set m_reportedInvalidSelectors = New Dictionary

    ResetSignalState
    ResetFileDialogCaptureState
    ResetDownloadCaptureState
    ResetNewContextCaptureState

    m_lastHudKey = ""
    m_lastHudCtx = ""
    m_lastHudRetryTick = 0
    m_contextLifecycleNeedsReconcile = True
End Sub

Public Sub ClearState()
    ResetVolatileContextState

    SetMainContextUnbound
    Set m_activeSubscriptions = New Dictionary
    Set m_blockingPatterns = New Collection

    If Not socket_ Is Nothing Then
        m_contextLifecycleDroppedEventsBaseline = CurrentLifecycleRelevantLossCount()
    Else
        m_contextLifecycleDroppedEventsBaseline = 0
    End If
    m_contextLifecycleNeedsReconcile = False
End Sub

' Session-boundary reset runs only when one session ends and another may begin.
' Flags are dropped before their collections so guarded append paths stay safe.
Private Sub ResetSessionBoundaryState()
    m_isRecording = False
    m_isProcessingQueue = False

    Set m_recorderLogs = Nothing
    Set m_recSuppByKey = Nothing
    m_recSuppOverflowCount = 0
    m_droppedHeadCount = 0
    m_recorderStartTime = 0

    m_navigationDetected = False
    m_navigationContextId = ""

    m_waitStackDepth = 0

    m_shadowUnlockerRegistered = False
    m_shadowUnlockerScriptId = vbNullString
    ResetLastSyncState
    ResetFileDialogCaptureState
    ResetDownloadCaptureState
    ResetNewContextCaptureState
    m_downloadBehaviorCleanupRequired = False
    m_downloadBehaviorStateUnknown = False

    m_unresolvedJsonThisWait = False
    m_unresolvedJsonFirstKeyThisWait = ""
    m_unresolvedJsonFirstSecThisWait = 0#
    m_hasCarryoverHintForNextWait = False
    m_carryoverNoteArmedThisWait = False
    m_carryoverNoteEmittedThisWait = False
End Sub

Private Sub SetMainContextUnbound()
    m_mainContextId = ""
    m_lostMainContextId = ""
    m_mainContextState = MainContextUnbound
End Sub

Private Sub MarkMainContextLost(ByVal contextId As String)
    If m_mainContextState <> MainContextPinned Then Exit Sub
    If StrComp(contextId, m_mainContextId, vbBinaryCompare) <> 0 Then Exit Sub

    m_lostMainContextId = m_mainContextId
    m_mainContextId = ""
    m_mainContextState = MainContextLost
End Sub

Private Sub RaiseMainContextLost()
    Dim lostId As String
    lostId = m_lostMainContextId
    If Len(lostId) = 0 Then lostId = "(unknown)"

    Err.Raise ERR_MAIN_CONTEXT_LOST, "BiDi:GetMainContextId", _
              "The pinned main browsing context was destroyed: " & lostId & _
              ". Call SetMainContextId with a surviving top-level context."
End Sub

Public Sub SetMainContextId(ByVal contextId As String)
    If Len(Trim$(contextId)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:SetMainContextId", "contextId must not be empty."
    End If

    DrainNonBlockingRecorderQueue
    MarkLifecycleCacheStaleIfDrops

    Dim jsonRes As String
    jsonRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, True)

    Dim jsonTree As Object
    Set jsonTree = ValidateProtocolResponse(jsonRes, "browsingContext.getTree")

    Dim contexts As Collection
    Set contexts = GetPathObject(jsonTree, "result.contexts")
    If contexts Is Nothing Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:SetMainContextId", _
                  "No top-level browsing contexts are available."
    End If

    ReconcileContextTreeCache contexts

    If Not TopLevelContextExistsExactId(contexts, contextId) Then
        If ContextTreeContainsExactId(contexts, contextId) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:SetMainContextId", _
                      "contextId identifies a child browsing context. The main context must be top-level."
        End If
        Err.Raise ERR_ACTION_FAILED, "BiDi:SetMainContextId", _
                  "Top-level context ID not found in browser tree: " & contextId
    End If

    m_mainContextId = contextId
    m_lostMainContextId = ""
    m_mainContextState = MainContextPinned
    m_validatedContexts(contextId) = True
End Sub

Public Function GetMainContextId() As String
    If Not m_isProcessingQueue Then DrainNonBlockingRecorderQueue
    MarkLifecycleCacheStaleIfDrops

    If m_mainContextState = MainContextLost Then RaiseMainContextLost

    If m_mainContextState = MainContextPinned And _
       Len(m_mainContextId) > 0 And _
       Not m_contextLifecycleNeedsReconcile Then
        GetMainContextId = m_mainContextId
        Exit Function
    End If

    GetMainContextId = UpdateMainContextId(True, True)
End Function

Public Function UpdateMainContextId(Optional ByVal forceRefresh As Boolean = False, _
                                    Optional ByVal Silent As Boolean = False) As String
    If Not m_isProcessingQueue Then DrainNonBlockingRecorderQueue
    MarkLifecycleCacheStaleIfDrops

    If m_mainContextState = MainContextLost Then RaiseMainContextLost

    If m_mainContextState = MainContextPinned And _
       Len(m_mainContextId) > 0 And _
       Not forceRefresh And _
       Not m_contextLifecycleNeedsReconcile Then
        UpdateMainContextId = m_mainContextId
        Exit Function
    End If

    Dim jsonRes As String
    jsonRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, Silent)

    Dim jsonTree As Object
    Set jsonTree = ValidateProtocolResponse(jsonRes, "browsingContext.getTree")

    Dim contexts As Collection
    Set contexts = GetPathObject(jsonTree, "result.contexts")
    If contexts Is Nothing Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", _
                  "Failed to read the top-level browsing context list."
    End If

    Dim pinnedId As String
    pinnedId = m_mainContextId
    ReconcileContextTreeCache contexts

    If m_mainContextState = MainContextPinned Then
        If TopLevelContextExistsExactId(contexts, pinnedId) Then
            m_mainContextId = pinnedId
            m_validatedContexts(pinnedId) = True
            UpdateMainContextId = pinnedId
            Exit Function
        End If

        MarkMainContextLost pinnedId
        RaiseMainContextLost
    End If

    If m_mainContextState = MainContextLost Then RaiseMainContextLost

    If contexts.count = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", _
                  "No top-level browsing context is available."
    End If

    If contexts.count <> 1 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", _
                  "Multiple top-level browsing contexts exist. Call SetMainContextId explicitly."
    End If

    Dim onlyContext As Dictionary
    On Error Resume Next
    Set onlyContext = contexts(1)
    Dim onlyContextErr As Long: onlyContextErr = Err.Number
    On Error GoTo 0
    If onlyContextErr <> 0 Or onlyContext Is Nothing Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UpdateMainContextId", _
                  "The only top-level browsing context has an invalid structure."
    End If

    Dim contextId As String
    contextId = GetPathString(onlyContext, "context")
    If Len(contextId) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", _
                  "The only top-level browsing context has no context ID."
    End If

    m_mainContextId = contextId
    m_lostMainContextId = ""
    m_mainContextState = MainContextPinned
    m_validatedContexts(contextId) = True
    UpdateMainContextId = contextId
End Function

Private Function ResolveTargetContext(ByVal contextId As String) As String
    If Not m_isProcessingQueue Then DrainNonBlockingRecorderQueue
    MarkLifecycleCacheStaleIfDrops

    If contextId = "" Then
        ResolveTargetContext = GetMainContextId()
        Exit Function
    End If

    If m_mainContextState = MainContextPinned Then
        If StrComp(contextId, m_mainContextId, vbBinaryCompare) = 0 And _
           Not m_contextLifecycleNeedsReconcile Then
            ResolveTargetContext = contextId
            Exit Function
        End If
    End If

    If m_validatedContexts Is Nothing Then Set m_validatedContexts = New Dictionary
    If Not m_contextLifecycleNeedsReconcile Then
        If m_validatedContexts.Exists(contextId) Then
            ResolveTargetContext = contextId
            Exit Function
        End If
    End If

    Dim treeJson As String
    treeJson = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, True)

    Dim jsonTree As Object
    Set jsonTree = ValidateProtocolResponse(treeJson, "browsingContext.getTree")

    Dim contexts As Collection
    Set contexts = GetPathObject(jsonTree, "result.contexts")

    If Not contexts Is Nothing Then
        ReconcileContextTreeCache contexts
        If ContextTreeContainsExactId(contexts, contextId) Then
            m_validatedContexts(contextId) = True
            ResolveTargetContext = contextId
            Exit Function
        End If
    End If

    Err.Raise ERR_ACTION_FAILED, "BiDi:ResolveTargetContext", _
              "Context ID not found in browser tree: " & contextId & _
              ". The tab/window may have been closed or the ID is invalid."
End Function

Private Function SaturatedLossSum(ByVal specificLoss As Long, _
                                  ByVal unknownLoss As Long) As Long
    Dim total As Double
    total = CDbl(specificLoss) + CDbl(unknownLoss)

    If total >= 2147483647# Then
        SaturatedLossSum = 2147483647
    Else
        SaturatedLossSum = CLng(total)
    End If
End Function

Private Function CurrentLifecycleRelevantLossCount() As Long
    If socket_ Is Nothing Then Exit Function

    CurrentLifecycleRelevantLossCount = SaturatedLossSum( _
        socket_.LifecycleEventLossCount, _
        socket_.UnknownCriticalEventLossCount)
End Function

Private Function CurrentDownloadRelevantLossCount() As Long
    If socket_ Is Nothing Then Exit Function

    CurrentDownloadRelevantLossCount = SaturatedLossSum( _
        socket_.DownloadEventLossCount, _
        socket_.UnknownCriticalEventLossCount)
End Function

Private Function CurrentFileDialogRelevantLossCount() As Long
    If socket_ Is Nothing Then Exit Function

    CurrentFileDialogRelevantLossCount = SaturatedLossSum( _
        socket_.FileDialogEventLossCount, _
        socket_.UnknownCriticalEventLossCount)
End Function

Private Sub MarkLifecycleCacheStaleIfDrops()
    If socket_ Is Nothing Then Exit Sub

    Dim currentDrops As Long
    currentDrops = CurrentLifecycleRelevantLossCount()

    If currentDrops <> m_contextLifecycleDroppedEventsBaseline Then
        Set m_validatedContexts = New Dictionary
        m_contextLifecycleNeedsReconcile = True
        m_contextLifecycleDroppedEventsBaseline = currentDrops
    End If
End Sub

Private Sub ReconcileContextTreeCache(ByVal contexts As Collection)
    Set m_validatedContexts = New Dictionary
    CacheContextTreeIds contexts

    If m_mainContextState = MainContextPinned Then
        If Not TopLevelContextExistsExactId(contexts, m_mainContextId) Then
            MarkMainContextLost m_mainContextId
        End If
    End If

    If Not socket_ Is Nothing Then
        m_contextLifecycleDroppedEventsBaseline = CurrentLifecycleRelevantLossCount()
    End If
    m_contextLifecycleNeedsReconcile = False
End Sub

Private Sub CacheContextTreeIds(ByVal contextList As Collection, _
                                Optional ByVal currentDepth As Long = 0)
    Const MAX_RECURSION_DEPTH As Long = 20
    On Error GoTo CacheError
    If currentDepth > MAX_RECURSION_DEPTH Then Exit Sub

    Dim ctx As Dictionary
    For Each ctx In contextList
        Dim contextId As String
        contextId = GetPathString(ctx, "context")
        If Len(contextId) > 0 Then m_validatedContexts(contextId) = True

        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    CacheContextTreeIds ctx("children"), currentDepth + 1
                End If
            End If
        End If
    Next ctx
    Exit Sub

CacheError:
    If Err.Number = 28 Then
        Debug.Print "BiDi Critical: Stack overflow detected in CacheContextTreeIds"
    ElseIf m_debugMode Then
        Debug.Print "BiDi Error in CacheContextTreeIds: " & Err.description
    End If
End Sub

Private Function TopLevelContextExistsExactId(ByVal contextList As Collection, _
                                              ByVal targetContextId As String) As Boolean
    On Error GoTo SearchError

    Dim ctx As Dictionary
    For Each ctx In contextList
        If StrComp(GetPathString(ctx, "context"), targetContextId, vbBinaryCompare) = 0 Then
            TopLevelContextExistsExactId = True
            Exit Function
        End If
    Next ctx
    Exit Function

SearchError:
    If m_debugMode Then Debug.Print "BiDi Error in TopLevelContextExistsExactId: " & Err.description
End Function

Private Function FindContextInfoExactId(ByVal contextList As Collection, _
                                        ByVal targetContextId As String, _
                                        Optional ByVal currentDepth As Long = 0) As Dictionary
    Const MAX_RECURSION_DEPTH As Long = 20
    On Error GoTo SearchError
    If currentDepth > MAX_RECURSION_DEPTH Then Exit Function

    Dim ctx As Dictionary
    For Each ctx In contextList
        If StrComp(GetPathString(ctx, "context"), targetContextId, vbBinaryCompare) = 0 Then
            Set FindContextInfoExactId = ctx
            Exit Function
        End If

        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    Dim childMatch As Dictionary
                    Set childMatch = FindContextInfoExactId(ctx("children"), targetContextId, currentDepth + 1)
                    If Not childMatch Is Nothing Then
                        Set FindContextInfoExactId = childMatch
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ctx
    Exit Function

SearchError:
    If Err.Number = 28 Then
        Debug.Print "BiDi Critical: Stack overflow detected in FindContextInfoExactId"
    ElseIf m_debugMode Then
        Debug.Print "BiDi Error in FindContextInfoExactId: " & Err.description
    End If
End Function

Private Function ContextTreeContainsExactId(ByVal contextList As Collection, _
                                            ByVal targetContextId As String, _
                                            Optional ByVal currentDepth As Long = 0) As Boolean
    Const MAX_RECURSION_DEPTH As Long = 20

    On Error GoTo SearchError

    If currentDepth > MAX_RECURSION_DEPTH Then
        If m_debugMode Then
            Debug.Print "BiDi Warning: Max context-tree recursion depth reached (" & _
                        MAX_RECURSION_DEPTH & ")"
        End If
        Exit Function
    End If

    Dim ctx As Dictionary
    For Each ctx In contextList
        If ctx.Exists("context") Then
            If StrComp(CStr(ctx("context")), targetContextId, vbBinaryCompare) = 0 Then
                ContextTreeContainsExactId = True
                Exit Function
            End If
        End If

        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    If ContextTreeContainsExactId(ctx("children"), targetContextId, _
                                                 currentDepth + 1) Then
                        ContextTreeContainsExactId = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ctx

    Exit Function

SearchError:
    If Err.Number = 28 Then ' Out of stack space
        Debug.Print "BiDi Critical: Stack overflow detected in ContextTreeContainsExactId"
    ElseIf m_debugMode Then
        Debug.Print "BiDi Error in ContextTreeContainsExactId: " & Err.description
    End If
End Function

Private Function GetScriptRealmId(Optional ByVal contextId As String = "", _
                                 Optional ByVal forceRefresh As Boolean = False, _
                                 Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = GetMainContextId()
    If targetCtx = "" Then Exit Function

    If m_cachedRealmIds Is Nothing Then Set m_cachedRealmIds = New Dictionary

    If Not forceRefresh Then
        If m_cachedRealmIds.Exists(targetCtx) Then
            GetScriptRealmId = m_cachedRealmIds(targetCtx)
            Exit Function
        End If
    End If

    Dim params As New Dictionary
    params.Add "context", targetCtx
    
    Dim jsonRes As String
    On Error Resume Next
    jsonRes = ExecuteBiDiCommand("script.getRealms", params, Silent, _
                                allowContextRetarget:=(targetCtx = m_mainContextId))
    Dim realmErr As Long: realmErr = Err.Number
    Dim realmErrDesc As String: realmErrDesc = Err.description
    On Error GoTo 0
    
    If realmErr <> 0 Then
        If m_debugMode Then
            Debug.Print "BiDi: getRealms failed for context " & Left$(targetCtx, 16) & _
                        "... (Err " & realmErr & ": " & realmErrDesc & ")"
        End If
        GetScriptRealmId = ""
        Exit Function
    End If
    
    Dim realms As Collection
    Set realms = GetPathObject(WebJsonConverter.ParseJson(jsonRes), "result.realms")
    
    Dim foundRealmId As String: foundRealmId = ""
    
    If Not realms Is Nothing Then
        If realms.count > 0 Then
            Dim r As Dictionary
            For Each r In realms
                If r.Exists("type") Then
                    If r("type") = "window" Then
                        foundRealmId = CStr(r("realm"))
                        Exit For
                    End If
                End If
            Next r
            If foundRealmId = "" Then foundRealmId = CStr(realms(1)("realm"))
        End If
    End If
    
    If foundRealmId <> "" Then
        m_cachedRealmIds(targetCtx) = foundRealmId
    End If
    
    GetScriptRealmId = foundRealmId
End Function

Private Function GetCdpSessionId(Optional ByVal forceUpdate As Boolean = False) As String
    If m_mainContextId = "" Then UpdateMainContextId True, True
    
    If Not forceUpdate And m_cdpSessionId <> "" Then
        GetCdpSessionId = m_cdpSessionId: Exit Function
    End If

    Dim params As New Dictionary: params.Add "context", m_mainContextId
    
    Dim strRes As String: strRes = ExecuteBiDiCommand("goog:cdp.getSession", params, True)
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    Dim sessionId As String: sessionId = GetPathString(json, "result.session")
    
    If sessionId <> "" Then
        m_cdpSessionId = sessionId: m_consecutiveCdpFailures = 0: GetCdpSessionId = m_cdpSessionId
    Else
        m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1: GetCdpSessionId = ""
    End If
End Function

Public Sub DetachCurrentCdpSession()
    If m_cdpSessionId = "" Then Exit Sub
    
    On Error Resume Next
    Dim params As New Dictionary
    params.Add "session", m_cdpSessionId
    ExecuteBiDiCommand "goog:cdp.detachSession", params
    On Error GoTo 0
    
    m_cdpSessionId = ""
    m_consecutiveCdpFailures = 0
End Sub

' AutoClicker uses script.addPreloadScript; XPath and timeout are inlined because the
' preload function is invoked without arguments.
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeoutMs As Long = 30000)
    Me.ExecuteSessionSubscribe "log.entryAdded", , True
    
    Dim vTimeoutMs As Long
    vTimeoutMs = SanitizeMs(timeoutMs, "timeoutMs", 30000)

    Dim sourceJs As String: sourceJs = GetActionJs("autoclicker")
    Dim safeXpath As String: safeXpath = WebJsonConverter.ConvertToJson(xpath)
    sourceJs = Replace(sourceJs, "{{XPATH}}", safeXpath)
    sourceJs = Replace(sourceJs, "{{TIMEOUT}}", CStr(vTimeoutMs))
    
    Dim params As New Dictionary: params.Add "functionDeclaration", sourceJs
    
    Dim strRes As String: strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True, allowAmbiguousRetry:=False)
    Dim scriptId As String
    scriptId = GetPathString(WebJsonConverter.ParseJson(strRes), "result.script")
    
    If scriptId <> "" Then
        Debug.Print "BiDi: AutoClicker registered for XPath: " & xpath & " (ID: " & scriptId & ")"
    Else
        Debug.Print "BiDi Error: AutoClicker registration failed."
    End If
End Sub

Private Function BuildShadowSelectorArg(ByVal selectorsArray As Variant) As Dictionary
    Dim values As New Collection, i As Long
    If IsArray(selectorsArray) Then
        For i = LBound(selectorsArray) To UBound(selectorsArray)
            values.Add NewLocalValue("string", CStr(selectorsArray(i)))
        Next i
    Else
        values.Add NewLocalValue("string", CStr(selectorsArray))
    End If

    Dim arg As New Dictionary
    arg.Add "type", "array"
    arg.Add "value", values
    Set BuildShadowSelectorArg = arg
End Function

Public Sub StartDiscoveryLog(Optional ByVal excludeImagesAndCss As Boolean = True)
    ' Stage all recorder-owned objects and scalars locally. No live recorder state is changed
    ' until session.subscribe succeeds, so a failed start leaves an existing recorder intact
    ' and never creates a half-enabled new recorder.
    Dim stagedLogs As New Collection
    Dim stagedSuppByKey As New Dictionary
    Dim stagedStartTime As Double
    Dim events(4) As String

    On Error GoTo StartFailed

    ' Recorder timestamps are seconds; NowMs values are divided by 1000 consistently.
    stagedStartTime = socket_.nowMs() / 1000

    events(0) = "network.responseCompleted"
    events(1) = "network.fetchError"
    events(2) = "browsingContext.load"
    events(3) = "browsingContext.domContentLoaded"
    events(4) = "network.beforeRequestSent"

    ' Subscribe before committing recorder state. ExecuteSessionSubscribe raises on failure;
    ' in that case all fields below remain exactly as they were on entry.
    Me.ExecuteSessionSubscribe events

    ' Commit the staged recorder as one logical transaction. Publish m_isRecording last so
    ' guarded append paths never observe an enabled recorder with uninitialized collections.
    m_isRecording = False
    Set m_recorderLogs = stagedLogs
    Set m_recSuppByKey = stagedSuppByKey
    m_recorderStartTime = stagedStartTime
    m_excludeResources = excludeImagesAndCss
    m_recSuppOverflowCount = 0
    m_droppedHeadCount = 0

    ' A reused wrapper starts recording with all carryover diagnostics cleared.
    m_unresolvedJsonThisWait = False
    m_unresolvedJsonFirstKeyThisWait = ""
    m_unresolvedJsonFirstSecThisWait = 0#
    m_hasCarryoverHintForNextWait = False
    m_carryoverNoteArmedThisWait = False
    m_carryoverNoteEmittedThisWait = False

    m_isRecording = True
    Debug.Print "BiDi Recorder: Started. (Filter Resources: " & excludeImagesAndCss & ")"
    Exit Sub

StartFailed:
    Dim errNum As Long: errNum = Err.Number
    Dim errSrc As String: errSrc = Err.source
    Dim errDesc As String: errDesc = Err.description
    On Error GoTo 0
    Err.Raise errNum, errSrc, errDesc
End Sub

Public Sub StopAndSaveDiscoveryLog(Optional ByVal filePath As String = ".\discovery_log.txt")
    ' Drain the queue before setting m_isRecording=False because event dispatch records only
    ' while that flag is True.
    Call ProcessRecorderQueue

    m_isRecording = False

    If m_recorderLogs Is Nothing Then Exit Sub

    AppendSuppressedSummary
    
    If m_recorderLogs.count > 0 Then
        ' Resolve relative paths against the folder of the document that owns the active VBA project.
        filePath = WebShared.getFullLocalPath(filePath, targetExists:=False)

        ' ADODB.Stream is used to write UTF-8 without a BOM.
        Dim stm As ADODB.stream
        Set stm = New ADODB.stream
        stm.Type = 2                      ' adTypeText
        stm.Charset = "utf-8"
        stm.Open

        stm.WriteText "=== BiDi Discovery Log (Optimized for AI Analysis) ===" & vbCrLf
        stm.WriteText "Timestamp | Type | Method/Status | MIME | URL/Detail" & vbCrLf
        stm.WriteText "CONFIG | minStableMs(default)=" & DEFAULT_IDLE_STABLE_MS & _
                      " | maxTimeoutMs(default)=" & DEFAULT_IDLE_TIMEOUT_MS & _
                      " | excludeResources=" & m_excludeResources & vbCrLf
        ' Ring-buffer head eviction is reported so a truncated timeline is never mistaken for
        ' a complete capture.
        If m_droppedHeadCount > 0 Then
            stm.WriteText "TRUNCATED | dropped " & m_droppedHeadCount & _
                          " oldest timeline line(s); only the most recent " & MAX_LOG_SIZE & _
                          " event(s) are shown. Earlier events existed but were evicted (ring buffer)." & vbCrLf
        End If
        stm.WriteText "LEGEND | [RES] carries: id=<reqId; a redirect chain reuses the SAME id, pairing its hops> key=<host+path, volatile query stripped, rpcids kept> idle=<skip|track: was this URL in the idle probes ignore set>" & vbCrLf
        stm.WriteText "LEGEND | [REQ]=request start (network.beforeRequestSent; recording-only, not wired to consensus). [REQ] pairs with [RES] on the shared id. Per-request latency (dur=) is not computed." & vbCrLf
        stm.WriteText "LEGEND | [SPA-ACT]=DOM mutation burst  [SPA-STABLE]=consensus  [SPA-WAIT-BEGIN/END]=one wait (END has outcome + slackMs)  [SUPPRESSED]=filtered-noise tally" & vbCrLf
        stm.WriteText "LEGEND | times are VBA poll-based (~100ms granularity), do not over-trust sub-100ms gaps. minStableMs/maxTimeoutMs above are DEFAULTS; per-action effective values appear on each [SPA-WAIT-BEGIN]." & vbCrLf
        stm.WriteText "LEGEND | [SPA-SIGNAL-ARM/HIT]=completion-signal gate (type=dom|net|vis; Arm*Signal APIs). A HIT opens the gate AND resets the DOM-quiet timer; quietAtHit=<quiet ms discarded by that reset>. net HITs are event-triggered and timestamped at VBA dispatch time; queue backlog can add positive lag. dom/vis HITs are poll-detected (~100ms). Gated waits carry signal=<armed types> on [SPA-WAIT-BEGIN] and signal=<type>:<hit|miss> per armed type on [SPA-WAIT-END]; ungated waits omit the field." & vbCrLf
        stm.WriteText "LEGEND | signal miss at TIMEOUT = the armed signal never arrived (selector/pattern error, action never fired, or the site produced no change). signal hit at TIMEOUT = signal arrived but post-signal quiet never settled. [SPA-SIGNAL-LOST] = navigation destroyed an armed dom/vis observer mid-wait; the gate stays closed (fail-safe) so that wait ends TIMEOUT with miss." & vbCrLf
        stm.WriteText "LEGEND | [SPA-CONSENSUS-SUSPECT]=this wait reached STABLE while a tracked json/API it saw had NO content render -> the STABLE may be premature (firstJson=<key>@<sec> is the earliest such response; note is advisory). [SPA-CARRYOVER-NOTE]=the IMMEDIATELY next wait then showed a content render that MAY be that previous action's delayed paint. Both are heads-up markers only: no owner is asserted, no gap is computed, consensus is unchanged. Confirm with an armed re-run (see section 3)." & vbCrLf
        stm.WriteText vbCrLf
        stm.WriteText "=== ANALYSIS REQUEST (for AI) ===" & vbCrLf
        stm.WriteText "Using ONLY the data in this log, produce three things for THIS site:" & vbCrLf
        stm.WriteText vbCrLf
        stm.WriteText "1) Network config suggestions" & vbCrLf
        stm.WriteText "   - IGNORE (bidi.AddIdleIgnoreNetworkPattern): background telemetry / high-freq APIs that" & vbCrLf
        stm.WriteText "     should NOT gate the SPA idle wait. IGNORE is non-destructive (the request still fires)," & vbCrLf
        stm.WriteText "     so it is the safe default for first-party telemetry (e.g. /log, generate_204, autocomplete)." & vbCrLf
        stm.WriteText "   - BLOCK (bidi.ExecuteEnableResourceBlocking): ONLY requests safe to kill (third-party" & vbCrLf
        stm.WriteText "     ads/analytics/trackers, or bulk resources not needed). BLOCK is page-altering. NEVER block" & vbCrLf
        stm.WriteText "     hosts that also serve app code/content (e.g. a CDN serving the page's JS), auth/SSO" & vbCrLf
        stm.WriteText "     endpoints, or anything page logic depends on. If unsure -> IGNORE, not BLOCK." & vbCrLf
        stm.WriteText "   - Derive patterns from URLs ACTUALLY seen below (do not guess). '*' = wildcard ""contains"";" & vbCrLf
        stm.WriteText "     no '*' = literal substring. Do not over-broaden (use ""/log?"" literal, NOT ""*/log*"", which" & vbCrLf
        stm.WriteText "     would also match /login). Match what was observed (e.g. generate_204, not gen_204)." & vbCrLf
        stm.WriteText vbCrLf
        stm.WriteText "2) Wait-time assessment" & vbCrLf
        stm.WriteText "   - Per [SPA-WAIT-END]: slackMs = domQuietMs - minStableMs (margin over the stability gate)." & vbCrLf
        stm.WriteText "     Flag any wait with slack < ~100ms as marginal -> consider raising minStableMs, especially" & vbCrLf
        stm.WriteText "     if that action is capture/timing-sensitive." & vbCrLf
        stm.WriteText "   - Check maxElapsed vs maxTimeoutMs (headroom). Any TIMEOUT/HALT = wait too short or a real stall." & vbCrLf
        stm.WriteText "   - Long elapsed caused by slow server responses is NOT a threshold problem; do not raise" & vbCrLf
        stm.WriteText "     minStableMs for that. Infer server latency from the timestamp gaps between consecutive" & vbCrLf
        stm.WriteText "     [RES]/[NAV] lines (per-request dur is not recorded; see LEGEND)." & vbCrLf
        stm.WriteText "   - Gated waits ([SPA-WAIT-BEGIN] carries signal=): a TIMEOUT with signal=..:miss is NOT a" & vbCrLf
        stm.WriteText "     threshold problem -- do not suggest raising minStableMs for it; fix the armed selector/" & vbCrLf
        stm.WriteText "     pattern or the action instead. quietAtHit > minStableMs on a net/vis HIT is a settle-to-" & vbCrLf
        stm.WriteText "     render gap the gate absorbed; for dom HITs use section 3 instead (a dom HIT is itself a" & vbCrLf
        stm.WriteText "     mutation, so its quietAtHit is structurally near zero and proves nothing)." & vbCrLf
        stm.WriteText vbCrLf
        stm.WriteText "3) Settle-to-render suspicion check (early-STABLE warning; the tool pre-computes this)" & vbCrLf
        stm.WriteText "   - The tool already emits two advisory markers; you do NOT re-derive them by pairing:" & vbCrLf
        stm.WriteText "     [SPA-CONSENSUS-SUSPECT] = this wait hit STABLE while a tracked json/API it saw produced" & vbCrLf
        stm.WriteText "       no content render -> treat that STABLE as possibly premature (NOT clean for capture)." & vbCrLf
        stm.WriteText "     [SPA-CARRYOVER-NOTE] = the immediately next wait then showed a content render that MAY be" & vbCrLf
        stm.WriteText "       the previous action's delayed paint. It is a heads-up, not an attribution." & vbCrLf
        stm.WriteText "   - Deliberately NOT provided (do not attempt to compute or demand them): an exact settle-to-" & vbCrLf
        stm.WriteText "     render gap, an owner pairing (which json produced which render), or an AMBIGUOUS/CONFIRMED" & vbCrLf
        stm.WriteText "     verdict. The markers only assert 'this run is not clean', never 'this render belongs to X'." & vbCrLf
        stm.WriteText "   - The check is adjacent-wait only: a hint lives for exactly ONE following wait, so a SUSPECT" & vbCrLf
        stm.WriteText "     with no NOTE in the very next wait is not chased further. Absence of a NOTE is not 'safe'." & vbCrLf
        stm.WriteText "   - What to do with a SUSPECT/NOTE (do NOT just raise minStableMs on an ungated wait):" & vbCrLf
        stm.WriteText "     re-run with an explicit gate from observed lines to confirm and to get the real timing --" & vbCrLf
        stm.WriteText "       bidi.ArmContentSignal ""<a targeted xpath seen in [SPA-ACT], e.g. //*[@id='table-body']>""" & vbCrLf
        stm.WriteText "       and/or bidi.ArmNetworkSignal ""<substring of the SUSPECT's firstJson key>""." & vbCrLf
        stm.WriteText "     With net+content armed, the HITs give a much cleaner measurement than ungated inference" & vbCrLf
        stm.WriteText "     (net HIT is event-triggered and timestamped at VBA dispatch time; queue backlog can add positive lag; dom/vis HIT is poll-detected at ~100ms) and resolve the owner" & vbCrLf
        stm.WriteText "     without guessing. Fallback: on a net-armed wait with no stable dom" & vbCrLf
        stm.WriteText "     selector, raising THAT action's minStableMs to (measured gap + margin) is valid per-action." & vbCrLf
        stm.WriteText "   - Beware a stale render tripping a naive gate: arming a bare content xpath can HIT on the" & vbCrLf
        stm.WriteText "     PREVIOUS action's late paint. Prefer a value-identifying selector for the expected content." & vbCrLf
        stm.WriteText vbCrLf
        stm.WriteText "Per-line features to use: id (same id across a redirect chain), key (normalized endpoint)," & vbCrLf
        stm.WriteText "idle (skip|track), [SPA-ACT] (DOM mutation), [SUPPRESSED] (already-filtered noise tally)." & vbCrLf
        stm.WriteText "=== END REQUEST ===" & vbCrLf
        stm.WriteText "--------------------------------------------------------" & vbCrLf

        Dim logItem As Variant
        For Each logItem In m_recorderLogs
            stm.WriteText logItem & vbCrLf
        Next logItem

        stm.position = 0
        stm.Type = 1                      ' adTypeBinary
        stm.position = 3                  ' skip EF BB BF
        Dim utf8Bytes() As Byte
        utf8Bytes = stm.Read
        stm.Close

        Dim outStm As ADODB.stream
        Set outStm = New ADODB.stream
        outStm.Type = 1                   ' adTypeBinary
        outStm.Open
        outStm.Write utf8Bytes
        outStm.SaveToFile filePath, 2     ' adSaveCreateOverWrite
        outStm.Close
        Debug.Print "BiDi Recorder: Log saved to " & filePath
    End If

    Set m_recorderLogs = Nothing
    Set m_recSuppByKey = Nothing
    m_recSuppOverflowCount = 0
End Sub

Public Sub RecordEventsForSeconds(ByVal seconds As Long)
    If seconds <= 0 Then Exit Sub

    ' Arm the handler and validate socket_ before starting the timed recorder loop.
    On Error GoTo ErrorHandler

    If socket_ Is Nothing Then
        Err.Raise 91, "BiDi:RecordEventsForSeconds", "Socket object is not available."
    End If

    Dim startMs As Double: startMs = socket_.nowMs()
    Dim durationMs As Double: durationMs = CDbl(seconds) * 1000#

    Do While (socket_.nowMs() - startMs) < durationMs
        ProcessRecorderQueue

        If durationMs > 5000 Then
            socket_.PumpMessages 100
            socket_.YieldExecution 100, False
        Else
            socket_.YieldExecution 100, socket_.AllowUiPump
        End If
    Loop

Cleanup:
    ' Always release the recorder processing lock after an error.
    m_isProcessingQueue = False
    
    Exit Sub

ErrorHandler:
    ' Use Resume before Cleanup to deactivate VBA's active error handler; otherwise an error
    ' inside Cleanup can terminate the call stack instead of being handled.
    Debug.Print "BiDi Warning in RecordEventsForSeconds: " & Err.description
    Resume Cleanup
End Sub

Public Sub ProcessRecorderQueue()
    ProcessRecorderQueueCore False
End Sub

Private Sub DrainNonBlockingRecorderQueue()
    ProcessRecorderQueueCore True
End Sub

Private Sub ProcessRecorderQueueCore(ByVal nonBlocking As Boolean)
    If m_isProcessingQueue Then Exit Sub
    
    m_isProcessingQueue = True
    Dim json As Object
    Dim processedCount As Long
    Dim consecutiveErrors As Long
    Dim droppedCount As Long
    Dim dropLogCount As Long
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    Const MAX_DROP_LOGS_PER_BATCH As Long = 5
    
    On Error GoTo Cleanup
    
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        Set json = Nothing

        Dim dropReason As String
        Dim messageConsumed As Boolean
        messageConsumed = socket_.TryGetParsedEvent(json, dropReason, MIN_MSG_SIZE, MAX_PARSED_EVENT_SIZE, nonBlocking)
        If Not messageConsumed Then Exit Do

        ' Count valid and dropped messages alike so malformed traffic cannot bypass the batch cap.
        processedCount = processedCount + 1

        If json Is Nothing Then
            droppedCount = droppedCount + 1
            If m_debugMode And dropLogCount < MAX_DROP_LOGS_PER_BATCH Then
                Debug.Print "BiDi: Dropped recorder event: " & dropReason
                dropLogCount = dropLogCount + 1
            End If
        Else
            On Error Resume Next
            DispatchGlobalEvents json

            If Err.Number <> 0 Then
                consecutiveErrors = consecutiveErrors + 1
                If m_debugMode Then
                    Debug.Print "BiDi: Dispatch error #" & consecutiveErrors & ": " & Err.description
                End If
                Err.Clear
            Else
                consecutiveErrors = 0
            End If
            On Error GoTo Cleanup ' Restore main handler

            ' Repeated dispatch failures trip a circuit breaker rather than looping indefinitely.
            If consecutiveErrors >= 10 Then
                Debug.Print "BiDi Critical: 10 consecutive dispatch errors. Abandoning queue."
                Exit Do
            End If
        End If

        If Not nonBlocking Then
            If (processedCount Mod 100) = 0 Then socket_.PumpMessages 100
        End If
    Loop

    If m_debugMode And droppedCount > dropLogCount Then
        Debug.Print "BiDi: " & (droppedCount - dropLogCount) & _
                    " additional recorder event drop(s) suppressed in this batch."
    End If

Cleanup:
    m_isProcessingQueue = False
    
    If Err.Number <> 0 Then
        Debug.Print "BiDi Critical Error in ProcessRecorderQueueCore: " & Err.description
    End If
End Sub


Private Sub FormatAndStoreLog(ByVal json As Object)
    ' Guard keys explicitly; blanket OERN would hide malformed event structures.

    If Not json.Exists("method") Then Exit Sub
    Dim method As String: method = json("method")

    Select Case method
        Case "network.responseCompleted", "network.fetchError", "browsingContext.load", "browsingContext.domContentLoaded", "log.entryAdded", "network.beforeRequestSent"
        Case Else
            Exit Sub
    End Select

    Dim params As Object
    Dim LogLine As String
    Dim relTime As String

    If Not json.Exists("params") Then Exit Sub
    Set params = json("params")

    relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"

    Select Case method
        Case "network.responseCompleted"
            If Not params.Exists("response") Then Exit Sub
            If Not IsObject(params("response")) Then Exit Sub
            Dim res As Object: Set res = params("response")
            If Not res.Exists("url") Then Exit Sub
            If Not res.Exists("mimeType") Then Exit Sub
            Dim resUrl As String: resUrl = CStr(res("url"))
            Dim mime As String: mime = CStr(res("mimeType"))
            Dim resKey As String: resKey = NormalizeEndpointKey(resUrl)
            Dim resId As String: resId = ""
            If params.Exists("request") Then
                If IsObject(params("request")) Then
                    If params("request").Exists("request") Then resId = CStr(params("request")("request"))
                End If
            End If

            If m_excludeResources Then
                If IsNoiseUrl(resUrl) Or IsNoiseMime(mime) Then
                    BumpSuppressed resKey
                    Exit Sub
                End If
            End If

            Dim resStatus As String: resStatus = ""
            If res.Exists("status") Then resStatus = CStr(res("status"))
            LogLine = "[" & relTime & "] [RES] " & resStatus & " [" & mime & "] | id=" & ShortId(resId) & _
                      " | key=" & resKey & " | idle=" & IIf(IsIdleIgnored(resUrl), "skip", "track")

            ' Only tracked API responses arriving during an active wait can create this wait's
            ' unresolved-render hint.
            If m_isRecording And m_waitStackDepth > 0 And Not IsIdleIgnored(resUrl) Then
                If InStr(1, LCase$(mime), "json") > 0 Then
                    If Not m_unresolvedJsonThisWait Then
                        m_unresolvedJsonFirstKeyThisWait = resKey
                        m_unresolvedJsonFirstSecThisWait = (socket_.nowMs() / 1000) - m_recorderStartTime
                    End If
                    m_unresolvedJsonThisWait = True
                End If
            End If

        Case "network.beforeRequestSent"
            If Not params.Exists("request") Then Exit Sub
            If Not IsObject(params("request")) Then Exit Sub
            Dim reqObj As Object: Set reqObj = params("request")
            If Not reqObj.Exists("url") Then Exit Sub
            Dim reqUrl As String: reqUrl = CStr(reqObj("url"))
            If m_excludeResources Then
                If IsNoiseUrl(reqUrl) Then BumpSuppressed NormalizeEndpointKey(reqUrl): Exit Sub
            End If
            Dim reqId As String: reqId = ""
            If reqObj.Exists("request") Then reqId = CStr(reqObj("request"))
            Dim reqMethod As String: reqMethod = ""
            If reqObj.Exists("method") Then reqMethod = CStr(reqObj("method"))
            LogLine = "[" & relTime & "] [REQ] " & reqMethod & " | id=" & ShortId(reqId) & _
                      " | key=" & NormalizeEndpointKey(reqUrl) & _
                      " | idle=" & IIf(IsIdleIgnored(reqUrl), "skip", "track")

        Case "network.fetchError"
            ' Log blocker matches as [BLOCK] and all other fetch failures as [ERR], even when
            ' Chromium supplies no error text.
            If Not params.Exists("request") Then Exit Sub
            If Not IsObject(params("request")) Then Exit Sub
            Dim ferReq As Object: Set ferReq = params("request")
            If Not ferReq.Exists("url") Then Exit Sub
            Dim ferUrl As String: ferUrl = CStr(ferReq("url"))
            If m_excludeResources Then
                If IsNoiseUrl(ferUrl) Then BumpSuppressed NormalizeEndpointKey(ferUrl): Exit Sub
            End If

            Dim wasBlocked As Boolean: wasBlocked = False
            If m_blockingPatterns.count > 0 Then
                Dim ferTok As Variant
                For Each ferTok In m_blockingPatterns
                    If UrlMatchesPattern(ferUrl, CStr(ferTok)) Then
                        wasBlocked = True
                        Exit For
                    End If
                Next
            End If

            If wasBlocked Then
                LogLine = "[" & relTime & "] [BLOCK] - " & NormalizeEndpointKey(ferUrl)
            Else
                Dim ferErr As String: ferErr = ""
                If params.Exists("errorText") Then ferErr = CStr(params("errorText"))
                LogLine = "[" & relTime & "] [ERR] " & ferErr & " - " & NormalizeEndpointKey(ferUrl)
            End If

        Case "browsingContext.load", "browsingContext.domContentLoaded"
            If Not params.Exists("url") Then Exit Sub
            ' Normalize navigation URLs so opaque authentication payloads cannot bloat the log.
            LogLine = "[" & relTime & "] [NAV] " & Replace(method, "browsingContext.", "") & " - " & NormalizeEndpointKey(CStr(params("url")))

            ' A document load clears all carryover state; same-document SPA transitions do not.
            m_unresolvedJsonThisWait = False: m_unresolvedJsonFirstKeyThisWait = "": m_unresolvedJsonFirstSecThisWait = 0#
            m_hasCarryoverHintForNextWait = False: m_carryoverNoteArmedThisWait = False: m_carryoverNoteEmittedThisWait = False

        Case "log.entryAdded"
            If Not params.Exists("level") Then Exit Sub
            If CStr(params("level")) = "error" Then
                Dim logText As String: logText = ""
                If params.Exists("text") Then logText = CStr(params("text"))
                If Len(logText) > MAX_CONSOLE_LOG_TEXT_SIZE Then _
                    logText = Left$(logText, MAX_CONSOLE_LOG_TEXT_SIZE) & "...(truncated)"
                LogLine = "[" & relTime & "] [CONSOLE-ERR] " & logText
            End If
    End Select

    If LogLine <> "" Then PushTimelineLine LogLine
End Sub

' PushTimelineLine is the sole timeline-line append point and accounts for evicted head lines.
' lines.
Private Sub PushTimelineLine(ByVal s As String)
    If m_recorderLogs Is Nothing Then Exit Sub
    If m_recorderLogs.count >= MAX_LOG_SIZE Then
        m_recorderLogs.Remove 1
        m_droppedHeadCount = m_droppedHeadCount + 1
    End If
    m_recorderLogs.Add s
End Sub

' Shared pattern semantics follow the Discovery Log contract: no '*' means a
' case-insensitive substring; '*' is the only wildcard and all other characters are literal.
Private Function UrlMatchesPattern(ByVal url As String, ByVal pat As String) As Boolean
    If Len(pat) = 0 Then Exit Function

    Dim u As String: u = LCase$(url)
    Dim p As String: p = LCase$(pat)
    If InStr(p, "*") = 0 Then
        UrlMatchesPattern = (InStr(1, u, p, vbBinaryCompare) > 0)
        Exit Function
    End If

    Dim parts() As String: parts = Split(p, "*")
    Dim searchPos As Long: searchPos = 1
    Dim i As Long
    Dim token As String
    Dim foundAt As Long

    For i = LBound(parts) To UBound(parts)
        token = parts(i)
        If Len(token) > 0 Then
            foundAt = InStr(searchPos, u, token, vbBinaryCompare)
            If foundAt = 0 Then Exit Function
            searchPos = foundAt + Len(token)
        End If
    Next i

    UrlMatchesPattern = True
End Function

Private Function IsNoiseUrl(ByVal url As String) As Boolean
    Dim lUrl As String: lUrl = LCase(url)
    
    If InStr(lUrl, "/_next/static/") > 0 Or _
       InStr(lUrl, "webworker.js") > 0 Or _
       InStr(lUrl, ".png") > 0 Or InStr(lUrl, ".jpg") > 0 Or InStr(lUrl, ".jpeg") > 0 Or _
       InStr(lUrl, ".gif") > 0 Or InStr(lUrl, ".webp") > 0 Or InStr(lUrl, ".svg") > 0 Or _
       InStr(lUrl, ".css") > 0 Or InStr(lUrl, ".woff") > 0 Or InStr(lUrl, ".woff2") > 0 Or _
       InStr(lUrl, ".ico") > 0 Or Left(lUrl, 5) = "data:" Then
        IsNoiseUrl = True
        Exit Function
    End If
    
    If InStr(lUrl, "google-analytics") > 0 Or _
       InStr(lUrl, "googletagmanager") > 0 Or _
       InStr(lUrl, "doubleclick") > 0 Or _
       InStr(lUrl, "pagead") > 0 Then
        IsNoiseUrl = True
        Exit Function
    End If

    ' Keep generic tracker endpoints segment-bound to avoid hiding business APIs.
    If UrlPathHasExactSegment(lUrl, "collect") Or _
       UrlPathHasExactSegment(lUrl, "pixel") Or _
       UrlPathHasExactSegment(lUrl, "beacon") Then
        IsNoiseUrl = True
        Exit Function
    End If

    If Not m_userIgnoreNetworkPatterns Is Nothing Then
        Dim pattern As Variant
        For Each pattern In m_userIgnoreNetworkPatterns
            If UrlMatchesPattern(lUrl, CStr(pattern)) Then
                IsNoiseUrl = True
                Exit Function
            End If
        Next pattern
    End If

    IsNoiseUrl = False
End Function

Private Function UrlPathHasExactSegment( _
    ByVal url As String, _
    ByVal segmentName As String) As Boolean

    Dim lUrl As String
    Dim lSegment As String
    Dim schemePos As Long
    Dim pathStart As Long
    Dim pathEnd As Long
    Dim queryPos As Long
    Dim fragmentPos As Long
    Dim pathText As String
    Dim needle As String
    Dim foundAt As Long
    Dim afterPos As Long
    Dim nextChar As String

    lUrl = LCase$(url)
    lSegment = LCase$(segmentName)
    If Len(lSegment) = 0 Then Exit Function

    schemePos = InStr(1, lUrl, "://", vbBinaryCompare)
    If schemePos > 0 Then
        pathStart = InStr(schemePos + 3, lUrl, "/")
        If pathStart = 0 Then Exit Function
    Else
        pathStart = 1
    End If

    pathEnd = Len(lUrl) + 1

    queryPos = InStr(pathStart, lUrl, "?")
    If queryPos > 0 And queryPos < pathEnd Then pathEnd = queryPos

    fragmentPos = InStr(pathStart, lUrl, "#")
    If fragmentPos > 0 And fragmentPos < pathEnd Then pathEnd = fragmentPos

    pathText = Mid$(lUrl, pathStart, pathEnd - pathStart)
    needle = "/" & lSegment
    foundAt = InStr(1, pathText, needle, vbBinaryCompare)

    Do While foundAt > 0
        afterPos = foundAt + Len(needle)
        If afterPos > Len(pathText) Then
            UrlPathHasExactSegment = True
            Exit Function
        End If

        nextChar = Mid$(pathText, afterPos, 1)
        If nextChar = "/" Or nextChar = ";" Then
            UrlPathHasExactSegment = True
            Exit Function
        End If

        foundAt = InStr(foundAt + 1, pathText, needle, vbBinaryCompare)
    Loop
End Function

Private Function IsNoiseMime(ByVal mime As String) As Boolean
    Dim lMime As String: lMime = LCase(mime)
    IsNoiseMime = (InStr(lMime, "image/") > 0 Or InStr(lMime, "font/") > 0 Or InStr(lMime, "css") > 0)
End Function

' Endpoint keys drop volatile query tokens but retain high-signal rpcids so repeated API
' families collapse without losing their discriminator.
Private Function NormalizeEndpointKey(ByVal url As String) As String
    On Error GoTo Fallback
    Dim s As String: s = url
    ' Collapse data: URIs before URL parsing; embedded namespace URLs inside payloads are
    ' not real hosts.
    If Left$(LCase$(s), 5) = "data:" Then
        Dim cm As Long: cm = InStr(s, ",")
        If cm > 0 Then NormalizeEndpointKey = Left$(s, cm) Else NormalizeEndpointKey = Left$(s, 40)
        Exit Function
    End If
    Dim p As Long: p = InStr(s, "://")
    If p > 0 Then s = Mid$(s, p + 3)
    Dim q As Long: q = InStr(s, "?")
    Dim pathPart As String, queryPart As String
    If q > 0 Then
        pathPart = Left$(s, q - 1)
        queryPart = Mid$(s, q + 1)
    Else
        pathPart = s
        queryPart = ""
    End If
    Dim keyStr As String: keyStr = pathPart
    Dim rp As String: rp = ExtractQueryToken(queryPart, "rpcids")
    If Len(rp) > 0 Then keyStr = keyStr & "?rpcids=" & rp
    Const MAX_KEY_LEN As Long = 160
    If Len(keyStr) > MAX_KEY_LEN Then keyStr = Left$(keyStr, MAX_KEY_LEN) & "..."
    NormalizeEndpointKey = keyStr
    Exit Function
Fallback:
    NormalizeEndpointKey = url
End Function

Private Function ExtractQueryToken(ByVal query As String, ByVal name As String) As String
    If Len(query) = 0 Then Exit Function
    Dim parts() As String: parts = Split(query, "&")
    Dim i As Long, pfx As String: pfx = LCase(name) & "="
    For i = LBound(parts) To UBound(parts)
        If LCase(Left$(parts(i), Len(pfx))) = pfx Then
            ExtractQueryToken = Mid$(parts(i), Len(pfx) + 1)
            Exit Function
        End If
    Next i
End Function

Private Function IsIdleIgnored(ByVal url As String) As Boolean
    Dim lUrl As String: lUrl = LCase(url)
    Dim defToks() As String: defToks = Split(DEFAULT_IGNORE_NETWORK, ",")
    Dim i As Long, t As String
    For i = LBound(defToks) To UBound(defToks)
        t = LCase(Trim$(defToks(i)))
        If Len(t) > 0 Then
            If UrlMatchesPattern(lUrl, t) Then IsIdleIgnored = True: Exit Function
        End If
    Next i
    If Not m_userIgnoreNetworkPatterns Is Nothing Then
        Dim pat As Variant
        For Each pat In m_userIgnoreNetworkPatterns
            If UrlMatchesPattern(lUrl, CStr(pat)) Then IsIdleIgnored = True: Exit Function
        Next pat
    End If
End Function

Private Function ShortId(ByVal idStr As String) As String
    If Len(idStr) = 0 Then ShortId = "-": Exit Function
    If Len(idStr) > 18 Then ShortId = Right$(idStr, 18) Else ShortId = idStr
End Function

Private Sub BumpSuppressed(ByVal keyStr As String)
    On Error Resume Next
    If m_recSuppByKey Is Nothing Then Exit Sub

    If m_recSuppByKey.Exists(keyStr) Then
        m_recSuppByKey.item(keyStr) = CLng(m_recSuppByKey.item(keyStr)) + 1
    ElseIf m_recSuppByKey.count < MAX_SUPPRESSED_KEYS Then
        m_recSuppByKey.Add keyStr, 1
    Else
        m_recSuppOverflowCount = m_recSuppOverflowCount + 1
    End If
    On Error GoTo 0
End Sub

Private Sub AppendSuppressedSummary()
    On Error Resume Next
    If m_recorderLogs Is Nothing Or m_recSuppByKey Is Nothing Then Exit Sub
    If m_recSuppByKey.count = 0 And m_recSuppOverflowCount = 0 Then Exit Sub

    Dim topKeys() As String
    Dim topCounts() As Long
    ReDim topKeys(0 To MAX_SUPPRESSED_SUMMARY_LINES - 1)
    ReDim topCounts(0 To MAX_SUPPRESSED_SUMMARY_LINES - 1)

    Dim total As Double: total = CDbl(m_recSuppOverflowCount)
    Dim topUsed As Long: topUsed = 0
    Dim k As Variant
    Dim countVal As Long
    Dim insertPos As Long
    Dim i As Long

    For Each k In m_recSuppByKey.keys
        countVal = CLng(m_recSuppByKey.item(k))
        total = total + CDbl(countVal)

        insertPos = topUsed
        For i = 0 To topUsed - 1
            If countVal > topCounts(i) Then
                insertPos = i
                Exit For
            End If
        Next i

        If insertPos < MAX_SUPPRESSED_SUMMARY_LINES Then
            If topUsed < MAX_SUPPRESSED_SUMMARY_LINES Then topUsed = topUsed + 1
            For i = topUsed - 1 To insertPos + 1 Step -1
                topCounts(i) = topCounts(i - 1)
                topKeys(i) = topKeys(i - 1)
            Next i
            topCounts(insertPos) = countVal
            topKeys(insertPos) = CStr(k)
        End If
    Next k

    Dim summaryLine As String
    summaryLine = "[end] [SUPPRESSED-TOTAL] " & Format$(total, "0") & _
                  " event(s) filtered, " & m_recSuppByKey.count & " distinct key(s)"
    If m_recSuppOverflowCount > 0 Then
        summaryLine = summaryLine & " (+" & m_recSuppOverflowCount & _
                      " event(s) not recorded by key, key limit reached)"
    End If
    summaryLine = summaryLine & " (excludeResources=" & m_excludeResources & ")"
    m_recorderLogs.Add summaryLine

    For i = 0 To topUsed - 1
        m_recorderLogs.Add "[end] [SUPPRESSED] count=" & topCounts(i) & " | key=" & topKeys(i)
    Next i

    If m_recSuppByKey.count > MAX_SUPPRESSED_SUMMARY_LINES Then
        m_recorderLogs.Add "[end] [SUPPRESSED] ... and " & _
                           (m_recSuppByKey.count - MAX_SUPPRESSED_SUMMARY_LINES) & _
                           " more recorded distinct key(s)"
    End If
    On Error GoTo 0
End Sub

Private Function QuerySpaProbe(Optional ByVal contextId As String = "") As Dictionary
    Dim strRes As String
    strRes = CallScript(GetProbeJs("query"), Nothing, contextId, False, False, 10000, _
                           Silent:=True, allowAmbiguousRetry:=True)
    
    Dim innerJsonStr As String: innerJsonStr = CStr(GetCleanResult(strRes))
    If innerJsonStr = "" Then Set QuerySpaProbe = Nothing: Exit Function

    On Error Resume Next
    Set QuerySpaProbe = WebJsonConverter.ParseJson(innerJsonStr)
    On Error GoTo 0
End Function

' Presence of window.__vbaIdleProbe is insufficient: a framework may replace Fetch/XHR or
' the input setter while leaving the shell object behind. Missing legacy health fields are
' intentionally unhealthy so the bounded reinjection path upgrades the page probe.
Private Function IsSpaProbeReady(ByVal probe As Dictionary) As Boolean
    On Error GoTo NotReady

    If probe Is Nothing Then Exit Function
    If Not probe.Exists("active") Then Exit Function
    If Not CBool(probe("active")) Then Exit Function
    If Not probe.Exists("version") Then Exit Function
    If CLng(probe("version")) <> SPA_PROBE_VERSION Then Exit Function
    If Not probe.Exists("health") Then Exit Function
    If Not IsObject(probe("health")) Then Exit Function

    Dim health As Dictionary
    Set health = probe("health")
    If health Is Nothing Then Exit Function

    If Not health.Exists("versionOk") Then Exit Function
    If Not health.Exists("fetchWrapped") Then Exit Function
    If Not health.Exists("xhrWrapped") Then Exit Function
    If Not health.Exists("observerActive") Then Exit Function
    If Not health.Exists("inputWrapped") Then Exit Function

    IsSpaProbeReady = CBool(health("versionOk")) And _
                      CBool(health("fetchWrapped")) And _
                      CBool(health("xhrWrapped")) And _
                      CBool(health("observerActive")) And _
                      CBool(health("inputWrapped"))
    Exit Function

NotReady:
    IsSpaProbeReady = False
End Function

Private Sub LogInvalidProbeSelectors(ByVal probe As Dictionary, ByVal contextId As String)
    On Error GoTo Done
    If probe Is Nothing Then Exit Sub
    If Not probe.Exists("invalidSelectors") Then Exit Sub
    If m_reportedInvalidSelectors Is Nothing Then Set m_reportedInvalidSelectors = New Dictionary

    Dim invalidItems As Collection
    Set invalidItems = probe("invalidSelectors")
    If invalidItems Is Nothing Then Exit Sub

    Dim item As Variant
    For Each item In invalidItems
        Dim selectorText As String: selectorText = CStr(item)
        If Len(selectorText) > 0 Then
            If Not m_reportedInvalidSelectors.Exists(selectorText) Then
                m_reportedInvalidSelectors.Add selectorText, True
                Dim safeText As String
                safeText = Replace(Replace(selectorText, vbCr, " "), vbLf, " ")
                If Len(safeText) > 200 Then safeText = Left$(safeText, 197) & "..."
                LogProbeDiagnosticLine "[SPA-FILTER-INVALID] selector=" & safeText & _
                                       " | ctx=" & contextId
            End If
        End If
    Next item
Done:
End Sub

' COMPLETION SIGNAL GATE
' SPA idle is negative evidence and can pass during a response-to-render gap. Arm* adds
' caller-selected positive evidence:
'   ArmContentSignal: existing subtree is rewritten.
'   ArmNetworkSignal: matching responseCompleted arrives.
'   ArmVisibilitySignal: previously absent/hidden element becomes visible.
' A hit opens the gate and restarts minStableMs. Arms are live immediately, multiple arms
' compose with AND, and the first following SPA-idle wait consumes them on any outcome.
' Do not insert another synchronizing call between Arm* and the intended action.

' ArmContentSignal requires an existing valid target. By default only childList and
' characterData count; attribute churn is opt-in.
Public Sub ArmContentSignal(ByVal signalXPath As String, _
                            Optional ByVal includeAttributes As Boolean = False, _
                            Optional ByVal contextId As String = "")
    If Len(Trim$(signalXPath)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ArmContentSignal", "signalXPath is empty."
    End If

    Dim targetCtx As String
    targetCtx = ResolveSignalContextForArm(contextId, "ArmContentSignal")

    Dim js As String
    js = "var p=window.__vbaIdleProbe;if(!p||!p.armDomSignal)return -9;" & _
         "return p.armDomSignal(" & JsStringLiteral(signalXPath) & "," & _
         LCase$(CStr(includeAttributes)) & ");"

    Dim rc As Long: rc = InvokeProbeArm(js, targetCtx)

    Select Case rc
        Case 0
            ClaimSignalOwner targetCtx
            If m_sigDomArmed Then LogSignalLine "[SPA-SIGNAL-ARM] type=dom | overwrite=true (previous dom arm replaced)"
            m_sigDomArmed = True: m_sigDomSeen = False: m_sigDomXPath = signalXPath
            m_sigLostWarned = False
            LogSignalLine "[SPA-SIGNAL-ARM] type=dom | xpath=" & signalXPath & " | ctx=" & targetCtx
        Case -1
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=not-found | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmContentSignal", _
                "ArmContentSignal: target element not found at arm time: " & signalXPath & vbCrLf & _
                "This method watches an EXISTING element for rewrites. To wait for an element " & _
                "that appears only after the action, use ArmVisibilitySignal instead."
        Case -2
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=invalid-xpath | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmContentSignal", _
                "ArmContentSignal: invalid XPath expression: " & signalXPath
        Case Else
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=probe-unavailable | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmContentSignal", _
                "ArmContentSignal: SPA probe unavailable in the target context (rc=" & rc & ")."
    End Select
End Sub

' ArmNetworkSignal uses the shared URL-pattern syntax and matches before ignore filters.
' Arrival is not rendering; increase the tail or combine with a DOM signal when needed.
Public Sub ArmNetworkSignal(ByVal urlPattern As String, _
                            Optional ByVal contextId As String = "")
    If Len(Trim$(urlPattern)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ArmNetworkSignal", "urlPattern is empty."
    End If

    Dim targetCtx As String
    targetCtx = ResolveSignalContextForArm(contextId, "ArmNetworkSignal")

    ' Ensure responseCompleted is subscribed even without the recorder; otherwise a network
    ' arm would silently never fire.
    Me.ExecuteSessionSubscribe Array("network.responseCompleted"), , True

    ClaimSignalOwner targetCtx
    If m_sigNetArmed Then LogSignalLine "[SPA-SIGNAL-ARM] type=net | overwrite=true (previous net arm replaced)"
    m_sigNetArmed = True: m_sigNetSeen = False: m_sigNetPattern = urlPattern
    m_sigNetContextMissWarned = False
    LogSignalLine "[SPA-SIGNAL-ARM] type=net | pattern=" & urlPattern & _
                  " | ctx=" & targetCtx
End Sub

' ArmVisibilitySignal rejects an already-visible target because that would silently
' behave as an ungated wait. The XPath is polled only until the first visibility latch.
Public Sub ArmVisibilitySignal(ByVal signalXPath As String, _
                               Optional ByVal contextId As String = "")
    If Len(Trim$(signalXPath)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:ArmVisibilitySignal", "signalXPath is empty."
    End If

    Dim targetCtx As String
    targetCtx = ResolveSignalContextForArm(contextId, "ArmVisibilitySignal")

    Dim js As String
    js = "var p=window.__vbaIdleProbe;if(!p||!p.armVisSignal)return -9;" & _
         "return p.armVisSignal(" & JsStringLiteral(signalXPath) & ");"

    Dim rc As Long: rc = InvokeProbeArm(js, targetCtx)

    Select Case rc
        Case 0
            ClaimSignalOwner targetCtx
            If m_sigVisArmed Then LogSignalLine "[SPA-SIGNAL-ARM] type=vis | overwrite=true (previous vis arm replaced)"
            m_sigVisArmed = True: m_sigVisSeen = False: m_sigVisXPath = signalXPath
            m_sigLostWarned = False
            LogSignalLine "[SPA-SIGNAL-ARM] type=vis | xpath=" & signalXPath & " | ctx=" & targetCtx
        Case -3
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=already-visible | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmVisibilitySignal", _
                "ArmVisibilitySignal: target is ALREADY visible at arm time: " & signalXPath & vbCrLf & _
                "Waiting for the appearance of something already visible is meaningless and " & _
                "would open the gate instantly. To wait for this element's content to be " & _
                "REWRITTEN, use ArmContentSignal instead."
        Case -2
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=invalid-xpath | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmVisibilitySignal", _
                "ArmVisibilitySignal: invalid XPath expression: " & signalXPath
        Case Else
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=probe-unavailable | xpath=" & signalXPath
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmVisibilitySignal", _
                "ArmVisibilitySignal: SPA probe unavailable in the target context (rc=" & rc & ")."
    End Select
End Sub

' DisarmSignal is abnormal-flow cleanup; normal arms are consumed automatically by the
' next SPA-idle wait.
Public Sub DisarmSignal(Optional ByVal contextId As String = "")
    Dim wasArmed As Boolean
    wasArmed = AnySignalArmed()

    Dim ownerCtx As String
    If wasArmed Then
        ownerCtx = m_signalOwnerContextId
        If Len(Trim$(contextId)) > 0 Then
            Dim requestedCtx As String: requestedCtx = ResolveTargetContext(contextId)
            If StrComp(requestedCtx, ownerCtx, vbBinaryCompare) <> 0 Then
                Err.Raise ERR_ACTION_FAILED, "BiDi:DisarmSignal", _
                    "DisarmSignal context does not own the active completion signals." & vbCrLf & _
                    "Owner: " & ownerCtx & vbCrLf & "Requested: " & requestedCtx
            End If
        End If
    Else
        ownerCtx = ResolveTargetContext(contextId)
    End If

    ResetSignalState

    ' Browser-side disarm is best-effort because a dead context already lost its observer.
    On Error Resume Next
    Call CallScript("var p=window.__vbaIdleProbe;if(p&&p.disarmSignal)p.disarmSignal();return 0;", _
                    Nothing, ownerCtx, False, False, 5000, _
                    Silent:=True, allowAmbiguousRetry:=True)
    On Error GoTo 0

    If wasArmed Then LogSignalLine "[SPA-SIGNAL-DISARM] scope=all | by=user | ctx=" & ownerCtx
End Sub

' Consume all arm types after WriteSpaWaitEnd on every SPA-wait outcome. Cleanup always
' targets the saved owner, even if a caller accidentally supplies a different context.
Private Sub ConsumeSignalArms(Optional ByVal contextId As String = "")
    If Not AnySignalArmed() Then Exit Sub

    Dim ownerCtx As String: ownerCtx = m_signalOwnerContextId
    If Len(ownerCtx) = 0 Then ownerCtx = ResolveTargetContext(contextId)

    ResetSignalState

    On Error Resume Next
    Call CallScript("var p=window.__vbaIdleProbe;if(p&&p.disarmSignal)p.disarmSignal();return 0;", _
                    Nothing, ownerCtx, False, False, 5000, _
                    Silent:=True, allowAmbiguousRetry:=True)
    On Error GoTo 0
End Sub

' Carryover render detection accepts targeted nested childList paths, not html/body root
' churn. Ancestor-ID paths are always considered targeted.
Private Function IsContentRenderPath(ByVal domPath As String) As Boolean
    Dim p As String: p = Trim$(domPath)
    If Len(p) = 0 Then IsContentRenderPath = False: Exit Function
    If InStr(p, "[@id=") > 0 Then IsContentRenderPath = True: Exit Function
    Dim segs() As String, i As Long, depth As Long
    segs = Split(p, "/")
    For i = LBound(segs) To UBound(segs)
        If Len(segs(i)) > 0 Then depth = depth + 1
    Next
    IsContentRenderPath = (depth >= 3)
End Function

Private Function AnySignalArmed() As Boolean
    AnySignalArmed = (m_sigDomArmed Or m_sigNetArmed Or m_sigVisArmed)
End Function

Private Function ResolveSignalContextForArm(ByVal contextId As String, _
                                            ByVal callerName As String) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    If Len(targetCtx) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  callerName & ": unable to resolve the target browsing context."
    End If

    If AnySignalArmed() Then
        If Len(m_signalOwnerContextId) = 0 Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                      callerName & ": active completion signals have no owner context."
        End If
        If StrComp(targetCtx, m_signalOwnerContextId, vbBinaryCompare) <> 0 Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                callerName & ": completion signals cannot span browsing contexts." & vbCrLf & _
                "Owner: " & m_signalOwnerContextId & vbCrLf & "Requested: " & targetCtx
        End If
    End If

    ResolveSignalContextForArm = targetCtx
End Function

Private Sub ClaimSignalOwner(ByVal targetCtx As String)
    If Len(m_signalOwnerContextId) = 0 Then m_signalOwnerContextId = targetCtx
End Sub

Private Sub ValidateSignalWaitContext(ByVal targetCtx As String, _
                                      Optional ByVal callerName As String = "WaitForNetworkIdleSpa")
    If Not AnySignalArmed() Then Exit Sub

    If Len(m_signalOwnerContextId) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Active completion signals have no owner context."
    End If

    If StrComp(targetCtx, m_signalOwnerContextId, vbBinaryCompare) <> 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
            "The requested context does not own the active completion signals." & vbCrLf & _
            "Owner: " & m_signalOwnerContextId & vbCrLf & "Requested: " & targetCtx
    End If
End Sub

Private Function BeginSignalCarryoverDiagnostic(ByVal actionName As String, _
                                                ByVal targetCtx As String, _
                                                ByRef domSeenBefore As Boolean, _
                                                ByRef netSeenBefore As Boolean, _
                                                ByRef visSeenBefore As Boolean) As Boolean
    If Not AnySignalArmed() Then Exit Function

    ValidateSignalWaitContext targetCtx, actionName

    domSeenBefore = m_sigDomSeen
    netSeenBefore = m_sigNetSeen
    visSeenBefore = m_sigVisSeen

    LogSignalLine "[SPA-SIGNAL-CARRYOVER] action=" & actionName & _
                  " | ownerCtx=" & m_signalOwnerContextId & _
                  " | state=" & BuildSignalEndDesc()
    BeginSignalCarryoverDiagnostic = True
End Function

Private Sub EndSignalCarryoverDiagnostic(ByVal actionName As String, _
                                         ByVal diagnosticActive As Boolean, _
                                         ByVal domSeenBefore As Boolean, _
                                         ByVal netSeenBefore As Boolean, _
                                         ByVal visSeenBefore As Boolean)
    If Not diagnosticActive Then Exit Sub

    Dim hitTypes As String
    If m_sigDomArmed And Not domSeenBefore And m_sigDomSeen Then hitTypes = "dom"
    If m_sigNetArmed And Not netSeenBefore And m_sigNetSeen Then
        If Len(hitTypes) > 0 Then hitTypes = hitTypes & ","
        hitTypes = hitTypes & "net"
    End If
    If m_sigVisArmed And Not visSeenBefore And m_sigVisSeen Then
        If Len(hitTypes) > 0 Then hitTypes = hitTypes & ","
        hitTypes = hitTypes & "vis"
    End If

    If Len(hitTypes) > 0 Then
        LogSignalLine "[SPA-SIGNAL-CARRYOVER-HIT] action=" & actionName & _
                      " | type=" & hitTypes & _
                      " | state=" & BuildSignalEndDesc()
    End If
End Sub

Private Sub ResetSignalState()
    m_sigDomArmed = False: m_sigDomSeen = False: m_sigDomXPath = ""
    m_sigNetArmed = False: m_sigNetSeen = False: m_sigNetPattern = ""
    m_sigVisArmed = False: m_sigVisSeen = False: m_sigVisXPath = ""
    m_signalOwnerContextId = ""
    m_sigLostWarned = False
    m_sigNetContextMissWarned = False
End Sub

Private Function InvokeProbeArm(ByVal jsBody As String, ByVal contextId As String) As Long
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)

    Call PrepareProbeForAction(targetCtx)

    Dim strRes As String
    strRes = CallScript(jsBody, Nothing, targetCtx, False, False, 10000, _
                           Silent:=True, allowAmbiguousRetry:=True)

    Dim v As Variant: v = GetCleanResult(strRes)
    If IsEmpty(v) Then
        InvokeProbeArm = -9
    Else
        InvokeProbeArm = CLng(v)
    End If
End Function

Private Function BuildSignalArmDesc() As String
    Dim s As String
    If m_sigDomArmed Then s = "dom"
    If m_sigNetArmed Then s = s & IIf(Len(s) > 0, ",", "") & "net"
    If m_sigVisArmed Then s = s & IIf(Len(s) > 0, ",", "") & "vis"
    BuildSignalArmDesc = s
End Function

Private Function BuildSignalEndDesc() As String
    Dim s As String
    If m_sigDomArmed Then s = "dom:" & IIf(m_sigDomSeen, "hit", "miss")
    If m_sigNetArmed Then s = s & IIf(Len(s) > 0, ",", "") & "net:" & IIf(m_sigNetSeen, "hit", "miss")
    If m_sigVisArmed Then s = s & IIf(Len(s) > 0, ",", "") & "vis:" & IIf(m_sigVisSeen, "hit", "miss")
    BuildSignalEndDesc = s
End Function

' HUD pending-signal text is derived from the same arm/latch flags as consensus, so a
' closed gate cannot be displayed as Stable.
Private Function BuildSignalPendingDesc() As String
    Dim s As String
    If m_sigDomArmed And Not m_sigDomSeen Then s = "dom"
    If m_sigNetArmed And Not m_sigNetSeen Then s = s & IIf(Len(s) > 0, ",", "") & "net"
    If m_sigVisArmed And Not m_sigVisSeen Then s = s & IIf(Len(s) > 0, ",", "") & "vis"
    BuildSignalPendingDesc = s
End Function

' LogSignalLine is the single path to both the ring-buffered log and developer trace.
Private Sub LogSignalLine(ByVal line As String)
    If m_isRecording Then
        Dim relTime As String
        relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
        PushTimelineLine "[" & relTime & "] " & line
    End If
    If m_debugMode Then Debug.Print "BiDi Signal: " & line
End Sub

Private Sub LogProbeDiagnosticLine(ByVal line As String)
    If m_isRecording Then
        Dim relTime As String
        relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
        PushTimelineLine "[" & relTime & "] " & line
    End If
    If m_debugMode Then Debug.Print "BiDi: " & line
End Sub

Private Function JsStringLiteral(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    JsStringLiteral = """" & t & """"
End Function

Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                 ByVal maxTimeoutMs As Long, _
                                 Optional ByVal contextId As String = "")
    
    ' Preserve the historical zero-means-default contract without emitting warnings.
    ' Positive values below 50 ms remain valid for backward compatibility.
    If minStableMs = 0 Then
        minStableMs = DEFAULT_IDLE_STABLE_MS
    Else
        minStableMs = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS, 1)
    End If

    If maxTimeoutMs = 0 Then
        maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    Else
        maxTimeoutMs = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS, 1)
    End If
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    ValidateSignalWaitContext targetCtx
    
    ProcessRecorderQueue

    Dim waitCtx As New Dictionary
    waitCtx.Add "contextId", targetCtx
    waitCtx.Add "minStableMs", minStableMs
    waitCtx.Add "stableStart", socket_.nowMs()
    waitCtx.Add "missingProbeCount", 0 ' Used for multi-stage JIT recovery
    ' Pre-seeded metrics are diagnostic only and guarantee a complete END line.
    waitCtx.Add "lastJsIdleDuration", 0#
    waitCtx.Add "lastInflFetch", 0
    waitCtx.Add "lastInflXhr", 0
    ' Network hits latch asynchronously. Reset stableStart once on the first poll that sees
    ' the hit; pre-wait hits need no extra reset because stableStart is already fresh.
    waitCtx.Add "sigNetAcked", m_sigNetSeen

    ' A carryover reservation becomes a one-wait arm at BEGIN and is cleared immediately;
    ' this wait's unresolved state is reset independently.
    m_carryoverNoteArmedThisWait = m_hasCarryoverHintForNextWait
    m_carryoverNoteEmittedThisWait = False
    m_hasCarryoverHintForNextWait = False
    m_unresolvedJsonThisWait = False
    m_unresolvedJsonFirstKeyThisWait = ""
    m_unresolvedJsonFirstSecThisWait = 0#
    
    ' Log effective per-operation minStableMs/maxTimeoutMs, not only class defaults.
    If m_isRecording Then
        Dim beginRelTime As String
        beginRelTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
        ' BEGIN includes armed types only for gated waits.
        Dim sigArmDesc As String: sigArmDesc = BuildSignalArmDesc()
        PushTimelineLine "[" & beginRelTime & "] [SPA-WAIT-BEGIN] minStableMs=" & minStableMs & _
                         " | maxTimeoutMs=" & maxTimeoutMs & _
                         IIf(Len(sigArmDesc) > 0, " | signal=" & sigArmDesc, "") & _
                         " | ctx=" & targetCtx
    End If

    MasterWaitEngine WaitMode_SpaIdle, maxTimeoutMs, 100, waitCtx
End Sub

Private Function EvaluateSpaIdleCondition(ByVal probe As Dictionary, _
                                          ByRef ctx As Dictionary, _
                                          ByVal currentMs As Double) As Boolean
    Dim jsNow As Double, lastActivityTs As Double, jsIdleDuration As Double
    Dim inflFetch As Long, inflXhr As Long
    Dim netIdle As Boolean, activityStable As Boolean

    Dim minStableMs As Long: minStableMs = CLng(ctx("minStableMs"))
    
    jsNow = CDbl(probe("now"))
    lastActivityTs = CDbl(probe("lastMutationTs"))
    inflFetch = CLng(probe("inflightFetchCount"))
    inflXhr = CLng(probe("inflightXhrCount"))
    
    If probe.Exists("inflightUrls") Then
        Set ctx("lastInflightUrls") = probe("inflightUrls")
    End If
    
    jsIdleDuration = jsNow - lastActivityTs

    ' Cached exit metrics are diagnostic and never influence consensus.
    ctx("lastJsIdleDuration") = jsIdleDuration
    ctx("lastInflFetch") = inflFetch
    ctx("lastInflXhr") = inflXhr

    ' Keep the pre-hit DOM-quiet duration so logs can show when a network hit prevented a
    ' premature consensus.
    m_lastProbeQuietMs = jsIdleDuration

    LogInvalidProbeSelectors probe, CStr(ctx("contextId"))

    If probe.Exists("changes") Then
        Dim changes As Collection: Set changes = probe("changes")
        If Not changes Is Nothing Then
            If changes.count > 0 Then
                Dim v As Variant, logSummary As String
                For Each v In changes: logSummary = logSummary & v & " ": Next
                logSummary = Trim(logSummary)
                
                If Len(logSummary) > 500 Then logSummary = Left$(logSummary, 497) & "..."
                If m_debugMode Then Debug.Print "[BiDi SPA Activity] " & logSummary
                
                If m_isRecording Then
                    Dim relTime As String
                    relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
                    PushTimelineLine "[" & relTime & "] [SPA-ACT] " & logSummary
                End If

                ' Only non-ignored targeted DOM childList mutations count as a content render; root churn
                ' and ATTR/PROP toggles do not.
                Dim ce As Variant, renderXPath As String: renderXPath = ""
                For Each ce In changes
                    If Left$(CStr(ce), 4) = "DOM:" Then
                        If IsContentRenderPath(Mid$(CStr(ce), 5)) Then
                            renderXPath = CStr(ce): Exit For
                        End If
                    End If
                Next
                If Len(renderXPath) > 0 Then
                    m_unresolvedJsonThisWait = False
                    ' An adjacent render may emit one non-committal carryover note; it never claims ownership.
                    If m_carryoverNoteArmedThisWait And Not m_carryoverNoteEmittedThisWait And m_isRecording Then
                        Dim rt As String
                        rt = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
                        PushTimelineLine "[" & rt & "] [SPA-CARRYOVER-NOTE]" & _
                                         " reason=prev-wait-unresolved-json | render=" & renderXPath & _
                                         " | note=this-render-may-belong-to-the-previous-action"
                        m_carryoverNoteEmittedThisWait = True
                        m_carryoverNoteArmedThisWait = False
                    End If
                End If
            End If
        End If
    End If

    ' Every signal hit opens the gate and resets stableStart, guaranteeing a full quiet
    ' window after the evidence rather than one polling interval.
    If m_sigDomArmed Or m_sigVisArmed Then
        ' Navigation destroys browser-side arms. Do not silently re-arm or un-gate; keep the
        ' gate closed, log SPA-SIGNAL-LOST, and let the wait time out visibly.
        Dim probeSigArmed As Boolean, probeVisArmed As Boolean
        If probe.Exists("sigArmed") Then probeSigArmed = CBool(probe("sigArmed"))
        If probe.Exists("visArmed") Then probeVisArmed = CBool(probe("visArmed"))
        If ((m_sigDomArmed And Not m_sigDomSeen And Not probeSigArmed) Or _
            (m_sigVisArmed And Not m_sigVisSeen And Not probeVisArmed)) And _
           Not m_sigLostWarned Then
            m_sigLostWarned = True
            LogSignalLine "[SPA-SIGNAL-LOST] reason=probe-reinjected (navigation?) | gate stays closed; expect TIMEOUT signal=miss"
        End If
    End If
    If m_sigDomArmed And Not m_sigDomSeen Then
        If probe.Exists("sigCount") Then
            If CLng(probe("sigCount")) > 0 Then
                m_sigDomSeen = True
                ctx("stableStart") = currentMs   ' reset: full window after the hit
                LogSignalLine "[SPA-SIGNAL-HIT] type=dom | count=" & CLng(probe("sigCount")) & _
                              " | quietAtHit=" & Format(jsIdleDuration, "0") & "ms"
            End If
        End If
    End If
    If m_sigVisArmed And Not m_sigVisSeen Then
        If probe.Exists("visSeen") Then
            If CBool(probe("visSeen")) Then
                m_sigVisSeen = True
                ctx("stableStart") = currentMs   ' reset: full window after the hit
                LogSignalLine "[SPA-SIGNAL-HIT] type=vis | quietAtHit=" & Format(jsIdleDuration, "0") & "ms"
            End If
        End If
    End If
    If m_sigNetArmed And m_sigNetSeen Then
        If ctx.Exists("sigNetAcked") Then
            If Not CBool(ctx("sigNetAcked")) Then
                ctx("sigNetAcked") = True
                ctx("stableStart") = currentMs   ' reset: full window after the hit
            End If
        End If
    End If

    netIdle = (inflFetch = 0) And (inflXhr = 0)

    activityStable = (jsIdleDuration >= CDbl(minStableMs))

    ' Multiple armed signal types compose with AND.
    Dim sigGateOpen As Boolean
    sigGateOpen = (Not m_sigDomArmed Or m_sigDomSeen) And _
                  (Not m_sigNetArmed Or m_sigNetSeen) And _
                  (Not m_sigVisArmed Or m_sigVisSeen)

    If netIdle And activityStable Then
        If Not sigGateOpen Then
            ' Do not reset stableStart while merely waiting for a signal; the hit itself owns the
            ' single reset.
            EvaluateSpaIdleCondition = False
            Exit Function
        End If
        If (currentMs - CDbl(ctx("stableStart"))) >= CDbl(minStableMs) Then
            If m_debugMode Then
                Call socket_.LogEvent("Consensus reached (" & minStableMs & "ms window satisfied)", "", ">>> [BiDi SPA Stable]")
            End If
            
            If m_isRecording Then
                Dim successRelTime As String
                successRelTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
                PushTimelineLine "[" & successRelTime & "] [SPA-STABLE] Consensus reached (minStableMs=" & minStableMs & ")"

                ' Consensus with a tracked response but no content render emits a diagnostic SUSPECT and
                ' reserves one adjacent-wait hint. It does not change the successful result.
                If m_unresolvedJsonThisWait Then
                    PushTimelineLine "[" & successRelTime & "] [SPA-CONSENSUS-SUSPECT]" & _
                                     " reason=track-json-no-content-render" & _
                                     " | firstJson=" & m_unresolvedJsonFirstKeyThisWait & "@" & Format(m_unresolvedJsonFirstSecThisWait, "0.00") & "s" & _
                                     " | note=next-content-render-may-be-stale"
                    m_hasCarryoverHintForNextWait = True
                End If
            End If
            
            EvaluateSpaIdleCondition = True ' Stability confirmed
        Else
            EvaluateSpaIdleCondition = False ' Cooling down...
        End If
    Else
        ctx("stableStart") = currentMs
        EvaluateSpaIdleCondition = False
    End If
End Function

Private Function ExecuteCDPCommand(ByVal cdpMethod As String, _
                                   ByVal cdpParams As Dictionary, _
                                   Optional ByVal useSession As Boolean = True, _
                                   Optional ByVal Silent As Boolean = False, _
                                   Optional ByVal sessionIdOverride As String = "", _
                                   Optional ByVal countFailure As Boolean = True) As Object
    Dim res As New Dictionary
    res.Add "Ok", False: res.Add "Json", "": res.Add "Error", ""

    If socket_ Is Nothing Then
        res("Error") = "Socket object is not set."
        Set ExecuteCDPCommand = res: Exit Function
    End If

    If countFailure And m_consecutiveCdpFailures >= MAX_CONSECUTIVE_CDP_FAILURES Then
        Call PerformFatalReset
        res("Error") = "Fatal: CDP instability. Context state has been cleared."
        Set ExecuteCDPCommand = res
        Exit Function
    End If

    Dim params As New Dictionary
    params.Add "method", cdpMethod
    params.Add "params", EnsureDictionary(cdpParams)

    If useSession Then
        Dim sessionId As String
        If Len(sessionIdOverride) > 0 Then
            sessionId = sessionIdOverride
        Else
            Dim sessionErrNum As Long
            Dim sessionErrDesc As String
            On Error Resume Next
            sessionId = GetCdpSessionId()
            sessionErrNum = Err.Number
            sessionErrDesc = Err.description
            On Error GoTo 0

            If sessionErrNum <> 0 Then
                If countFailure And sessionErrNum <> ERR_MAIN_CONTEXT_LOST Then
                    m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
                End If
                res("Error") = "CDP session error: " & sessionErrDesc
                Set ExecuteCDPCommand = res
                Exit Function
            End If
        End If

        If sessionId <> "" Then
            params.Add "session", sessionId
        Else
            If countFailure Then m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
            res("Error") = "CDP session unavailable."
            Set ExecuteCDPCommand = res: Exit Function
        End If
    End If

    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("goog:cdp.sendCommand", params, Silent)
    Dim cdpCallErr As Long: cdpCallErr = Err.Number
    Dim cdpCallDesc As String: cdpCallDesc = Err.description
    On Error GoTo 0
    
    If cdpCallErr <> 0 Then
        res("Error") = "Socket Error: " & cdpCallDesc
        If countFailure Then m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
        Set ExecuteCDPCommand = res: Exit Function
    End If

    Dim parsedCdpResponse As Object
    Dim cdpResponseError As String

    If Not IsCdpTransportResponseSuccess(strRes, parsedCdpResponse, cdpResponseError) Then
        res("Error") = cdpResponseError
        res("Json") = strRes
        If countFailure Then m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
        socket_.LogEvent "CDP Command FAILED: " & cdpMethod, strRes, "[CDP:ERR]"
    Else
        res("Ok") = True
        res("Json") = strRes
        If countFailure Then m_consecutiveCdpFailures = 0
        If m_debugMode And Not Silent Then
            socket_.LogEvent "CDP Command OK: " & cdpMethod, strRes, "[CDP:OK]"
        End If
    End If

    Set ExecuteCDPCommand = res
End Function

Private Function IsCdpTransportResponseSuccess( _
    ByVal jsonRaw As String, _
    ByRef parsedJson As Object, _
    ByRef errorText As String) As Boolean

    Set parsedJson = Nothing
    errorText = vbNullString

    On Error Resume Next
    Set parsedJson = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0

    If parsedJson Is Nothing Then
        errorText = "CDP Error: response is not valid JSON."
        Exit Function
    End If

    If StrComp(TypeName(parsedJson), "Dictionary", vbBinaryCompare) <> 0 Then
        errorText = "CDP Error: response root is not a JSON object."
        Exit Function
    End If

    If parsedJson.Exists("error") Then
        Dim errType As String
        Dim errMsg As String

        errType = GetPathString(parsedJson, "error")
        errMsg = GetPathString(parsedJson, "message")

        errorText = "CDP Error"
        If Len(errType) > 0 Then errorText = errorText & ": [" & errType & "]"
        If Len(errMsg) > 0 Then errorText = errorText & " " & errMsg
        Exit Function
    End If

    IsCdpTransportResponseSuccess = True
End Function

Private Sub PerformFatalReset()
    Debug.Print "BiDi Critical: Performing Fatal Reset due to protocol instability."
    ResetVolatileContextState
    m_consecutiveCdpFailures = 0
End Sub

' Returns False when the protocol or JavaScript injection failed. Wait-loop callers must
' not raise for a single injection failure; missingProbeCount / 505 owns persistent failure.
Private Function StartSpaIdleProbe(Optional ByVal contextId As String = "") As Boolean
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId(True, True)

    Dim targetRealm As String
    targetRealm = GetScriptRealmId(targetCtx, True, True)

    Dim p As New Dictionary
    p.Add "expression", GetProbeJs("inject")
    p.Add "target", BuildContextTarget(targetCtx, targetRealm)
    p.Add "awaitPromise", False

    Dim rawRes As String
    Dim validated As Object
    Dim injectErrNum As Long
    Dim injectErrDesc As String

    ' Transport failures retain their previous behavior and propagate. Only a returned BiDi
    ' protocol error or JavaScript exception is converted into a recorded probe miss.
    rawRes = ExecuteBiDiCommand("script.evaluate", p, True, _
                                allowContextRetarget:=(targetCtx = m_mainContextId))

    On Error GoTo InjectFailed
    Set validated = ValidateProtocolResponse(rawRes, "StartSpaIdleProbe")
    MarkProbeFilterCurrent targetCtx

    StartSpaIdleProbe = True
    Exit Function

InjectFailed:
    ' A swallowed response-level injection error otherwise reappears as an absorbed SPA timeout
    ' and misleads timeout tuning. Recording must not alter the wait control flow.
    injectErrNum = Err.Number
    injectErrDesc = Err.description
    On Error Resume Next
    If m_debugMode Then
        Debug.Print "BiDi: [PROBE-INJECT-FAIL] ctx=" & Left$(targetCtx, 16) & _
                    "... Err " & injectErrNum & ": " & injectErrDesc
    End If
    PushTimelineLine "[" & Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & _
                     "s] [PROBE-INJECT-FAIL] " & injectErrNum & " " & injectErrDesc
    Err.Clear
    On Error GoTo 0
    StartSpaIdleProbe = False
End Function

Private Function FormatJavaScriptError(ByVal details As Object) As String
    Dim msg As String
    Dim stack As String
    Dim Frames As Collection
    Dim frame As Object
    Dim i As Long

    msg = "JS Exception: " & details("text")
    
    If details.Exists("lineNumber") Then
        msg = msg & " [Line: " & details("lineNumber") & ", Col: " & details("columnNumber") & "]"
    End If
    If details.Exists("exception") Then
        If details("exception").Exists("description") Then
            msg = msg & vbCrLf & "Description: " & details("exception")("description")
        End If
    End If
    If details.Exists("stackTrace") Then
        Set Frames = details("stackTrace")("callFrames")
        stack = "--- JS Stack Trace ---"
        For i = 1 To Frames.count
            Set frame = Frames(i)
            stack = stack & vbCrLf & "  at " & frame("functionName") & " (" & frame("url") & ":" & frame("lineNumber") & ":" & frame("columnNumber") & ")"
        Next i
        msg = msg & vbCrLf & stack
    End If

    FormatJavaScriptError = msg
End Function

' Filter changes are pushed into an existing healthy probe on the next probe preparation;
' installed hooks read p.selectors dynamically, so they do not need to be rebuilt.
Public Sub AddIdleIgnoreSelector(ByVal cssSelector As String)
    If Len(Trim$(cssSelector)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:AddIdleIgnoreSelector", "cssSelector is empty."
    End If

    If m_userIgnoreSelectors Is Nothing Then Set m_userIgnoreSelectors = New Collection
    m_userIgnoreSelectors.Add cssSelector
    BumpIdleFilterRevision
End Sub

' Network hooks read p.ignoreNetwork dynamically for the same reason.
Public Sub AddIdleIgnoreNetworkPattern(ByVal urlPattern As String)
    If Len(Trim$(urlPattern)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:AddIdleIgnoreNetworkPattern", "urlPattern is empty."
    End If

    If m_userIgnoreNetworkPatterns Is Nothing Then Set m_userIgnoreNetworkPatterns = New Collection
    m_userIgnoreNetworkPatterns.Add urlPattern
    BumpIdleFilterRevision
End Sub

' Clear removes user additions; built-in defaults are restored on the next probe preparation.
Public Sub ClearIdleFilters()
    Set m_userIgnoreSelectors = Nothing
    Set m_userIgnoreNetworkPatterns = Nothing
    BumpIdleFilterRevision
End Sub

' Filter revisions are wrapper-wide, while each browsing context owns a separate probe.
' Track the revision applied to every context so updating one tab cannot mark another current.
Private Sub BumpIdleFilterRevision()
    If m_idleFilterRevision = 2147483647 Then
        m_idleFilterRevision = 1
        Set m_probeFilterRevisions = New Dictionary
    Else
        m_idleFilterRevision = m_idleFilterRevision + 1
    End If
End Sub

Private Function IsProbeFilterCurrent(ByVal contextId As String) As Boolean
    If m_probeFilterRevisions Is Nothing Then Set m_probeFilterRevisions = New Dictionary
    If Not m_probeFilterRevisions.Exists(contextId) Then Exit Function
    IsProbeFilterCurrent = (CLng(m_probeFilterRevisions(contextId)) = m_idleFilterRevision)
End Function

Private Sub MarkProbeFilterCurrent(ByVal contextId As String)
    If Len(contextId) = 0 Then Exit Sub
    If m_probeFilterRevisions Is Nothing Then Set m_probeFilterRevisions = New Dictionary
    m_probeFilterRevisions(contextId) = m_idleFilterRevision
End Sub

' Update only the mutable arrays. Existing healthy Fetch/XHR/DOM hooks remain unchanged.
Private Function PushIdleFiltersToProbe(ByVal contextId As String) As Boolean
    On Error GoTo PushFailed

    Dim js As String
    js = "var p=window.__vbaIdleProbe;if(!p||p.version!==" & SPA_PROBE_VERSION & ")return false;" & _
         "p.selectors=" & MergeToJsArray(DEFAULT_IGNORE_SELECTORS, m_userIgnoreSelectors) & ";" & _
         "p.ignoreNetwork=" & MergeToJsArray(DEFAULT_IGNORE_NETWORK, m_userIgnoreNetworkPatterns) & ";" & _
         "if(!(p.invalidSelectors instanceof Set))p.invalidSelectors=new Set();else p.invalidSelectors.clear();" & _
         "return true;"

    Dim strRes As String
    strRes = CallScript(js, Nothing, contextId, False, False, 10000, _
                        Silent:=True, allowAmbiguousRetry:=True)

    Dim result As Variant
    result = GetCleanResult(strRes)
    If Not IsEmpty(result) Then PushIdleFiltersToProbe = CBool(result)
    If PushIdleFiltersToProbe Then MarkProbeFilterCurrent contextId
    Exit Function

PushFailed:
    PushIdleFiltersToProbe = False
End Function

Private Function MergeToJsArray(ByVal defaultItems As String, ByVal userItems As Collection) As String
    Dim items As New Collection
    Dim dArr() As String: dArr = Split(defaultItems, ",")
    Dim i As Long
    
    For i = LBound(dArr) To UBound(dArr)
        items.Add Trim(dArr(i))
    Next i
    
    If Not userItems Is Nothing Then
        Dim v As Variant
        For Each v In userItems
            items.Add CStr(v)
        Next v
    End If
    Dim result As String: result = "["
    For i = 1 To items.count
        result = result & JsStringLiteral(CStr(items(i)))
        If i < items.count Then result = result & ","
    Next i
    result = result & "]"
    
    MergeToJsArray = result
End Function

' 604/600/601/517 mean dispatch outcome is unknown. Promise.race does not cancel the
' losing browser task, so actions that duplicate state or overlap a survivor must not replay.
Private Function IsAmbiguousOutcome(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case 604, 601, 600, 517
            IsAmbiguousOutcome = True
    End Select
End Function

Private Function IsAmbiguousDispatchFailure(ByVal errNum As Long, _
                                            ByVal errSrc As String, _
                                            ByVal errDesc As String) As Boolean
    If IsAmbiguousOutcome(errNum) Then
        IsAmbiguousDispatchFailure = True
        Exit Function
    End If

    If errNum <> ERR_BIDI_PROTOCOL Then Exit Function

    If StrComp(errSrc, "BiDi:ExecuteBiDiCommand", vbTextCompare) = 0 And _
       InStr(1, errDesc, "Environment desync (Context Lost)", vbTextCompare) > 0 Then
        IsAmbiguousDispatchFailure = True
        Exit Function
    End If

    Dim d As String: d = LCase$(errDesc)
    If InStr(d, "inspected target navigated or closed") > 0 Or _
       InStr(d, "execution context was destroyed") > 0 Or _
       (InStr(d, "context") > 0 And _
        (InStr(d, "not found") > 0 Or InStr(d, "cannot find") > 0 Or _
         InStr(d, "destroyed") > 0 Or InStr(d, "closed") > 0)) Then
        IsAmbiguousDispatchFailure = True
    End If
End Function

' Automatic navigation retry covers selected Chromium network failures only and remains
' unsafe for one-time or side-effecting URLs.
Private Function IsRetryableNavFailure(ByVal errDesc As String) As Boolean
    Dim d As String: d = LCase$(errDesc)
    Dim codes As Variant
    codes = Array("connection_reset", "connection_closed", "connection_aborted", _
                  "connection_refused", "connection_failed", "connection_timed_out", _
                  "timed_out", "empty_response", "socket_not_connected", _
                  "address_unreachable", "network_changed")
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        If InStr(d, "net::err_" & codes(i)) > 0 Then
            IsRetryableNavFailure = True
            Exit Function
        End If
    Next i
End Function

' Non-idempotent actions may duplicate user-visible or server-visible state.
Private Function IsNonIdempotentActionKey(ByVal key As String) As Boolean
    Select Case LCase$(key)
        Case "click", "input", "input_keys", "select", "select_text", "shadow_click"
            IsNonIdempotentActionKey = True
    End Select
End Function

' A surviving lazy_scroll can reset page geometry while its retry is still measuring
' scrollHeight, causing an early stable decision and incomplete collection.
Private Function IsUnsafeToOverlapActionKey(ByVal key As String) As Boolean
    IsUnsafeToOverlapActionKey = (LCase$(key) = "lazy_scroll")
End Function

Private Function IsUnsafeToRetryAfterAmbiguousOutcome(ByVal key As String) As Boolean
    IsUnsafeToRetryAfterAmbiguousOutcome = _
        IsNonIdempotentActionKey(key) Or IsUnsafeToOverlapActionKey(key)
End Function

Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String: js = ""
    
    Dim rp As String
    rp = "if(window.__vbaIdleProbe) window.__vbaIdleProbe.lastMutationTs = performance.now(); "
    
    Select Case LCase(actionType)
        Case "click"
            js = js & "try { const e=arguments[0]; if(!e||!e.isConnected)throw new Error('STALE_ELEMENT'); "
            js = js & "e.scrollIntoView({block:'center'}); e.focus(); if(!e.isConnected)throw new Error('STALE_ELEMENT'); "
            js = js & rp & "e.click(); return true; } catch(err) { throw err; }"
                 
        Case "input"
            js = js & "try { const e=arguments[0],v=arguments[1]; if(!e||!e.isConnected)throw new Error('STALE_ELEMENT'); "
            js = js & "e.scrollIntoView({block:'center',inline:'center'}); if(!e.isConnected)throw new Error('STALE_ELEMENT'); e.click(); e.focus(); " & rp & "e.select(); "
            js = js & "const s=document.execCommand('insertText',false,v); if(!s){e.value='';e.value=v;e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));} "
            js = js & "e.blur(); return true;} catch(err){throw err;}"
                 
        Case "select"
            js = js & "try { const s=arguments[0],v=arguments[1]; if(!s||!s.isConnected)throw new Error('STALE_ELEMENT'); "
            js = js & "s.scrollIntoView({block:'center',inline:'center'}); s.focus(); if(!s.isConnected)throw new Error('STALE_ELEMENT'); " & rp & "const f=Array.from(s.options).some(o=>o.value===String(v)); if(!f)throw new Error('Option value not found: '+v); s.value=v; "
            js = js & "s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); s.blur(); return true;} catch(err){throw err;}"

        Case "select_text"
            js = js & "try { const s=arguments[0],t=arguments[1],n=(str)=>str.replace(/[\s\u00A0]+/g,' ').trim(); if(!s||!s.isConnected)throw new Error('STALE_ELEMENT'); let f=false; const target=n(t); "
            js = js & "for(let i=0;i<s.options.length;i++){if(n(s.options[i].text)===target){s.value=s.options[i].value;f=true;break;}} "
            js = js & "if(!f)throw new Error('Option text not found: '+t); s.scrollIntoView({block:'center'}); s.focus(); if(!s.isConnected)throw new Error('STALE_ELEMENT'); " & rp
            js = js & "s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); s.blur(); return true;} catch(err){throw err;}"

        Case "get_text"
            js = js & "const e=arguments[0]; if(!e||!e.isConnected)throw new Error('Element is no longer connected.'); "
            js = js & "return(e.textContent||'').replace(/[\s\u00A0]+/g,' ').trim();"

        Case "get_attribute"
            js = js & "const e=arguments[0],n=arguments[1]; if(!e||!e.isConnected)throw new Error('Element is no longer connected.'); "
            js = js & "return e.getAttribute(n);"

        Case "get_property"
            js = js & "const e=arguments[0],n=arguments[1]; if(!e||!e.isConnected)throw new Error('Element is no longer connected.'); "
            js = js & "const v=e[n]; if(v===undefined)throw new Error('Property is undefined: '+n); "
            js = js & "if(v===null||typeof v==='string'||typeof v==='boolean'||typeof v==='number')return v; "
            js = js & "throw new Error('Property is not a supported scalar value: '+n);"

        Case "get_outer_html"
            js = js & "const e=arguments[0]; if(!e||!e.isConnected)throw new Error('Element is no longer connected.'); "
            js = js & "return e.outerHTML;"

        Case "visibility_check"
            js = js & "const e=arguments[0]; if(!e||!e.isConnected)return false; const s=window.getComputedStyle(e),r=e.getBoundingClientRect(); "
            js = js & "return(s.display!=='none'&&s.visibility!=='hidden'&&parseFloat(s.opacity||'1')>0&&(r.width>0||r.height>0||e.getClientRects().length>0));"
   
        Case "shadow_click"
            js = js & "const sels=arguments[0],tout=arguments[1],start=Date.now(); let lastS=null; "
            js = js & "const fd=(r,s)=>{let n=r.querySelector(s);if(n)return n;for(let c of r.querySelectorAll('*')){if(c.shadowRoot&&(n=fd(c.shadowRoot,s)))return n;}return null;}; "
            js = js & "while(Date.now()-start<tout){ hb(); let el=document; lastS=null; "
            js = js & "for(const s of sels){el=fd(el.shadowRoot||el,s);if(!el){lastS=s;break;}} "
            js = js & "if(el){el.scrollIntoView({block:'center'});el.focus();" & rp & "el.click();return true;} "
            js = js & "if(Date.now()-start+100>tout)throw new Error('ShadowSearchError: Failed at ['+lastS+'] within '+tout+'ms'); "
            js = js & "await new Promise(r=>setTimeout(r,100));} throw new Error('ShadowSearchError: Failed at ['+lastS+'] within '+tout+'ms'); "
            
      Case "input_keys"

            js = js & "try{let e=arguments[0],v=arguments[1];if(!e||!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');"
            js = js & "e.scrollIntoView({block:'center',inline:'center'});if(!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');e.click();e.focus();"
            js = js & "if(!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');let prev=null,stb=0;for(let t=Date.now();Date.now()-t<3000;){hb();let a=document.activeElement;"
            js = js & "if(a&&a.isConnected&&a.tagName==='INPUT'&&(a.type==='text'||a.type===''||!a.type)&&typeof a.selectionStart==='number'){if(a===prev){stb++;if(stb>=2){e=a;break;}}else{prev=a;stb=1;}}else{prev=null;stb=0;}await new Promise(r=>setTimeout(r,15));}"
            js = js & "if(!e||!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');(e.select?e.select():document.execCommand('selectAll'));await new Promise(r=>setTimeout(r,80));document.execCommand('delete');await new Promise(r=>setTimeout(r,80));"
            js = js & "if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(e.value){if(e.select)e.select();document.execCommand('forwardDelete');await new Promise(r=>setTimeout(r,80));}if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');"
            js = js & "if(e.value){e.value='';e.dispatchEvent(new Event('input',{bubbles:true}));await new Promise(r=>setTimeout(r,80));}if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');" & rp & "await new Promise(r=>setTimeout(r,40));"
            js = js & "for(let c of v){hb();if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');let s=(typeof e.selectionStart==='number'?e.selectionStart:0);document.execCommand('insertText',false,c);for(let k=0;k<80;k++){let a=document.activeElement;if(a&&a.isConnected&&a!==e&&a.tagName==='INPUT'&&(a.type==='text'||!a.type))e=a;if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(document.activeElement===e&&typeof e.selectionStart==='number'&&e.selectionStart>=s+1)break;await new Promise(r=>setTimeout(r,10));}}"
            js = js & "if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(v.length>0&&(!e.value||e.value.length<v.length))throw new Error('Wiz Input Validation Failed: '+e.value);return true;}catch(err){throw err;}"

        Case "autoclicker"
            js = js & "function(){const x={{XPATH}},t={{TIMEOUT}};"
            js = js & "const f=()=>{const e=document.evaluate(x,document,null,9,null).singleNodeValue;"
            js = js & "if(e&&(e.offsetWidth>0||e.offsetHeight>0)){e.click();"
            js = js & "console.log('BiDi-AutoClicker: Target clicked');return 1}return 0};"
            js = js & "if(f())return;const o=new MutationObserver(()=>{if(f()){clearTimeout(tm);o.disconnect()}});"
            js = js & "const tm=setTimeout(()=>o.disconnect(),t);"
            js = js & "o.observe(document.body||document,{childList:1,subtree:1,attributes:1});}"

        Case "lazy_scroll"
            js = js & "try{const tout=arguments[0],poll=arguments[1],start=Date.now();"
            js = js & "const se=document.scrollingElement||document.documentElement||document.body;"
            js = js & "let lastH=se.scrollHeight,stable=0;"
            js = js & "while(Date.now()-start<tout){hb();"
            js = js & "se.scrollTo(0,se.scrollHeight);"      ' cross the lazy-load trigger line
            js = js & "await new Promise(r=>setTimeout(r,poll));"  ' let fetch -> DOM settle
            js = js & "let curH=se.scrollHeight;"
            js = js & "if(curH===lastH){stable++;if(stable>=2)break;}else{lastH=curH;stable=0;}}"
            js = js & "se.scrollTo(0,0);return true;}catch(err){throw err;}"  ' reset geometry

    End Select
    GetActionJs = js
End Function

' JS AUTHORING CONTRACT
' Long CallScript tasks must yield with await and call hb() inside loops. Never use a
' synchronous blocking loop: it blocks the browser event loop and prevents a BiDi response.
' VBA cannot forcibly terminate such browser-side JavaScript.
Public Function CallScript(ByVal js As String, _
                           Optional ByVal args As Collection = Nothing, _
                           Optional ByVal contextId As String = "", _
                           Optional ByVal awaitPromise As Boolean = False, _
                           Optional ByVal forceRealmTarget As Boolean = False, _
                           Optional ByVal safetyTimeoutMs As Long = 10000, _
                           Optional ByVal Silent As Boolean = False, _
                           Optional ByVal allowAmbiguousRetry As Boolean = False) As String
    
    Dim params As New Dictionary, targetDic As New Dictionary
    If args Is Nothing Then Set args = New Collection
    Dim resolvedCtx As String: resolvedCtx = ResolveTargetContext(contextId)
    
    If resolvedCtx = "" Then
        CallScript = "{""error"":""no_context"",""message"":""Failed to resolve contextId.""}"
        Exit Function
    End If

    Dim vSafetyTimeoutMs As Long
    vSafetyTimeoutMs = SanitizeMs(safetyTimeoutMs, "safetyTimeoutMs", 10000)
    Dim finalJs As String: finalJs = WrapInSafetyShell(js, vSafetyTimeoutMs)

    ' Strict realm contract: forceRealmTarget=True must never fall back to context, because
    ' the caller explicitly requires realm-scoped execution. BuildContextTarget is the
    ' best-effort counterpart used by infrastructure scripts such as the SPA probe.
    If forceRealmTarget Then
        Dim realmId As String
        realmId = GetScriptRealmId(resolvedCtx, False, Silent)
        If Len(realmId) = 0 Then
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:CallScript", _
                      "forceRealmTarget=True was requested, but no realm could be resolved."
        End If
        targetDic.Add "realm", realmId
    Else
        targetDic.Add "context", resolvedCtx
    End If
    
    params.Add "functionDeclaration", finalJs
    params.Add "arguments", args
    params.Add "target", targetDic
    ' CallScript always awaits the Safety Shell promise; otherwise BiDi returns a Promise
    ' object instead of the resolved envelope. Use PostScript for send-only execution.
    params.Add "awaitPromise", True
    
    Dim rawRes As String: rawRes = ExecuteBiDiCommand("script.callFunction", params, Silent, allowAmbiguousRetry, _
                                allowContextRetarget:=(resolvedCtx = m_mainContextId), _
                                commandTimeoutMs:=ResolveTransportTimeout(vSafetyTimeoutMs))
    Dim json As Object: Set json = ValidateProtocolResponse(rawRes, "script.callFunction")
    Dim procObj As Object: Set procObj = UnwrapSafetyShell(json, "script.callFunction")
    CallScript = WebJsonConverter.ConvertToJson(procObj)
End Function

Private Function ResolveTransportTimeout(ByVal operationTimeoutMs As Long) As Long
    Dim candidate As Double
    candidate = CDbl(operationTimeoutMs) + CDbl(TRANSPORT_TIMEOUT_MARGIN_MS)

    If candidate < DEFAULT_COMMAND_TIMEOUT_MS Then candidate = DEFAULT_COMMAND_TIMEOUT_MS
    If candidate > MAX_TRANSPORT_TIMEOUT_MS Then candidate = MAX_TRANSPORT_TIMEOUT_MS

    ResolveTransportTimeout = CLng(candidate)
End Function

Private Function ResolveNavigationTransportTimeout(ByVal maxTimeoutMs As Long) As Long
    Dim candidate As Long
    candidate = ResolveTransportTimeout(maxTimeoutMs)
    If candidate < NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS Then
        candidate = NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS
    End If
    ResolveNavigationTransportTimeout = candidate
End Function

Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeoutMs As Double) As Boolean
    If timeoutMs <= 0 Then Exit Function
    IsTimeout = socket_.elapsedMs(startTicks) > timeoutMs
End Function

Private Sub WaitHeartbeat(Optional ByVal sleepMs As Long = 50)
    ProcessRecorderQueue

    socket_.PumpMessages 100

    Dim vSleep As Long: vSleep = sleepMs
    If vSleep < 50 Then vSleep = 50

    socket_.YieldExecution vSleep, False
End Sub

Private Function GetProbeJs(ByVal probeType As String) As String
    Dim js As String: js = ""
    Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, m_userIgnoreSelectors)
    Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, m_userIgnoreNetworkPatterns)

    Select Case LCase$(probeType)
        Case "inject"
            ' A probe is reusable only when its version and all tracked hooks remain healthy.
            ' On repair, known wrappers are unwrapped to their saved originals; legacy wrappers
            ' are adopted once so an in-page upgrade does not stack duplicate Fetch/XHR hooks.
            js = js & "(function(){var V=" & SPA_PROBE_VERSION & ",selectors=" & selArray & ",ignoreNetwork=" & netArray & ",p=window.__vbaIdleProbe;"
            js = js & "var proto=(typeof XMLHttpRequest==='function')?XMLHttpRequest.prototype:null;"
            js = js & "var getInputDesc=()=>{try{return Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')}catch(e){return null}};"
            js = js & "var probeHealth=()=>{var d=getInputDesc();return{versionOk:!!p&&p.version===V,fetchWrapped:(typeof window.fetch!=='function')||(!!p&&!!p.fetchWrapper&&window.fetch===p.fetchWrapper),xhrWrapped:(!proto)||(!!p&&!!p.xhrOpenWrapper&&!!p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper),observerActive:!!p&&!!p.baseObserver,inputWrapped:!!d&&(!!p&&!!p.inputSetterWrapper&&d.set===p.inputSetterWr"
            js = js & "apper)}};"
            js = js & "if(p){p.selectors=selectors;p.ignoreNetwork=ignoreNetwork;var h0=probeHealth();if(h0.versionOk&&h0.fetchWrapped&&h0.xhrWrapped&&h0.observerActive&&h0.inputWrapped){window.__vbaSpaProbeInstalled=true;return true}}"
            js = js & "var legacy=!!(window.__vbaSpaProbeInstalled&&p&&(p.version===undefined||p.version===null));"
            js = js & "p=p||{inflightFetchCount:0,inflightXhrCount:0,lastMutationTs:performance.now(),changeList:[],inflightUrlCounts:new Map()};window.__vbaIdleProbe=p;window.__vbaSpaProbeInstalled=true;p.version=V;p.selectors=selectors;p.ignoreNetwork=ignoreNetwork;"
            js = js & "if(typeof p.inflightFetchCount!=='number')p.inflightFetchCount=0;if(typeof p.inflightXhrCount!=='number')p.inflightXhrCount=0;if(typeof p.lastMutationTs!=='number')p.lastMutationTs=performance.now();if(!Array.isArray(p.changeList))p.changeList=[];"
            js = js & "if(!(p.inflightUrlCounts instanceof Map)){p.inflightUrlCounts=new Map();if(p.inflightUrls&&p.inflightUrls.forEach)p.inflightUrls.forEach((u)=>p.inflightUrlCounts.set(u,1))}"
            js = js & "p.invalidSelectors=new Set();"
            js = js & "if(p.sigCount===undefined){p.sigCount=0;p.sigArmed=false;p.sigObs=null;p.visArmed=false;p.visSeen=false;p.visXPath=''}"
            js = js & "var gX=function(e){if(!e||e.nodeType!==1)return '';var s=[];while(e&&e.nodeType===1){if(e.id){s.unshift('//*[@id=\''+e.id+'\']');return s.join('/')}var i=1,b=e.previousSibling;while(b){if(b.nodeType===1&&b.tagName===e.tagName)i++;b=b.previousSibling}s.unshift(e.tagName.toLowerCase()+'['+i+']');e=e.parentNode}return '/'+s.join('/')};"
            js = js & "p.vChk=function(e){if(!e||!e.isConnected)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return(s.display!=='none'&&s.visibility!=='hidden'&&parseFloat(s.opacity||'1')>0&&(r.width>0||r.height>0||e.getClientRects().length>0))};"
            js = js & "p.armDomSignal=function(x,a){var n;try{n=document.evaluate(x,document,null,9,null).singleNodeValue}catch(e){return -2}if(!n)return -1;if(p.sigObs){try{p.sigObs.disconnect()}catch(e){}}p.sigCount=0;p.sigArmed=true;p.sigObs=new MutationObserver(function(ms){p.sigCount+=ms.length});p.sigObs.observe(n,{childList:true,characterData:true,subtree:true,attributes:!!a});return 0};"
            js = js & "p.armVisSignal=function(x){var n=null;try{n=document.evaluate(x,document,null,9,null).singleNodeValue}catch(e){return -2}if(n&&p.vChk(n))return -3;p.visXPath=x;p.visArmed=true;p.visSeen=false;return 0};"
            js = js & "p.disarmSignal=function(){if(p.sigObs){try{p.sigObs.disconnect()}catch(e){}p.sigObs=null}p.sigArmed=false;p.sigCount=0;p.visArmed=false;p.visSeen=false;p.visXPath=''};"
            js = js & "var uT=()=>{p.lastMutationTs=performance.now()},lg=(e)=>{if(e){uT();if(p.changeList.indexOf(e)===-1&&p.changeList.length<50)p.changeList.push(e)}};p.touch=uT;p.logChange=lg;"
            js = js & "var iN=(n,s)=>{if(!n)return 1;var t=(n.nodeType===1)?n:n.parentElement;if(!t)return 0;for(var i=0;i<s.length;i++){var x=s[i],k=String(x);if(p.invalidSelectors.has(k))continue;try{if(t.closest(x)!==null)return 1}catch(e){p.invalidSelectors.add(k)}}return 0};"
            js = js & "try{var d=getInputDesc();if(legacy&&!p.inputSetterWrapper&&d&&typeof d.set==='function'){p.inputSetterWrapper=d.set;p.inputGetterWrapper=d.get||null}else if(!(p.inputSetterWrapper&&d&&d.set===p.inputSetterWrapper&&!p.inputOriginalSet)){var bS=(p.inputSetterWrapper&&d&&d.set===p.inputSetterWrapper&&p.inputOriginalSet)?p.inputOriginalSet:(d&&d.set),bG=(p.inputGetterWrapper&&d&&d.get===p.inputGetterWrapper&&p.inputOrigi"
            js = js & "nalGet)?p.inputOriginalGet:(d&&d.get);if(bS&&bG){var sW=function(v){if(!iN(this,p.selectors||[]))lg('PROP:'+gX(this)+'[value]');return bS.call(this,v)},gW=function(){return bG.call(this)};Object.defineProperty(HTMLInputElement.prototype,'value',{set:sW,get:gW,configurable:true});p.inputOriginalSet=bS;p.inputOriginalGet=bG;p.inputSetterWrapper=sW;p.inputGetterWrapper=gW}}}catch(e){p.lastHealthError='input:'+String(e&&e.message||e)}"
            js = js & "try{if(p.baseObserver&&p.baseObserver.disconnect)p.baseObserver.disconnect()}catch(e){}p.baseObserver=null;"
            js = js & "try{p.baseObserver=new MutationObserver((ms)=>{ms.forEach((m)=>{if(iN(m.target,p.selectors||[]))return;var t=(m.type==='characterData')?m.target.parentElement:m.target,path=gX(t),entry=(m.type==='attributes')?'ATTR:'+path+'['+m.attributeName+']':'DOM:'+path;lg(entry)})});p.baseObserver.observe(document,{childList:true,characterData:true,subtree:true,attributes:true})}catch(e){p.baseObserver="
            js = js & "null;p.lastHealthError='observer:'+String(e&&e.message||e)}"
            js = js & "var __vbaMatch=function(u,pat){u=String(u||'').toLowerCase();pat=String(pat||'').toLowerCase();if(!pat)return false;if(pat.indexOf('*')<0)return u.indexOf(pat)!==-1;var a=pat.split('*'),pos=0;for(var i=0;i<a.length;i++){var q=a[i];if(!q)continue;var n=u.indexOf(q,pos);if(n<0)return false;pos=n+q.length}return true};"
            js = js & "if(typeof window.fetch==='function'){if(legacy&&!p.fetchWrapper&&window.fetch.__vbaWrapped){p.fetchWrapper=window.fetch}else if(!(p.fetchWrapper&&window.fetch===p.fetchWrapper&&!p.fetchOriginal)){var oF=(p.fetchWrapper&&window.fetch===p.fetchWrapper&&p.fetchOriginal)?p.fetchOriginal:window.fetch;p.fetchGen=(p.fetchGen||0)+1;var myFetchGen=p.fetchGen;var fW=function(u,o){if(myFetchGen!==p.fetchGen)return oF.apply(this,arguments);var r=(typeof u==='string')?u:(u&&u.url)||'';if((p.ignoreNetwork||[]).some((pat)=>__vbaMatch(r,pat)))return oF"
            js = js & ".apply(this,arguments);p.inflightFetchCount++;p.inflightUrlCounts.set(r,(p.inflightUrlCounts.get(r)||0)+1);var dec=()=>{p.inflightFetchCount=Math.max(0,p.inflightFetchCount-1);var n=(p.inflightUrlCounts.get(r)||0)-1;if(n<=0)p.inflightUrlCounts.delete(r);else p.inflightUrlCounts.set(r,n);uT()};try{return oF.apply(this,arguments).then((r2)=>{dec();return r2}).catch((e)=>{dec();throw e})}catch(e2){dec();throw e2}};fW.__"
            js = js & "vbaWrapped=true;fW.__vbaProbeVersion=V;fW.__vbaProbeGeneration=myFetchGen;p.fetchOriginal=oF;p.fetchWrapper=fW;window.fetch=fW}}else{p.fetchOriginal=null;p.fetchWrapper=null}"
            js = js & "if(proto){if(legacy&&!p.xhrOpenWrapper&&proto.__vbaWrapped){p.xhrOpenWrapper=proto.open;p.xhrSendWrapper=proto.send}else if(!(p.xhrOpenWrapper&&p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper&&(!p.xhrOpenOriginal||!p.xhrSendOriginal))){var oO=(p.xhrOpenWrapper&&proto.open===p.xhrOpenWrapper&&p.xhrOpenOriginal)?p.xhrOpenOriginal:proto.open,oS=(p.xhrSendWrapper&&proto.send===p.xhrSendWrap"
            js = js & "per&&p.xhrSendOriginal)?p.xhrSendOriginal:proto.send;p.xhrGen=(p.xhrGen||0)+1;var myXhrGen=p.xhrGen;var xO=function(m,u){if(myXhrGen!==p.xhrGen)return oO.apply(this,arguments);this.__vbaUrl=u;this.__vbaHooked=true;return oO.apply(this,arguments)},xS=function(){if(myXhrGen!==p.xhrGen)return oS.apply(this,arguments);var s=this,u=this.__vbaUrl||'';if((p.ignoreNetwork||[]).some((pat)=>__vbaMatch(u,pat))){s.__vbaHooked=false;return oS.apply(this,arguments)}p.inflightXhrCount++;p.inflightUrlCounts.set(u,(p.inflightUrlCounts.get(u)||0)+1);var dn=()=>{if(s.__vbaHooked){p.inflightXhrCoun"
            js = js & "t=Math.max(0,p.inflightXhrCount-1);var n=(p.inflightUrlCounts.get(u)||0)-1;if(n<=0)p.inflightUrlCounts.delete(u);else p.inflightUrlCounts.set(u,n);s.__vbaHooked=false;uT()}s.removeEventListener('loadend',dn)};this.addEventListener('loadend',dn);try{return oS.apply(this,arguments)}catch(e2){dn();throw e2}};xO.__vbaProbeGeneration=myXhrGen;xS.__vbaProbeGeneration=myXhrGen;p.xhrOpenOriginal=oO;p.xhrSendOriginal=oS;p.xhrOpenWrapper=xO;p.xhrSendWrapper=xS;proto.open=xO;proto.send=xS;pro"
            js = js & "to.__vbaWrapped=true;proto.__vbaProbeVersion=V}}"
            js = js & "var h=probeHealth();return h.versionOk&&h.fetchWrapped&&h.xhrWrapped&&h.observerActive&&h.inputWrapped})();"

        Case "query"
            ' Query reports hook-by-hook health. VBA treats any missing/mismatched field as
            ' inactive, resets the stable window, and routes through bounded probe recovery.
            js = js & "var V=" & SPA_PROBE_VERSION & ";"
            js = js & "var p=window.__vbaIdleProbe;if(!p)return JSON.stringify({active:false,version:0,health:{versionOk:false,fetchWrapped:false,xhrWrapped:false,observerActive:false,inputWrapped:false}});"
            js = js & "var proto=(typeof XMLHttpRequest==='function')?XMLHttpRequest.prototype:null,d=null;try{d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')}catch(e){}"
            js = js & "var h={versionOk:p.version===V,fetchWrapped:(typeof window.fetch!=='function')||(!!p.fetchWrapper&&window.fetch===p.fetchWrapper),xhrWrapped:(!proto)||(!!p.xhrOpenWrapper&&!!p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper),observerActive:!!p.baseObserver,inputWrapped:!!d&&(!!p.inputSetterWrapper&&d.set===p.inputSetterWrapper)};"
            js = js & "var healthy=h.versionOk&&h.fetchWrapped&&h.xhrWrapped&&h.observerActive&&h.inputWrapped;"
            js = js & "if(p.visArmed&&!p.visSeen&&p.visXPath){try{var vn=document.evaluate(p.visXPath,document,null,9,null).singleNodeValue;if(vn&&p.vChk&&p.vChk(vn))p.visSeen=true}catch(e){}}"
            js = js & "var res={active:healthy,version:p.version||0,health:h,now:performance.now(),lastMutationTs:p.lastMutationTs,inflightFetchCount:p.inflightFetchCount,inflightXhrCount:p.inflightXhrCount,changes:p.changeList,inflightUrls:Array.from((p.inflightUrlCounts||new Map()).keys()),invalidSelectors:Array.from(p.invalidSelectors||[]),sigArmed:!!p.sigArmed,sigCount:p.sigCount||0,visArmed:!!p.visArmed,visSeen:!!p.visSeen};p.changeList=[];return JSON.stringify(res);"
    End Select
    GetProbeJs = js
End Function

' Best-effort target for infrastructure scripts such as the SPA probe. Prefers a resolved
' realm and falls back to the browsing context when realm discovery is unavailable.
' Strict realm callers must use CallScript(forceRealmTarget:=True) instead.
Private Function BuildContextTarget(ByVal contextId As String, _
                                   Optional ByVal realmId As String = "") As Dictionary
    Dim target As New Dictionary
    
    If realmId <> "" Then
        target.Add "realm", realmId
    Else
        target.Add "context", contextId
    End If
    
    Set BuildContextTarget = target
End Function

Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    If Len(Trim$(cssSelectors)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:IgnoreSelectors", "cssSelectors is empty."
    End If

    If m_userIgnoreSelectors Is Nothing Then Set m_userIgnoreSelectors = New Collection
    m_userIgnoreSelectors.Add cssSelectors
    BumpIdleFilterRevision
End Property

Public Property Let IgnoreNetworkPatterns(ByVal urlPattern As String)
    If Len(Trim$(urlPattern)) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:IgnoreNetworkPatterns", "urlPattern is empty."
    End If

    If m_userIgnoreNetworkPatterns Is Nothing Then Set m_userIgnoreNetworkPatterns = New Collection
    m_userIgnoreNetworkPatterns.Add urlPattern
    BumpIdleFilterRevision
End Property

Private Function EnsureDictionary(ByVal dic As Dictionary) As Dictionary
    If dic Is Nothing Then
        Set EnsureDictionary = New Dictionary
    Else
        Set EnsureDictionary = dic
    End If
End Function

' MasterWaitEngine owns URL-contains and SPA-idle polling. m_waitStackDepth tracks
' nested waits, and every exit converges on cleanup.
Private Function MasterWaitEngine(ByVal mode As waitMode, _
                                  ByVal timeoutMs As Long, _
                                  ByVal intervalMs As Long, _
                                  ByRef ctx As Dictionary) As Boolean
    
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim vInterval As Long: vInterval = SanitizeMs(intervalMs, "intervalMs", 150, 100)
    Dim errNum As Long, errDesc As String, errSrc As String
    Dim startTick As Currency: startTick = socket_.NowTicks()
    Dim success As Boolean: success = False

    m_waitStackDepth = m_waitStackDepth + 1
    
    If m_waitStackDepth = 1 Then
        m_lastTelemetryTick = 0
        
        ' A new HUD instance invalidates the fingerprint and retry backoff from
        ' the previous outer wait. The first telemetry update must reach the new DOM.
        m_lastHudKey = ""
        m_lastHudRetryTick = 0
        
        On Error Resume Next
        Call Me.CallScript(GetHudJs("init"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), _
                           False, False, 2000, Silent:=True, allowAmbiguousRetry:=True)
        On Error GoTo 0
    End If
    On Error GoTo ErrorHandler
    Do
        ' Queue processing and Office pumping occur only at this safe wrapper wait boundary.
        WaitHeartbeat vInterval
        
        If socket_.elapsedMs(m_lastTelemetryTick) > 100 Then
            Call UpdateWaitTelemetry(mode, ctx)
            m_lastTelemetryTick = socket_.NowTicks()
        End If
        
        On Error Resume Next
        success = EvaluateWaitCondition(mode, ctx)
        If Err.Number <> 0 Then
            errNum = Err.Number: errDesc = Err.description: errSrc = Err.source
            On Error GoTo ErrorHandler: Exit Do
        End If
        On Error GoTo ErrorHandler
        
        If success Then Exit Do
        
        If IsTimeout(startTick, CDbl(vTimeout)) Then
            errNum = 504
            errDesc = "Wait condition timed out (" & vTimeout & "ms)."
            errSrc = "LogicalTimeout"
            Exit Do
        End If
    Loop

CleanExit:
    m_waitStackDepth = m_waitStackDepth - 1
    If m_waitStackDepth < 0 Then m_waitStackDepth = 0

    On Error Resume Next
    ' Keep the HUD while an outer nested wait is still active.
    If m_waitStackDepth = 0 Then
        Call Me.CallScript(GetHudJs("remove"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), _
                           False, False, 1000, Silent:=True, allowAmbiguousRetry:=True)
    End If
    
    On Error GoTo 0
    
    If errNum <> 0 Then
        ' HALT escalation applies only to logical timeout 504 with pending requests; other errors
        ' retain their original codes for recovery or best-effort handling.
        On Error Resume Next
        Dim inflightDetail As String: inflightDetail = ""
        If ctx.Exists("lastInflightUrls") Then
            Dim urls As Collection: Set urls = ctx("lastInflightUrls")
            If Not urls Is Nothing Then
                Dim urlCount As Long: urlCount = urls.count
                Dim ui As Long
                For ui = 1 To urlCount
                    If ui <= 10 Then
                        inflightDetail = inflightDetail & vbCrLf & "  " & CStr(urls(ui))
                    End If
                Next ui
                If urlCount > 10 Then
                    inflightDetail = inflightDetail & vbCrLf & "  ... and " & (urlCount - 10) & " more"
                End If
            End If
        End If
        On Error GoTo 0
        
        ' Pending counts are authoritative for SPA waits. URL details are diagnostic only because
        ' concurrent requests to the same URL may collapse in the JavaScript Set.
        Dim pendingFetch As Long
        Dim pendingXhr As Long
        If mode = WaitMode_SpaIdle Then
            If ctx.Exists("lastInflFetch") Then pendingFetch = CLng(ctx("lastInflFetch"))
            If ctx.Exists("lastInflXhr") Then pendingXhr = CLng(ctx("lastInflXhr"))
        End If
        Dim hasPendingNetwork As Boolean
        hasPendingNetwork = (pendingFetch > 0 Or pendingXhr > 0)
        
        ' Record failure outcome before raising so saved logs distinguish timeout, halt, and
        ' non-504 error.
        Dim endStatus As String
        If errNum = 504 And hasPendingNetwork Then
            endStatus = "HALT"
        ElseIf errNum = 504 Then
            endStatus = "TIMEOUT"
        Else
            endStatus = "ERROR"
        End If
        On Error Resume Next
        Call WriteSpaWaitEnd(endStatus, mode, ctx, startTick)
        ' Consume one-shot arms on success and failure, after END logging, to prevent stale gates.
        If mode = WaitMode_SpaIdle Then ConsumeSignalArms CStr(ctx("contextId"))
        Err.Clear
        On Error GoTo 0

        If errNum = 504 And hasPendingNetwork Then
            ' A 504 with pending requests becomes non-absorbable ERR_CONSENSUS_HALT.
            Debug.Print String(60, "=")
            Debug.Print "[" & Format(Now, "hh:mm:ss") & "] CONSENSUS HALT"
            Debug.Print "  Pending counts: Fetch=" & pendingFetch & "; XHR=" & pendingXhr
            If Len(inflightDetail) > 0 Then
                Debug.Print "  Pending URLs:" & inflightDetail
            Else
                Debug.Print "  Pending URLs: (details unavailable)"
            End If
            Debug.Print String(60, "=")
            
            Dim haltMsg As String
            haltMsg = "CONSENSUS HALT: SPA stability consensus was not reached." & vbCrLf & _
                      "Pending Fetch=" & pendingFetch & "; XHR=" & pendingXhr & "." & vbCrLf & _
                      "Check Immediate Window for details."
            
            Err.Raise ERR_CONSENSUS_HALT, "BiDi:MasterWaitEngine:" & errSrc, haltMsg
        Else
            ' DOM-only 504 and non-504 errors retain their codes; TrySyncPhase may absorb them.
            Err.Raise errNum, "BiDi:MasterWaitEngine:" & errSrc, errDesc
        End If
    End If
    
    If success Then
        On Error Resume Next
        Call WriteSpaWaitEnd("STABLE", mode, ctx, startTick)
        If mode = WaitMode_SpaIdle Then ConsumeSignalArms CStr(ctx("contextId"))
        On Error GoTo 0
    End If

    MasterWaitEngine = success
    Exit Function

ErrorHandler:
    errNum = Err.Number: errDesc = Err.description: errSrc = Err.source
    Resume CleanExit
End Function

' SPA-WAIT-END records outcome, elapsed time, achieved DOM quiet/slack, and network
' counters. It is the diagnostic pair of SPA-WAIT-BEGIN and runs only while recording.
Private Sub WriteSpaWaitEnd(ByVal outcome As String, _
                            ByVal mode As waitMode, _
                            ByRef ctx As Dictionary, _
                            ByVal startTick As Currency)
    If Not m_isRecording Then Exit Sub
    If mode <> WaitMode_SpaIdle Then Exit Sub
    If ctx Is Nothing Then Exit Sub
    If m_recorderLogs Is Nothing Then Exit Sub

    Dim minStableMs As Long
    If ctx.Exists("minStableMs") Then minStableMs = CLng(ctx("minStableMs"))

    Dim domQuietMs As Double
    If ctx.Exists("lastJsIdleDuration") Then domQuietMs = CDbl(ctx("lastJsIdleDuration"))

    Dim inflFetch As Long, inflXhr As Long
    If ctx.Exists("lastInflFetch") Then inflFetch = CLng(ctx("lastInflFetch"))
    If ctx.Exists("lastInflXhr") Then inflXhr = CLng(ctx("lastInflXhr"))

    Dim elapsedMs As Double: elapsedMs = socket_.elapsedMs(startTick)
    Dim slackMs As Double: slackMs = domQuietMs - CDbl(minStableMs)

    Dim relTime As String
    relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"

    ' Signal END status distinguishes miss (evidence never arrived) from hit-but-unsettled.
    ' Build it before arms are consumed.
    Dim sigEndDesc As String: sigEndDesc = BuildSignalEndDesc()

    PushTimelineLine "[" & relTime & "] [SPA-WAIT-END] " & outcome & _
                     " | elapsedMs=" & Format(elapsedMs, "0") & _
                     " | domQuietMs=" & Format(domQuietMs, "0") & _
                     " | slackMs=" & Format(slackMs, "0") & _
                     " | minStableMs=" & minStableMs & _
                     " | fetch=" & inflFetch & _
                     " | xhr=" & inflXhr & _
                     IIf(Len(sigEndDesc) > 0, " | signal=" & sigEndDesc, "")

    m_carryoverNoteArmedThisWait = False
    m_carryoverNoteEmittedThisWait = False
End Sub

Private Function EvaluateWaitCondition(ByVal mode As waitMode, ByRef ctx As Dictionary) As Boolean
    Const MAX_RECOVERY_ATTEMPTS As Long = 8

    If ctx Is Nothing Then Exit Function

    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))

    Select Case mode
        Case WaitMode_UrlContains
            EvaluateWaitCondition = (InStr(GetCurrentUrl(targetCtx, True), CStr(ctx("partialUrl"))) > 0)

        Case WaitMode_SpaIdle
            Dim probe As Dictionary: Set probe = QuerySpaProbe(targetCtx)
            Dim isProbeActive As Boolean: isProbeActive = IsSpaProbeReady(probe)

            If isProbeActive Then
                ctx("missingProbeCount") = 0

                EvaluateWaitCondition = EvaluateSpaIdleCondition(probe, ctx, socket_.nowMs())

                Set probe = Nothing
            Else
                ' When the probe disappears, clear cached inflight URLs so stale data cannot misclassify
                ' a later timeout as a network HALT.
                If ctx.Exists("lastInflightUrls") Then ctx.Remove "lastInflightUrls"

                ' navigationStarted permits predictive probe reinjection without an extra failed query;
                ' the flag is one-shot.
                If m_navigationDetected And m_navigationContextId = targetCtx Then
                    m_navigationDetected = False
                    m_navigationContextId = ""

                    On Error Resume Next
                    Call UpdateMainContextId(True, True)
                    On Error GoTo 0

                    ctx("stableStart") = socket_.nowMs()
                    ctx("missingProbeCount") = 0

                    If m_debugMode Then
                        Debug.Print "BiDi: Predictive probe injection (navigationStarted shortcut)"
                    End If
                    Call StartSpaIdleProbe(targetCtx)

                    EvaluateWaitCondition = False
                    Exit Function
                End If

                Dim bCount As Long: bCount = CLng(ctx("missingProbeCount")) + 1
                ctx("missingProbeCount") = bCount

                ctx("stableStart") = socket_.nowMs()

                ' Repeated inability to establish the probe raises logical panic 505.
                If bCount > MAX_RECOVERY_ATTEMPTS Then
                    Err.Raise 505, "BiDi:ConsensusError", _
                        "SPA Consensus Panic: Stability probe failed to initialize after " & _
                        MAX_RECOVERY_ATTEMPTS & " attempts. The page state is too volatile."
                End If

                If m_debugMode Then Debug.Print "BiDi: SPA Probe inactive. Attempting re-injection (#" & bCount & ")..."
                Call StartSpaIdleProbe(targetCtx)

                EvaluateWaitCondition = False
            End If
    End Select
End Function

' ExecuteBaseAction is the three-stage interaction funnel; retry classification, not
' procedure success alone, decides whether another attempt is safe.
Private Function ExecuteBaseAction(ByVal contextId As String, _
                                    ByVal jsKey As String, _
                                    ByVal args As Collection, _
                                    Optional ByVal waitNetwork As Boolean = True, _
                                    Optional ByVal minStable As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal awaitPromise As Boolean = True, _
                                    Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                    Optional ByVal maxTimeout As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal checkUrl As Boolean = False, _
                                    Optional ByVal retryXPath As String = "", _
                                    Optional ByVal retrySearchTimeoutMs As Long = 10000, _
                                    Optional ByVal retryCallerName As String = "") As String
    ResetLastSyncState
    Dim retryCount As Long
    Dim localErrNum As Long
    Dim localErrDesc As String
    Dim localErrSrc As String
    Dim callResult As String
    Dim status As Long

    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim vStable As Long: vStable = SanitizeMs(minStable, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vTimeout As Long: vTimeout = SanitizeMs(maxTimeout, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim currentArgs As Collection
    Set currentArgs = args

    Do While retryCount <= maxRetries
        status = TryPreparePhase(targetCtx, waitNetwork, (retryCount > 0), _
                                 localErrNum, localErrDesc, localErrSrc)
        If status = 0 Then
            status = TryExecutePhase(targetCtx, jsKey, currentArgs, awaitPromise, vTimeout, _
                                     callResult, localErrNum, localErrDesc, localErrSrc)
        End If

        If status = 0 Then
            Call TrySyncPhase(targetCtx, waitNetwork, vStable, vTimeout, checkUrl)
            ExecuteBaseAction = callResult
            Exit Function
        End If
        Dim staleFailure As Boolean
        staleFailure = IsStaleElementFailure(status, localErrSrc, localErrDesc)

        If status = ERR_ACTION_FAILED And Not staleFailure Then Exit Do

        ' Never replay an ambiguous action that could duplicate state or overlap.
        If IsUnsafeToRetryAfterAmbiguousOutcome(jsKey) And _
           IsAmbiguousDispatchFailure(status, localErrSrc, localErrDesc) Then Exit Do

        retryCount = retryCount + 1
        If retryCount <= maxRetries Then
            If m_debugMode Then
                Debug.Print "BiDi: Action '" & jsKey & "' failed (Status:" & status & "). " & _
                            "Attempting retry " & retryCount & "/" & maxRetries & "..."
            End If

            WaitHeartbeat 500

            If staleFailure And Len(retryXPath) > 0 Then
                Dim freshSharedId As String
                Dim resolverName As String

                resolverName = retryCallerName
                If Len(resolverName) = 0 Then resolverName = "ExecuteBaseAction"

                freshSharedId = ResolveXPathOrThrow( _
                    retryXPath, targetCtx, retrySearchTimeoutMs, resolverName)

                Set currentArgs = ReplaceFirstElementReference(currentArgs, freshSharedId)

                If m_debugMode Then
                    Debug.Print "BiDi: Stale element re-resolved from XPath: " & retryXPath
                End If
            End If
        End If
    Loop

    If localErrNum = 0 Then localErrNum = ERR_ACTION_FAILED
    If Len(localErrDesc) = 0 Then localErrDesc = "Unexpected execution state in ExecuteBaseAction."
    Err.Raise localErrNum, localErrSrc, localErrDesc
End Function

' INVARIANT: call only at a top-level action or session boundary. Never call from
' a path reachable during TrySyncPhase/SyncAfterAction; an inner reset would erase
' the outer action's synchronization result.
Private Sub ResetLastSyncState()
    m_lastSyncAttempted = False
    m_lastSyncSucceeded = False
    m_lastSyncErrorNumber = 0
    m_lastSyncErrorSource = vbNullString
    m_lastSyncErrorDescription = vbNullString
End Sub

Private Sub RecordSyncFailure(ByVal errorNumber As Long, _
                              ByVal errorSource As String, _
                              ByVal errorDescription As String)
    m_lastSyncSucceeded = False
    m_lastSyncErrorNumber = errorNumber
    m_lastSyncErrorSource = errorSource
    m_lastSyncErrorDescription = errorDescription
End Sub

Private Sub RunRequiredSpaSync(ByVal contextId As String, _
                               ByVal minStableMs As Long, _
                               ByVal maxTimeoutMs As Long)
    m_lastSyncAttempted = True
    On Error GoTo Failed
    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, contextId
    m_lastSyncSucceeded = True
    Exit Sub

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    RecordSyncFailure errorNumber, errorSource, errorDescription
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub SyncAfterAction(ByVal contextId As String, _
                            ByVal minStableMs As Long, _
                            ByVal maxTimeoutMs As Long, _
                            Optional ByVal checkUrl As Boolean = False)
    
    WaitHeartbeat 50

    Dim urlStable As Boolean: urlStable = True
    Dim urlErrNum As Long
    Dim urlErrSrc As String
    Dim urlErrDesc As String
    
    If checkUrl Then
        If m_debugMode Then Debug.Print ">>> [Sync] Page navigation mode active."

        urlStable = False
        On Error Resume Next
        Err.Clear
        urlStable = WaitForUrlStability(contextId, 200, 3000)
        urlErrNum = Err.Number
        urlErrSrc = Err.source
        urlErrDesc = Err.description
        Err.Clear
        UpdateMainContextId True
        If urlErrNum = 0 And Err.Number <> 0 Then
            urlErrNum = Err.Number
            urlErrSrc = Err.source
            urlErrDesc = Err.description
        End If
        On Error GoTo 0
    End If
    ProcessRecorderQueue
    
    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, contextId

    ' URL stabilization is diagnostic/best-effort for the completed action, but it must not
    ' be reported as a successful sync when the bounded URL window actually expired.
    If urlErrNum <> 0 Then Err.Raise urlErrNum, urlErrSrc, urlErrDesc
    If checkUrl And Not urlStable Then
        Err.Raise 504, "LogicalTimeout", _
            "URL and readyState did not remain stable within 3000ms."
    End If
End Sub

' Polling commands distinguish transient document/context churn from fatal protocol errors.
' Only the former may continue; malformed responses and contract errors remain visible.
Private Function TryGetPollingResponse( _
    ByVal jsonRaw As String, _
    ByVal methodName As String, _
    ByRef parsedJson As Object, _
    ByRef pendingErrorSummary As String, _
    ByRef lastLoggedErrorSummary As String) As Boolean

    Set parsedJson = Nothing

    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0

    If json Is Nothing Then
        Err.Raise vbObjectError + 512, "BiDi:" & methodName, _
                  "Response is not valid JSON."
    End If

    If json.Exists("error") Then
        Dim errType As String: errType = GetPathString(json, "error")
        Dim errMsg As String: errMsg = GetPathString(json, "message")
        Dim summary As String: summary = "[" & errType & "] " & errMsg

        If IsRetryablePollingProtocolError(methodName, errType, errMsg) Then
            pendingErrorSummary = summary
            If m_debugMode And summary <> lastLoggedErrorSummary Then
                Debug.Print "BiDi: Transient protocol error in " & methodName & _
                            "; polling continues. " & summary
            End If
            lastLoggedErrorSummary = summary
            Exit Function
        End If
    End If

    Set parsedJson = ValidateProtocolResponse(jsonRaw, methodName)
    pendingErrorSummary = vbNullString
    TryGetPollingResponse = True
End Function

Private Function IsRetryablePollingProtocolError( _
    ByVal methodName As String, _
    ByVal errType As String, _
    ByVal errMsg As String) As Boolean

    Dim normalizedMethod As String: normalizedMethod = LCase$(methodName)
    Dim normalizedType As String: normalizedType = LCase$(Trim$(errType))

    If normalizedMethod = "browsingcontext.locatenodes" Then
        If normalizedType = "no such frame" Or _
           normalizedType = "no such node" Or _
           normalizedType = "stale element reference" Then
            IsRetryablePollingProtocolError = True
            Exit Function
        End If
    ElseIf normalizedMethod = "browsingcontext.gettree" Then
        If normalizedType = "no such frame" Or normalizedType = "no such node" Then
            IsRetryablePollingProtocolError = True
            Exit Function
        End If
    Else
        Exit Function
    End If

    If normalizedType <> "unknown error" Then Exit Function

    Dim msg As String: msg = LCase$(errMsg)
    IsRetryablePollingProtocolError = _
        (InStr(msg, "inspected target navigated or closed") > 0) Or _
        (InStr(msg, "execution context was destroyed") > 0) Or _
        (InStr(msg, "document was unloaded") > 0) Or _
        (InStr(msg, "target navigated") > 0) Or _
        (InStr(msg, "node is detached from document") > 0) Or _
        (InStr(msg, "cannot find context") > 0) Or _
        ((InStr(msg, "context") > 0) And _
         ((InStr(msg, "destroyed") > 0) Or _
          (InStr(msg, "gone") > 0) Or _
          (InStr(msg, "closed") > 0) Or _
          (InStr(msg, "not found") > 0)))
End Function

Private Sub RaiseIfPollingEndedOnProtocolError( _
    ByVal methodName As String, _
    ByVal pendingErrorSummary As String)

    If Len(pendingErrorSummary) = 0 Then Exit Sub

    Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:" & methodName, _
              "Polling ended while the remote end was returning a retryable protocol error. " & _
              "Last error: " & pendingErrorSummary
End Sub

Private Function ValidateProtocolResponse(ByVal jsonRaw As String, ByVal methodName As String) As Object
    Dim json As Object

    ' Limit OERN to ParseJson; malformed JSON becomes Nothing, while later structural errors
    ' must remain visible.
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0

    If json Is Nothing Then
        Err.Raise vbObjectError + 512, "BiDi:" & methodName, "Response is not valid JSON."
    End If

    If json.Exists("error") Then
        Dim errType As String: errType = json("error")
        Dim errMsg As String: errMsg = GetPathString(json, "message")
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:" & methodName, "[" & errType & "] " & errMsg
    End If

    Dim excDetails As Object
    Set excDetails = GetPathObject(json, "result.exceptionDetails")

    If Not excDetails Is Nothing Then
        Dim jsErrDetail As String
        jsErrDetail = GetPathString(json, "result.exceptionDetails.text")

        If m_debugMode Then
            Debug.Print "--- BIDI JS ERROR REPORT | Method: " & methodName & " ---"
            Debug.Print FormatJavaScriptError(excDetails)
        End If

        Err.Raise ERR_JS_EXCEPTION, "BiDi:JSException:" & methodName, jsErrDetail
    End If

    Set ValidateProtocolResponse = json
End Function

Private Function UnwrapSafetyShell(ByVal json As Object, ByVal methodName As String) As Object
    Dim resultObj As New Dictionary ' Container for normalized output

    Dim rawVal As Variant
    Assign rawVal, GetPathValue(json, "result.result.value")

    Dim isWrapped As Boolean: isWrapped = False

    If VarType(rawVal) = vbString Then
        Dim sVal As String: sVal = Trim(CStr(rawVal))
        If Left(sVal, 1) = "{" And Right(sVal, 1) = "}" Then
            Dim innerJson As Object
            On Error Resume Next
            Set innerJson = WebJsonConverter.ParseJson(sVal)
            On Error GoTo 0

            If Not innerJson Is Nothing Then
                If innerJson.Exists("status") Then
                    isWrapped = True
                    If innerJson("status") = "error" Then
                        Dim shellMsg As String: shellMsg = CStr(innerJson("message"))

                        ' Map the Safety Shell VBA_TIMEOUT token to error 517.
                        If shellMsg = "VBA_TIMEOUT" Then
                            Err.Raise 517, "BiDi:SafetyTimeout:" & methodName, _
                                "Safety Shell timeout: the wrapper stopped waiting, but the browser-side task may still be running."
                        Else
                            Err.Raise ERR_ACTION_FAILED, "BiDi:JSError:" & methodName, shellMsg
                        End If
                    End If

                    Call AssignValue(resultObj, "value", innerJson("value"))
                End If
            End If
        End If
    End If

    If Not isWrapped Then
        If VarType(rawVal) = vbEmpty Then
            If InStr(1, methodName, "script.", vbTextCompare) > 0 Then
                Dim resultType As String
                resultType = GetPathString(json, "result.result.type")

                Select Case resultType
                    Case "undefined", "null"
                    Case ""
                        Dim hasScriptId As String: hasScriptId = GetPathString(json, "result.script")
                        Dim hasRealms As String:   hasRealms = GetPathString(json, "result.realms.0.realm")

                        If hasScriptId = "" And hasRealms = "" Then
                            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                                "Structural failure: No recognizable result or metadata in script response."
                        End If
                    Case Else
                        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                            "Protocol violation: Result type '" & resultType & "' reported without value."
                End Select
            End If
        End If

        Call AssignValue(resultObj, "value", rawVal)
    End If

    resultObj("status") = "ok"
    Set UnwrapSafetyShell = resultObj
End Function

Private Sub AssignValue(ByRef dic As Dictionary, ByVal key As String, ByRef val As Variant)
    If IsObject(val) Then
        Set dic(key) = val
    Else
        dic(key) = val
    End If
End Sub

Private Function ResolveXPathOrThrow(ByVal xpath As String, _
                                     ByVal contextId As String, _
                                     ByVal timeoutMs As Long, _
                                     ByVal methodName As String) As String
    Dim sid As String
    Dim targetCtx As String
    
    If InStr(xpath, Chr$(0)) > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "XPath contains null byte (Chr$(0)) which would corrupt the WebSocket message. Sanitize input before calling."
    End If
    Dim vTimeout As Long
    If timeoutMs = 0 Then
        vTimeout = 0
    Else
        vTimeout = SanitizeMs(timeoutMs, "searchTimeoutMs", 10000)
    End If
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId()
    End If
    sid = GetFindElementSharedIdByXPath(xpath, vTimeout, targetCtx)
    
    If sid = "" Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "Element not found within " & IIf(vTimeout = 0, "0 (immediate)", vTimeout) & "ms for XPath: " & xpath
    End If
    
    ResolveXPathOrThrow = sid
End Function

Private Sub PrepareProbeForAction(ByVal contextId As String)
    Dim probe As Dictionary
    Dim isReady As Boolean

    Set probe = QuerySpaProbe(contextId)
    isReady = IsSpaProbeReady(probe)

    If Not isReady Then
        StartSpaIdleProbe contextId
        WaitHeartbeat 150
    ElseIf Not IsProbeFilterCurrent(contextId) Then
        ' A healthy probe can accept new arrays in place. If the lightweight push fails,
        ' use the existing bounded reinjection path, which also refreshes the arrays.
        If Not PushIdleFiltersToProbe(contextId) Then
            StartSpaIdleProbe contextId
            WaitHeartbeat 150
        End If
    End If
End Sub

' SanitizeMs substitutes defaults for non-positive values, applies the minimum polling
' floor, and caps prolonged waits. Exact-zero APIs branch before calling it.
Private Function SanitizeMs(ByVal ms As Long, ByVal argName As String, _
                             ByVal defaultMs As Long, _
                             Optional ByVal minMs As Long = 50) As Long
    Dim vMs As Long: vMs = ms
    
    If vMs <= 0 Then
        If m_debugMode Then
            Debug.Print "BiDi Warning: Invalid value (" & ms & ") for '" & argName & "'. Using default: " & defaultMs
        End If
        vMs = defaultMs
    End If
    If vMs < minMs Then vMs = minMs

    If vMs > 600000 Then vMs = 600000

    SanitizeMs = vMs
End Function

Public Function WaitForUrlStability(Optional ByVal contextId As String = "", _
                                    Optional ByVal stableDurationMs As Long = 500, _
                                    Optional ByVal timeoutMs As Long = 10000) As Boolean

    If stableDurationMs <= 0 Then stableDurationMs = 500
    If timeoutMs <= 0 Then timeoutMs = 10000

    Dim vStableMs As Long: vStableMs = SanitizeMs(stableDurationMs, "stableDurationMs", 500)
    Dim vTimeoutMs As Long: vTimeoutMs = SanitizeMs(timeoutMs, "timeoutMs", 10000)
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim startMs As Double: startMs = socket_.nowMs()
    Dim stableStartMs As Double: stableStartMs = socket_.nowMs()
    Dim lastStateKey As String: lastStateKey = ""
    Dim stabilized As Boolean: stabilized = False
    Dim checkJs As String: checkJs = "return window.location.href + '|' + document.readyState;"
    
    Do
        Dim currentResult As String
        currentResult = CStr(GetCleanResult(CallScript(checkJs, Nothing, targetCtx, False, False, vTimeoutMs, _
                                                       Silent:=True, allowAmbiguousRetry:=True)))
        
        Dim isMoving As Boolean
        isMoving = (currentResult <> lastStateKey) Or (Right$(currentResult, 8) <> "complete")
        
        If isMoving Then
            lastStateKey = currentResult
            stableStartMs = socket_.nowMs()
        Else
            If (socket_.nowMs() - stableStartMs) >= CDbl(vStableMs) Then
                stabilized = True
                Exit Do
            End If
        End If
        
        If (socket_.nowMs() - startMs) > CDbl(vTimeoutMs) Then Exit Do
        
        If m_isRecording Then
            WaitHeartbeat 50
        Else
            socket_.YieldExecution 50, True
        End If
    Loop

    WaitForUrlStability = stabilized
End Function

' The Safety Shell bounds how long VBA waits, not how long the browser task runs.
' Promise.race does not cancel the losing task.
Private Function WrapInSafetyShell(ByVal jsBody As String, ByVal timeoutMs As Long) As String
    Dim js As String
    
    js = "async function(){ "
    js = js & "const _t=" & timeoutMs & "; let _la=performance.now(); const hb=()=>{_la=performance.now()}; "
    
    ' _tk must remain an arrow function so jsBody arguments resolve to the outer BiDi
    ' function parameters; a regular function silently breaks all action arguments.
    js = js & "const _tk = async () => { " & jsBody & " }; "
    
    ' The watchdog uses the monotonic browser clock and rejects when hb() silence exceeds its limit.
    js = js & "let _wdTimer; const _wd = new Promise((_,r)=>{ _wdTimer=setInterval(()=>{if(performance.now()-_la>_t){clearInterval(_wdTimer);r(new Error('VBA_TIMEOUT'))}},100);}); "
    
    js = js & "try{ const res = await Promise.race([_tk(), _wd]); clearInterval(_wdTimer); return JSON.stringify({status:'ok', value:(typeof res === 'undefined' ? null : res)}); "
    js = js & "}catch(e){ clearInterval(_wdTimer); return JSON.stringify({status:'error', message:e.message}); } }"
    
    WrapInSafetyShell = js
End Function

Private Sub HandleConsoleLog(ByVal json As Object)
    ' Do not add blanket OERN here; missing paths are already handled and other errors matter.
    Dim text As String: text = GetPathString(json, "params.text")
    
    If InStr(text, "BiDi-AutoClicker") > 0 Then
        Debug.Print ">>> [BROWSER] " & text
    End If
End Sub

Public Sub Shutdown(Optional ByVal ForceClose As Boolean = True)
    ' Explicit callers and Class_Terminate share this path. Once cleanup completes, a later
    ' destructor call must not repeat session teardown or emit a second completion line.
    If m_shutdownCompleted Then Exit Sub
    If socket_ Is Nothing Then
        m_shutdownCompleted = True
        Exit Sub
    End If

    ' Shutdown is best-effort from this point; cleanup must continue even after detach or
    ' close failure.
    On Error Resume Next

    ' Download behavior is browser-side state and may survive wrapper/WebSocket teardown. Clear
    ' only when this wrapper attempted to change it; unrelated browser defaults remain untouched.
    If m_downloadBehaviorCleanupRequired Then
        Call ClearDownloadBehaviorBestEffort(True)
    End If

    ' File-chooser interception is browser-side state and must be explicitly released BEFORE
    ' detaching or force-closing the transport. Keep a loud warning if confirmation fails.
    If m_fileChooserInterceptMayBeActive Then
        If Not ReleaseTrackedFileChooserInterceptionBestEffort(True) Then
            If m_debugMode Then
                Debug.Print "BiDi CRITICAL: Shutdown could not confirm native file chooser " & _
                            "interception cleanup. Reconnect before further chooser use."
            End If
        End If
    End If
    
    ' ForceClose skips acknowledged CDP detach because teardown should not spend a receive
    ' timeout on a suspended browser. The tunnel dies with the BiDi socket. Graceful callers
    ' may use Shutdown ForceClose:=False.
    If ForceClose Then
        m_cdpSessionId = ""
    ElseIf m_cdpSessionId <> "" Then
        Dim params As New Dictionary
        params.Add "session", m_cdpSessionId
        ExecuteBiDiCommand "goog:cdp.detachSession", params, True
        m_cdpSessionId = ""
    End If
    
    ' Order matters: drop flags before releasing collections, otherwise a guarded append
    ' path can observe m_isRecording=True with m_recorderLogs already Nothing.
    ResetSessionBoundaryState

    ClearState

    If ForceClose Then
        socket_.ForceClose
    Else
        ' Graceful shutdown must explicitly end the transport even when another object
        ' still holds a reference to the communicator.
        socket_.Disconnect
    End If

    ' Retain the disconnected communicator so this wrapper remains reusable, as ConnectTo's
    ' own error message promises. ForceClose and Disconnect are both idempotent.
    m_shutdownCompleted = True
    
    If m_debugMode Then Debug.Print "BiDi: Shutdown sequence completed."
End Sub

Private Sub UpdateWaitTelemetry(ByVal mode As waitMode, ByRef ctx As Dictionary)
    On Error Resume Next
    
    Dim msg As String, colorCode As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))
    
    Select Case mode
        Case WaitMode_UrlContains
            msg = "BiDi: Waiting for URL..." & vbLf & "Pattern: " & ctx("partialUrl")
            colorCode = "#00FF00" ' Green
            
        Case WaitMode_SpaIdle
            Dim idleTime As Double: idleTime = (socket_.nowMs() - CDbl(ctx("stableStart"))) / 1000
            Dim targetIdle As Double: targetIdle = CLng(ctx("minStableMs")) / 1000

            ' HUD pending types may lag DOM/visibility by one poll because telemetry precedes probe
            ' evaluation; network latches during the preceding event drain.
            Dim sigPending As String: sigPending = BuildSignalPendingDesc()

            ' A network hit is latched before its stableStart reset is acknowledged on the next poll.
            ' Treat that gap as synchronizing to prevent a false green Stable flash.
            Dim netResetPending As Boolean
            netResetPending = m_sigNetArmed And m_sigNetSeen
            If netResetPending And ctx.Exists("sigNetAcked") Then netResetPending = Not CBool(ctx("sigNetAcked"))

            If idleTime < targetIdle Then
                colorCode = "#FF4444" ' Red (Busy: DOM/network still churning)
                msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Synchronizing..."
            ElseIf Len(sigPending) > 0 Then
                ' DOM quiet with a closed signal gate is amber Awaiting signal, not Stable.
                colorCode = "#FFAA00" ' Amber (Gated: awaiting signal)
                msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Awaiting signal (" & sigPending & ")"
            ElseIf netResetPending Then
                colorCode = "#FF4444" ' Red (post-hit reset imminent)
                msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Synchronizing..."
            Else
                colorCode = "#00FF00" ' Green (Stable: gate open or ungated)
                msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Stable"
            End If
    End Select
    
    ' Fingerprint unchanged HUD state to suppress redundant socket traffic.
    Dim currentHudKey As String: currentHudKey = msg & "|" & colorCode
    
    If targetCtx <> m_lastHudCtx Then
        m_lastHudKey = ""
        m_lastHudCtx = targetCtx
        m_lastHudRetryTick = 0
    End If
    
    If currentHudKey = m_lastHudKey Then Exit Sub
    
    If m_lastHudRetryTick > 0 Then
        If socket_.elapsedMs(m_lastHudRetryTick) < 500 Then Exit Sub
    End If
    
    Dim safeMsg As String: safeMsg = Replace(Replace(msg, "\", "\\"), "'", "\'")
    safeMsg = Replace(Replace(safeMsg, vbCr, ""), vbLf, "\n")
    Dim safeColor As String: safeColor = Replace(colorCode, "'", "\'")
    Dim js As String: js = ""
    js = js & "(function(){ var b=document.body; if(!b) return; var MSG='" & safeMsg & "', CLR='" & safeColor & "';"
    js = js & " var h=document.getElementById('vba-bidi-hud'); if(!h){ h=document.createElement('div'); h.id='vba-bidi-hud';"
    js = js & " h.style.cssText='position:fixed;top:15px;right:15px;z-index:2147483647;padding:10px;font-family:monospace;font-size:12px;background:rgba(0,0,0,0.85);color:#00FF00;border:1px solid #00FF00;border-radius:6px;pointer-events:none;box-shadow:0 4px 15px rgba(0,0,0,0.5);min-width:180px;white-space:pre-wrap;'; b.appendChild(h);}"
    js = js & " h.innerText=MSG; h.style.borderColor=CLR; h.style.color=CLR; return true;})()"
    
    Call Me.CallScript(js, Nothing, targetCtx, False, False, 1000, _
                       Silent:=True, allowAmbiguousRetry:=True)
    
    If Err.Number = 0 Then
        m_lastHudKey = currentHudKey
        m_lastHudRetryTick = 0
    Else
        m_lastHudRetryTick = socket_.NowTicks()
    End If
    
    On Error GoTo 0
End Sub

Private Function GetHudJs(ByVal snippetType As String) As String
    Dim js As String: js = ""
    Select Case LCase(snippetType)
        Case "init"
            js = js & "(function(){if(document.getElementById('vba-bidi-hud')) return;"
            js = js & "const h=document.createElement('div'); h.id='vba-bidi-hud';"
            js = js & "h.style.cssText='position:fixed;top:15px;right:15px;z-index:2147483647;padding:10px;font-family:monospace;font-size:12px;background:rgba(0,0,0,0.85);color:#00FF00;border:1px solid #00FF00;border-radius:6px;pointer-events:none;box-shadow:0 4px 15px rgba(0,0,0,0.5);min-width:180px;white-space:pre-wrap;';"
            js = js & "h.innerText='BiDi: Initializing...'; document.body.appendChild(h);})();"
            GetHudJs = js
                       
        Case "remove"
            GetHudJs = "(function(){const h=document.getElementById('vba-bidi-hud'); if(h)h.remove();})();"
    End Select
End Function

' Returns the registered preload script ID. Raises on protocol or malformed-response
' failure; failure is no longer represented by an empty string.
Public Function ExecuteEnableGlobalShadowUnlocker() As String
    If m_shadowUnlockerRegistered And Len(m_shadowUnlockerScriptId) > 0 Then
        ExecuteEnableGlobalShadowUnlocker = m_shadowUnlockerScriptId
        If m_debugMode Then
            Debug.Print "BiDi: Global Shadow DOM Unlocker already registered (ID: " & _
                        m_shadowUnlockerScriptId & ")"
        End If
        Exit Function
    End If

    If m_shadowUnlockerRegistered Then
        m_shadowUnlockerRegistered = False
        m_shadowUnlockerScriptId = vbNullString
    End If
    Dim params As New Dictionary, js As String
    
    ' The attachShadow hook forces mode:'open' while preserving other options; dropping
    ' slotAssignment or delegatesFocus can break components.
    js = "function() { "
    js = js & "if (Element.prototype._attachShadow) return; "
    js = js & "Element.prototype._attachShadow = Element.prototype.attachShadow; "
    js = js & "Element.prototype.attachShadow = function(options) { "
    js = js & "  return this._attachShadow(Object.assign({}, options || {}, { mode: 'open' })); "
    js = js & "}; }"
    
    params.Add "functionDeclaration", js

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True, _
                               allowAmbiguousRetry:=False)

    Dim json As Object
    Set json = ValidateProtocolResponse(strRes, "script.addPreloadScript")

    Dim scriptId As String
    scriptId = GetPathString(json, "result.script")
    If Len(scriptId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:script.addPreloadScript", _
            "script.addPreloadScript succeeded without returning result.script."
    End If

    ' Commit capability state only after full response validation. A later failed duplicate
    ' registration must not erase a preload script that is already active in this session.
    m_shadowUnlockerRegistered = True
    m_shadowUnlockerScriptId = scriptId
    ExecuteEnableGlobalShadowUnlocker = scriptId
    
    If m_debugMode Then
        Debug.Print "BiDi: Global Shadow DOM Unlocker registered (ID: " & scriptId & ")"
    End If
End Function

Private Function TryEnableGlobalShadowUnlocker( _
    ByRef failureNumber As Long, _
    ByRef failureSource As String, _
    ByRef failureDescription As String) As Boolean
    failureNumber = 0
    failureSource = vbNullString
    failureDescription = vbNullString

    On Error GoTo Failed
    Call ExecuteEnableGlobalShadowUnlocker
    TryEnableGlobalShadowUnlocker = True
    Exit Function

Failed:
    failureNumber = Err.Number
    failureSource = Err.source
    failureDescription = Err.description
    Err.Clear
End Function

Private Function TryPreparePhase(ByVal ctx As String, _
                                  ByVal waitNet As Boolean, _
                                  ByVal isRetry As Boolean, _
                                  ByRef outErrNum As Long, _
                                  ByRef outErrDesc As String, _
                                  ByRef outErrSrc As String) As Long
    Dim capturedErr As Long
    
    If waitNet Then
        On Error Resume Next
        Call PrepareProbeForAction(ctx)
        capturedErr = Err.Number
        If capturedErr <> 0 Then
            outErrNum = capturedErr: outErrDesc = Err.description: outErrSrc = Err.source
            On Error GoTo 0
            TryPreparePhase = capturedErr: Exit Function
        End If
        On Error GoTo 0
    End If
    
    If isRetry Then
        On Error Resume Next
        Call UpdateMainContextId(True, True) ' forceRefresh:=True, Silent:=True
        capturedErr = Err.Number
        If capturedErr <> 0 Then
            outErrNum = capturedErr: outErrDesc = Err.description: outErrSrc = Err.source
            On Error GoTo 0
            TryPreparePhase = capturedErr: Exit Function
        End If
        On Error GoTo 0
    End If
    
    TryPreparePhase = 0
End Function

' Trusted pointer clicks are opt-in. Synthetic element.click() remains the default proven
' path for Wiz/Lit SPAs.
Public Property Get UseTrustedClick() As Boolean
    UseTrustedClick = m_useTrustedClick
End Property
Public Property Let UseTrustedClick(ByVal value As Boolean)
    m_useTrustedClick = value
End Property

' DebugMode controls wrapper trace only; it does not affect Discovery Log or communicator
' wire tracing.
Public Property Get DebugMode() As Boolean
    DebugMode = m_debugMode
End Property
Public Property Let DebugMode(ByVal value As Boolean)
    m_debugMode = value
End Property

Public Property Get ShadowUnlockerRegistered() As Boolean
    ShadowUnlockerRegistered = m_shadowUnlockerRegistered
End Property
Public Property Get ShadowUnlockerScriptId() As String
    ShadowUnlockerScriptId = m_shadowUnlockerScriptId
End Property
Public Property Get LastSyncAttempted() As Boolean
    LastSyncAttempted = m_lastSyncAttempted
End Property
Public Property Get LastSyncSucceeded() As Boolean
    LastSyncSucceeded = m_lastSyncSucceeded
End Property
Public Property Get LastSyncErrorNumber() As Long
    LastSyncErrorNumber = m_lastSyncErrorNumber
End Property
Public Property Get LastSyncErrorSource() As String
    LastSyncErrorSource = m_lastSyncErrorSource
End Property
Public Property Get LastSyncErrorDescription() As String
    LastSyncErrorDescription = m_lastSyncErrorDescription
End Property

Private Function TryExecutePhase(ByVal ctx As String, _
                                  ByVal key As String, _
                                  ByVal args As Collection, _
                                  ByVal await As Boolean, _
                                  ByVal TIMEOUT As Long, _
                                  ByRef outResult As String, _
                                  ByRef outErrNum As Long, _
                                  ByRef outErrDesc As String, _
                                  ByRef outErrSrc As String) As Long
    On Error Resume Next
    
    ' Trusted clicks can leave focus state that breaks some autocomplete flows, so use them
    ' only for sites that reject synthetic clicks.
    If key = "click" And m_useTrustedClick Then
        outResult = PerformTrustedClick(ctx, args, TIMEOUT)
    Else
        outResult = CallScript(GetActionJs(key), args, ctx, await, , TIMEOUT, allowAmbiguousRetry:=Not IsUnsafeToRetryAfterAmbiguousOutcome(key))
    End If
    
    TryExecutePhase = Err.Number
    If Err.Number <> 0 Then
        outErrNum = Err.Number
        outErrDesc = Err.description
        outErrSrc = Err.source
    End If
    On Error GoTo 0
End Function

' Post-action synchronization is best-effort after action success, except
' ERR_CONSENSUS_HALT which must stop the caller.
Private Sub TrySyncPhase(ByVal ctx As String, _
                         ByVal waitNet As Boolean, _
                         ByVal stable As Long, _
                         ByVal TIMEOUT As Long, _
                         ByVal chkUrl As Boolean)
    Dim syncErrNum As Long
    Dim syncErrDesc As String
    Dim syncErrSrc As String

    If waitNet Then
        m_lastSyncAttempted = True

        On Error Resume Next
        Err.Clear
        Call SyncAfterAction(ctx, stable, TIMEOUT, chkUrl)
        syncErrNum = Err.Number
        syncErrDesc = Err.description
        syncErrSrc = Err.source
        On Error GoTo 0

        m_lastSyncSucceeded = (syncErrNum = 0)
        If syncErrNum <> 0 Then RecordSyncFailure syncErrNum, syncErrSrc, syncErrDesc
    End If
    
    ' Never absorb ERR_CONSENSUS_HALT.
    If syncErrNum = ERR_CONSENSUS_HALT Then
        Err.Raise syncErrNum, syncErrSrc, syncErrDesc
    End If
    
    ' Ordinary synchronization failure does not undo the completed interaction.
    If syncErrNum <> 0 And m_debugMode Then
        Debug.Print "[" & Format(Now, "hh:mm:ss") & "] BiDi Warning: Post-action Sync failed (Status: " & syncErrNum & "). DOM mutations only (no pending requests). Consider extending idle timeout."
    End If
End Sub

' PostScript is disposable send-only infrastructure; completion must not be inferred.
Public Sub PostScript(ByVal js As String, _
                      Optional ByVal args As Collection = Nothing, _
                      Optional ByVal contextId As String = "")
    
    Dim params As New Dictionary, targetDic As New Dictionary
    Dim resolvedCtx As String: resolvedCtx = ResolveTargetContext(contextId)
    
    If resolvedCtx = "" Then Exit Sub

    Dim finalJs As String: finalJs = WrapInSafetyShell(js, 1000)

    targetDic.Add "context", resolvedCtx
    params.Add "functionDeclaration", finalJs

    Dim safeArgs As Collection
    If args Is Nothing Then
        Set safeArgs = New Collection
    Else
        Set safeArgs = args
    End If
    params.Add "arguments", safeArgs

    params.Add "target", targetDic
    params.Add "awaitPromise", True

    ' PostScript id 2 belongs to the communicator's reserved band; normal IDs cannot collide
    ' and its response is consumed internally.
    Dim requestDic As New Dictionary
    requestDic.Add "id", POST_SCRIPT_ID
    requestDic.Add "method", "script.callFunction"
    requestDic.Add "params", params
    
    ' This path bypasses response correlation but does not pump Office messages.
    socket_.PostScript WebJsonConverter.ConvertToJson(requestDic)
End Sub

' Converts arbitrary text to an XPath 1.0 string literal. The returned value is an XPath
' expression such as 'text', "text", or concat(...), not the original unescaped text.
Public Function XPathLiteral(ByVal text As String) As String
    XPathLiteral = EscapeXPathString(text)
End Function

' Escape XPath literals with single quotes, double quotes, or concat() when both quote
' types occur.
Private Function EscapeXPathString(ByVal text As String) As String
    If InStr(text, "'") = 0 Then
        EscapeXPathString = "'" & text & "'"
    ElseIf InStr(text, """") = 0 Then
        EscapeXPathString = """" & text & """"
    Else
        Dim parts() As String: parts = Split(text, "'")
        Dim result As String: result = "concat("
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If i > LBound(parts) Then result = result & ",""'"","
            result = result & "'" & parts(i) & "'"
        Next i
        result = result & ")"
        EscapeXPathString = result
    End If
End Function

Private Sub Assign(ByRef dest As Variant, ByRef src As Variant)
    If IsObject(src) Then Set dest = src Else dest = src
End Sub

' GetPathValue traverses dot paths through Dictionaries and Collections; path indexes are
' zero-based while VBA Collection indexes are one-based.
Private Function InternalTraverseFromObject(ByVal root As Object, ByVal path As String) As Variant
    If root Is Nothing Or path = "" Then Exit Function
    
    Dim keys() As String: keys = Split(path, ".")
    Dim current As Object: Set current = root
    Dim i As Long, idx As Long
    
    For i = 0 To UBound(keys) - 1
        If TypeName(current) = "Dictionary" Then
            If Not current.Exists(keys(i)) Then Exit Function
            If IsObject(current(keys(i))) Then Set current = current(keys(i)) Else Exit Function
        ElseIf TypeName(current) = "Collection" And IsNumeric(keys(i)) Then
            idx = CLng(keys(i)) + 1
            If idx < 1 Or idx > current.count Then Exit Function
            If IsObject(current(idx)) Then Set current = current(idx) Else Exit Function
        Else: Exit Function
        End If
        If current Is Nothing Then Exit Function
    Next i
    
    Dim lastKey As String: lastKey = keys(UBound(keys))
    If TypeName(current) = "Dictionary" Then
        If current.Exists(lastKey) Then
            If IsObject(current(lastKey)) Then Set InternalTraverseFromObject = current(lastKey) _
            Else InternalTraverseFromObject = current(lastKey)
        End If
    ElseIf TypeName(current) = "Collection" And IsNumeric(lastKey) Then
        idx = CLng(lastKey) + 1
        If idx >= 1 And idx <= current.count Then
            If IsObject(current(idx)) Then Set InternalTraverseFromObject = current(idx) _
            Else InternalTraverseFromObject = current(idx)
        End If
    End If
End Function

Private Function GetPathValue(ByVal root As Object, ByVal path As String) As Variant
    Assign GetPathValue, InternalTraverseFromObject(root, path)
End Function

Private Function GetPathString(ByVal root As Object, ByVal path As String) As String
    Dim v As Variant: Assign v, GetPathValue(root, path)
    If Not IsObject(v) Then GetPathString = CStr(v & "")
End Function

Private Function GetPathObject(ByVal root As Object, ByVal path As String) As Object
    ' Verify Object type explicitly before Set; do not hide type errors with OERN.
    Dim v As Variant: Assign v, GetPathValue(root, path)
    If IsObject(v) Then Set GetPathObject = v
End Function

Private Function GetCleanResult(ByVal jsonRaw As String) As Variant
    ' Only ParseJson is tolerant of malformed input; path traversal and assignment errors
    ' indicate structural bugs and must propagate.
    If Len(jsonRaw) = 0 Then GetCleanResult = Empty: Exit Function
    
    Dim parsed As Object
    On Error Resume Next
    Set parsed = WebJsonConverter.ParseJson(jsonRaw)
    Dim parseErr As Long: parseErr = Err.Number
    On Error GoTo 0
    
    If parseErr <> 0 Or parsed Is Nothing Then GetCleanResult = Empty: Exit Function
    
    Dim v As Variant: Assign v, GetPathValue(parsed, "value")
    If IsObject(v) Then
        If v Is Nothing Then GetCleanResult = Empty Else Set GetCleanResult = v
    Else: GetCleanResult = v
    End If
End Function

Public Sub ExecuteWizSelectDropdown(ByVal triggerXPath As String, _
                                    ByVal optionText As String, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "")
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Call Me.ExecuteClickByXPath(triggerXPath, searchTimeoutMs, False, , maxTimeoutMs, targetCtx)
    
    Dim optionXPath As String
    optionXPath = "//*[@role='option'][contains(., " & EscapeXPathString(optionText) & ")]"
    
    Call Me.ExecuteClickByXPath(optionXPath, searchTimeoutMs, True, , maxTimeoutMs, targetCtx)
End Sub
