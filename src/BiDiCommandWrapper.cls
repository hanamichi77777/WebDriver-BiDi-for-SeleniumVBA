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
Private Const BIDI_WRAPPER_VERSION As String = "4.4"
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
'   - SendAndReceive pumps Office messages only when AllowUiPump is explicitly enabled (default False).
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
' The generation guard also lets a later inert preload contain an addPreloadScript whose
' response was lost, so an orphaned unlocker cannot affect future realms indefinitely.
Private m_shadowUnlockerRegistered As Boolean
Private m_shadowUnlockerScriptId As String
Private m_shadowUnlockerSuppressorScriptId As String
Private m_shadowUnlockerContainmentRisk As Boolean
Private m_lastShadowUnlockerGuardGeneration As Double

' AutoClicker preload scripts are session-scoped browser state. Track them by canonical locator key so
' duplicate XPath or Shadow registrations cannot create multiple observers/clicks in future realms.
Private m_autoClickerScriptIds As Dictionary
Private m_autoClickerTimeouts As Dictionary
' A7: every AutoClicker preload participates in a per-realm generation guard keyed by locator.
' A later generation supersedes older/unknown duplicate preload scripts without allowing duplicate clicks.
Private m_autoClickerGenerations As Dictionary
' A suppressor is an inert higher-generation preload intentionally left session-scoped after unregister.
' It prevents an older script whose addPreloadScript response was lost from becoming active again.
Private m_autoClickerSuppressorScriptIds As Dictionary
' XPath entries here mean an ambiguous registration could not yet be proven contained.
Private m_autoClickerContainmentRisks As Dictionary
Private m_lastAutoClickerGuardGeneration As Double

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

' A6: one navigation transaction may correlate its browsingContext.navigate result with the
' main-document network.responseStarted event carrying the same BiDi navigation id. This is
' the only source used for HTTP status; browsingContext.navigate itself has no HTTP status.
Private m_navHttpCaptureInProgress As Boolean
Private m_navHttpCaptureArmed As Boolean
Private m_navHttpCaptureOwnerContextId As String
Private m_navHttpCaptureExpectedNavigationId As String
Private m_navHttpCaptureExpectedUrl As String
Private m_navHttpCaptureStatusKnown As Boolean
Private m_navHttpCaptureStatus As Long
Private m_navHttpCaptureStatusText As String
Private m_navHttpCaptureResponseUrl As String
Private m_navHttpCaptureRedirectCount As Long
Private m_navHttpCaptureLossBaseline As Long

' N7: per-attempt observation boundary for a navigate command that Chromium reports as
' superseded by another navigation. This state never authorizes a retry of the original URL;
' it only proves that a later navigationStarted event occurred in the same target context.
Private m_navSupersessionWatchArmed As Boolean
Private m_navSupersessionOwnerContextId As String
Private m_navSupersessionObservedCount As Long
Private m_navSupersessionNavigationId As String
Private m_navSupersessionUrl As String

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
' Context destruction removes event/context keys from the active delta registry, but the server-
' issued subscription ID remains ours until session.unsubscribe confirms release.
Private m_retiredSubscriptionIds As Dictionary
' If transport teardown prevents a confirmed session.unsubscribe, retain the issued IDs plus
' the BiDi session identity. A same-session reconnect must discharge them before subscribing again.
Private m_pendingSubscriptionCleanupIds As Dictionary
Private m_pendingSubscriptionCleanupSessionIdentity As String
Private m_currentSessionIdentity As String
Private m_isProcessingQueue As Boolean

' Events can be successfully received and parsed, then fail inside DispatchGlobalEvents.
' Such messages are already consumed from the FIFO, so track them as upper-layer event loss.
Private m_dispatchEventLossCount As Long
Private m_dispatchLifecycleEventLossCount As Long
Private m_dispatchDownloadEventLossCount As Long
Private m_dispatchFileDialogEventLossCount As Long
Private m_dispatchUnknownCriticalEventLossCount As Long

' Chromium BiDi may leave fetchError.errorText empty for blocked requests;
' classify [BLOCK] by matching the URL against the installed block patterns.
Private m_blockingPatterns As Collection

Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_COMMAND_TIMEOUT_MS As Long = 30000
Private Const NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS As Long = 60000
Private Const COMMAND_CORRELATION_TIMEOUT_MARGIN_MS As Long = 5000
Private Const DEFAULT_STALE_RETRIES As Long = 3
Private Const SHORT_WAIT_MS As Long = 500
' Discovery Log ring-buffer limit.
Private Const MAX_LOG_SIZE As Long = 5000
' Manual Discovery DOM sampling uses the existing SPA probe. Keep the normal cadence aligned
' with the probe/recorder's ~100ms granularity, but back off after a failed probe recovery so
' navigation/realm churn cannot become a tight script.evaluate retry loop.
Private Const DISCOVERY_DOM_POLL_MS As Long = 100
Private Const DISCOVERY_PROBE_RETRY_MS As Long = 500
' Bounded critical-incident ring independent of Discovery Log recording.
Private Const MAX_FORENSIC_INCIDENT_LOG_SIZE As Long = 512
Private Const MAX_FORENSIC_INCIDENT_LINE_CHARS As Long = 2048
Private Const MAX_SUPPRESSED_KEYS As Long = 2000
Private Const MAX_SUPPRESSED_SUMMARY_LINES As Long = 25
Private Const MAX_CONSOLE_LOG_TEXT_SIZE As Long = 2000
Private Const MIN_MSG_SIZE As Long = 10     ' Ignore extremely short messages
' Reserved send-only ID; communicator discards its response.
Private Const POST_SCRIPT_ID As Long = 2

Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore, #vba-bidi-hud, [id^='sizzle']"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"
Private Const SPA_PROBE_VERSION As Long = 41 ' Internal probe generation/quiet/filter/health-diagnostic contract
Private m_userIgnoreSelectors As Collection
Private m_userIgnoreNetworkPatterns As Collection
Private m_idleFilterRevision As Long          ' increments whenever VBA-side ignore filters change
Private m_probeFilterRevisions As Dictionary ' context ID -> last filter revision pushed into that probe
Private m_reportedInvalidSelectors As Dictionary ' selector -> one-shot diagnostic emitted this session

Private m_recorderLogs As Collection
Private m_isRecording As Boolean
Private m_recorderStartTime As Double
Private m_excludeResources As Boolean
Private m_discoveryDomCaptureEnabled As Boolean ' manual Discovery: poll SPA probe into [SPA-ACT]
Private m_discoveryManualPumpActive As Boolean  ' True only inside RecordEventsForSeconds
Private m_discoveryDomContextId As String       ' current top-level context used by manual DOM sampling
Private m_nextDiscoveryDomPollMs As Double      ' nowMs deadline; 0 = sample on next recorder pump
' beforeRequestSent is recorder-only; no latency pairing is retained.
Private m_recSuppByKey As Dictionary       ' normalized key -> suppressed-event count
Private m_recSuppOverflowCount As Long      ' suppressed events not keyed after the key cap
Private m_droppedHeadCount As Long          ' count of oldest timeline lines evicted by the MAX_LOG_SIZE ring buffer (0 = none)

' Critical forensic incidents are retained even when Discovery Log recording is OFF.
' This ring is intentionally low-volume and survives recorder start/stop and ClearState/FatalReset;
' callers may clear it explicitly through ClearForensicIncidentLog.
Private m_forensicIncidentLogs As Collection
Private m_forensicIncidentDroppedHeadCount As Long

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

' Unified download capture. One engine handles a String XPath, a one-item Array/Collection,
' or multiple triggers. The outer Dictionary is keyed by a protocol correlation key and each
' value is a Dictionary containing one download transaction. Correlation prefers the BiDi
' download UUID; context+navigation is the measured Edge fallback.
Private Const DOWNLOAD_CORRELATION_UUID As String = "download"
Private Const DOWNLOAD_CORRELATION_CONTEXT_NAV As String = "context+navigation"
Private Const DOWNLOAD_CORRELATION_UNCORRELATABLE As String = "uncorrelatable"
Private Const DEFAULT_DOWNLOAD_TIMEOUT_MS As Long = 30000
Private m_downloadArmed As Boolean
Private m_downloadOwnerContextId As String
Private m_downloadContextMissWarned As Boolean
Private m_downloadDroppedEventsAtArm As Long
Private m_downloadStartedCount As Long
Private m_downloadUncorrelatableCount As Long
Private m_downloadFirstUncorrelatableContext As String

' The same transaction store is used whether expectedCount is 1 or greater.
Private m_downloadBatchActive As Boolean
Private m_downloadBatchExpectedCount As Long
Private m_downloadBatchDispatchCount As Long
Private m_downloadBatchOverStarted As Boolean
Private m_downloadBatchFirstExtraCorrelationKey As String
Private m_downloadBatchFilenameCollision As Boolean
Private m_downloadBatchCollisionFilename As String
Private m_downloadBatchCollisionFirstCorrelationKey As String
Private m_downloadBatchCollisionSecondCorrelationKey As String
Private m_downloadBatchTransactions As Dictionary

' B1: durable structured snapshot of the most recent download execution failure.
' It is owned by the download batch APIs, independent of live capture
' state, and survives ResetDownloadCaptureState/ClearState until replaced or explicitly cleared.
Private m_lastDownloadFailureDiagnostic As Dictionary

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

' Exclusive ownership for wrapper-local one-shot capture state. VBA DoEvents may re-enter this
' object from Application.OnTime, UserForm/worksheet events, Win32 timers, etc. A second
' composite capture operation must be rejected before it can reset the first operation's state.
Private m_exclusiveCaptureOperationOwner As String

' B4: Fatal Reset must never erase one-shot transaction state while an exclusive capture owns it.
' A fatal reset requested during capture is applied only after the owning operation releases its
' state in its normal success/failure cleanup path.
Private m_fatalResetDeferred As Boolean
Private m_fatalResetDeferredOwner As String
Private m_fatalResetDeferredDroppedArms As String

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
Private Const ERR_MAIN_CONTEXT_LOST As Long = vbObjectError + 1023
Private Const ERR_NEW_CONTEXT_TIMEOUT As Long = vbObjectError + 1024
Private Const ERR_NEW_CONTEXT_MULTIPLE As Long = vbObjectError + 1025
Private Const ERR_NEW_CONTEXT_DESTROYED As Long = vbObjectError + 1026
Private Const ERR_REENTRANT_CAPTURE As Long = vbObjectError + 1027
Private Const ERR_DOWNLOAD_BATCH_COUNT_MISMATCH As Long = vbObjectError + 1028
Private Const ERR_DOWNLOAD_BATCH_FILENAME_COLLISION As Long = vbObjectError + 1029

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

Public Enum BiDiElementAction
    ElementActionClick = 0
    ElementActionInput = 1
    ElementActionInputKeys = 2
    ElementActionSelect = 3
    ElementActionSelectText = 4
    ElementActionGetText = 5
    ElementActionGetAttribute = 6
    ElementActionGetProperty = 7
    ElementActionGetOuterHtml = 8
    ElementActionVisibilityCheck = 9
End Enum

Private Enum ElementLocatorKind
    ElementLocatorNone = 0
    ElementLocatorXPath = 1
    ElementLocatorShadowPath = 2
End Enum

Private Enum waitMode
    WaitMode_UrlContains
    WaitMode_SpaIdle
End Enum
' Extension contracts: result JS uses the Safety Shell; long loops yield and call hb();
' public timeouts use shared validation; wrapper waits yield outside correlation only.

' Download / new-context / file-dialog / navigation-status capture uses class-level state.
' DoEvents inside WaitHeartbeat intentionally keeps Excel responsive, but it can re-enter VBA.
' These helpers make those composite transactions mutually exclusive without disabling the Office
' message pump, and let B4 defer Fatal Reset until the owning transaction has cleaned up.
Private Sub EnterExclusiveCaptureOperation(ByVal methodName As String)
    If Len(m_exclusiveCaptureOperationOwner) > 0 Then
        Err.Raise ERR_REENTRANT_CAPTURE, "BiDi:" & methodName, _
                  "ReentrantCaptureBlocked: '" & methodName & _
                  "' cannot start while '" & m_exclusiveCaptureOperationOwner & _
                  "' owns wrapper capture state. The nested call was rejected before shared " & _
                  "capture state was modified."
    End If

    m_exclusiveCaptureOperationOwner = methodName
End Sub

Private Sub LeaveExclusiveCaptureOperation(ByVal methodName As String)
    If StrComp(m_exclusiveCaptureOperationOwner, methodName, vbBinaryCompare) = 0 Then
        m_exclusiveCaptureOperationOwner = vbNullString

        ' B4: transaction cleanup is complete; a fatal reset deferred during ownership can now
        ' safely clear volatile capture/context state without corrupting the operation in flight.
        If m_fatalResetDeferred Then ApplyDeferredFatalReset
    End If
End Sub

Private Sub RejectIfExclusiveCaptureActive(ByVal methodName As String)
    If Len(m_exclusiveCaptureOperationOwner) = 0 Then Exit Sub

    Err.Raise ERR_REENTRANT_CAPTURE, "BiDi:" & methodName, _
              "ReentrantCaptureBlocked: '" & methodName & _
              "' cannot run while '" & m_exclusiveCaptureOperationOwner & _
              "' owns wrapper capture state."
End Sub

Private Sub AbandonExclusiveCaptureForShutdownBestEffort()
    Dim abandonedOwner As String
    abandonedOwner = m_exclusiveCaptureOperationOwner

    Dim hadDeferredFatalReset As Boolean
    hadDeferredFatalReset = m_fatalResetDeferred

    If Len(abandonedOwner) = 0 And Not hadDeferredFatalReset Then Exit Sub

    ' Shutdown itself clears volatile/session state, so a deferred fatal reset has no later
    ' transaction to protect. Drop both ownership and deferred-reset residue before teardown.
    m_exclusiveCaptureOperationOwner = vbNullString
    m_fatalResetDeferred = False
    m_fatalResetDeferredOwner = vbNullString
    m_fatalResetDeferredDroppedArms = vbNullString
    If hadDeferredFatalReset Then m_consecutiveCdpFailures = 0

    On Error Resume Next
    LogSignalLine "[SHUTDOWN-CAPTURE-ABANDON] owner=" & _
                  IIf(Len(abandonedOwner) > 0, abandonedOwner, "(none)") & _
                  " | deferredFatalReset=" & CStr(hadDeferredFatalReset)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    Set socket_ = New BiDiSocketCommunicator
    Set m_activeSubscriptions = New Dictionary
    Set m_retiredSubscriptionIds = New Dictionary
    Set m_pendingSubscriptionCleanupIds = New Dictionary
    Set m_cachedRealmIds = New Dictionary
    Set m_validatedContexts = New Dictionary
    Set m_probeFilterRevisions = New Dictionary
    Set m_reportedInvalidSelectors = New Dictionary
    Set m_autoClickerScriptIds = New Dictionary
    Set m_autoClickerTimeouts = New Dictionary
    Set m_autoClickerGenerations = New Dictionary
    Set m_autoClickerSuppressorScriptIds = New Dictionary
    Set m_autoClickerContainmentRisks = New Dictionary
    m_lastAutoClickerGuardGeneration = 0#
    Set m_blockingPatterns = New Collection
    Set m_newContextBaseline = New Dictionary
    Set m_forensicIncidentLogs = New Collection
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

    ' A previously disconnected transport can still belong to a live BiDi session. Preserve any
    ' subscription IDs whose remote cleanup was not confirmed before dropping local session state.
    CaptureTrackedSubscriptionsForDeferredCleanup

    ResetSessionBoundaryState
    ClearStateCore False

    Dim requestedSessionIdentity As String
    requestedSessionIdentity = GetBiDiSessionIdentity(webSocketUrl)

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
    m_currentSessionIdentity = requestedSessionIdentity

    ' A normal Shutdown can lose only the acknowledgement, not necessarily the unsubscribe itself.
    ' Reconcile those issued IDs before creating any new subscriptions in the same BiDi session.
    RecoverPendingSubscriptionCleanupOnConnect requestedSessionIdentity
    
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

    ' Initialization may already have created session-side state before a later mandatory step
    ' failed. Contain/restore Shadow DOM hooks and release subscriptions while the socket is live.
    Call ReleaseGlobalShadowUnlockerBestEffort
    If Not ReleaseTrackedSubscriptionsBeforeClearBestEffort(2000) Then
        CaptureTrackedSubscriptionsForDeferredCleanup
    End If

    socket_.ForceClose
    ResetSessionBoundaryState
    ClearStateCore False
    m_currentSessionIdentity = vbNullString
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
    m_currentSessionIdentity = vbNullString
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
            ObserveNavigationStartedForActiveNavigate json

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

        Case "network.responseStarted"
            HandleNavigationResponseStartedEvent json

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
            LogSignalLine "[NEW-CONTEXT-IGNORED] ctx=" & eventCtx & _
                          " | reason=foreign-opener | opener=" & openerId & _
                          " | owner=" & m_newContextOwnerContextId
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

        LogSignalLine "[NEW-CONTEXT-CAPTURED] ctx=" & eventCtx & _
                      " | owner=" & m_newContextOwnerContextId & _
                      " | correlation=" & correlationMode & _
                      " | opener=" & IIf(Len(openerId) > 0, openerId, "(none)")
        Exit Sub
    End If

    If StrComp(eventCtx, m_newContextId, vbBinaryCompare) = 0 Then Exit Sub

    m_newContextCandidateCount = m_newContextCandidateCount + 1
    m_newContextAmbiguous = True
    m_newContextArmed = False
    LogSignalLine "[NEW-CONTEXT-AMBIGUOUS] firstCtx=" & m_newContextId & _
                  " | secondCtx=" & eventCtx & _
                  " | candidateCount=" & CStr(m_newContextCandidateCount)
End Sub

Private Sub HandleContextDestroyedEvent(ByVal json As Object)
    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then Exit Sub

    InvalidateContextState eventCtx, True

    If m_newContextArmed Then
        If StrComp(eventCtx, m_newContextOwnerContextId, vbBinaryCompare) = 0 Then
            m_newContextOwnerDestroyed = True
            LogSignalLine "[NEW-CONTEXT-OWNER-DESTROYED] ctx=" & eventCtx
        End If
    End If

    If m_newContextSeen Then
        If StrComp(eventCtx, m_newContextId, vbBinaryCompare) = 0 Then
            m_newContextCandidateDestroyed = True
            m_newContextArmed = False
            LogSignalLine "[NEW-CONTEXT-CANDIDATE-DESTROYED] ctx=" & eventCtx & _
                          " | owner=" & m_newContextOwnerContextId
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
    ' Remove dead-context keys from the active delta registry, but retain their server-issued
    ' subscription IDs for the next ClearState/Shutdown unsubscribe.
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
                Dim retiredId As String
                retiredId = CStr(m_activeSubscriptions(key))
                EnsureRetiredSubscriptionRegistry
                If Len(retiredId) > 0 Then
                    If Not m_retiredSubscriptionIds.Exists(retiredId) Then _
                        m_retiredSubscriptionIds.Add retiredId, True
                End If
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
    LogSignalLine "[NEW-CONTEXT-ARM] owner=" & ownerContextId & _
                  " | baselineCount=" & CStr(m_newContextBaseline.count) & _
                  " | droppedBase=" & CStr(m_newContextDroppedEventsAtArm)
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

        Dim droppedSinceArm As Long
        droppedSinceArm = NewContextDroppedEventsSinceArm()
        If droppedSinceArm > 0 Then
            LogSignalLine "[NEW-CONTEXT-LOSS] owner=" & m_newContextOwnerContextId & _
                          " | droppedSinceArm=" & CStr(droppedSinceArm)
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
                LogSignalLine "[NEW-CONTEXT-SETTLED] ctx=" & m_newContextId & _
                              " | correlation=" & m_newContextCorrelationMode & _
                              " | candidateCount=" & CStr(m_newContextCandidateCount)
                Exit Sub
            End If
        ElseIf IsTimeout(startTick, CDbl(timeoutMs)) Then
            m_newContextArmed = False

            Dim ownerNote As String
            If m_newContextOwnerDestroyed Then
                ownerNote = " Owner context was destroyed during the wait."
            End If

            LogSignalLine "[NEW-CONTEXT-TIMEOUT] owner=" & m_newContextOwnerContextId & _
                          " | timeoutMs=" & CStr(timeoutMs) & _
                          " | foreignIgnored=" & CStr(m_newContextForeignIgnoredCount) & _
                          " | ownerDestroyed=" & LCase$(CStr(m_newContextOwnerDestroyed))

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

Private Function TryBuildDownloadBatchCorrelationKey(ByVal p As Object, _
                                                     ByVal eventCtx As String, _
                                                     ByRef correlationKey As String, _
                                                     ByRef correlationMode As String, _
                                                     ByRef downloadId As String, _
                                                     ByRef navigationId As String) As Boolean
    correlationKey = vbNullString
    correlationMode = DOWNLOAD_CORRELATION_UNCORRELATABLE
    downloadId = vbNullString
    navigationId = vbNullString

    If TryGetDownloadCorrelationValue(p, "download", downloadId) Then
        correlationMode = DOWNLOAD_CORRELATION_UUID
        correlationKey = "D|" & downloadId
        TryBuildDownloadBatchCorrelationKey = True
        Exit Function
    End If

    If TryGetDownloadCorrelationValue(p, "navigation", navigationId) Then
        correlationMode = DOWNLOAD_CORRELATION_CONTEXT_NAV
        correlationKey = "N|" & eventCtx & "|" & navigationId
        TryBuildDownloadBatchCorrelationKey = True
    End If
End Function

Private Function FindDownloadBatchCorrelationKeyBySuggestedFilename( _
                    ByVal suggestedFilename As String) As String
    If Len(suggestedFilename) = 0 Then Exit Function
    If m_downloadBatchTransactions Is Nothing Then Exit Function

    Dim key As Variant
    Dim tx As Dictionary
    Dim existingFilename As String

    For Each key In m_downloadBatchTransactions.keys
        Set tx = m_downloadBatchTransactions(key)
        If tx.Exists("suggestedFilename") Then
            existingFilename = CStr(tx("suggestedFilename"))
            If Len(existingFilename) > 0 Then
                ' Chromium download destinations are Windows paths in the supported SeleniumVBA
                ' environment, so filename collision detection is intentionally case-insensitive.
                If StrComp(existingFilename, suggestedFilename, vbTextCompare) = 0 Then
                    FindDownloadBatchCorrelationKeyBySuggestedFilename = CStr(key)
                    Exit Function
                End If
            End If
        End If
    Next key
End Function

Private Sub HandleDownloadBatchWillBeginEvent(ByVal json As Object, ByVal p As Object)
    If Not m_downloadBatchActive Or Not m_downloadArmed Then Exit Sub

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then
        LogSignalLine "[DOWNLOAD-BATCH-WILL-BEGIN-IGNORED] reason=missing-context"
        Exit Sub
    End If

    If StrComp(eventCtx, m_downloadOwnerContextId, vbBinaryCompare) <> 0 Then
        If Not m_downloadContextMissWarned Then
            m_downloadContextMissWarned = True
            LogSignalLine "[DOWNLOAD-BATCH-CTX-MISS] expectedCtx=" & m_downloadOwnerContextId & _
                          " | actualCtx=" & eventCtx
        End If
        Exit Sub
    End If

    Dim correlationKey As String
    Dim correlationMode As String
    Dim downloadId As String
    Dim navigationId As String
    If Not TryBuildDownloadBatchCorrelationKey(p, eventCtx, correlationKey, correlationMode, downloadId, navigationId) Then
        m_downloadUncorrelatableCount = m_downloadUncorrelatableCount + 1
        If Len(m_downloadFirstUncorrelatableContext) = 0 Then
            m_downloadFirstUncorrelatableContext = eventCtx
        End If
        LogSignalLine "[DOWNLOAD-BATCH-WILL-BEGIN-IGNORED] reason=uncorrelatable | ctx=" & eventCtx & _
                      " | count=" & CStr(m_downloadUncorrelatableCount)
        Exit Sub
    End If

    If m_downloadBatchTransactions Is Nothing Then Set m_downloadBatchTransactions = New Dictionary

    ' Duplicate delivery of the same protocol identity must not create another transaction.
    If m_downloadBatchTransactions.Exists(correlationKey) Then
        LogSignalLine "[DOWNLOAD-BATCH-WILL-BEGIN-IGNORED] reason=duplicate-correlation | key=" & correlationKey
        Exit Sub
    End If

    Dim suggestedFilename As String
    suggestedFilename = GetPathString(json, "params.suggestedFilename")

    ' A second transaction with the same browser-suggested filename is unsafe when a global
    ' destination folder is active: measured Edge behavior can report complete for both while
    ' leaving only the last artifact at that path. Latch the first observed collision. The second
    ' transaction is still registered because it has already started in the browser.
    Dim collisionFirstKey As String
    collisionFirstKey = FindDownloadBatchCorrelationKeyBySuggestedFilename(suggestedFilename)

    Dim tx As New Dictionary
    tx.Add "correlationKey", correlationKey
    tx.Add "correlationMode", correlationMode
    tx.Add "context", eventCtx
    tx.Add "downloadId", downloadId
    tx.Add "navigationId", navigationId
    tx.Add "suggestedFilename", suggestedFilename
    tx.Add "url", GetPathString(json, "params.url")
    tx.Add "state", "started"
    tx.Add "startedSeen", True
    tx.Add "endSeen", False
    tx.Add "status", vbNullString
    tx.Add "filePath", vbNullString

    m_downloadBatchTransactions.Add correlationKey, tx
    m_downloadStartedCount = m_downloadBatchTransactions.count

    If Len(collisionFirstKey) > 0 And Not m_downloadBatchFilenameCollision Then
        m_downloadBatchFilenameCollision = True
        m_downloadBatchCollisionFilename = suggestedFilename
        m_downloadBatchCollisionFirstCorrelationKey = collisionFirstKey
        m_downloadBatchCollisionSecondCorrelationKey = correlationKey

        LogSignalLine "[DOWNLOAD-BATCH-FILENAME-COLLISION] filename=" & suggestedFilename & _
                      " | firstKey=" & collisionFirstKey & _
                      " | secondKey=" & correlationKey
    End If

    If m_downloadStartedCount > m_downloadBatchExpectedCount Then
        m_downloadBatchOverStarted = True
        If Len(m_downloadBatchFirstExtraCorrelationKey) = 0 Then
            m_downloadBatchFirstExtraCorrelationKey = correlationKey
        End If
    End If

    LogSignalLine "[DOWNLOAD-BATCH-START-ACCEPTED] key=" & correlationKey & _
                  " | ctx=" & eventCtx & _
                  " | correlation=" & correlationMode & _
                  " | count=" & CStr(m_downloadStartedCount) & _
                  " | expected=" & CStr(m_downloadBatchExpectedCount)
End Sub

Private Sub HandleDownloadBatchEndEvent(ByVal json As Object, ByVal p As Object)
    If Not m_downloadBatchActive Then Exit Sub

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then
        LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=missing-context"
        Exit Sub
    End If

    If StrComp(eventCtx, m_downloadOwnerContextId, vbBinaryCompare) <> 0 Then
        LogSignalLine "[DOWNLOAD-BATCH-END-CTX-MISS] expectedCtx=" & m_downloadOwnerContextId & _
                      " | actualCtx=" & eventCtx
        Exit Sub
    End If

    Dim correlationKey As String
    Dim correlationMode As String
    Dim downloadId As String
    Dim navigationId As String
    If Not TryBuildDownloadBatchCorrelationKey(p, eventCtx, correlationKey, correlationMode, downloadId, navigationId) Then
        LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=uncorrelatable | ctx=" & eventCtx
        Exit Sub
    End If

    If m_downloadBatchTransactions Is Nothing Then Exit Sub
    If Not m_downloadBatchTransactions.Exists(correlationKey) Then
        LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=unknown-correlation | key=" & correlationKey
        Exit Sub
    End If

    Dim tx As Dictionary
    Set tx = m_downloadBatchTransactions(correlationKey)
    If CBool(tx("endSeen")) Then
        LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=duplicate-terminal | key=" & correlationKey
        Exit Sub
    End If

    Dim terminalStatus As String
    If Not TryGetDownloadCorrelationValue(p, "status", terminalStatus) Then
        LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=missing-status | key=" & correlationKey
        Exit Sub
    End If

    Select Case LCase$(terminalStatus)
        Case "complete", "canceled"
        Case Else
            LogSignalLine "[DOWNLOAD-BATCH-END-IGNORED] reason=unknown-status | key=" & correlationKey & _
                          " | status=" & terminalStatus
            Exit Sub
    End Select

    Dim filePath As String
    Call TryGetDownloadCorrelationValue(p, "filepath", filePath)

    tx("endSeen") = True
    tx("status") = terminalStatus
    tx("filePath") = filePath
    tx("state") = LCase$(terminalStatus)

    LogSignalLine "[DOWNLOAD-BATCH-END-MATCHED] key=" & correlationKey & _
                  " | status=" & terminalStatus & _
                  " | filepath=" & filePath
End Sub

Private Function DownloadBatchTerminalCount() As Long
    If m_downloadBatchTransactions Is Nothing Then Exit Function

    Dim key As Variant
    Dim tx As Dictionary
    For Each key In m_downloadBatchTransactions.keys
        Set tx = m_downloadBatchTransactions(key)
        If CBool(tx("endSeen")) Then DownloadBatchTerminalCount = DownloadBatchTerminalCount + 1
    Next key
End Function

Private Function AllDownloadBatchTransactionsTerminal() As Boolean
    If m_downloadBatchTransactions Is Nothing Then Exit Function
    If m_downloadBatchTransactions.count <> m_downloadBatchExpectedCount Then Exit Function

    Dim key As Variant
    Dim tx As Dictionary
    For Each key In m_downloadBatchTransactions.keys
        Set tx = m_downloadBatchTransactions(key)
        If Not CBool(tx("endSeen")) Then Exit Function
    Next key

    AllDownloadBatchTransactionsTerminal = True
End Function

Private Sub ResetDownloadBatchState()
    m_downloadBatchActive = False
    m_downloadBatchExpectedCount = 0
    m_downloadBatchDispatchCount = 0
    m_downloadBatchOverStarted = False
    m_downloadBatchFirstExtraCorrelationKey = vbNullString
    m_downloadBatchFilenameCollision = False
    m_downloadBatchCollisionFilename = vbNullString
    m_downloadBatchCollisionFirstCorrelationKey = vbNullString
    m_downloadBatchCollisionSecondCorrelationKey = vbNullString
    Set m_downloadBatchTransactions = Nothing
End Sub

Private Sub ArmDownloadBatch(ByVal contextId As String, ByVal expectedCount As Long)
    ResetDownloadCaptureState
    m_downloadOwnerContextId = contextId
    m_downloadBatchExpectedCount = expectedCount
    Set m_downloadBatchTransactions = New Dictionary
    m_downloadDroppedEventsAtArm = CurrentDownloadRelevantLossCount()
    m_downloadBatchActive = True
    m_downloadArmed = True

    LogSignalLine "[DOWNLOAD-BATCH-ARM] ctx=" & contextId & _
                  " | expected=" & CStr(expectedCount) & _
                  " | droppedBase=" & CStr(m_downloadDroppedEventsAtArm)
End Sub

Private Sub RaiseDownloadBatchCountMismatch(ByVal methodName As String)
    m_downloadArmed = False
    m_downloadBatchActive = False
    Err.Raise ERR_DOWNLOAD_BATCH_COUNT_MISMATCH, "BiDi:" & methodName, _
              "DownloadBatchCountMismatch: more correlatable downloads started than expected. " & _
              "The trigger clicks are not replayed. Expected=" & CStr(m_downloadBatchExpectedCount) & _
              "; Started=" & CStr(m_downloadStartedCount) & _
              "; FirstExtraCorrelationKey=" & _
              IIf(Len(m_downloadBatchFirstExtraCorrelationKey) > 0, m_downloadBatchFirstExtraCorrelationKey, "(none)") & "."
End Sub

Private Sub RaiseDownloadBatchFilenameCollision(ByVal methodName As String)
    m_downloadArmed = False
    m_downloadBatchActive = False
    Err.Raise ERR_DOWNLOAD_BATCH_FILENAME_COLLISION, "BiDi:" & methodName, _
              "DownloadBatchFilenameCollision: multiple downloads in the same batch reported the " & _
              "same suggested filename. Concurrent downloads may target the same destination path " & _
              "and overwrite an earlier artifact. The collision is already browser-observed, so " & _
              "already-started downloads may have modified disk. Remaining trigger clicks are not " & _
              "dispatched after the collision is observed. Filename=" & _
              IIf(Len(m_downloadBatchCollisionFilename) > 0, m_downloadBatchCollisionFilename, "(missing)") & _
              "; FirstCorrelationKey=" & _
              IIf(Len(m_downloadBatchCollisionFirstCorrelationKey) > 0, m_downloadBatchCollisionFirstCorrelationKey, "(none)") & _
              "; SecondCorrelationKey=" & _
              IIf(Len(m_downloadBatchCollisionSecondCorrelationKey) > 0, m_downloadBatchCollisionSecondCorrelationKey, "(none)") & "."
End Sub

Private Sub WaitForDownloadBatch(ByVal timeoutMs As Long, ByVal methodName As String)
    Dim stageTick As Currency
    stageTick = socket_.NowTicks()

    ' Stage 1: collect exactly the expected number of correlatable starts. All trigger clicks have
    ' already been dispatched, but no download completion has been awaited, so transfers can overlap.
    Do
        ProcessRecorderQueue

        Dim downloadLossStart As Long
        downloadLossStart = DownloadDroppedEventsSinceArm()
        If downloadLossStart > 0 Then
            m_downloadArmed = False
            m_downloadBatchActive = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                      "Async event loss occurred while waiting for download batch start; reliable " & _
                      "batch correlation cannot be determined. The trigger clicks are not replayed. " & _
                      "AsyncEventsDroppedSinceArm=" & CStr(downloadLossStart) & "."
        End If

        If m_downloadBatchFilenameCollision Then
            RaiseDownloadBatchFilenameCollision methodName
        End If

        If m_downloadBatchOverStarted Or m_downloadStartedCount > m_downloadBatchExpectedCount Then
            RaiseDownloadBatchCountMismatch methodName
        End If

        If m_downloadStartedCount = m_downloadBatchExpectedCount Then
            ' Freeze the start set after draining the current recorder queue. Terminal events for
            ' accepted transactions continue to be processed while the batch remains active.
            m_downloadArmed = False
            Exit Do
        End If

        If IsTimeout(stageTick, CDbl(timeoutMs)) Then
            m_downloadArmed = False
            m_downloadBatchActive = False
            Err.Raise ERR_DOWNLOAD_START_TIMEOUT, "BiDi:" & methodName, _
                      "DownloadBatchStartTimeout: expected " & CStr(m_downloadBatchExpectedCount) & _
                      " correlatable downloads but accepted " & CStr(m_downloadStartedCount) & _
                      " within " & CStr(timeoutMs) & "ms. The trigger clicks are not replayed. " & _
                      "UncorrelatableStarts=" & CStr(m_downloadUncorrelatableCount) & _
                      "; AsyncQueueDepthAtTimeout=" & CStr(DownloadEventQueueDepth()) & "."
        End If
        WaitHeartbeat 50
    Loop

    stageTick = socket_.NowTicks()

    ' Stage 2: wait for every accepted transaction to reach a browser-authoritative terminal state.
    Do
        ProcessRecorderQueue

        Dim downloadLossCompletion As Long
        downloadLossCompletion = DownloadDroppedEventsSinceArm()
        If downloadLossCompletion > 0 Then
            m_downloadBatchActive = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                      "Async event loss occurred while waiting for download batch completion; reliable " & _
                      "batch results cannot be determined. Downloads may already exist on disk. " & _
                      "AsyncEventsDroppedSinceArm=" & CStr(downloadLossCompletion) & "."
        End If

        If m_downloadBatchFilenameCollision Then RaiseDownloadBatchFilenameCollision methodName
        If m_downloadBatchOverStarted Then RaiseDownloadBatchCountMismatch methodName
        If AllDownloadBatchTransactionsTerminal() Then Exit Sub

        If IsTimeout(stageTick, CDbl(timeoutMs)) Then
            m_downloadBatchActive = False
            Err.Raise ERR_DOWNLOAD_COMPLETION_TIMEOUT, "BiDi:" & methodName, _
                      "DownloadBatchCompletionTimeout: only " & CStr(DownloadBatchTerminalCount()) & _
                      " of " & CStr(m_downloadBatchExpectedCount) & _
                      " downloads reached complete/canceled within " & CStr(timeoutMs) & "ms. " & _
                      "The trigger clicks are not replayed. AsyncQueueDepthAtTimeout=" & _
                      CStr(DownloadEventQueueDepth()) & "."
        End If
        WaitHeartbeat 50
    Loop
End Sub

Private Function BuildDownloadBatchResult() As Dictionary
    Dim result As New Dictionary
    Dim downloads As New Dictionary
    Dim completeCount As Long
    Dim canceledCount As Long

    If Not m_downloadBatchTransactions Is Nothing Then
        Dim key As Variant
        Dim tx As Dictionary
        Dim snapshot As Dictionary

        For Each key In m_downloadBatchTransactions.keys
            Set tx = m_downloadBatchTransactions(key)
            Set snapshot = CloneFlatDictionary(tx)
            downloads.Add CStr(key), snapshot

            Select Case LCase$(CStr(tx("status")))
                Case "complete": completeCount = completeCount + 1
                Case "canceled": canceledCount = canceledCount + 1
            End Select
        Next key
    End If

    If completeCount = m_downloadBatchExpectedCount And canceledCount = 0 Then
        result.Add "status", "complete"
    Else
        result.Add "status", "failed"
    End If

    result.Add "expectedCount", m_downloadBatchExpectedCount
    result.Add "dispatchCount", m_downloadBatchDispatchCount
    result.Add "startedCount", m_downloadStartedCount
    result.Add "terminalCount", DownloadBatchTerminalCount()
    result.Add "completeCount", completeCount
    result.Add "canceledCount", canceledCount
    result.Add "context", m_downloadOwnerContextId
    result.Add "asyncEventsDroppedSinceArm", DownloadDroppedEventsSinceArm()
    result.Add "downloads", downloads

    Set BuildDownloadBatchResult = result
End Function

Private Function NormalizeDownloadTriggerXpaths(ByVal triggerXpaths As Variant, _
                                                ByVal methodName As String) As Collection
    Dim result As New Collection
    Dim item As Variant
    Dim s As String

    On Error GoTo InvalidInput

    If IsObject(triggerXpaths) Then
        If TypeName(triggerXpaths) <> "Collection" Then GoTo InvalidInput
        For Each item In triggerXpaths
            If IsObject(item) Or IsNull(item) Then GoTo InvalidInput
            s = Trim$(CStr(item))
            If Len(s) = 0 Then GoTo InvalidInput
            result.Add s
        Next item
    ElseIf IsArray(triggerXpaths) Then
        Dim i As Long
        For i = LBound(triggerXpaths) To UBound(triggerXpaths)
            item = triggerXpaths(i)
            If IsObject(item) Or IsNull(item) Then GoTo InvalidInput
            s = Trim$(CStr(item))
            If Len(s) = 0 Then GoTo InvalidInput
            result.Add s
        Next i
    Else
        If IsNull(triggerXpaths) Then GoTo InvalidInput
        s = Trim$(CStr(triggerXpaths))
        If Len(s) = 0 Then GoTo InvalidInput
        result.Add s
    End If

    If result.count = 0 Then GoTo InvalidInput
    Set NormalizeDownloadTriggerXpaths = result
    Exit Function

InvalidInput:
    Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
              "triggerXpaths must be a non-empty String, Variant array, or Collection of non-empty XPath strings."
End Function

Private Function NormalizeShadowDownloadTriggerPaths(ByVal triggerSelectorPaths As Variant, _
                                                     ByVal methodName As String) As Collection
    Dim result As New Collection
    Dim item As Variant

    On Error GoTo InvalidInput

    If IsObject(triggerSelectorPaths) Then
        If TypeName(triggerSelectorPaths) <> "Collection" Then GoTo InvalidInput
        For Each item In triggerSelectorPaths
            ValidateShadowDownloadTriggerPath item, methodName
            result.Add item
        Next item
    Else
        ValidateShadowDownloadTriggerPath triggerSelectorPaths, methodName
        result.Add triggerSelectorPaths
    End If

    If result.count = 0 Then GoTo InvalidInput
    Set NormalizeShadowDownloadTriggerPaths = result
    Exit Function

InvalidInput:
    Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
              "triggerSelectorPaths must be one non-empty Shadow selector path, or a " & _
              "Collection of non-empty Shadow selector paths."
End Function

Private Sub ValidateShadowDownloadTriggerPath(ByVal selectorPath As Variant, _
                                              ByVal methodName As String)
    Dim item As Variant
    Dim i As Long
    Dim s As String

    If IsObject(selectorPath) Or IsNull(selectorPath) Then GoTo InvalidInput

    If IsArray(selectorPath) Then
        If LBound(selectorPath) > UBound(selectorPath) Then GoTo InvalidInput
        For i = LBound(selectorPath) To UBound(selectorPath)
            item = selectorPath(i)
            If IsObject(item) Or IsNull(item) Then GoTo InvalidInput
            s = Trim$(CStr(item))
            If Len(s) = 0 Then GoTo InvalidInput
        Next i
    Else
        s = Trim$(CStr(selectorPath))
        If Len(s) = 0 Then GoTo InvalidInput
    End If
    Exit Sub

InvalidInput:
    Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
              "Each Shadow download trigger must be a non-empty String or Variant array " & _
              "of non-empty selector strings."
End Sub


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

    HandleDownloadBatchWillBeginEvent json, p
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

    HandleDownloadBatchEndEvent json, p
End Sub

Private Sub ResetDownloadCaptureState()
    ResetDownloadBatchState
    m_downloadArmed = False
    m_downloadOwnerContextId = vbNullString
    m_downloadContextMissWarned = False
    m_downloadDroppedEventsAtArm = 0
    m_downloadStartedCount = 0
    m_downloadUncorrelatableCount = 0
    m_downloadFirstUncorrelatableContext = vbNullString
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

Private Function ClassifyDownloadFailureKind(ByVal errNum As Long, _
                                             ByVal errDesc As String, _
                                             ByVal captureInitialized As Boolean) As String
    Select Case errNum
        Case ERR_DOWNLOAD_START_TIMEOUT
            ClassifyDownloadFailureKind = "start-timeout"
            Exit Function
        Case ERR_DOWNLOAD_COMPLETION_TIMEOUT
            ClassifyDownloadFailureKind = "completion-timeout"
            Exit Function
        Case ERR_DOWNLOAD_BATCH_COUNT_MISMATCH
            ClassifyDownloadFailureKind = "batch-count-mismatch"
            Exit Function
        Case ERR_DOWNLOAD_BATCH_FILENAME_COLLISION
            ClassifyDownloadFailureKind = "batch-filename-collision"
            Exit Function
    End Select

    If InStr(1, errDesc, "Async event loss occurred while waiting for download batch start", vbTextCompare) > 0 Then
        ClassifyDownloadFailureKind = "event-loss-start"
    ElseIf InStr(1, errDesc, "Async event loss occurred while waiting for download batch completion", vbTextCompare) > 0 Then
        ClassifyDownloadFailureKind = "event-loss-completion"
    ElseIf captureInitialized Then
        ClassifyDownloadFailureKind = "post-arm-failure"
    Else
        ClassifyDownloadFailureKind = "pre-arm-failure"
    End If
End Function

Private Sub CaptureLastDownloadFailureDiagnostic(ByVal errNum As Long, _
                                                 ByVal errSrc As String, _
                                                 ByVal errDesc As String, _
                                                 ByVal captureInitialized As Boolean)
    ' This routine runs from an active error handler. It must never replace the primary failure.
    On Error Resume Next

    Dim droppedSinceArm As Long
    Dim queueDepth As Long
    droppedSinceArm = -1
    queueDepth = -1

    If captureInitialized And Not socket_ Is Nothing Then
        Err.Clear
        droppedSinceArm = DownloadDroppedEventsSinceArm()
        If Err.Number <> 0 Then
            droppedSinceArm = -1
            Err.Clear
        End If

        queueDepth = DownloadEventQueueDepth()
        If Err.Number <> 0 Then
            queueDepth = -1
            Err.Clear
        End If
    End If

    Dim d As New Dictionary
    d.Add "errorNumber", errNum
    d.Add "errorSource", errSrc
    d.Add "errorDescription", errDesc
    d.Add "failureKind", ClassifyDownloadFailureKind(errNum, errDesc, captureInitialized)
    d.Add "captureInitialized", captureInitialized
    d.Add "ownerContext", m_downloadOwnerContextId
    d.Add "uncorrelatableStarts", m_downloadUncorrelatableCount
    d.Add "firstUncorrelatableContext", m_downloadFirstUncorrelatableContext
    d.Add "asyncEventsDroppedSinceArm", droppedSinceArm
    d.Add "eventQueueDepthAtFailure", queueDepth
    d.Add "batchActive", m_downloadBatchActive
    d.Add "batchExpectedCount", m_downloadBatchExpectedCount
    d.Add "batchDispatchCount", m_downloadBatchDispatchCount
    If m_downloadBatchTransactions Is Nothing Then
        d.Add "batchStartedCount", 0
        d.Add "batchTerminalCount", 0
    Else
        d.Add "batchStartedCount", m_downloadBatchTransactions.count
        d.Add "batchTerminalCount", DownloadBatchTerminalCount()
    End If
    d.Add "batchOverStarted", m_downloadBatchOverStarted
    d.Add "batchFirstExtraCorrelationKey", m_downloadBatchFirstExtraCorrelationKey
    d.Add "batchFilenameCollision", m_downloadBatchFilenameCollision
    d.Add "batchCollisionFilename", m_downloadBatchCollisionFilename
    d.Add "batchCollisionFirstCorrelationKey", m_downloadBatchCollisionFirstCorrelationKey
    d.Add "batchCollisionSecondCorrelationKey", m_downloadBatchCollisionSecondCorrelationKey

    Set m_lastDownloadFailureDiagnostic = d
    Err.Clear
    On Error GoTo 0
End Sub

Private Function CloneFlatDictionary(ByVal source As Dictionary) As Dictionary
    If source Is Nothing Then Exit Function

    Dim snapshot As New Dictionary
    Dim key As Variant
    For Each key In source.keys
        snapshot.Add CStr(key), source(key)
    Next key

    Set CloneFlatDictionary = snapshot
End Function

' Returns a detached structured snapshot of the most recent download execution failure.
' Nothing means no failure has been captured since construction or the last explicit clear.
Public Function GetLastDownloadFailureDiagnostic() As Dictionary
    If m_lastDownloadFailureDiagnostic Is Nothing Then Exit Function
    Set GetLastDownloadFailureDiagnostic = CloneFlatDictionary(m_lastDownloadFailureDiagnostic)
End Function

Public Sub ClearLastDownloadFailureDiagnostic()
    RejectIfExclusiveCaptureActive "ClearLastDownloadFailureDiagnostic"
    Set m_lastDownloadFailureDiagnostic = Nothing
End Sub

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

' Denies downloads globally for the current browser session. A resulting canceled transaction is
' aggregated by the download batch transaction into the returned result.
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

' Executes one or more XPath download triggers through the shared transaction engine.
' triggerXpaths accepts one String XPath, a Variant array, or a Collection of XPath strings.
Public Function ExecuteDownloadsByXPath(ByVal triggerXpaths As Variant, _
                                        Optional ByVal searchTimeoutMs As Long = 10000, _
                                        Optional ByVal timeoutMs As Long = DEFAULT_DOWNLOAD_TIMEOUT_MS, _
                                        Optional ByVal contextId As String = "") As Dictionary
    Dim triggers As Collection
    Set triggers = NormalizeDownloadTriggerXpaths(triggerXpaths, "ExecuteDownloadsByXPath")

    Set ExecuteDownloadsByXPath = ExecuteDownloadsCore( _
        ElementLocatorXPath, triggers, searchTimeoutMs, timeoutMs, contextId, _
        "ExecuteDownloadsByXPath")
End Function

' Executes one or more Shadow DOM download triggers through the same transaction engine.
' A String or Variant array represents one Shadow selector path. For multiple triggers, pass a
' Collection whose items are individual String or Variant-array selector paths.
Public Function ExecuteShadowDownloads(ByVal triggerSelectorPaths As Variant, _
                                       Optional ByVal searchTimeoutMs As Long = 10000, _
                                       Optional ByVal timeoutMs As Long = DEFAULT_DOWNLOAD_TIMEOUT_MS, _
                                       Optional ByVal contextId As String = "", _
                                       Optional ByVal requireUnlocker As Boolean = False) As Dictionary
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowDownloads"

    Dim triggers As Collection
    Set triggers = NormalizeShadowDownloadTriggerPaths( _
        triggerSelectorPaths, "ExecuteShadowDownloads")

    Set ExecuteShadowDownloads = ExecuteDownloadsCore( _
        ElementLocatorShadowPath, triggers, searchTimeoutMs, timeoutMs, contextId, _
        "ExecuteShadowDownloads")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, _
                             "ExecuteShadowDownloads"
End Function

' Runs the download batch transaction for either XPath or Shadow trigger locators.
Private Function ExecuteDownloadsCore(ByVal triggerLocatorKind As ElementLocatorKind, _
                                      ByVal triggers As Collection, _
                                      ByVal searchTimeoutMs As Long, _
                                      ByVal timeoutMs As Long, _
                                      ByVal contextId As String, _
                                      ByVal callerName As String) As Dictionary
    Dim methodName As String
    methodName = callerName

    Dim exclusiveCaptureEntered As Boolean
    EnterExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = True
    On Error GoTo Failed

    ResetDownloadCaptureState

    Dim downloadCaptureInitialized As Boolean
    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeoutMs, "timeoutMs", DEFAULT_DOWNLOAD_TIMEOUT_MS, 1)

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    If m_downloadBehaviorStateUnknown Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & methodName, _
                  "DownloadBehaviorStateUnknown: browser.setDownloadBehavior cleanup could not " & _
                  "be confirmed. Call SetDownloadFolder, SetDownloadDenied, or " & _
                  "ClearDownloadBehavior before starting another download."
    End If

    Dim carryoverDiagnosticActive As Boolean
    Dim carryDomSeenBefore As Boolean
    Dim carryNetSeenBefore As Boolean
    Dim carryVisSeenBefore As Boolean
    carryoverDiagnosticActive = BeginSignalCarryoverDiagnostic( _
        methodName, targetCtx, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore)

    Me.ExecuteSessionSubscribe Array( _
        "browsingContext.downloadWillBegin", _
        "browsingContext.downloadEnd"), , True

    ' Resolve every trigger before arming so a locator failure causes no download side effect.
    Dim resolvedTriggerIds As New Collection
    Dim trigger As Variant
    For Each trigger In triggers
        resolvedTriggerIds.Add ResolveElementSharedId( _
            triggerLocatorKind, trigger, targetCtx, searchTimeoutMs, methodName)
    Next trigger

    ProcessRecorderQueue
    ArmDownloadBatch targetCtx, resolvedTriggerIds.count
    downloadCaptureInitialized = True

    Dim triggerSharedId As Variant
    Dim clickArgs As Collection
    Dim clickResult As String
    Dim clickErrNum As Long
    Dim clickErrSrc As String
    Dim clickErrDesc As String

    For Each triggerSharedId In resolvedTriggerIds
        ProcessRecorderQueue
        If m_downloadBatchFilenameCollision Then RaiseDownloadBatchFilenameCollision methodName

        Set clickArgs = NewElementArgs(CStr(triggerSharedId))

        On Error Resume Next
        Err.Clear
        clickResult = PerformTrustedClick(targetCtx, clickArgs, vTimeout)
        clickErrNum = Err.Number
        clickErrSrc = Err.source
        clickErrDesc = Err.description
        Err.Clear
        On Error GoTo Failed

        If clickErrNum <> 0 Then
            ProcessRecorderQueue
            m_downloadArmed = False
            m_downloadBatchActive = False
            Err.Raise clickErrNum, clickErrSrc, clickErrDesc
        End If

        m_downloadBatchDispatchCount = m_downloadBatchDispatchCount + 1

        ProcessRecorderQueue
        If m_downloadBatchFilenameCollision Then RaiseDownloadBatchFilenameCollision methodName
    Next triggerSharedId

    WaitForDownloadBatch vTimeout, methodName

    EndSignalCarryoverDiagnostic methodName, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore

    Dim result As Dictionary
    Set result = BuildDownloadBatchResult()

    m_downloadBatchActive = False
    LeaveExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = False

    Set ExecuteDownloadsCore = result
    Exit Function

Failed:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrSrc As String: savedErrSrc = Err.source
    Dim savedErrDesc As String: savedErrDesc = Err.description

    CaptureLastDownloadFailureDiagnostic savedErrNum, savedErrSrc, savedErrDesc, downloadCaptureInitialized

    On Error Resume Next
    m_downloadArmed = False
    m_downloadBatchActive = False
    EndSignalCarryoverDiagnostic methodName, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    If exclusiveCaptureEntered Then LeaveExclusiveCaptureOperation methodName
    On Error GoTo 0

    Err.Raise savedErrNum, savedErrSrc, savedErrDesc
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
    Set ExecuteOpenNewContextByXPath = ExecuteOpenNewContextCore( _
        ElementLocatorXPath, triggerXPath, searchTimeoutMs, timeoutMs, ownerContextId, _
        "ExecuteOpenNewContextByXPath")
End Function

' Opens a new top-level browsing context from a Shadow DOM trigger.
Public Function ExecuteShadowOpenNewContext(ByVal selectorPath As Variant, _
                                            Optional ByVal searchTimeoutMs As Long = 10000, _
                                            Optional ByVal timeoutMs As Long = 10000, _
                                            Optional ByVal ownerContextId As String = "", _
                                            Optional ByVal requireUnlocker As Boolean = False) As Dictionary
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowOpenNewContext"

    Set ExecuteShadowOpenNewContext = ExecuteOpenNewContextCore( _
        ElementLocatorShadowPath, selectorPath, searchTimeoutMs, timeoutMs, ownerContextId, _
        "ExecuteShadowOpenNewContext")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, _
                             "ExecuteShadowOpenNewContext"
End Function

' Runs the new-context capture transaction for either XPath or Shadow trigger locators.
Private Function ExecuteOpenNewContextCore(ByVal triggerLocatorKind As ElementLocatorKind, _
                                           ByVal triggerLocator As Variant, _
                                           ByVal searchTimeoutMs As Long, _
                                           ByVal timeoutMs As Long, _
                                           ByVal ownerContextId As String, _
                                           ByVal callerName As String) As Dictionary
    Dim methodName As String
    methodName = callerName

    Dim exclusiveCaptureEntered As Boolean
    EnterExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = True
    On Error GoTo Failed

    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 10000, 1)

    Dim ownerCtx As String
    ownerCtx = ResolveTargetContext(ownerContextId)

    Dim carryoverDiagnosticActive As Boolean
    Dim carryDomSeenBefore As Boolean
    Dim carryNetSeenBefore As Boolean
    Dim carryVisSeenBefore As Boolean
    carryoverDiagnosticActive = BeginSignalCarryoverDiagnostic( _
        methodName, ownerCtx, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore)

    ' Resolve before arming so lookup cannot consume the capture boundary.
    Dim triggerSharedId As String
    triggerSharedId = ResolveElementSharedId( _
        triggerLocatorKind, triggerLocator, ownerCtx, searchTimeoutMs, methodName)

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
    On Error GoTo Failed

    If clickErrNum <> 0 Then
        Err.Raise clickErrNum, clickErrSrc, clickErrDesc
    End If

    WaitForNewContextCapture vTimeout, methodName

    Dim result As Dictionary
    Set result = BuildNewContextResult()
    ResetNewContextCaptureState
    EndSignalCarryoverDiagnostic methodName, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore

    LeaveExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = False

    Set ExecuteOpenNewContextCore = result
    Exit Function

Failed:
    Dim captureErrNum As Long: captureErrNum = Err.Number
    Dim captureErrSrc As String: captureErrSrc = Err.source
    Dim captureErrDesc As String: captureErrDesc = Err.description

    On Error Resume Next
    ResetNewContextCaptureState
    EndSignalCarryoverDiagnostic methodName, carryoverDiagnosticActive, carryDomSeenBefore, carryNetSeenBefore, carryVisSeenBefore
    If exclusiveCaptureEntered Then LeaveExclusiveCaptureOperation methodName
    On Error GoTo 0

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

    ' Preserve dispatch failures before attempting to validate any response text. handleUserPrompt
    ' is state-changing, so an ambiguous transport outcome must never be replaced by a parse error.
    Dim rawRes As String
    Dim commandErr As Long
    Dim commandDesc As String
    Dim commandSrc As String

    On Error Resume Next
    Err.Clear
    rawRes = ExecuteBiDiCommand("browsingContext.handleUserPrompt", params, _
                                allowAmbiguousRetry:=False, allowContextRetarget:=(targetCtx = m_mainContextId))
    commandErr = Err.Number
    commandDesc = Err.description
    commandSrc = Err.source
    On Error GoTo 0

    If commandErr <> 0 Then Err.Raise commandErr, commandSrc, commandDesc

    ExecuteHandleUserPrompt = rawRes

    ' Absence of a prompt is the only expected protocol error. Classify it from the protocol
    ' error field itself rather than from localized/free-form Err.Description text.
    If IsExpectedNoUserPromptResponse(rawRes) Then Exit Function

    ValidateProtocolResponse rawRes, "browsingContext.handleUserPrompt"
End Function

Private Function IsExpectedNoUserPromptResponse(ByVal jsonRaw As String) As Boolean
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    Dim parseErr As Long: parseErr = Err.Number
    On Error GoTo 0

    If parseErr <> 0 Or json Is Nothing Then Exit Function
    If Not json.Exists("error") Then Exit Function
    If IsNull(json("error")) Or IsObject(json("error")) Then Exit Function

    Dim errType As String: errType = LCase$(CStr(json("error")))
    IsExpectedNoUserPromptResponse = _
        (StrComp(errType, "no such alert", vbBinaryCompare) = 0 Or _
         StrComp(errType, "no user prompt", vbBinaryCompare) = 0)
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

 ' A6 truthful HTTP-status contract
' --------------------------------
' Keep the existing public API name and return type. With requireHttpStatus=True (default),
' browsingContext.navigate itself has no HTTP status, so this method temporarily subscribes to
' network.responseStarted and returns only a real correlated main-document response.status.
' Non-HTTP navigation or missing correlation raises HttpStatusUnavailable. With
' requireHttpStatus=False, the same navigation/settlement path succeeds without requesting or
' fabricating an HTTP status and returns vbNullString.
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal requireHttpStatus As Boolean = True) As String

    Const METHOD_NAME As String = "ExecuteNavigateAndGetStatus"
    Const MAX_NAV_RETRIES As Long = 2

    ' N6: Treat this whole navigation transaction as exclusive. Even navigate-only mode must not
    ' re-enter while another one-shot capture owns class state, and it must not run inside an active
    ' navigation-status capture whose expected navigation ID could otherwise be rebound.
    Dim exclusiveCaptureEntered As Boolean
    EnterExclusiveCaptureOperation METHOD_NAME
    exclusiveCaptureEntered = True
    On Error GoTo Failed

    ' After acquiring exclusive ownership, a True flag can only be stale state from an abnormal
    ' host interruption; no live navigation-status transaction can coexist with this owner.
    If m_navHttpCaptureInProgress Then
        Err.Raise ERR_REENTRANT_CAPTURE, "BiDi:" & METHOD_NAME, _
                  "ReentrantNavigationStatusBlocked: navigation HTTP-status capture state is already active."
    End If

    If requireHttpStatus Then m_navHttpCaptureInProgress = True

    Dim temporaryResponseStartedSubscriptionId As String
    Dim vMinStable As Long
    Dim vMaxTimeout As Long
    Dim targetCtx As String
    vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    targetCtx = ResolveTargetContext(contextId)

    If requireHttpStatus Then
        ' N1: responseStarted is transaction-scoped unless another owner already subscribed it.
        ' Only a subscription created here is released during Cleanup.
        AcquireNavigationResponseStartedSubscription temporaryResponseStartedSubscriptionId

        ' responseStarted carries response.status and the navigation id. Drain stale events before
        ' arming so only events after the arm boundary can satisfy this transaction.
        ProcessRecorderQueue
        ArmNavigationHttpCapture targetCtx
    Else
        ' N2: navigate-only mode performs the existing navigation + optional SPA settlement without
        ' creating any responseStarted subscription and without fabricating an HTTP status.
        ResetNavigationHttpCaptureState
    End If

    Dim modeStr As String
    If waitNetworkIdle Then
        modeStr = "complete"
    Else
        modeStr = "none"
    End If

    Dim attempt As Long
    Dim rawNavigate As String
    Dim navErr As Long
    Dim navDesc As String
    Dim navSrc As String
    Dim navigationAccepted As Boolean
    Dim navigationAcceptedViaSupersession As Boolean

    For attempt = 0 To MAX_NAV_RETRIES
        On Error Resume Next
        Err.Clear
        rawNavigate = ExecuteBrowsingContextNavigate(targetUrl, modeStr, _
                                                     vMinStable, vMaxTimeout, targetCtx)
        navErr = Err.Number
        navDesc = Err.description
        navSrc = Err.source
        Err.Clear
        On Error GoTo Failed

        If navErr = ERR_CONSENSUS_HALT Then Err.Raise navErr, navSrc, navDesc

        If navErr = 0 Then
            navigationAccepted = True
            Exit For
        End If

        ' N7: Chromium can reject the original browsingContext.navigate after the page itself
        ' starts another navigation. Re-sending targetUrl would create a real duplicate navigation.
        ' Instead, prove a successor navigationStarted event in this context and adopt it.
        If IsConcurrentNavigationSupersession(navErr, navSrc, navDesc) Then
            If TryAdoptConcurrentNavigation(targetCtx, targetUrl, waitNetworkIdle, vMinStable, vMaxTimeout) Then
                navigationAccepted = True
                navigationAcceptedViaSupersession = True
                Exit For
            End If
        End If

        ' Preserve the existing ordinary-navigation retry policy. A retry happens only for the
        ' explicitly classified Chromium network errors; ambiguous/post-commit failures stop.
        If IsRetryableNavFailure(navDesc) And attempt < MAX_NAV_RETRIES Then
            WaitHeartbeat 1000
        Else
            Exit For
        End If
    Next attempt

    If Not navigationAccepted Then
        If navErr = 0 Then navErr = ERR_ACTION_FAILED
        If Len(navSrc) = 0 Then navSrc = "BiDi:" & METHOD_NAME
        If Len(navDesc) = 0 Then navDesc = "Navigation command failed."
        Err.Raise navErr, navSrc, navDesc
    End If

    Dim resultText As String

    If Not requireHttpStatus Then
        resultText = vbNullString
        GoTo succeeded
    End If

    ' ExecuteBrowsingContextNavigate binds the navigation id before any optional settlement wait.
    ' For waitMode=none, drain now so buffered responseStarted events can satisfy the capture.
    ProcessRecorderQueue

    If m_navHttpCaptureStatusKnown Then
        resultText = CStr(m_navHttpCaptureStatus)
        GoTo succeeded
    End If

    Dim statusSource As String
    Dim lossDelta As Long
    lossDelta = NavigationHttpCaptureLossSinceArm()

    If Len(m_navHttpCaptureExpectedNavigationId) = 0 Then
        statusSource = "no-navigation-id"
    ElseIf lossDelta > 0 Then
        statusSource = "correlated-response-not-observed; async-event-loss-present"
    Else
        statusSource = "no-correlated-http-response"
    End If

    Dim navigationOutcomeText As String
    If navigationAcceptedViaSupersession Then
        navigationOutcomeText = "the original browsingContext.navigate was superseded and a later navigation was adopted"
    Else
        navigationOutcomeText = "browsingContext.navigate succeeded"
    End If

    Err.Raise ERR_ACTION_FAILED, "BiDi:" & METHOD_NAME, _
              "HttpStatusUnavailable: " & navigationOutcomeText & _
              ", but no correlated HTTP main-document response status was observed. Source=" & statusSource & _
              "; Navigation=" & m_navHttpCaptureExpectedNavigationId & _
              "; URL=" & m_navHttpCaptureExpectedUrl & "."

succeeded:
    ExecuteNavigateAndGetStatus = resultText
    GoTo Cleanup

Failed:
    Dim savedErrNum As Long
    Dim savedErrSrc As String
    Dim savedErrDesc As String
    savedErrNum = Err.Number
    savedErrSrc = Err.source
    savedErrDesc = Err.description
    Resume Cleanup

Cleanup:
    ' Cleanup must never replace an already-determined navigation outcome. In particular,
    ' temporary unsubscribe, capture reset, or deferred-reset logging may fail secondarily.
    On Error Resume Next

    ReleaseNavigationResponseStartedSubscriptionBestEffort temporaryResponseStartedSubscriptionId
    ResetNavigationSupersessionWatch
    ResetNavigationHttpCaptureState
    m_navHttpCaptureInProgress = False

    If exclusiveCaptureEntered Then
        LeaveExclusiveCaptureOperation METHOD_NAME
        exclusiveCaptureEntered = False
    End If

    Err.Clear
    On Error GoTo 0

    If savedErrNum <> 0 Then
        Err.Raise savedErrNum, savedErrSrc, savedErrDesc
    End If
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

    ' History traversal is state-changing. A command-correlation timeout, disconnect, or
    ' context-loss response can mean that the browser already accepted the traversal, so never resend it,
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


' Public compatibility facade. Basic operations delegate to the unified element core.
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

    Dim noValue As Variant
    ExecuteIsElementVisible = CBool(ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionVisibilityCheck, False, noValue, _
        vSearchTimeout, True, waitNetworkIdle, _
        SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS), _
        SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS), _
        False, contextId, "ExecuteIsElementVisible"))
End Function

Public Function ExecuteGetTextByXPath(ByVal xpath As String, _
                                      Optional ByVal searchTimeoutMs As Long = 10000, _
                                      Optional ByVal contextId As String = "") As String
    Dim noValue As Variant
    ExecuteGetTextByXPath = CStr(ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionGetText, False, noValue, _
        searchTimeoutMs, False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, _
        False, contextId, "ExecuteGetTextByXPath"))
End Function

Public Function ExecuteGetAttributeByXPath(ByVal xpath As String, _
                                           ByVal attributeName As String, _
                                           Optional ByVal searchTimeoutMs As Long = 10000, _
                                           Optional ByVal contextId As String = "") As Variant
    ExecuteGetAttributeByXPath = ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionGetAttribute, True, attributeName, _
        searchTimeoutMs, False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, _
        False, contextId, "ExecuteGetAttributeByXPath")
End Function

Public Function ExecuteGetPropertyByXPath(ByVal xpath As String, _
                                          ByVal propertyName As String, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal contextId As String = "") As Variant
    ExecuteGetPropertyByXPath = ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionGetProperty, True, propertyName, _
        searchTimeoutMs, False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, _
        False, contextId, "ExecuteGetPropertyByXPath")
End Function

Public Function ExecuteGetOuterHtmlByXPath(ByVal xpath As String, _
                                           Optional ByVal searchTimeoutMs As Long = 10000, _
                                           Optional ByVal contextId As String = "") As String
    Dim noValue As Variant
    ExecuteGetOuterHtmlByXPath = CStr(ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionGetOuterHtml, False, noValue, _
        searchTimeoutMs, False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, _
        False, contextId, "ExecuteGetOuterHtmlByXPath"))
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
               errDesc = "STALE_ELEMENT_PRE_INPUT" Or _
               errDesc = "Element is no longer connected." Then
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

Private Function RequireBooleanResult(ByVal value As Variant, _
                                      ByVal callerName As String, _
                                      ByVal resultName As String) As Boolean
    If IsEmpty(value) Or IsNull(value) Or IsObject(value) Or VarType(value) <> vbBoolean Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "The " & resultName & " result was not a Boolean value."
    End If
    RequireBooleanResult = CBool(value)
End Function

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

Private Sub AcquireNavigationResponseStartedSubscription(ByRef ownedSubscriptionId As String)
    Const SUB_KEY As String = "network.responseStarted"
    ownedSubscriptionId = vbNullString

    If m_activeSubscriptions Is Nothing Then Set m_activeSubscriptions = New Dictionary

    ' N1: Borrow an existing global responseStarted subscription without taking ownership of it.
    ' Only a subscription created by this navigation transaction may be removed by its cleanup.
    If m_activeSubscriptions.Exists(SUB_KEY) Then Exit Sub

    Call Me.ExecuteSessionSubscribe(Array(SUB_KEY), Silent:=True)

    If Not m_activeSubscriptions.Exists(SUB_KEY) Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:AcquireNavigationResponseStartedSubscription", _
                  "session.subscribe succeeded but the responseStarted subscription was not " & _
                  "recorded in the local subscription registry."
    End If

    ownedSubscriptionId = CStr(m_activeSubscriptions(SUB_KEY))
    If Len(ownedSubscriptionId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:AcquireNavigationResponseStartedSubscription", _
                  "The responseStarted subscription registry entry has an empty subscription ID."
    End If
End Sub

Private Sub ReleaseNavigationResponseStartedSubscriptionBestEffort( _
    ByVal ownedSubscriptionId As String)

    Const SUB_KEY As String = "network.responseStarted"
    If Len(ownedSubscriptionId) = 0 Then Exit Sub
    If m_activeSubscriptions Is Nothing Then Exit Sub
    If Not m_activeSubscriptions.Exists(SUB_KEY) Then Exit Sub

    ' If registry ownership changed, never unsubscribe a subscription this transaction no longer owns.
    If StrComp(CStr(m_activeSubscriptions(SUB_KEY)), ownedSubscriptionId, vbBinaryCompare) <> 0 Then
        Exit Sub
    End If

    If socket_ Is Nothing Then Exit Sub
    If socket_.state <> StateConnected Then Exit Sub

    On Error GoTo Failed

    Dim ids(0 To 0) As String
    ids(0) = ownedSubscriptionId

    Dim params As New Dictionary
    params.Add "subscriptions", ids

    Dim raw As String
    raw = ExecuteBiDiCommand( _
              "session.unsubscribe", _
              params, _
              True, _
              allowAmbiguousRetry:=False, _
              allowRecoveryRetry:=False)

    Call ValidateProtocolResponse(raw, "session.unsubscribe")

    ' Remove the local key only after remote unsubscribe was positively acknowledged.
    If m_activeSubscriptions.Exists(SUB_KEY) Then
        If StrComp(CStr(m_activeSubscriptions(SUB_KEY)), ownedSubscriptionId, vbBinaryCompare) = 0 Then
            m_activeSubscriptions.Remove SUB_KEY
        End If
    End If

    If m_debugMode Then
        Debug.Print "BiDi: Temporary network.responseStarted subscription released."
    End If
    Exit Sub

Failed:
    ' Navigation outcome/status remains primary. Keep the registry entry so ClearState can retry
    ' cleanup later instead of pretending the remote subscription is gone.
    If m_debugMode Then
        Debug.Print "BiDi Warning: temporary network.responseStarted unsubscribe was not confirmed. " & _
                    "Err " & CStr(Err.Number) & ": " & Err.description
    End If
    Err.Clear
End Sub

' browsingContext.navigate waits to interactive at the protocol layer. waitMode controls
' only the additional bounded wrapper-side settlement after the command returns. A5 caps command
' correlation ownership independently from the browser-side operation timeout.
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

    ' N7: arm only after target resolution and immediately before dispatch. ResolveTargetContext
    ' may drain recorder events, so arming any earlier would let a stale navigationStarted event
    ' cross the per-attempt boundary. No Office yield occurs between this drain/arm and send.
    DrainNonBlockingRecorderQueue
    ArmNavigationSupersessionWatch targetCtx

    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params, _
                                allowAmbiguousRetry:=False, _
                                allowContextRetarget:=allowRetarget, _
                                commandTimeoutMs:=ResolveNavigationTransportTimeout(maxTimeoutMs))
    Dim navJson As Object
    Set navJson = ValidateProtocolResponse(strRes, "browsingContext.navigate")

    ' Bind A6 HTTP-status capture before any settlement wait drains buffered network events.
    ' The navigation id is the correlation key defined by WebDriver BiDi network events.
    BindNavigationHttpCapture targetCtx, _
                              GetPathString(navJson, "result.navigation"), _
                              GetPathString(navJson, "result.url")

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

Private Sub ResetNavigationHttpCaptureState()
    m_navHttpCaptureArmed = False
    m_navHttpCaptureOwnerContextId = vbNullString
    m_navHttpCaptureExpectedNavigationId = vbNullString
    m_navHttpCaptureExpectedUrl = vbNullString
    m_navHttpCaptureStatusKnown = False
    m_navHttpCaptureStatus = 0
    m_navHttpCaptureStatusText = vbNullString
    m_navHttpCaptureResponseUrl = vbNullString
    m_navHttpCaptureRedirectCount = 0
    m_navHttpCaptureLossBaseline = 0
End Sub

Private Sub ArmNavigationHttpCapture(ByVal contextId As String)
    ResetNavigationHttpCaptureState
    m_navHttpCaptureOwnerContextId = contextId
    m_navHttpCaptureLossBaseline = CurrentTotalAsyncEventLossCount()
    m_navHttpCaptureArmed = True
End Sub

Private Sub BindNavigationHttpCapture(ByVal contextId As String, _
                                      ByVal navigationId As String, _
                                      ByVal resultUrl As String)
    If Not m_navHttpCaptureArmed Then Exit Sub
    If StrComp(contextId, m_navHttpCaptureOwnerContextId, vbBinaryCompare) <> 0 Then Exit Sub

    ' A client-side/meta navigation that supersedes the command has a different navigation id.
    ' Never carry a status captured for the earlier document into the successor transaction.
    If Len(navigationId) > 0 And _
       StrComp(navigationId, m_navHttpCaptureExpectedNavigationId, vbBinaryCompare) <> 0 Then
        m_navHttpCaptureStatusKnown = False
        m_navHttpCaptureStatus = 0
        m_navHttpCaptureStatusText = vbNullString
        m_navHttpCaptureResponseUrl = vbNullString
        m_navHttpCaptureRedirectCount = 0
    End If

    m_navHttpCaptureExpectedNavigationId = navigationId
    m_navHttpCaptureExpectedUrl = resultUrl
End Sub

Private Sub ResetNavigationSupersessionWatch()
    m_navSupersessionWatchArmed = False
    m_navSupersessionOwnerContextId = vbNullString
    m_navSupersessionObservedCount = 0
    m_navSupersessionNavigationId = vbNullString
    m_navSupersessionUrl = vbNullString
End Sub

Private Sub ArmNavigationSupersessionWatch(ByVal contextId As String)
    ResetNavigationSupersessionWatch
    m_navSupersessionOwnerContextId = contextId
    m_navSupersessionWatchArmed = True
End Sub

Private Sub ObserveNavigationStartedForActiveNavigate(ByVal json As Object)
    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If Len(eventCtx) = 0 Then Exit Sub

    Dim navigationId As String
    Dim navigationUrl As String
    navigationId = GetPathString(json, "params.navigation")
    navigationUrl = GetPathString(json, "params.url")

    If m_navSupersessionWatchArmed Then
        If StrComp(eventCtx, m_navSupersessionOwnerContextId, vbBinaryCompare) = 0 Then
            If m_navSupersessionObservedCount < 2147483647 Then _
                m_navSupersessionObservedCount = m_navSupersessionObservedCount + 1
            m_navSupersessionNavigationId = navigationId
            m_navSupersessionUrl = navigationUrl
        End If
    End If

    ' Bind from navigationStarted as well as the command result. This is essential when Chromium
    ' returns an error for the command because a later page-driven navigation superseded it.
    If m_navHttpCaptureArmed Then
        If StrComp(eventCtx, m_navHttpCaptureOwnerContextId, vbBinaryCompare) = 0 And _
           Len(navigationId) > 0 Then
            BindNavigationHttpCapture eventCtx, navigationId, navigationUrl
        End If
    End If
End Sub

Private Function IsConcurrentNavigationSupersession(ByVal errNum As Long, _
                                                    ByVal errSrc As String, _
                                                    ByVal errDesc As String) As Boolean
    If errNum <> ERR_BIDI_PROTOCOL Then Exit Function
    If InStr(1, errSrc, "browsingContext.navigate", vbTextCompare) = 0 Then Exit Function

    Dim d As String
    d = LCase$(errDesc)
    IsConcurrentNavigationSupersession = _
        (InStr(d, "navigation canceled by concurrent navigation") > 0 Or _
         InStr(d, "navigation cancelled by concurrent navigation") > 0)
End Function

Private Function TryAdoptConcurrentNavigation(ByVal targetCtx As String, _
                                              ByVal requestedUrl As String, _
                                              ByVal waitNetworkIdle As Boolean, _
                                              ByVal minStableMs As Long, _
                                              ByVal maxTimeoutMs As Long) As Boolean
    Const ADOPTION_EVENT_WAIT_MS As Long = 2000

    Dim startedAt As Currency
    startedAt = socket_.NowTicks()

    Do
        ' Non-blocking drain may pump the socket but never Office/VBA messages. Event dispatch
        ' updates the per-attempt watch and, when armed, the HTTP-status navigation binding.
        DrainNonBlockingRecorderQueue

        If m_navSupersessionWatchArmed And _
           StrComp(m_navSupersessionOwnerContextId, targetCtx, vbBinaryCompare) = 0 Then
            ' Two navigationStarted events prove that the command navigation was superseded.
            ' One event is also sufficient when its URL is visibly different from targetUrl; this
            ' covers implementations that do not surface the canceled first navigation event.
            If m_navSupersessionObservedCount >= 2 Or _
               (m_navSupersessionObservedCount >= 1 And Len(m_navSupersessionUrl) > 0 And _
                StrComp(Trim$(m_navSupersessionUrl), Trim$(requestedUrl), vbBinaryCompare) <> 0) Then
                If m_debugMode Then
                    Debug.Print "BiDi: Adopted concurrent successor navigation. id=" & _
                                m_navSupersessionNavigationId & " | url=" & m_navSupersessionUrl
                End If

                If waitNetworkIdle Then
                    SettleAfterNavigation targetCtx, minStableMs, maxTimeoutMs
                End If

                ' Settlement may have drained the final responseStarted event. Drain once more
                ' for waitNetworkIdle=False and for responses arriving at the settlement tail.
                DrainNonBlockingRecorderQueue
                TryAdoptConcurrentNavigation = True
                Exit Function
            End If
        End If

        If socket_.elapsedMs(startedAt) >= ADOPTION_EVENT_WAIT_MS Then Exit Do
        socket_.YieldExecution 10, False
    Loop
End Function

Private Function NavigationHttpCaptureLossSinceArm() As Long
    If Not m_navHttpCaptureArmed Then Exit Function

    Dim delta As Long
    delta = CurrentTotalAsyncEventLossCount() - m_navHttpCaptureLossBaseline
    If delta < 0 Then delta = 0
    NavigationHttpCaptureLossSinceArm = delta
End Function

Private Function IsHttpNavigationUrl(ByVal value As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Trim$(value))

    IsHttpNavigationUrl = _
        (Left$(normalized, 7) = "http://" Or Left$(normalized, 8) = "https://")
End Function

Private Sub HandleNavigationResponseStartedEvent(ByVal json As Object)
    If Not m_navHttpCaptureArmed Then Exit Sub
    If Len(m_navHttpCaptureExpectedNavigationId) = 0 Then Exit Sub

    ' Chromium may emit network.responseStarted with a synthetic status (commonly 200) for
    ' non-HTTP main documents such as data:, about:, file:, or blob:. ExecuteNavigateAndGetStatus
    ' promises an HTTP status, so a non-HTTP final navigation can never satisfy this capture.
    If Not IsHttpNavigationUrl(m_navHttpCaptureExpectedUrl) Then Exit Sub

    Dim eventCtx As String
    eventCtx = GetPathString(json, "params.context")
    If StrComp(eventCtx, m_navHttpCaptureOwnerContextId, vbBinaryCompare) <> 0 Then Exit Sub

    Dim navigationId As String
    navigationId = GetPathString(json, "params.navigation")
    If StrComp(navigationId, m_navHttpCaptureExpectedNavigationId, vbBinaryCompare) <> 0 Then Exit Sub

    Dim responseObj As Object
    Set responseObj = GetPathObject(json, "params.response")
    If responseObj Is Nothing Then Exit Sub

    Dim responseUrl As String
    If responseObj.Exists("url") Then
        If Not IsObject(responseObj("url")) And Not IsNull(responseObj("url")) Then
            responseUrl = CStr(responseObj("url"))
        End If
    End If

    ' The network event itself must also identify an HTTP(S) response. This prevents a synthetic
    ' response.status attached to a non-HTTP scheme from being exposed as an HTTP status code.
    If Not IsHttpNavigationUrl(responseUrl) Then Exit Sub

    If Not responseObj.Exists("status") Then Exit Sub
    If IsObject(responseObj("status")) Or IsNull(responseObj("status")) Then Exit Sub
    If Not IsNumeric(responseObj("status")) Then Exit Sub

    Dim redirectCount As Long
    Dim redirectText As String
    redirectText = GetPathString(json, "params.redirectCount")
    If IsNumeric(redirectText) Then redirectCount = CLng(redirectText)

    ' A redirect chain reuses the same navigation id. Keep the greatest redirectCount, i.e. the
    ' latest main-document response in that navigation transaction.
    If m_navHttpCaptureStatusKnown Then
        If redirectCount < m_navHttpCaptureRedirectCount Then Exit Sub
    End If

    m_navHttpCaptureStatusKnown = True
    m_navHttpCaptureStatus = CLng(responseObj("status"))
    m_navHttpCaptureRedirectCount = redirectCount

    If responseObj.Exists("statusText") Then
        If Not IsObject(responseObj("statusText")) And Not IsNull(responseObj("statusText")) Then
            m_navHttpCaptureStatusText = CStr(responseObj("statusText"))
        End If
    End If

    m_navHttpCaptureResponseUrl = responseUrl
End Sub

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

Private Function ExecuteInputValueCore(ByVal locatorKind As ElementLocatorKind, _
                                       ByVal locator As Variant, _
                                       ByVal valueToSet As String, _
                                       ByVal searchTimeoutMs As Long, _
                                       ByVal waitNetworkIdle As Boolean, _
                                       ByVal minStableMs As Long, _
                                       ByVal maxTimeoutMs As Long, _
                                       ByVal contextId As String, _
                                       ByVal useKeyEvents As Boolean, _
                                       ByVal useNativeKeys As Boolean, _
                                       ByVal callerName As String) As String
    ResetLastSyncState
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveElementSharedId( _
        locatorKind, locator, targetCtx, searchTimeoutMs, callerName)

    Dim args As Collection: Set args = NewElementArgs(sharedId, valueToSet)
    If useNativeKeys Then
        Dim nativeKeyActions As Collection
        Set nativeKeyActions = BuildNativeKeyActions(valueToSet, callerName)

        Dim focusArgs As Collection: Set focusArgs = NewElementArgs(sharedId)
        Call ExecuteBaseAction(targetCtx, "click", focusArgs, False, minStableMs, _
                               True, DEFAULT_STALE_RETRIES, maxTimeoutMs, False, _
                               locatorKind, locator, searchTimeoutMs, callerName)
        ExecuteInputValueCore = PerformNativeKeyInput( _
            targetCtx, nativeKeyActions, maxTimeoutMs, waitNetworkIdle, minStableMs)
        If m_debugMode Then Debug.Print "BiDi: Action 'native_keys' completed."
    ElseIf Not useKeyEvents Then
        Dim inputResult As String
        inputResult = ExecuteBaseAction( _
            targetCtx, "input", args, waitNetworkIdle, minStableMs, True, _
            DEFAULT_STALE_RETRIES, maxTimeoutMs, False, locatorKind, locator, _
            searchTimeoutMs, callerName)

        LogInputValueMismatchBestEffort locatorKind, locator, valueToSet, _
                                        searchTimeoutMs, targetCtx, callerName
        ExecuteInputValueCore = inputResult
    Else
        Dim keySafetyMs As Long: keySafetyMs = maxTimeoutMs
        If keySafetyMs < 20000 Then keySafetyMs = 20000

        Dim keyResult As String
        keyResult = ExecuteBaseAction( _
            targetCtx, "input_keys", args, False, minStableMs, False, _
            DEFAULT_STALE_RETRIES, keySafetyMs, False, locatorKind, locator, _
            searchTimeoutMs, callerName)

        Call TrySyncPhase(targetCtx, waitNetworkIdle, _
                          SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS), _
                          SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS), False)
        ExecuteInputValueCore = keyResult
        If m_debugMode Then Debug.Print "BiDi: Action 'input_keys' completed."
    End If
End Function

Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal checkUrl As Boolean = False) As String
    Dim noValue As Variant
    ExecuteClickByXPath = CStr(ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, ElementActionClick, False, noValue, _
        searchTimeoutMs, True, waitNetworkIdle, minStableMs, maxTimeoutMs, _
        checkUrl, contextId, "ExecuteClickByXPath", True))
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
    ExecuteInputValueByXPath = ExecuteInputValueCore( _
        ElementLocatorXPath, xpath, valueToSet, searchTimeoutMs, waitNetworkIdle, _
        minStableMs, maxTimeoutMs, contextId, useKeyEvents, useNativeKeys, _
        "ExecuteInputValueByXPath")
End Function

Public Function ExecuteSetFilesByXPath(ByVal xpath As String, _
                                       ByVal filePaths As Variant, _
                                       Optional ByVal searchTimeoutMs As Long = 10000, _
                                       Optional ByVal waitNetworkIdle As Boolean = True, _
                                       Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                       Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                       Optional ByVal contextId As String = "") As String
    ExecuteSetFilesByXPath = ExecuteSetFilesCore( _
        ElementLocatorXPath, xpath, filePaths, searchTimeoutMs, waitNetworkIdle, _
        minStableMs, maxTimeoutMs, contextId, "ExecuteSetFilesByXPath")
End Function

' Shares input.setFiles across locator types.
Private Function ExecuteSetFilesCore(ByVal locatorKind As ElementLocatorKind, _
                                     ByVal locator As Variant, _
                                     ByVal filePaths As Variant, _
                                     ByVal searchTimeoutMs As Long, _
                                     ByVal waitNetworkIdle As Boolean, _
                                     ByVal minStableMs As Long, _
                                     ByVal maxTimeoutMs As Long, _
                                     ByVal contextId As String, _
                                     ByVal callerName As String) As String
    ResetLastSyncState

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    Dim sharedId As String
    sharedId = ResolveElementSharedId( _
        locatorKind, locator, targetCtx, searchTimeoutMs, callerName)

    Dim paths As Variant
    paths = NormalizeFilePaths(filePaths)

    ' Validate ownership before input.setFiles changes browser state.
    If waitNetworkIdle Then
        ValidateSignalWaitContext targetCtx, callerName
    End If

    ExecuteSetFilesCore = PerformSetFiles( _
        targetCtx, sharedId, paths, maxTimeoutMs, waitNetworkIdle, minStableMs)

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
' The CDP session is resolved from the SAME BiDi browsing context as the trigger locator and is
' held for cleanup so iframe/popup contexts do not silently fall back to the main-context session.
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
    ExecuteSetFileSelectionViaDialog = ExecuteSetFileSelectionViaDialogCore( _
        ElementLocatorXPath, triggerXPath, filePaths, searchTimeoutMs, fileDialogTimeoutMs, _
        minStableMs, maxTimeoutMs, contextId, "ExecuteSetFileSelectionViaDialog")
End Function

' Opens a native file chooser from a Shadow DOM trigger and applies the selected files.
Public Function ExecuteShadowSetFileSelectionViaDialog(ByVal selectorPath As Variant, _
                                                        ByVal filePaths As Variant, _
                                                        Optional ByVal searchTimeoutMs As Long = 10000, _
                                                        Optional ByVal fileDialogTimeoutMs As Long = 10000, _
                                                        Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                                        Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                                        Optional ByVal contextId As String = "", _
                                                        Optional ByVal requireUnlocker As Boolean = False) As String
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowSetFileSelectionViaDialog"

    ExecuteShadowSetFileSelectionViaDialog = ExecuteSetFileSelectionViaDialogCore( _
        ElementLocatorShadowPath, selectorPath, filePaths, searchTimeoutMs, fileDialogTimeoutMs, _
        minStableMs, maxTimeoutMs, contextId, "ExecuteShadowSetFileSelectionViaDialog")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, _
                             "ExecuteShadowSetFileSelectionViaDialog"
End Function

' Runs the native file-chooser transaction for either XPath or Shadow locators.
Private Function ExecuteSetFileSelectionViaDialogCore(ByVal triggerLocatorKind As ElementLocatorKind, _
                                                       ByVal triggerLocator As Variant, _
                                                       ByVal filePaths As Variant, _
                                                       ByVal searchTimeoutMs As Long, _
                                                       ByVal fileDialogTimeoutMs As Long, _
                                                       ByVal minStableMs As Long, _
                                                       ByVal maxTimeoutMs As Long, _
                                                       ByVal contextId As String, _
                                                       ByVal callerName As String) As String
    Dim methodName As String
    methodName = callerName

    ' Acquire before any chooser/capture state mutation.
    Dim exclusiveCaptureEntered As Boolean
    EnterExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = True
    On Error GoTo Failed

    ResetLastSyncState

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    If AnySignalArmed() Then
        If Len(m_signalOwnerContextId) = 0 Or _
           StrComp(targetCtx, m_signalOwnerContextId, vbBinaryCompare) <> 0 Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                      "CompletionSignalContextMismatch: The active completion signal does not belong " & _
                      "to the file-selection context." & vbCrLf & _
                      "Signal owner: " & m_signalOwnerContextId & vbCrLf & _
                      "Target: " & targetCtx
        End If
    End If

    If m_fileChooserInterceptMayBeActive Then
        If Not ReleaseTrackedFileChooserInterceptionBestEffort(True) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
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
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "NoFilesSupplied: No file paths were supplied for file selection."
    End If

    Dim vDialogTimeout As Long
    Dim vMinStable As Long
    Dim vMaxTimeout As Long
    vDialogTimeout = SanitizeMs(fileDialogTimeoutMs, "fileDialogTimeoutMs", 10000)
    vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    Me.ExecuteSessionSubscribe Array("input.fileDialogOpened"), , True

    ' Resolve before arming so locator work cannot consume an unrelated chooser event.
    Dim triggerSharedId As String
    triggerSharedId = ResolveElementSharedId( _
        triggerLocatorKind, triggerLocator, targetCtx, searchTimeoutMs, callerName)

    Dim fileChooserCdpSessionId As String
    fileChooserCdpSessionId = GetCdpSessionIdForContext(targetCtx)
    BeginNativeFileChooserInterception targetCtx, fileChooserCdpSessionId

    ProcessRecorderQueue
    ArmFileDialogCapture targetCtx

    Dim clickArgs As Collection
    Set clickArgs = NewElementArgs(triggerSharedId)

    ' The chooser trigger requires real user activation.
    Dim clickResult As String
    clickResult = PerformTrustedClick(targetCtx, clickArgs, vMaxTimeout, False)

    If Not WaitForFileDialogCapture(vDialogTimeout) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "FileDialogTimeout: input.fileDialogOpened was not observed within " & _
                  vDialogTimeout & "ms."
    End If

    If Not m_fileDlgMultipleKnown Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & callerName, _
                  "FileDialogMissingMultiple: input.fileDialogOpened did not include the " & _
                  "required 'multiple' field."
    End If

    If Not m_fileDlgHasElement Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "FileDialogHasNoInputElement: input.fileDialogOpened had no HTML " & _
                  "file-input element. Elementless picker paths are outside this API contract."
    End If

    If pathCount > 1 And Not m_fileDlgMultiple Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "MultipleFilesNotAllowed: the opened file input does not allow multiple files."
    End If

    Dim capturedSharedId As String
    capturedSharedId = m_fileDlgSharedId

    Dim setFilesResult As String
    setFilesResult = PerformSetFiles(targetCtx, capturedSharedId, paths, _
                                     vMaxTimeout, False, vMinStable)

    Dim selectionApplied As Boolean
    selectionApplied = True

    ResetFileDialogCaptureState

    If Not ReleaseTrackedFileChooserInterceptionBestEffort(True) Then
        Me.DisarmSignal
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "PostSelectionInterceptionCleanupFailed: input.setFiles succeeded, but native " & _
                  "file-chooser interception could not be confirmed disabled. The selection may " & _
                  "already have triggered application work; do not retry the whole operation. " & _
                  "Call Shutdown or reconnect before further file-chooser use."
    End If

    RunRequiredSpaSync targetCtx, vMinStable, vMaxTimeout
    ExecuteSetFileSelectionViaDialogCore = setFilesResult

    LeaveExclusiveCaptureOperation methodName
    exclusiveCaptureEntered = False

    If m_debugMode Then Debug.Print "BiDi: Action 'set_file_selection_via_dialog' completed."
    Exit Function

Failed:
    Dim primaryErrNum As Long
    Dim primaryErrSrc As String
    Dim primaryErrDesc As String
    primaryErrNum = Err.Number
    primaryErrSrc = Err.source
    primaryErrDesc = Err.description

    Dim cleanupConfirmed As Boolean
    On Error Resume Next
    ResetFileDialogCaptureState
    cleanupConfirmed = ReleaseTrackedFileChooserInterceptionBestEffort(True)
    Me.DisarmSignal
    If exclusiveCaptureEntered Then LeaveExclusiveCaptureOperation methodName
    On Error GoTo 0

    If primaryErrNum = 0 Then primaryErrNum = ERR_ACTION_FAILED

    If selectionApplied And InStr(1, primaryErrDesc, "PostSelectionInterceptionCleanupFailed:", vbBinaryCompare) = 0 Then
        Dim underlyingSrc As String
        underlyingSrc = primaryErrSrc
        If Len(underlyingSrc) = 0 Then underlyingSrc = "(unknown)"

        primaryErrSrc = "BiDi:" & callerName
        primaryErrDesc = _
            "PostSelectionCompletionFailed: input.setFiles succeeded before the completion " & _
            "wait failed. File selection may already have triggered application work; do not " & _
            "blindly retry the whole file-selection operation." & vbCrLf & _
            "Underlying source: " & underlyingSrc & vbCrLf & _
            "Underlying error: " & primaryErrDesc
    ElseIf Len(primaryErrSrc) = 0 Then
        primaryErrSrc = "BiDi:" & callerName
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

        Dim droppedSinceArm As Long
        droppedSinceArm = FileDialogDroppedEventsSinceArm()
        If droppedSinceArm > 0 Then
            LogSignalLine "[FILE-DIALOG-LOSS] ctx=" & m_fileDlgOwnerContextId & _
                          " | droppedSinceArm=" & CStr(droppedSinceArm)
            m_fileDlgArmed = False
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ExecuteSetFileSelectionViaDialog", _
                      "Async event loss occurred after the file-dialog arm boundary; reliable " & _
                      "file-input capture cannot be determined. AsyncEventsDroppedSinceArm=" & _
                      CStr(droppedSinceArm) & "."
        End If

        If m_fileDlgSeen Then
            WaitForFileDialogCapture = True
            Exit Function
        End If

        If IsTimeout(startTick, CDbl(timeoutMs)) Then
            LogSignalLine "[FILE-DIALOG-TIMEOUT] ctx=" & m_fileDlgOwnerContextId & _
                          " | timeoutMs=" & CStr(timeoutMs)
            Exit Do
        End If
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
    If json Is Nothing Then
        failureDetail = "Malformed or empty browser.setDownloadBehavior response during residual cleanup."
        GoTo Failed
    End If

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
        If Err.Number <> 0 Then
            failureDetail = "Error #" & CStr(Err.Number) & ": " & Err.description
        Else
            failureDetail = "Residual cleanup failed without a VBA error code or protocol diagnostic."
        End If
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

Private Function BuildNativeKeyActions(ByVal text As String, _
                                       ByVal callerName As String) As Collection

    Dim keyActions As New Collection
    Dim i As Long
    Dim ch As String
    Dim codeUnit As Long
    Dim nextCodeUnit As Long
    Dim skipNext As Boolean

    For i = 1 To Len(text)
        If skipNext Then
            skipNext = False
        Else
            ch = Mid$(text, i, 1)

            ' AscW returns a signed Integer. Normalize to an unsigned UTF-16 code unit.
            codeUnit = CLng(AscW(ch)) And &HFFFF&

            If codeUnit >= &HD800& And codeUnit <= &HDBFF& Then
                If i >= Len(text) Then
                    Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                        "InvalidNativeKeyText: malformed UTF-16 at code-unit position " & _
                        CStr(i) & "; high surrogate U+" & Right$("0000" & Hex$(codeUnit), 4) & _
                        " is not followed by a low surrogate."
                End If

                nextCodeUnit = CLng(AscW(Mid$(text, i + 1, 1))) And &HFFFF&
                If nextCodeUnit < &HDC00& Or nextCodeUnit > &HDFFF& Then
                    Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                        "InvalidNativeKeyText: malformed UTF-16 at code-unit position " & _
                        CStr(i) & "; high surrogate U+" & Right$("0000" & Hex$(codeUnit), 4) & _
                        " is followed by U+" & Right$("0000" & Hex$(nextCodeUnit), 4) & _
                        " instead of a low surrogate."
                End If

                ' Send one complete supplementary scalar as one WebDriver key value.
                AppendKeyStroke keyActions, Mid$(text, i, 2)
                skipNext = True

            ElseIf codeUnit >= &HDC00& And codeUnit <= &HDFFF& Then
                Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                    "InvalidNativeKeyText: malformed UTF-16 at code-unit position " & _
                    CStr(i) & "; lone low surrogate U+" & Right$("0000" & Hex$(codeUnit), 4) & "."

            Else
                Select Case codeUnit
                    Case 13 ' CR -> Enter (also swallow following LF if present: CRLF)
                        AppendKeyStroke keyActions, ChrW(&HE007&) ' U+E007 = Enter (W3C key code)
                        If i < Len(text) Then
                            nextCodeUnit = CLng(AscW(Mid$(text, i + 1, 1))) And &HFFFF&
                            If nextCodeUnit = 10 Then skipNext = True
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
        End If
    Next i

    Set BuildNativeKeyActions = keyActions
End Function

Private Function PerformNativeKeyInput(ByVal contextId As String, _
                                       ByVal keyActions As Collection, _
                                       Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                       Optional ByVal waitNetworkIdle As Boolean = True, _
                                       Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String

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

    ' A hit-test failure is not equivalent to a genuine hit=0 result. The strict helper raises
    ' on transport/protocol/parse/contract failures, so synthetic fallback is considered only
    ' after a successfully evaluated hit=0 verdict.
    Dim onTarget As Boolean
    onTarget = RequireTrustedHitTestResult(contextId, args)

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

Private Function RequireTrustedHitTestResult(ByVal contextId As String, _
                                             ByVal args As Collection) As Boolean
    Dim rawRes As String
    rawRes = CallScript(HitTestJs(), args, contextId, True, False, 5000, _
                        Silent:=False, allowAmbiguousRetry:=True)

    Dim cleanValue As Variant
    Assign cleanValue, GetCleanResult(rawRes)

    If IsEmpty(cleanValue) Or IsNull(cleanValue) Or IsObject(cleanValue) Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:PerformTrustedClick", _
                  "TrustedClickHitTestIndeterminate: hit-test completed without a usable verdict; " & _
                  "synthetic fallback was not attempted."
    End If

    Dim verdict As String: verdict = CStr(cleanValue)
    Select Case verdict
        Case "hit=1"
            RequireTrustedHitTestResult = True
        Case "hit=0"
            RequireTrustedHitTestResult = False
        Case Else
            Err.Raise ERR_ACTION_FAILED, "BiDi:PerformTrustedClick", _
                      "TrustedClickHitTestIndeterminate: unexpected hit-test verdict '" & verdict & _
                      "'; synthetic fallback was not attempted."
    End Select
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

Private Function ExecuteSelectValueCore(ByVal locatorKind As ElementLocatorKind, _
                                        ByVal locator As Variant, _
                                        ByVal valueOrText As String, _
                                        ByVal selectByText As Boolean, _
                                        ByVal searchTimeoutMs As Long, _
                                        ByVal waitNetworkIdle As Boolean, _
                                        ByVal maxTimeoutMs As Long, _
                                        ByVal contextId As String, _
                                        ByVal callerName As String) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim sharedId As String
    sharedId = ResolveElementSharedId( _
        locatorKind, locator, targetCtx, searchTimeoutMs, callerName)

    Dim elementArgs As Collection: Set elementArgs = NewElementArgs(sharedId)
    Dim isStandardSelect As Boolean
    isStandardSelect = RequireBooleanResult(GetCleanResult(CallScript( _
        "return arguments[0].tagName==='SELECT';", elementArgs, targetCtx, _
        False, False, 2000, Silent:=True, allowAmbiguousRetry:=True)), _
        callerName, "standard-select probe")

    If isStandardSelect Then
        Dim jsKey As String
        If selectByText Then jsKey = "select_text" Else jsKey = "select"
        ExecuteSelectValueCore = ExecuteBaseAction( _
            targetCtx, jsKey, NewElementArgs(sharedId, valueOrText), waitNetworkIdle, _
            DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs, _
            False, locatorKind, locator, searchTimeoutMs, callerName)
    Else
        Call ExecuteBaseAction(targetCtx, "click", elementArgs, False, _
                               DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, _
                               maxTimeoutMs, False, locatorKind, locator, _
                               searchTimeoutMs, callerName)

        ExecuteSelectValueCore = ExecuteUniqueVisibleRoleOptionClick( _
            targetCtx, valueOrText, searchTimeoutMs, waitNetworkIdle, maxTimeoutMs, _
            callerName)
    End If
End Function

Public Function ExecuteSelectValueByXPath(ByVal xpath As String, _
                                          ByVal valueOrText As String, _
                                          Optional ByVal selectByText As Boolean = False, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                          Optional ByVal contextId As String = "") As String
    ExecuteSelectValueByXPath = ExecuteSelectValueCore( _
        ElementLocatorXPath, xpath, valueOrText, selectByText, searchTimeoutMs, _
        waitNetworkIdle, maxTimeoutMs, contextId, "ExecuteSelectValueByXPath")
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

' Custom role-option resolution traverses both the document tree and open ShadowRoots.
' It preserves exact normalized text, strict visibility, and ambiguity rejection.
Private Function TryResolveUniqueVisibleRoleOptionDeepOnce(ByVal optionText As String, _
                                                            ByVal targetCtx As String, _
                                                            ByRef state As String, _
                                                            ByRef candidateCount As Long, _
                                                            ByRef visibleCount As Long) As String
    Dim params As New Dictionary
    Dim targetDic As New Dictionary
    Dim args As New Collection
    Dim functionDeclaration As String

    functionDeclaration = "function(t){" & _
        "const norm=s=>String(s==null?'':s).replace(/\s+/g,' ').trim();" & _
        "const vis=e=>{if(!e||!e.isConnected)return false;const s=getComputedStyle(e),r=e.getBoundingClientRect();" & _
        "return s.display!=='none'&&s.visibility!=='hidden'&&parseFloat(s.opacity||'1')>0&&" & _
        "(r.width>0||r.height>0||e.getClientRects().length>0);};" & _
        "const matches=[];const walk=r=>{for(const e of r.querySelectorAll('[role=option]')){" & _
        "if(norm(e.textContent)===norm(t))matches.push(e);}" & _
        "for(const e of r.querySelectorAll('*')){if(e.shadowRoot)walk(e.shadowRoot);}};" & _
        "walk(document);const visible=matches.filter(vis);" & _
        "if(visible.length===1)return visible[0];" & _
        "if(visible.length>1)return 'A|'+matches.length+'|'+visible.length;" & _
        "if(matches.length>0)return 'H|'+matches.length+'|0';" & _
        "return 'N|0|0';}"

    args.Add NewLocalValue("string", optionText)
    targetDic.Add "context", targetCtx

    params.Add "functionDeclaration", functionDeclaration
    params.Add "arguments", args
    params.Add "target", targetDic
    params.Add "awaitPromise", False

    Dim rawRes As String
    rawRes = ExecuteBiDiCommand("script.callFunction", params, True, True, True, _
                                (targetCtx = m_mainContextId), ResolveTransportTimeout(5000))

    Dim json As Object
    Set json = ValidateProtocolResponse(rawRes, "script.callFunction")

    Dim resultType As String
    resultType = LCase$(GetPathString(json, "result.result.type"))

    If resultType = "node" Then
        Dim sid As String
        sid = GetPathString(json, "result.result.sharedId")
        If Len(sid) = 0 Then
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveUniqueVisibleRoleOptionDeepOnce", _
                      "Deep role-option resolver returned a node without sharedId."
        End If

        state = "unique"
        candidateCount = 1
        visibleCount = 1
        TryResolveUniqueVisibleRoleOptionDeepOnce = sid
        Exit Function
    End If

    If resultType <> "string" Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveUniqueVisibleRoleOptionDeepOnce", _
                  "Deep role-option resolver returned unexpected remote value type: " & resultType
    End If

    Dim token As String
    token = GetPathString(json, "result.result.value")

    Dim parts() As String
    parts = Split(token, "|")
    If UBound(parts) <> 2 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveUniqueVisibleRoleOptionDeepOnce", _
                  "Deep role-option resolver returned an invalid status token: " & token
    End If

    If Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveUniqueVisibleRoleOptionDeepOnce", _
                  "Deep role-option resolver returned invalid counts: " & token
    End If

    candidateCount = CLng(parts(1))
    visibleCount = CLng(parts(2))

    Select Case parts(0)
        Case "A"
            state = "ambiguous"
        Case "H"
            state = "hidden"
        Case "N"
            state = "notfound"
        Case Else
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveUniqueVisibleRoleOptionDeepOnce", _
                      "Deep role-option resolver returned an unknown status token: " & token
    End Select
End Function

Private Function ResolveUniqueVisibleRoleOptionByText(ByVal optionText As String, _
                                                       ByVal targetCtx As String, _
                                                       ByVal searchTimeoutMs As Long, _
                                                       ByVal callerName As String) As String
    If Len(optionText) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "DropdownOptionTextEmpty: option text must not be empty."
    End If
    If InStr(optionText, Chr$(0)) > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Dropdown option text contains null byte (Chr$(0)). Sanitize input before calling."
    End If

    Dim isImmediate As Boolean: isImmediate = (searchTimeoutMs <= 0)
    Dim vTimeoutMs As Long
    If isImmediate Then
        vTimeoutMs = 0
    Else
        vTimeoutMs = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 10000)
    End If

    Dim startTick As Currency: startTick = socket_.NowTicks()
    Dim lastState As String
    Dim lastCandidateCount As Long
    Dim lastVisibleCount As Long

    Do
        Dim sid As String
        sid = TryResolveUniqueVisibleRoleOptionDeepOnce( _
            optionText, targetCtx, lastState, lastCandidateCount, lastVisibleCount)

        If Len(sid) > 0 Then
            ResolveUniqueVisibleRoleOptionByText = sid
            Exit Function
        End If

        If lastState = "ambiguous" Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                      "AmbiguousDropdownOption: exact normalized text matched " & _
                      CStr(lastVisibleCount) & " visible role=option elements across document/Shadow DOM. Text=" & optionText
        End If

        If isImmediate Then Exit Do
        WaitHeartbeat 100
        If IsTimeout(startTick, CDbl(vTimeoutMs)) Then Exit Do
    Loop

    If lastCandidateCount > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "DropdownOptionNotVisible: exact normalized text matched " & _
                  CStr(lastCandidateCount) & " role=option element(s) across document/Shadow DOM, but none was visible. Text=" & optionText
    End If

    Dim timeoutLabel As String
    If vTimeoutMs = 0 Then
        timeoutLabel = "0 (immediate)"
    Else
        timeoutLabel = CStr(vTimeoutMs)
    End If

    Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
              "DropdownOptionNotFound: no role=option with exact normalized text was found within " & _
              timeoutLabel & "ms across document/Shadow DOM. Text=" & optionText
End Function

Private Function ExecuteUniqueVisibleRoleOptionClick(ByVal targetCtx As String, _
                                                     ByVal optionText As String, _
                                                     ByVal searchTimeoutMs As Long, _
                                                     ByVal waitNetworkIdle As Boolean, _
                                                     ByVal maxTimeoutMs As Long, _
                                                     ByVal callerName As String) As String
    Dim staleRetryCount As Long

    Do
        Dim optionSharedId As String
        optionSharedId = ResolveUniqueVisibleRoleOptionByText( _
            optionText, targetCtx, searchTimeoutMs, callerName)

        Dim actionRes As String
        Dim actionErr As Long
        Dim actionErrSrc As String
        Dim actionErrDesc As String

        On Error Resume Next
        Err.Clear
        actionRes = ExecuteBaseAction( _
            targetCtx, "click", NewElementArgs(optionSharedId), waitNetworkIdle, _
            DEFAULT_IDLE_STABLE_MS, True, 0, maxTimeoutMs)
        actionErr = Err.Number
        actionErrSrc = Err.source
        actionErrDesc = Err.description
        On Error GoTo 0

        If actionErr = 0 Then
            ExecuteUniqueVisibleRoleOptionClick = actionRes
            Exit Function
        End If

        If Not IsStaleElementFailure(actionErr, actionErrSrc, actionErrDesc) Or _
           staleRetryCount >= DEFAULT_STALE_RETRIES Then
            Err.Raise actionErr, actionErrSrc, actionErrDesc
        End If

        staleRetryCount = staleRetryCount + 1
        If m_debugMode Then
            Debug.Print "BiDi: Dropdown option became stale. Re-resolving unique visible option " & _
                        staleRetryCount & "/" & DEFAULT_STALE_RETRIES & ": " & optionText
        End If
        WaitHeartbeat 100
    Loop
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

Private Function TryResolveShadowElementSharedIdOnce(ByVal selectorsArray As Variant, _
                                                       ByVal targetCtx As String) As String
    Dim params As New Dictionary
    Dim targetDic As New Dictionary
    Dim args As New Collection
    Dim functionDeclaration As String

    ' Node references must bypass CallScript because the Safety Shell JSON-stringifies results.
    ' Compatibility contract: each selector is deep-searched through the current root and all open
    ' descendant ShadowRoots. When an array is supplied, the next selector starts from the matched
    ' element's ShadowRoot when present. This preserves the v4.2/v4.3 single-selector behavior used
    ' by live-site samples such as ExecuteShadowClick "#utility-sign-in button".
    functionDeclaration = "function(sels){" & _
        "const fd=(r,s)=>{let n=r.querySelector(String(s));if(n)return n;" & _
        "for(const c of r.querySelectorAll('*')){if(c.shadowRoot){n=fd(c.shadowRoot,s);if(n)return n;}}return null;};" & _
        "let el=document;for(const s of sels){el=fd(el.shadowRoot||el,s);if(!el)return null;}return el;}"

    args.Add BuildShadowSelectorArg(selectorsArray)
    targetDic.Add "context", targetCtx

    params.Add "functionDeclaration", functionDeclaration
    params.Add "arguments", args
    params.Add "target", targetDic
    params.Add "awaitPromise", False

    Dim rawRes As String
    rawRes = ExecuteBiDiCommand("script.callFunction", params, True, True, True, _
                                (targetCtx = m_mainContextId), ResolveTransportTimeout(5000))

    Dim json As Object
    Set json = ValidateProtocolResponse(rawRes, "script.callFunction")

    Dim resultType As String
    resultType = LCase$(GetPathString(json, "result.result.type"))

    If resultType = "null" Or resultType = "undefined" Then Exit Function

    If resultType <> "node" Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveShadowElementSharedIdOnce", _
                  "Shadow resolver returned unexpected remote value type: " & resultType
    End If

    Dim sid As String
    sid = GetPathString(json, "result.result.sharedId")
    If Len(sid) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:TryResolveShadowElementSharedIdOnce", _
                  "Shadow resolver returned a node without sharedId."
    End If

    TryResolveShadowElementSharedIdOnce = sid
End Function

Private Function ResolveShadowElementSharedId(ByVal selectorsArray As Variant, _
                                                          ByVal targetCtx As String, _
                                                          ByVal searchTimeoutMs As Long, _
                                                          ByVal callerName As String, _
                                                          Optional ByVal required As Boolean = True) As String
    Dim isImmediate As Boolean
    isImmediate = (searchTimeoutMs <= 0)

    Dim vTimeoutMs As Long
    If isImmediate Then
        vTimeoutMs = 0
    Else
        vTimeoutMs = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 5000)
    End If

    Dim startTick As Currency
    startTick = socket_.NowTicks()

    Do
        Dim sid As String
        sid = TryResolveShadowElementSharedIdOnce(selectorsArray, targetCtx)
        If Len(sid) > 0 Then
            ResolveShadowElementSharedId = sid
            Exit Function
        End If

        If isImmediate Then Exit Do
        WaitHeartbeat 100
        If IsTimeout(startTick, CDbl(vTimeoutMs)) Then Exit Do
    Loop

    Dim timeoutLabel As String
    If isImmediate Then
        timeoutLabel = "0 (immediate)"
    Else
        timeoutLabel = CStr(vTimeoutMs)
    End If

    If required Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Shadow element not found within " & timeoutLabel & "ms."
    End If
End Function

Private Sub ResolveElementActionSpec(ByVal actionType As BiDiElementAction, _
                                     ByRef actionKey As String, _
                                     ByRef requiresValue As Boolean, _
                                     ByRef defaultWaitNetworkIdle As Boolean)
    actionKey = vbNullString
    requiresValue = False
    defaultWaitNetworkIdle = False

    Select Case actionType
        Case ElementActionClick
            actionKey = "click"
            defaultWaitNetworkIdle = True
        Case ElementActionInput
            actionKey = "input"
            requiresValue = True
            defaultWaitNetworkIdle = True
        Case ElementActionInputKeys
            actionKey = "input_keys"
            requiresValue = True
            defaultWaitNetworkIdle = True
        Case ElementActionSelect
            actionKey = "select"
            requiresValue = True
            defaultWaitNetworkIdle = True
        Case ElementActionSelectText
            actionKey = "select_text"
            requiresValue = True
            defaultWaitNetworkIdle = True
        Case ElementActionGetText
            actionKey = "get_text"
        Case ElementActionGetAttribute
            actionKey = "get_attribute"
            requiresValue = True
        Case ElementActionGetProperty
            actionKey = "get_property"
            requiresValue = True
        Case ElementActionGetOuterHtml
            actionKey = "get_outer_html"
        Case ElementActionVisibilityCheck
            actionKey = "visibility_check"
        Case Else
            Err.Raise ERR_ACTION_FAILED, "BiDi:ResolveElementActionSpec", _
                      "Unsupported element action value: " & CStr(actionType)
    End Select
End Sub

Private Function ExecuteElementActionCore(ByVal locatorKind As ElementLocatorKind, _
                                          ByVal locator As Variant, _
                                          ByVal actionType As BiDiElementAction, _
                                          ByVal hasValue As Boolean, _
                                          ByVal value As Variant, _
                                          ByVal searchTimeoutMs As Long, _
                                          ByVal hasWaitNetworkIdle As Boolean, _
                                          ByVal waitNetworkIdleValue As Boolean, _
                                          ByVal minStableMs As Long, _
                                          ByVal maxTimeoutMs As Long, _
                                          ByVal checkUrl As Boolean, _
                                          ByVal contextId As String, _
                                          ByVal callerName As String, _
                                          Optional ByVal returnRaw As Boolean = False) As Variant
    Dim actionKey As String
    Dim requiresValue As Boolean
    Dim defaultWaitNetworkIdle As Boolean
    ResolveElementActionSpec actionType, actionKey, requiresValue, defaultWaitNetworkIdle

    If requiresValue Then
        If Not hasValue Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                      "Action '" & actionKey & "' requires a value argument."
        End If
    ElseIf hasValue Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Action '" & actionKey & "' does not accept a value argument."
    End If

    If actionType = ElementActionGetAttribute Then
        ValidateDomMemberName CStr(value), "attributeName", callerName
    ElseIf actionType = ElementActionGetProperty Then
        ValidateDomMemberName CStr(value), "propertyName", callerName
    End If

    Dim shouldWaitNetworkIdle As Boolean
    If hasWaitNetworkIdle Then
        shouldWaitNetworkIdle = waitNetworkIdleValue
    Else
        shouldWaitNetworkIdle = defaultWaitNetworkIdle
    End If

    If checkUrl And actionType <> ElementActionClick Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "checkUrl is supported only for ElementActionClick."
    End If

    Dim targetCtx As String
    targetCtx = ResolveTargetContext(contextId)

    Dim elementRequired As Boolean
    elementRequired = (actionType <> ElementActionVisibilityCheck)

    Dim sid As String
    sid = ResolveElementSharedId( _
        locatorKind, locator, targetCtx, searchTimeoutMs, callerName, elementRequired)

    If Len(sid) = 0 Then
        ExecuteElementActionCore = False
        Exit Function
    End If

    Dim args As Collection
    If requiresValue Then
        Set args = NewElementArgs(sid, CStr(value))
    Else
        Set args = NewElementArgs(sid)
    End If

    Dim rawResult As String
    rawResult = ExecuteBaseAction( _
        targetCtx, actionKey, args, shouldWaitNetworkIdle, minStableMs, True, _
        DEFAULT_STALE_RETRIES, maxTimeoutMs, checkUrl, locatorKind, locator, _
        searchTimeoutMs, callerName)

    If returnRaw Then
        ExecuteElementActionCore = rawResult
        Exit Function
    End If

    Dim cleanResult As Variant
    cleanResult = GetCleanResult(rawResult)

    Select Case actionType
        Case ElementActionGetText
            ExecuteElementActionCore = RequireStringResult( _
                cleanResult, callerName, "text")
        Case ElementActionGetAttribute
            ExecuteElementActionCore = RequireScalarResult( _
                cleanResult, callerName, _
                "The attribute result was not a scalar value or Null.")
        Case ElementActionGetProperty
            ExecuteElementActionCore = RequireScalarResult( _
                cleanResult, callerName, _
                "The property result was not a supported scalar value or Null.")
        Case ElementActionGetOuterHtml
            ExecuteElementActionCore = RequireStringResult( _
                cleanResult, callerName, "outerHTML")
        Case ElementActionVisibilityCheck
            ExecuteElementActionCore = RequireBooleanResult( _
                cleanResult, callerName, "visibility")
        Case Else
            ExecuteElementActionCore = cleanResult
    End Select
End Function

Private Sub ValidateShadowUnlockerRequirement(ByVal requireUnlocker As Boolean, _
                                                   ByVal callerName As String)
    If requireUnlocker And Not m_shadowUnlockerRegistered Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:" & callerName, _
            "Shadow Unlocker is not registered for this session. Reconnect with " & _
            "requireShadowUnlocker:=True, or verify script.addPreloadScript support."
    End If
End Sub

Private Sub ReraiseShadowPublicError(ByVal errorNumber As Long, _
                                     ByVal errorSource As String, _
                                     ByVal errorDescription As String, _
                                     ByVal callerName As String)
    If Not m_shadowUnlockerRegistered And _
       InStr(1, errorDescription, "Shadow element not found", vbTextCompare) > 0 Then
        errorDescription = errorDescription & vbCrLf & _
            "Diagnostic: Shadow Unlocker is not registered. The target may be inside a " & _
            "closed shadow root. Open shadow roots remain accessible without the Unlocker."
    End If

    If Len(errorSource) = 0 Then errorSource = "BiDi:" & callerName
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

' Centralizes Shadow facade validation and diagnostics.
Private Function ExecuteShadowElementActionCore(ByVal selectorPath As Variant, _
                                                ByVal actionType As BiDiElementAction, _
                                                ByVal hasValue As Boolean, _
                                                ByVal value As Variant, _
                                                ByVal searchTimeoutMs As Long, _
                                                ByVal hasWaitNetworkIdle As Boolean, _
                                                ByVal waitNetworkIdleValue As Boolean, _
                                                ByVal minStableMs As Long, _
                                                ByVal maxTimeoutMs As Long, _
                                                ByVal checkUrl As Boolean, _
                                                ByVal contextId As String, _
                                                ByVal requireUnlocker As Boolean, _
                                                ByVal callerName As String, _
                                                Optional ByVal returnRaw As Boolean = False) As Variant
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, callerName

    ExecuteShadowElementActionCore = ExecuteElementActionCore( _
        ElementLocatorShadowPath, selectorPath, actionType, hasValue, value, _
        searchTimeoutMs, hasWaitNetworkIdle, waitNetworkIdleValue, minStableMs, _
        maxTimeoutMs, checkUrl, contextId, callerName, returnRaw)
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, callerName
End Function

' Basic element operations use BiDiElementAction for both XPath and Shadow locators.
' Executes a basic element action against an XPath locator.
Public Function ExecuteElementActionByXPath(ByVal xpath As String, _
                                            ByVal actionType As BiDiElementAction, _
                                            Optional ByVal value As Variant, _
                                            Optional ByVal searchTimeoutMs As Long = 10000, _
                                            Optional ByVal waitNetworkIdle As Variant, _
                                            Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                            Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                            Optional ByVal contextId As String = "", _
                                            Optional ByVal checkUrl As Boolean = False) As Variant
    Dim hasValue As Boolean
    Dim actionValue As Variant
    hasValue = Not IsMissing(value)
    If hasValue Then actionValue = value

    Dim hasWaitOverride As Boolean
    Dim waitValue As Boolean
    hasWaitOverride = Not IsMissing(waitNetworkIdle)
    If hasWaitOverride Then
        If VarType(waitNetworkIdle) <> vbBoolean Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteElementActionByXPath", _
                      "waitNetworkIdle must be Boolean."
        End If
        waitValue = CBool(waitNetworkIdle)
    End If

    ExecuteElementActionByXPath = ExecuteElementActionCore( _
        ElementLocatorXPath, xpath, actionType, hasValue, actionValue, searchTimeoutMs, _
        hasWaitOverride, waitValue, minStableMs, maxTimeoutMs, checkUrl, contextId, _
        "ExecuteElementActionByXPath")
End Function

' Executes the same basic element actions against a Shadow DOM locator path.
Public Function ExecuteShadowAction(ByVal selectorPath As Variant, _
                                    ByVal actionType As BiDiElementAction, _
                                    Optional ByVal value As Variant, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal waitNetworkIdle As Variant, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal checkUrl As Boolean = False, _
                                    Optional ByVal requireUnlocker As Boolean = False) As Variant
    Dim hasValue As Boolean
    Dim actionValue As Variant
    hasValue = Not IsMissing(value)
    If hasValue Then actionValue = value

    Dim hasWaitOverride As Boolean
    Dim waitValue As Boolean
    hasWaitOverride = Not IsMissing(waitNetworkIdle)
    If hasWaitOverride Then
        If VarType(waitNetworkIdle) <> vbBoolean Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteShadowAction", _
                      "waitNetworkIdle must be Boolean."
        End If
        waitValue = CBool(waitNetworkIdle)
    End If

    ExecuteShadowAction = ExecuteShadowElementActionCore( _
        selectorPath, actionType, hasValue, actionValue, searchTimeoutMs, _
        hasWaitOverride, waitValue, minStableMs, maxTimeoutMs, checkUrl, _
        contextId, requireUnlocker, "ExecuteShadowAction")
End Function

' Compatibility facade for Shadow DOM click.
Public Function ExecuteShadowClick(ByVal selectorsArray As Variant, _
                                   Optional ByVal searchTimeoutMs As Long = 10000, _
                                   Optional ByVal waitNetworkIdle As Boolean = True, _
                                   Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                   Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                   Optional ByVal contextId As String = "", _
                                   Optional ByVal requireUnlocker As Boolean = False, _
                                   Optional ByVal checkUrl As Boolean = False) As String
    Dim noValue As Variant
    ExecuteShadowClick = CStr(ExecuteShadowElementActionCore( _
        selectorsArray, ElementActionClick, False, noValue, searchTimeoutMs, _
        True, waitNetworkIdle, minStableMs, maxTimeoutMs, checkUrl, contextId, _
        requireUnlocker, "ExecuteShadowClick", True))
End Function

' Compatibility facade for Shadow DOM input.
Public Function ExecuteShadowInputValue(ByVal selectorPath As Variant, _
                                        ByVal valueToSet As String, _
                                        Optional ByVal searchTimeoutMs As Long = 10000, _
                                        Optional ByVal waitNetworkIdle As Boolean = True, _
                                        Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                        Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                        Optional ByVal contextId As String = "", _
                                        Optional ByVal useKeyEvents As Boolean = True, _
                                        Optional ByVal useNativeKeys As Boolean = False, _
                                        Optional ByVal requireUnlocker As Boolean = False) As String
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowInputValue"

    ExecuteShadowInputValue = ExecuteInputValueCore( _
        ElementLocatorShadowPath, selectorPath, valueToSet, searchTimeoutMs, _
        waitNetworkIdle, minStableMs, maxTimeoutMs, contextId, useKeyEvents, _
        useNativeKeys, "ExecuteShadowInputValue")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, "ExecuteShadowInputValue"
End Function

Public Function ExecuteShadowSelectValue(ByVal selectorPath As Variant, _
                                         ByVal valueOrText As String, _
                                         Optional ByVal selectByText As Boolean = False, _
                                         Optional ByVal searchTimeoutMs As Long = 10000, _
                                         Optional ByVal waitNetworkIdle As Boolean = True, _
                                         Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                         Optional ByVal contextId As String = "", _
                                         Optional ByVal requireUnlocker As Boolean = False) As String
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowSelectValue"

    ExecuteShadowSelectValue = ExecuteSelectValueCore( _
        ElementLocatorShadowPath, selectorPath, valueOrText, selectByText, _
        searchTimeoutMs, waitNetworkIdle, maxTimeoutMs, contextId, _
        "ExecuteShadowSelectValue")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, _
                             "ExecuteShadowSelectValue"
End Function

Public Function ExecuteShadowSetFiles(ByVal selectorPath As Variant, _
                                      ByVal filePaths As Variant, _
                                      Optional ByVal searchTimeoutMs As Long = 10000, _
                                      Optional ByVal waitNetworkIdle As Boolean = True, _
                                      Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                      Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                      Optional ByVal contextId As String = "", _
                                      Optional ByVal requireUnlocker As Boolean = False) As String
    On Error GoTo Failed
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteShadowSetFiles"

    ExecuteShadowSetFiles = ExecuteSetFilesCore( _
        ElementLocatorShadowPath, selectorPath, filePaths, searchTimeoutMs, _
        waitNetworkIdle, minStableMs, maxTimeoutMs, contextId, _
        "ExecuteShadowSetFiles")
    Exit Function

Failed:
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorSource As String: errorSource = Err.source
    Dim errorDescription As String: errorDescription = Err.description
    ReraiseShadowPublicError errorNumber, errorSource, errorDescription, _
                             "ExecuteShadowSetFiles"
End Function

Public Function ExecuteShadowGetText(ByVal selectorPath As Variant, _
                                     Optional ByVal searchTimeoutMs As Long = 10000, _
                                     Optional ByVal contextId As String = "", _
                                     Optional ByVal requireUnlocker As Boolean = False) As String
    Dim noValue As Variant
    ExecuteShadowGetText = CStr(ExecuteShadowElementActionCore( _
        selectorPath, ElementActionGetText, False, noValue, searchTimeoutMs, _
        False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, False, _
        contextId, requireUnlocker, "ExecuteShadowGetText"))
End Function

Public Function ExecuteShadowGetAttribute(ByVal selectorPath As Variant, _
                                          ByVal attributeName As String, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal contextId As String = "", _
                                          Optional ByVal requireUnlocker As Boolean = False) As Variant
    ExecuteShadowGetAttribute = ExecuteShadowElementActionCore( _
        selectorPath, ElementActionGetAttribute, True, attributeName, searchTimeoutMs, _
        False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, False, _
        contextId, requireUnlocker, "ExecuteShadowGetAttribute")
End Function

Public Function ExecuteShadowGetProperty(ByVal selectorPath As Variant, _
                                         ByVal propertyName As String, _
                                         Optional ByVal searchTimeoutMs As Long = 10000, _
                                         Optional ByVal contextId As String = "", _
                                         Optional ByVal requireUnlocker As Boolean = False) As Variant
    ExecuteShadowGetProperty = ExecuteShadowElementActionCore( _
        selectorPath, ElementActionGetProperty, True, propertyName, searchTimeoutMs, _
        False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, False, _
        contextId, requireUnlocker, "ExecuteShadowGetProperty")
End Function

Public Function ExecuteShadowGetOuterHtml(ByVal selectorPath As Variant, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal contextId As String = "", _
                                          Optional ByVal requireUnlocker As Boolean = False) As String
    Dim noValue As Variant
    ExecuteShadowGetOuterHtml = CStr(ExecuteShadowElementActionCore( _
        selectorPath, ElementActionGetOuterHtml, False, noValue, searchTimeoutMs, _
        False, False, DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, False, _
        contextId, requireUnlocker, "ExecuteShadowGetOuterHtml"))
End Function

Public Function ExecuteShadowIsElementVisible(ByVal selectorPath As Variant, _
                                              Optional ByVal searchTimeoutMs As Long = 10000, _
                                              Optional ByVal waitNetworkIdle As Boolean = False, _
                                              Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                              Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                              Optional ByVal contextId As String = "", _
                                              Optional ByVal requireUnlocker As Boolean = False) As Boolean
    Dim vSearchTimeout As Long
    If searchTimeoutMs = 0 Then
        vSearchTimeout = 0
    Else
        vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    End If

    Dim noValue As Variant
    ExecuteShadowIsElementVisible = CBool(ExecuteShadowElementActionCore( _
        selectorPath, ElementActionVisibilityCheck, False, noValue, vSearchTimeout, _
        True, waitNetworkIdle, _
        SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS), _
        SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS), _
        False, contextId, requireUnlocker, "ExecuteShadowIsElementVisible"))
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
    Dim scrollResult As String
    scrollResult = ExecuteBaseAction(targetCtx, "lazy_scroll", args, _
                                     waitNetworkIdle, minStableMs, True, _
                                     DEFAULT_STALE_RETRIES, vMaxTimeout)

    ' The JS result remains Boolean for compatibility. False now means the bounded scroll budget
    ' expired before two consecutive equal-height samples proved convergence. The action itself
    ' still completed normally; surface the distinction as an advisory diagnostic rather than an
    ' exception because infinite/virtualized feeds may be intentionally non-convergent.
    Dim convergenceValue As Variant
    convergenceValue = GetCleanResult(scrollResult)
    If Not IsEmpty(convergenceValue) Then
        If VarType(convergenceValue) = vbBoolean Then
            If Not CBool(convergenceValue) Then
                LogLazyScrollUnstableBestEffort targetCtx, vScrollTimeout, scrollPollMs
            End If
        End If
    End If

    ExecuteLazyLoadScroll = scrollResult
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
    Dim recoveryAttemptCount As Long
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

        ' A response that exceeded the receive cap will not become smaller through context
        ' recovery/retry. Surface it immediately while preserving ambiguous-dispatch semantics.
        If localErrNum = 608 Then
            LogSignalLine "[TRANSPORT-OVERSIZE] " & socket_.LastTransportOversizeDiagnostic
            Exit Do
        End If

        ' The command-correlation ceiling protects Excel responsiveness; it is not a transport
        ' failure and recovery/replay would only duplicate or overlap a browser operation that may
        ' still be running. Surface it immediately even when the caller otherwise permits retries.
        If localErrNum = 604 And _
           InStr(1, localErrDesc, "CommandCorrelationTimeout:", vbBinaryCompare) > 0 Then
            LogSignalLine "[CORRELATION-TIMEOUT] method=" & method & _
                          " | requestedWaitMs=" & CStr(socket_.LastCommandRequestedWaitMs) & _
                          " | effectiveWaitMs=" & CStr(socket_.LastCommandEffectiveWaitMs) & _
                          " | capMs=" & CStr(socket_.MaxCommandCorrelationWaitMs)
            Exit Do
        End If

        ' Ambiguous outcomes for state-changing commands include 608, 604, 600, 601, 517, and
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

            recoveryAttemptCount = recoveryAttemptCount + 1
            m_isRecovering = True

            ' F1: Wrapper recovery is a critical phase. Disable the communicator's guarded
            ' Office yields as well as using a non-Office backoff so callbacks cannot re-enter
            ' this wrapper while its context/retry state is transient. Restore the caller setting
            ' before releasing m_isRecovering, even when refresh fails.
            Dim recoveryErr As Long
            Dim recoveryErrDesc As String
            Dim recoverySavedAllowUiPump As Boolean
            recoverySavedAllowUiPump = socket_.AllowUiPump

            On Error Resume Next
            socket_.AllowUiPump = False
            WaitRecoveryBackoffNoOfficePump 800 * (retryCount + 1)
            recoveryErr = Err.Number
            recoveryErrDesc = Err.description
            Err.Clear

            If recoveryErr = 0 Then
                If m_debugMode Then
                    Debug.Print "BiDi [RECOVERY]: Attempting to refresh context (" & recoveryAttemptCount & "/" & MAX_CONTEXT_RECOVERY_RETRIES & ")..."
                End If

                Call UpdateMainContextId(True, True) ' forceRefresh:=True, Silent:=True
                recoveryErr = Err.Number
                recoveryErrDesc = Err.description
                Err.Clear
            End If

            socket_.AllowUiPump = recoverySavedAllowUiPump
            Err.Clear
            On Error GoTo 0
            m_isRecovering = False

            If recoveryErr = 0 Then
                ' Retarget only when the caller explicitly permits main-context recovery.
                ' A failed transactional rewrite is terminal: retrying a partially retargeted
                ' command would be less safe than surfacing the recovery failure.
                If m_mainContextId <> "" And allowContextRetarget Then
                    If Not UpdateContextAndRealm(params, m_mainContextId) Then
                        retryCount = retryCount + 1
                        localErrNum = ERR_BIDI_PROTOCOL
                        localErrDesc = "Context recovery retarget failed; retry parameters were left unchanged."
                        Debug.Print "BiDi [RECOVERY FAIL]: " & localErrDesc
                        Exit Do
                    End If
                End If
                retryCount = retryCount + 1
            Else
                Debug.Print "BiDi [RECOVERY FAIL]: " & recoveryErrDesc
                If Len(recoveryErrDesc) > 0 Then
                    localErrDesc = localErrDesc & " | Recovery failed: " & recoveryErrDesc
                End If
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    Err.Raise IIf(localErrNum = 0, ERR_ACTION_FAILED, localErrNum), _
              "BiDi:ExecuteBiDiCommand", _
              "Command '" & method & "' failed after " & recoveryAttemptCount & " recovery attempt(s). " & _
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

Private Function UpdateContextAndRealmRecursive(ByVal node As Variant, _
                                                ByVal newCtx As String, _
                                                Optional ByVal newRealm As String = "", _
                                                Optional ByVal currentDepth As Long = 0) As Boolean
    Const MAX_RETARGET_RECURSION_DEPTH As Long = 32

    On Error GoTo ErrH

    ' Retarget parameters originate from JSON-like Dictionary/Collection graphs. Reject
    ' pathological depth before mutating deeper nodes so recovery cannot silently succeed
    ' with only part of a parameter graph rewritten.
    If currentDepth > MAX_RETARGET_RECURSION_DEPTH Then Exit Function

    If Not IsObject(node) Then
        UpdateContextAndRealmRecursive = True
        Exit Function
    End If

    If TypeName(node) = "Dictionary" Then
        Dim k As Variant
        For Each k In node.keys
            If StrComp(CStr(k), "context", vbTextCompare) = 0 Then
                node(k) = newCtx
            ElseIf StrComp(CStr(k), "realm", vbTextCompare) = 0 And Len(newRealm) > 0 Then
                node(k) = newRealm
            Else
                If Not UpdateContextAndRealmRecursive(node(k), newCtx, newRealm, currentDepth + 1) Then
                    Exit Function
                End If
            End If
        Next k
        UpdateContextAndRealmRecursive = True
        Exit Function
    End If

    If TypeName(node) = "Collection" Then
        Dim i As Long
        For i = 1 To node.count
            If Not UpdateContextAndRealmRecursive(node(i), newCtx, newRealm, currentDepth + 1) Then
                Exit Function
            End If
        Next i
        UpdateContextAndRealmRecursive = True
        Exit Function
    End If

    ' JSON-compatible retarget graphs should contain only Dictionary/Collection objects.
    ' An unexpected object is a contract failure, not a successful no-op.
    UpdateContextAndRealmRecursive = False
    Exit Function

ErrH:
    UpdateContextAndRealmRecursive = False
End Function

Private Function UpdateContextAndRealm(ByRef params As Dictionary, ByVal newCtx As String) As Boolean
    On Error GoTo ErrorHandler

    Dim newRealm As String
    newRealm = GetScriptRealmId(newCtx, True, True) ' Silent:=True

    If newRealm = "" Then Exit Function

    ' Transactional retarget: never mutate the live retry parameters until the complete
    ' cloned JSON graph has been rewritten successfully. This prevents a child failure
    ' or recursion guard from leaving a mixed old/new context+realm graph for retry.
    Dim stagedJson As String
    stagedJson = WebJsonConverter.ConvertToJson(params)

    Dim stagedObj As Object
    Set stagedObj = WebJsonConverter.ParseJson(stagedJson)
    If stagedObj Is Nothing Then Exit Function
    If TypeName(stagedObj) <> "Dictionary" Then Exit Function

    Dim stagedNode As Variant
    Set stagedNode = stagedObj
    If Not UpdateContextAndRealmRecursive(stagedNode, newCtx, newRealm, 0) Then Exit Function

    Dim stagedParams As Dictionary
    Set stagedParams = stagedObj
    Set params = stagedParams
    UpdateContextAndRealm = True
    Exit Function

ErrorHandler:
    UpdateContextAndRealm = False
End Function

Private Sub ResetVolatileContextState()
    ResetNavigationHttpCaptureState

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

Private Function ReleaseTrackedSubscriptionsBeforeClearBestEffort( _
    Optional ByVal commandTimeoutMs As Long = 0) As Boolean

    ReleaseTrackedSubscriptionsBeforeClearBestEffort = True

    If m_activeSubscriptions Is Nothing Then Set m_activeSubscriptions = New Dictionary
    EnsureRetiredSubscriptionRegistry
    If m_activeSubscriptions.count = 0 And m_retiredSubscriptionIds.count = 0 Then Exit Function

    ' Once IDs have been issued, a disconnected socket cannot prove that the BiDi session
    ' forgot them. Report "not confirmed" so the caller can retain deferred cleanup ownership.
    If socket_ Is Nothing Then
        ReleaseTrackedSubscriptionsBeforeClearBestEffort = False
        Exit Function
    End If
    If socket_.state <> StateConnected Then
        ReleaseTrackedSubscriptionsBeforeClearBestEffort = False
        Exit Function
    End If

    On Error GoTo Failed

    ' Multiple event/context registry keys may share one subscription ID because one
    ' session.subscribe call can register several events at once. Unsubscribe unique IDs only.
    Dim uniqueIds As New Dictionary
    Dim key As Variant
    Dim subId As String

    For Each key In m_activeSubscriptions.keys
        subId = CStr(m_activeSubscriptions(key))
        If Len(subId) > 0 Then
            If Not uniqueIds.Exists(subId) Then uniqueIds.Add subId, True
        End If
    Next key

    For Each key In m_retiredSubscriptionIds.keys
        subId = CStr(key)
        If Len(subId) > 0 Then
            If Not uniqueIds.Exists(subId) Then uniqueIds.Add subId, True
        End If
    Next key

    If uniqueIds.count = 0 Then Exit Function

    Dim ids() As String
    Dim i As Long
    Dim idValue As Variant
    ReDim ids(0 To uniqueIds.count - 1)

    i = 0
    For Each idValue In uniqueIds.keys
        ids(i) = CStr(idValue)
        i = i + 1
    Next idValue

    Dim params As New Dictionary
    params.Add "subscriptions", ids

    Dim raw As String
    raw = ExecuteBiDiCommand( _
              "session.unsubscribe", _
              params, _
              True, _
              allowAmbiguousRetry:=False, _
              allowRecoveryRetry:=False, _
              allowContextRetarget:=False, _
              commandTimeoutMs:=commandTimeoutMs)

    Call ValidateProtocolResponse(raw, "session.unsubscribe")
    Exit Function

Failed:
    Dim cleanupErrNum As Long
    Dim cleanupErrDesc As String
    cleanupErrNum = Err.Number
    cleanupErrDesc = Err.description
    Err.Clear

    ' One unknown ID makes a by-ID batch invalid without removing any of the others. Fall back
    ' to one-ID requests only for that deterministic response; each single unknown ID means the
    ' desired absent state is already true. Do not expand ambiguous transport failures here.
    If cleanupErrNum = ERR_BIDI_PROTOCOL And _
       InStr(1, cleanupErrDesc, "invalid argument", vbTextCompare) > 0 Then
        Dim fallbackIds As Variant
        fallbackIds = uniqueIds.keys

        Dim fallbackTimeoutMs As Long
        fallbackTimeoutMs = commandTimeoutMs
        If fallbackTimeoutMs <= 0 Then fallbackTimeoutMs = DEFAULT_COMMAND_TIMEOUT_MS

        Dim fallbackFailure As String
        Dim j As Long
        For j = LBound(fallbackIds) To UBound(fallbackIds)
            If Not TryReleaseDeferredSubscriptionId(CStr(fallbackIds(j)), _
                                                     fallbackTimeoutMs, fallbackFailure) Then
                GoTo FallbackFailed
            End If
        Next j

        ReleaseTrackedSubscriptionsBeforeClearBestEffort = True
        Exit Function
    End If

FallbackFailed:
    ReleaseTrackedSubscriptionsBeforeClearBestEffort = False
    If m_debugMode Then
        Debug.Print "BiDi Warning: remote subscription cleanup was not confirmed. " & _
                    "Err " & CStr(cleanupErrNum) & ": " & cleanupErrDesc & _
                    IIf(Len(fallbackFailure) > 0, " | fallback=" & fallbackFailure, vbNullString)
    End If
End Function

Private Sub EnsureRetiredSubscriptionRegistry()
    If m_retiredSubscriptionIds Is Nothing Then _
        Set m_retiredSubscriptionIds = New Dictionary
End Sub

Private Sub EnsurePendingSubscriptionCleanupRegistry()
    If m_pendingSubscriptionCleanupIds Is Nothing Then _
        Set m_pendingSubscriptionCleanupIds = New Dictionary
End Sub

Private Function GetBiDiSessionIdentity(ByVal webSocketUrl As String) As String
    ' Ports and localhost address-family spelling may vary while the WebDriver session path is
    ' stable. Prefer the /session/... suffix so reconnect cleanup is tied to the actual session.
    Dim value As String
    value = Trim$(webSocketUrl)

    Dim p As Long
    p = InStr(1, value, "/session/", vbTextCompare)
    If p > 0 Then value = Mid$(value, p)

    Dim cut As Long
    cut = InStr(1, value, "?", vbBinaryCompare)
    If cut > 0 Then value = Left$(value, cut - 1)
    cut = InStr(1, value, "#", vbBinaryCompare)
    If cut > 0 Then value = Left$(value, cut - 1)

    GetBiDiSessionIdentity = value
End Function

Private Sub CaptureTrackedSubscriptionsForDeferredCleanup()
    If m_activeSubscriptions Is Nothing Then Set m_activeSubscriptions = New Dictionary
    EnsureRetiredSubscriptionRegistry
    If m_activeSubscriptions.count = 0 And m_retiredSubscriptionIds.count = 0 Then Exit Sub
    If Len(m_currentSessionIdentity) = 0 Then Exit Sub

    EnsurePendingSubscriptionCleanupRegistry

    If m_pendingSubscriptionCleanupIds.count > 0 Then
        If StrComp(m_pendingSubscriptionCleanupSessionIdentity, _
                   m_currentSessionIdentity, vbBinaryCompare) <> 0 Then
            If m_debugMode Then
                Debug.Print "BiDi Warning: replacing deferred subscription cleanup ownership " & _
                            "because the tracked BiDi session identity changed."
            End If
            Set m_pendingSubscriptionCleanupIds = New Dictionary
        End If
    End If

    m_pendingSubscriptionCleanupSessionIdentity = m_currentSessionIdentity

    Dim key As Variant
    Dim subId As String
    For Each key In m_activeSubscriptions.keys
        subId = CStr(m_activeSubscriptions(key))
        If Len(subId) > 0 Then
            If Not m_pendingSubscriptionCleanupIds.Exists(subId) Then _
                m_pendingSubscriptionCleanupIds.Add subId, True
        End If
    Next key

    For Each key In m_retiredSubscriptionIds.keys
        subId = CStr(key)
        If Len(subId) > 0 Then
            If Not m_pendingSubscriptionCleanupIds.Exists(subId) Then _
                m_pendingSubscriptionCleanupIds.Add subId, True
        End If
    Next key
End Sub

Private Function TryReleaseDeferredSubscriptionId(ByVal subscriptionId As String, _
                                                   ByVal commandTimeoutMs As Long, _
                                                   ByRef failureDescription As String) As Boolean
    failureDescription = vbNullString
    On Error GoTo Failed

    Dim ids(0 To 0) As String
    ids(0) = subscriptionId

    Dim params As New Dictionary
    params.Add "subscriptions", ids

    Dim raw As String
    raw = ExecuteBiDiCommand( _
              "session.unsubscribe", _
              params, _
              True, _
              allowAmbiguousRetry:=False, _
              allowRecoveryRetry:=False, _
              allowContextRetarget:=False, _
              commandTimeoutMs:=commandTimeoutMs)

    Dim json As Object
    Set json = WebJsonConverter.ParseJson(raw)
    If json Is Nothing Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:session.unsubscribe", _
                  "session.unsubscribe returned malformed JSON during deferred cleanup."
    End If

    If json.Exists("error") Then
        ' By-ID unsubscribe reports invalid argument when this one ID is no longer known. For a
        ' single server-issued ID that is exactly the cleanup condition we need to establish.
        If StrComp(LCase$(CStr(json("error"))), "invalid argument", vbBinaryCompare) = 0 Then
            TryReleaseDeferredSubscriptionId = True
            Exit Function
        End If
    End If

    Call ValidateProtocolResponse(raw, "session.unsubscribe")
    TryReleaseDeferredSubscriptionId = True
    Exit Function

Failed:
    failureDescription = "Err " & CStr(Err.Number) & ": " & Err.description
    Err.Clear
End Function

Private Sub RecoverPendingSubscriptionCleanupOnConnect(ByVal requestedSessionIdentity As String)
    EnsurePendingSubscriptionCleanupRegistry
    If m_pendingSubscriptionCleanupIds.count = 0 Then Exit Sub

    If StrComp(m_pendingSubscriptionCleanupSessionIdentity, _
               requestedSessionIdentity, vbBinaryCompare) <> 0 Then
        If m_debugMode Then
            Debug.Print "BiDi: Deferred subscription IDs belong to a different WebDriver session; " & _
                        "local ownership record discarded."
        End If
        Set m_pendingSubscriptionCleanupIds = New Dictionary
        m_pendingSubscriptionCleanupSessionIdentity = vbNullString
        Exit Sub
    End If

    Dim ids As Variant
    ids = m_pendingSubscriptionCleanupIds.keys

    Dim i As Long
    Dim failure As String
    For i = LBound(ids) To UBound(ids)
        If TryReleaseDeferredSubscriptionId(CStr(ids(i)), 3000, failure) Then
            If m_pendingSubscriptionCleanupIds.Exists(CStr(ids(i))) Then _
                m_pendingSubscriptionCleanupIds.Remove CStr(ids(i))
        Else
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:ConnectTo", _
                "Same-session reconnect could not confirm cleanup of deferred subscription ID " & _
                CStr(ids(i)) & ". Start a new WebDriver session or retry ConnectTo. " & failure
        End If
    Next i

    m_pendingSubscriptionCleanupSessionIdentity = vbNullString
    If m_debugMode Then Debug.Print "BiDi: Deferred subscription cleanup completed on ConnectTo."
End Sub

Private Sub ClearStateCore(ByVal releaseRemoteSubscriptions As Boolean)
    Dim subscriptionsReleased As Boolean
    subscriptionsReleased = True

    If releaseRemoteSubscriptions Then
        subscriptionsReleased = ReleaseTrackedSubscriptionsBeforeClearBestEffort()

        If Not subscriptionsReleased Then
            ' If the socket is already gone, preserve only the issued IDs for a same-session
            ' reconnect. If it is still connected, retain the full active registry so delta
            ' subscribe logic cannot create duplicates while cleanup remains uncertain.
            If socket_ Is Nothing Then
                CaptureTrackedSubscriptionsForDeferredCleanup
                Set m_activeSubscriptions = New Dictionary
                Set m_retiredSubscriptionIds = New Dictionary
            ElseIf socket_.state <> StateConnected Then
                CaptureTrackedSubscriptionsForDeferredCleanup
                Set m_activeSubscriptions = New Dictionary
                Set m_retiredSubscriptionIds = New Dictionary
            End If
        End If
    End If

    ResetVolatileContextState

    SetMainContextUnbound
    If Not releaseRemoteSubscriptions Or subscriptionsReleased Then
        Set m_activeSubscriptions = New Dictionary
        Set m_retiredSubscriptionIds = New Dictionary
    End If
    Set m_blockingPatterns = New Collection

    If Not socket_ Is Nothing Then
        m_contextLifecycleDroppedEventsBaseline = CurrentLifecycleRelevantLossCount()
    Else
        m_contextLifecycleDroppedEventsBaseline = 0
    End If
    m_contextLifecycleNeedsReconcile = False
End Sub

Public Sub ClearState()
    RejectIfExclusiveCaptureActive "ClearState"
    ClearStateCore True
End Sub

' Session-boundary reset runs only when one session ends and another may begin.
' Flags are dropped before their collections so guarded append paths stay safe.
Private Sub ResetSessionBoundaryState()
    m_isRecording = False
    m_isProcessingQueue = False

    ' Dispatch-loss counters are session-scoped, matching the communicator loss counters.
    m_dispatchEventLossCount = 0
    m_dispatchLifecycleEventLossCount = 0
    m_dispatchDownloadEventLossCount = 0
    m_dispatchFileDialogEventLossCount = 0
    m_dispatchUnknownCriticalEventLossCount = 0

    Set m_recorderLogs = Nothing
    Set m_recSuppByKey = Nothing
    m_recSuppOverflowCount = 0
    m_droppedHeadCount = 0
    m_recorderStartTime = 0
    m_discoveryDomCaptureEnabled = False
    m_discoveryManualPumpActive = False
    m_discoveryDomContextId = vbNullString
    m_nextDiscoveryDomPollMs = 0#

    m_navigationDetected = False
    m_navigationContextId = ""
    ResetNavigationSupersessionWatch

    m_waitStackDepth = 0

    m_shadowUnlockerRegistered = False
    m_shadowUnlockerScriptId = vbNullString
    m_shadowUnlockerSuppressorScriptId = vbNullString
    m_shadowUnlockerContainmentRisk = False
    m_lastShadowUnlockerGuardGeneration = 0#
    Set m_autoClickerScriptIds = New Dictionary
    Set m_autoClickerTimeouts = New Dictionary
    Set m_autoClickerGenerations = New Dictionary
    Set m_autoClickerSuppressorScriptIds = New Dictionary
    Set m_autoClickerContainmentRisks = New Dictionary
    m_lastAutoClickerGuardGeneration = 0#
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

Private Function CurrentTotalAsyncEventLossCount() As Long
    If socket_ Is Nothing Then
        CurrentTotalAsyncEventLossCount = m_dispatchEventLossCount
        Exit Function
    End If

    CurrentTotalAsyncEventLossCount = SaturatedLossSum( _
        socket_.TotalAsyncEventLossCount, _
        m_dispatchEventLossCount)
End Function

Private Function CurrentLifecycleRelevantLossCount() As Long
    Dim dispatchLoss As Long
    dispatchLoss = SaturatedLossSum( _
        m_dispatchLifecycleEventLossCount, _
        m_dispatchUnknownCriticalEventLossCount)

    If socket_ Is Nothing Then
        CurrentLifecycleRelevantLossCount = dispatchLoss
        Exit Function
    End If

    CurrentLifecycleRelevantLossCount = SaturatedLossSum( _
        SaturatedLossSum(socket_.LifecycleEventLossCount, socket_.UnknownCriticalEventLossCount), _
        dispatchLoss)
End Function

Private Function CurrentDownloadRelevantLossCount() As Long
    Dim dispatchLoss As Long
    dispatchLoss = SaturatedLossSum( _
        m_dispatchDownloadEventLossCount, _
        m_dispatchUnknownCriticalEventLossCount)

    If socket_ Is Nothing Then
        CurrentDownloadRelevantLossCount = dispatchLoss
        Exit Function
    End If

    CurrentDownloadRelevantLossCount = SaturatedLossSum( _
        SaturatedLossSum(socket_.DownloadEventLossCount, socket_.UnknownCriticalEventLossCount), _
        dispatchLoss)
End Function

Private Function CurrentFileDialogRelevantLossCount() As Long
    Dim dispatchLoss As Long
    dispatchLoss = SaturatedLossSum( _
        m_dispatchFileDialogEventLossCount, _
        m_dispatchUnknownCriticalEventLossCount)

    If socket_ Is Nothing Then
        CurrentFileDialogRelevantLossCount = dispatchLoss
        Exit Function
    End If

    CurrentFileDialogRelevantLossCount = SaturatedLossSum( _
        SaturatedLossSum(socket_.FileDialogEventLossCount, socket_.UnknownCriticalEventLossCount), _
        dispatchLoss)
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
    
    ' GetCdpSessionId is a value/cache getter for the main ExecuteCDPCommand path.
    ' Failure increment and successful-command reset for that path are owned by
    ' ExecuteCDPCommand. Other explicit session-state paths may reset the shared counter.
    If sessionId <> "" Then
        m_cdpSessionId = sessionId
        GetCdpSessionId = m_cdpSessionId
    Else
        GetCdpSessionId = ""
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

' AutoClicker uses script.addPreloadScript. addPreloadScript is state-changing and a lost
' response can mean that the browser installed a preload whose script ID the wrapper never learned.
' A7 makes browser-side behavior generation-idempotent. XPath and Shadow paths share the same
' wrapper registry and preload lifecycle; only locator resolution differs inside the preload.
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeoutMs As Long = 30000)
    ExecuteRegisterAutoClickerCore BuildAutoClickerXPathKey(xpath), timeoutMs, _
                                   "ExecuteRegisterAutoClickerByXPath"
End Sub

Public Sub ExecuteRegisterShadowAutoClicker(ByVal selectorPath As Variant, _
                                            Optional ByVal timeoutMs As Long = 30000, _
                                            Optional ByVal requireUnlocker As Boolean = False)
    ValidateShadowUnlockerRequirement requireUnlocker, "ExecuteRegisterShadowAutoClicker"
    ExecuteRegisterAutoClickerCore _
        BuildAutoClickerShadowKey(selectorPath, "ExecuteRegisterShadowAutoClicker"), _
        timeoutMs, "ExecuteRegisterShadowAutoClicker"
End Sub

Private Sub ExecuteRegisterAutoClickerCore(ByVal registrationKey As String, _
                                           ByVal timeoutMs As Long, _
                                           ByVal callerName As String)
    Dim vTimeoutMs As Long
    vTimeoutMs = SanitizeMs(timeoutMs, "timeoutMs", 30000)

    EnsureAutoClickerRegistry

    If m_autoClickerScriptIds.Exists(registrationKey) Then
        Dim existingTimeout As Long
        existingTimeout = CLng(m_autoClickerTimeouts(registrationKey))

        If existingTimeout = vTimeoutMs Then
            If m_debugMode Then
                Debug.Print "BiDi: AutoClicker already registered for " & _
                            AutoClickerLocatorSummary(registrationKey) & _
                            " (ID: " & CStr(m_autoClickerScriptIds(registrationKey)) & ")"
            End If
            Exit Sub
        End If

        Dim oldScriptId As String
        oldScriptId = CStr(m_autoClickerScriptIds(registrationKey))

        Dim removeErrNum As Long
        Dim removeErrSource As String
        Dim removeErrDesc As String
        If Not TryRemovePreloadScript(oldScriptId, True, removeErrNum, removeErrSource, removeErrDesc) Then
            Dim removeRaiseNum As Long
            removeRaiseNum = removeErrNum
            If removeRaiseNum = 0 Then removeRaiseNum = ERR_BIDI_PROTOCOL

            Err.Raise removeRaiseNum, "BiDi:" & callerName, _
                      "AutoClicker replacement aborted because the previous tracked preload script " & _
                      "could not be confirmed removed. " & AutoClickerLocatorSummary(registrationKey) & _
                      AppendErrorDescription(removeErrDesc)
        End If

        RemoveTrackedAutoClickerRegistration registrationKey

        Dim replacementSuppressorId As String
        Dim suppressErrNum As Long
        Dim suppressErrSource As String
        Dim suppressErrDesc As String
        If Not TryInstallAutoClickerSuppressor(registrationKey, replacementSuppressorId, _
                                               suppressErrNum, suppressErrSource, suppressErrDesc) Then
            Dim suppressRaiseNum As Long
            suppressRaiseNum = suppressErrNum
            If suppressRaiseNum = 0 Then suppressRaiseNum = ERR_BIDI_PROTOCOL

            Err.Raise suppressRaiseNum, "BiDi:" & callerName, _
                      "AutoClicker replacement stopped after removing the previous tracked preload " & _
                      "because residual-orphan containment could not be confirmed. " & _
                      AutoClickerLocatorSummary(registrationKey) & _
                      AppendErrorDescription(suppressErrDesc)
        End If
    ElseIf m_autoClickerContainmentRisks.Exists(registrationKey) Then
        Dim riskSuppressorId As String
        Dim riskErrNum As Long
        Dim riskErrSource As String
        Dim riskErrDesc As String
        If Not TryInstallAutoClickerSuppressor(registrationKey, riskSuppressorId, _
                                               riskErrNum, riskErrSource, riskErrDesc) Then
            Dim riskRaiseNum As Long
            riskRaiseNum = riskErrNum
            If riskRaiseNum = 0 Then riskRaiseNum = ERR_BIDI_PROTOCOL

            Err.Raise riskRaiseNum, "BiDi:" & callerName, _
                      "AutoClicker registration blocked because a prior ambiguous preload could not " & _
                      "be contained. " & AutoClickerLocatorSummary(registrationKey) & _
                      AppendErrorDescription(riskErrDesc)
        End If
    End If

    Me.ExecuteSessionSubscribe "log.entryAdded", , True

    Dim generation As Double
    generation = NextAutoClickerGuardGeneration()

    Dim scriptId As String
    Dim addErrNum As Long
    Dim addErrSource As String
    Dim addErrDesc As String

    If Not TryAddAutoClickerPreload(registrationKey, vTimeoutMs, generation, True, scriptId, _
                                    addErrNum, addErrSource, addErrDesc) Then
        Dim finalDesc As String
        finalDesc = addErrDesc

        If IsAmbiguousDispatchFailure(addErrNum, addErrSource, addErrDesc) Then
            LogSignalLine "[AUTOCLICKER-AMBIGUOUS] " & AutoClickerSignalLocator(registrationKey) & _
                          " | generation=" & AutoClickerGenerationLiteral(generation) & _
                          " | error=" & CStr(addErrNum)

            Dim containmentId As String
            Dim containmentErrNum As Long
            Dim containmentErrSource As String
            Dim containmentErrDesc As String

            If TryInstallAutoClickerSuppressor(registrationKey, containmentId, _
                                               containmentErrNum, containmentErrSource, containmentErrDesc) Then
                finalDesc = finalDesc & vbCrLf & _
                            "AutoClickerAmbiguousContained: a higher-generation inert preload was " & _
                            "installed, so any untracked registration from the ambiguous dispatch " & _
                            "cannot click in future realms."
            Else
                MarkAutoClickerContainmentRisk registrationKey
                finalDesc = finalDesc & vbCrLf & _
                            "AutoClickerAmbiguousUncontained: browser-side registration state may " & _
                            "remain active because the containment preload could not be confirmed." & _
                            AppendErrorDescription(containmentErrDesc)
            End If
        End If

        Dim addRaiseNum As Long
        addRaiseNum = addErrNum
        If addRaiseNum = 0 Then addRaiseNum = ERR_BIDI_PROTOCOL

        Dim addRaiseSource As String
        addRaiseSource = addErrSource
        If Len(addRaiseSource) = 0 Then addRaiseSource = "BiDi:" & callerName

        Err.Raise addRaiseNum, addRaiseSource, finalDesc
    End If

    m_autoClickerScriptIds.Add registrationKey, scriptId
    m_autoClickerTimeouts.Add registrationKey, vTimeoutMs
    m_autoClickerGenerations.Add registrationKey, generation
    If m_autoClickerContainmentRisks.Exists(registrationKey) Then _
        m_autoClickerContainmentRisks.Remove registrationKey

    BestEffortRemoveKnownAutoClickerSuppressor registrationKey

    If m_debugMode Then
        Debug.Print "BiDi: AutoClicker registered for " & AutoClickerLocatorSummary(registrationKey) & _
                    " (ID: " & scriptId & ", timeoutMs=" & CStr(vTimeoutMs) & _
                    ", generation=" & AutoClickerGenerationLiteral(generation) & ")"
    End If
End Sub

Public Function ExecuteUnregisterAutoClickerByXPath(ByVal xpath As String) As Boolean
    ExecuteUnregisterAutoClickerByXPath = ExecuteUnregisterAutoClickerCore( _
        BuildAutoClickerXPathKey(xpath), "ExecuteUnregisterAutoClickerByXPath")
End Function

Public Function ExecuteUnregisterShadowAutoClicker(ByVal selectorPath As Variant) As Boolean
    ExecuteUnregisterShadowAutoClicker = ExecuteUnregisterAutoClickerCore( _
        BuildAutoClickerShadowKey(selectorPath, "ExecuteUnregisterShadowAutoClicker"), _
        "ExecuteUnregisterShadowAutoClicker")
End Function

Private Function ExecuteUnregisterAutoClickerCore(ByVal registrationKey As String, _
                                                  ByVal callerName As String) As Boolean
    EnsureAutoClickerRegistry

    Dim hadTrackedRegistration As Boolean
    Dim hadContainmentRisk As Boolean
    hadTrackedRegistration = m_autoClickerScriptIds.Exists(registrationKey)
    hadContainmentRisk = m_autoClickerContainmentRisks.Exists(registrationKey)

    If Not hadTrackedRegistration And Not hadContainmentRisk Then Exit Function

    If hadTrackedRegistration Then
        Dim scriptId As String
        scriptId = CStr(m_autoClickerScriptIds(registrationKey))

        Dim errNum As Long
        Dim errSource As String
        Dim errDesc As String
        If Not TryRemovePreloadScript(scriptId, True, errNum, errSource, errDesc) Then
            Dim removeRaiseNum As Long
            removeRaiseNum = errNum
            If removeRaiseNum = 0 Then removeRaiseNum = ERR_BIDI_PROTOCOL

            Err.Raise removeRaiseNum, "BiDi:" & callerName, _
                      "AutoClicker tracked preload removal could not be confirmed. " & _
                      AutoClickerLocatorSummary(registrationKey) & AppendErrorDescription(errDesc)
        End If

        RemoveTrackedAutoClickerRegistration registrationKey
    End If

    Dim suppressorId As String
    Dim suppressErrNum As Long
    Dim suppressErrSource As String
    Dim suppressErrDesc As String
    If Not TryInstallAutoClickerSuppressor(registrationKey, suppressorId, _
                                           suppressErrNum, suppressErrSource, suppressErrDesc) Then
        MarkAutoClickerContainmentRisk registrationKey

        Dim suppressRaiseNum As Long
        suppressRaiseNum = suppressErrNum
        If suppressRaiseNum = 0 Then suppressRaiseNum = ERR_BIDI_PROTOCOL

        Err.Raise suppressRaiseNum, "BiDi:" & callerName, _
                  "AutoClicker tracked preload was removed, but residual-orphan containment could " & _
                  "not be confirmed. " & AutoClickerLocatorSummary(registrationKey) & _
                  AppendErrorDescription(suppressErrDesc)
    End If

    ExecuteUnregisterAutoClickerCore = True

    If m_debugMode Then
        Debug.Print "BiDi: AutoClicker unregistered/contained for " & _
                    AutoClickerLocatorSummary(registrationKey) & _
                    " (suppressor ID: " & suppressorId & ")"
    End If
End Function

Private Function AppendErrorDescription(ByVal description As String) As String
    If Len(description) > 0 Then AppendErrorDescription = vbCrLf & description
End Function

Private Function BuildAutoClickerXPathKey(ByVal xpath As String) As String
    BuildAutoClickerXPathKey = "X|" & xpath
End Function

Private Function BuildAutoClickerShadowKey(ByVal selectorPath As Variant, _
                                           ByVal callerName As String) As String
    BuildAutoClickerShadowKey = "S|" & BuildShadowSelectorJson(selectorPath, callerName)
End Function

Private Function BuildShadowSelectorJson(ByVal selectorPath As Variant, _
                                         ByVal callerName As String) As String
    Dim result As String
    result = "["

    Dim itemCount As Long
    Dim selector As String
    Dim i As Long

    If IsArray(selectorPath) Then
        For i = LBound(selectorPath) To UBound(selectorPath)
            selector = CStr(selectorPath(i))
            ValidateShadowSelector selector, callerName
            If itemCount > 0 Then result = result & ","
            result = result & WebJsonConverter.ConvertToJson(selector)
            itemCount = itemCount + 1
        Next i
    Else
        selector = CStr(selectorPath)
        ValidateShadowSelector selector, callerName
        result = result & WebJsonConverter.ConvertToJson(selector)
        itemCount = 1
    End If

    If itemCount = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Shadow selector path must contain at least one selector."
    End If

    BuildShadowSelectorJson = result & "]"
End Function

Private Sub ValidateShadowSelector(ByVal selector As String, _
                                              ByVal callerName As String)
    If Len(selector) = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Shadow selector path contains an empty selector."
    End If

    If InStr(selector, Chr$(0)) > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                  "Shadow selector contains a null byte."
    End If
End Sub

Private Function AutoClickerLocatorSummary(ByVal registrationKey As String) As String
    If Left$(registrationKey, 2) = "X|" Then
        AutoClickerLocatorSummary = "XPath=" & Mid$(registrationKey, 3)
    ElseIf Left$(registrationKey, 2) = "S|" Then
        AutoClickerLocatorSummary = "ShadowPath=" & Mid$(registrationKey, 3)
    Else
        AutoClickerLocatorSummary = "LocatorKey=" & registrationKey
    End If
End Function

Private Function AutoClickerSignalLocator(ByVal registrationKey As String) As String
    If Left$(registrationKey, 2) = "X|" Then
        AutoClickerSignalLocator = "xpath=" & Mid$(registrationKey, 3)
    ElseIf Left$(registrationKey, 2) = "S|" Then
        AutoClickerSignalLocator = "shadowPath=" & Mid$(registrationKey, 3)
    Else
        AutoClickerSignalLocator = "locatorKey=" & registrationKey
    End If
End Function

Private Sub EnsureAutoClickerRegistry()
    If m_autoClickerScriptIds Is Nothing Then Set m_autoClickerScriptIds = New Dictionary
    If m_autoClickerTimeouts Is Nothing Then Set m_autoClickerTimeouts = New Dictionary
    If m_autoClickerGenerations Is Nothing Then Set m_autoClickerGenerations = New Dictionary
    If m_autoClickerSuppressorScriptIds Is Nothing Then Set m_autoClickerSuppressorScriptIds = New Dictionary
    If m_autoClickerContainmentRisks Is Nothing Then Set m_autoClickerContainmentRisks = New Dictionary
End Sub

Private Sub RemoveTrackedAutoClickerRegistration(ByVal registrationKey As String)
    If m_autoClickerScriptIds.Exists(registrationKey) Then m_autoClickerScriptIds.Remove registrationKey
    If m_autoClickerTimeouts.Exists(registrationKey) Then m_autoClickerTimeouts.Remove registrationKey
    If m_autoClickerGenerations.Exists(registrationKey) Then m_autoClickerGenerations.Remove registrationKey
End Sub

Private Sub MarkAutoClickerContainmentRisk(ByVal registrationKey As String)
    If Not m_autoClickerContainmentRisks.Exists(registrationKey) Then
        m_autoClickerContainmentRisks.Add registrationKey, True
    End If
End Sub

' QueryPerformanceCounter is system-wide and monotonic across wrapper instances. Convert its
' millisecond value to an integer-like microsecond generation; the local floor guarantees strict
' increase even when two registrations happen inside one timer quantum.
Private Function NextAutoClickerGuardGeneration() As Double
    Dim candidate As Double
    candidate = Fix(socket_.nowMs() * 1000#)

    If candidate <= m_lastAutoClickerGuardGeneration Then
        candidate = m_lastAutoClickerGuardGeneration + 1#
    End If

    m_lastAutoClickerGuardGeneration = candidate
    NextAutoClickerGuardGeneration = candidate
End Function

Private Function AutoClickerGenerationLiteral(ByVal generation As Double) As String
    AutoClickerGenerationLiteral = Format$(generation, "0")
End Function

Private Function BuildAutoClickerPreloadSource(ByVal registrationKey As String, _
                                               ByVal timeoutMs As Long, _
                                               ByVal generation As Double, _
                                               ByVal enabled As Boolean) As String
    Dim locatorMode As String
    Dim locatorJson As String
    Dim guardKeyJson As String

    If Left$(registrationKey, 2) = "X|" Then
        locatorMode = "xpath"
        locatorJson = WebJsonConverter.ConvertToJson(Mid$(registrationKey, 3))
        ' Keep the historical raw-XPath browser guard key so older XPath preloads remain containable.
        guardKeyJson = locatorJson
    ElseIf Left$(registrationKey, 2) = "S|" Then
        locatorMode = "shadow"
        locatorJson = Mid$(registrationKey, 3)
        guardKeyJson = WebJsonConverter.ConvertToJson("shadow|" & locatorJson)
    Else
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:BuildAutoClickerPreloadSource", _
                  "Invalid AutoClicker registration key."
    End If

    Dim enabledLiteral As String
    If enabled Then
        enabledLiteral = "true"
    Else
        enabledLiteral = "false"
    End If

    Dim sourceJs As String
    sourceJs = GetActionJs("autoclicker")
    sourceJs = Replace(sourceJs, "{{KEY}}", guardKeyJson)
    sourceJs = Replace(sourceJs, "{{MODE}}", WebJsonConverter.ConvertToJson(locatorMode))
    sourceJs = Replace(sourceJs, "{{LOCATOR}}", locatorJson)
    sourceJs = Replace(sourceJs, "{{TIMEOUT}}", CStr(timeoutMs))
    sourceJs = Replace(sourceJs, "{{GENERATION}}", AutoClickerGenerationLiteral(generation))
    sourceJs = Replace(sourceJs, "{{ENABLED}}", enabledLiteral)
    BuildAutoClickerPreloadSource = sourceJs
End Function

Private Function TryAddAutoClickerPreload(ByVal registrationKey As String, _
                                          ByVal timeoutMs As Long, _
                                          ByVal generation As Double, _
                                          ByVal enabled As Boolean, _
                                          ByRef scriptId As String, _
                                          ByRef failureNumber As Long, _
                                          ByRef failureSource As String, _
                                          ByRef failureDescription As String) As Boolean
    scriptId = vbNullString
    failureNumber = 0
    failureSource = vbNullString
    failureDescription = vbNullString

    On Error GoTo Failed

    Dim params As New Dictionary
    params.Add "functionDeclaration", _
               BuildAutoClickerPreloadSource(registrationKey, timeoutMs, generation, enabled)

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True, _
                                allowAmbiguousRetry:=False)

    Dim json As Object
    Set json = ValidateProtocolResponse(strRes, "script.addPreloadScript")

    scriptId = GetPathString(json, "result.script")
    If Len(scriptId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:script.addPreloadScript", _
                  "AutoClicker preload registration succeeded without returning result.script."
    End If

    TryAddAutoClickerPreload = True
    Exit Function

Failed:
    failureNumber = Err.Number
    failureSource = Err.source
    failureDescription = Err.description
    Err.Clear
End Function

Private Function TryInstallAutoClickerSuppressor(ByVal registrationKey As String, _
                                                 ByRef scriptId As String, _
                                                 ByRef failureNumber As Long, _
                                                 ByRef failureSource As String, _
                                                 ByRef failureDescription As String) As Boolean
    EnsureAutoClickerRegistry

    Dim previousSuppressorId As String
    If m_autoClickerSuppressorScriptIds.Exists(registrationKey) Then
        previousSuppressorId = CStr(m_autoClickerSuppressorScriptIds(registrationKey))
    End If

    Dim generation As Double
    generation = NextAutoClickerGuardGeneration()

    If Not TryAddAutoClickerPreload(registrationKey, 0, generation, False, scriptId, _
                                    failureNumber, failureSource, failureDescription) Then
        MarkAutoClickerContainmentRisk registrationKey
        LogSignalLine "[AUTOCLICKER-CONTAINMENT-FAILED] " & _
                      AutoClickerSignalLocator(registrationKey) & _
                      " | generation=" & AutoClickerGenerationLiteral(generation) & _
                      " | error=" & CStr(failureNumber)
        Exit Function
    End If

    If m_autoClickerSuppressorScriptIds.Exists(registrationKey) Then
        m_autoClickerSuppressorScriptIds(registrationKey) = scriptId
    Else
        m_autoClickerSuppressorScriptIds.Add registrationKey, scriptId
    End If

    If m_autoClickerContainmentRisks.Exists(registrationKey) Then _
        m_autoClickerContainmentRisks.Remove registrationKey

    If Len(previousSuppressorId) > 0 And _
       StrComp(previousSuppressorId, scriptId, vbBinaryCompare) <> 0 Then
        Dim oldErrNum As Long
        Dim oldErrSource As String
        Dim oldErrDesc As String
        Call TryRemovePreloadScript(previousSuppressorId, False, oldErrNum, oldErrSource, oldErrDesc)
    End If

    LogSignalLine "[AUTOCLICKER-CONTAINED] " & AutoClickerSignalLocator(registrationKey) & _
                  " | generation=" & AutoClickerGenerationLiteral(generation) & _
                  " | suppressorId=" & scriptId

    TryInstallAutoClickerSuppressor = True
End Function

Private Sub BestEffortRemoveKnownAutoClickerSuppressor(ByVal registrationKey As String)
    EnsureAutoClickerRegistry
    If Not m_autoClickerSuppressorScriptIds.Exists(registrationKey) Then Exit Sub

    Dim scriptId As String
    scriptId = CStr(m_autoClickerSuppressorScriptIds(registrationKey))

    Dim errNum As Long
    Dim errSource As String
    Dim errDesc As String
    If TryRemovePreloadScript(scriptId, False, errNum, errSource, errDesc) Then
        m_autoClickerSuppressorScriptIds.Remove registrationKey
    Else
        LogSignalLine "[AUTOCLICKER-SUPPRESSOR-RESIDUAL] " & _
                      AutoClickerSignalLocator(registrationKey) & _
                      " | suppressorId=" & scriptId & _
                      " | error=" & CStr(errNum)
    End If
End Sub

Private Function TryRemovePreloadScript(ByVal scriptId As String, _
                                        ByVal allowRecoveryRetry As Boolean, _
                                        ByRef failureNumber As Long, _
                                        ByRef failureSource As String, _
                                        ByRef failureDescription As String) As Boolean
    failureNumber = 0
    failureSource = vbNullString
    failureDescription = vbNullString

    On Error GoTo Failed

    Dim params As New Dictionary
    params.Add "script", scriptId

    ' Removal is idempotent for wrapper purposes: if an ambiguous first attempt actually
    ' succeeded, a retry can legitimately return "no such script", which is treated below
    ' as confirmation that the preload is absent.
    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.removePreloadScript", params, True, _
                                allowAmbiguousRetry:=True, _
                                allowRecoveryRetry:=allowRecoveryRetry, _
                                allowContextRetarget:=False)

    Dim json As Object
    Set json = WebJsonConverter.ParseJson(strRes)
    If json Is Nothing Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:script.removePreloadScript", _
                  "script.removePreloadScript returned malformed JSON."
    End If

    If json.Exists("error") Then
        If StrComp(LCase$(CStr(json("error"))), "no such script", vbBinaryCompare) = 0 Then
            TryRemovePreloadScript = True
            Exit Function
        End If
    End If

    Call ValidateProtocolResponse(strRes, "script.removePreloadScript")
    TryRemovePreloadScript = True
    Exit Function

Failed:
    failureNumber = Err.Number
    failureSource = Err.source
    failureDescription = Err.description
    Err.Clear
End Function

' Shutdown cleanup removes each known active preload, then leaves a higher-generation inert
' suppressor in the browser session. Suppressors are intentionally not removed during wrapper
' shutdown: they are the containment mechanism for any older preload whose script ID was lost.
' The browser session itself destroys all preload scripts when WebDriver ends.
Private Function ReleaseAllAutoClickersBestEffort() As Long
    EnsureAutoClickerRegistry

    Dim work As New Dictionary
    Dim key As Variant
    For Each key In m_autoClickerScriptIds.keys
        If Not work.Exists(CStr(key)) Then work.Add CStr(key), True
    Next key
    For Each key In m_autoClickerContainmentRisks.keys
        If Not work.Exists(CStr(key)) Then work.Add CStr(key), True
    Next key

    For Each key In work.keys
        Dim registrationKey As String
        registrationKey = CStr(key)

        Dim removalFailed As Boolean
        If m_autoClickerScriptIds.Exists(registrationKey) Then
            Dim activeId As String
            activeId = CStr(m_autoClickerScriptIds(registrationKey))

            Dim removeErrNum As Long
            Dim removeErrSource As String
            Dim removeErrDesc As String
            If TryRemovePreloadScript(activeId, False, removeErrNum, removeErrSource, removeErrDesc) Then
                RemoveTrackedAutoClickerRegistration registrationKey
            Else
                removalFailed = True
                If m_debugMode Then
                    Debug.Print "BiDi CRITICAL: AutoClicker tracked preload cleanup could not be confirmed. " & _
                                AutoClickerLocatorSummary(registrationKey) & " | ID=" & activeId & _
                                AppendErrorDescription(Replace(Replace(removeErrDesc, vbCr, " "), vbLf, " "))
                End If
            End If
        End If

        Dim suppressorId As String
        Dim suppressErrNum As Long
        Dim suppressErrSource As String
        Dim suppressErrDesc As String
        If Not TryInstallAutoClickerSuppressor(registrationKey, suppressorId, _
                                               suppressErrNum, suppressErrSource, suppressErrDesc) Then
            ReleaseAllAutoClickersBestEffort = ReleaseAllAutoClickersBestEffort + 1
        ElseIf removalFailed Then
            ReleaseAllAutoClickersBestEffort = ReleaseAllAutoClickersBestEffort + 1
        End If
    Next key
End Function

Private Function BuildShadowSelectorArg(ByVal selectorsArray As Variant) As Dictionary
    Dim values As New Collection
    Dim i As Long
    Dim selector As String

    If IsArray(selectorsArray) Then
        For i = LBound(selectorsArray) To UBound(selectorsArray)
            selector = CStr(selectorsArray(i))
            ValidateShadowSelector selector, "BuildShadowSelectorArg"
            values.Add NewLocalValue("string", selector)
        Next i
    Else
        selector = CStr(selectorsArray)
        ValidateShadowSelector selector, "BuildShadowSelectorArg"
        values.Add NewLocalValue("string", selector)
    End If

    If values.count = 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:BuildShadowSelectorArg", _
                  "Shadow selector path must contain at least one selector."
    End If

    Dim arg As New Dictionary
    arg.Add "type", "array"
    arg.Add "value", values
    Set BuildShadowSelectorArg = arg
End Function

Public Sub StartDiscoveryLog(Optional ByVal excludeImagesAndCss As Boolean = True, _
                             Optional ByVal captureDomActivity As Boolean = True)
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
    m_discoveryDomCaptureEnabled = captureDomActivity
    m_discoveryManualPumpActive = False
    m_discoveryDomContextId = vbNullString
    m_nextDiscoveryDomPollMs = 0#

    ' A reused wrapper starts recording with all carryover diagnostics cleared.
    m_unresolvedJsonThisWait = False
    m_unresolvedJsonFirstKeyThisWait = ""
    m_unresolvedJsonFirstSecThisWait = 0#
    m_hasCarryoverHintForNextWait = False
    m_carryoverNoteArmedThisWait = False
    m_carryoverNoteEmittedThisWait = False

    m_isRecording = True

    ' Prime manual DOM capture before control returns to the caller. Querying first drains any
    ' stale pre-recording changeList entries; if no healthy probe exists, the same call injects
    ' one. This is best-effort: network/lifecycle recording remains useful even if a page realm
    ' temporarily refuses the probe.
    If m_discoveryDomCaptureEnabled Then
        Dim probeReady As Boolean
        probeReady = CaptureDiscoverySpaActivity(m_discoveryDomContextId, discardChanges:=True)
        ScheduleNextDiscoveryDomPoll probeReady
    End If

    Debug.Print "BiDi Recorder: Started. (Filter Resources: " & excludeImagesAndCss & _
                ", DOM Activity: " & captureDomActivity & ")"
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
    ' while that flag is True. RecordEventsForSeconds owns the final manual DOM sample; Stop does
    ' not consume the SPA probe outside that explicit manual recording window.
    Call ProcessRecorderQueue

    m_isRecording = False

    If m_recorderLogs Is Nothing Then
        m_discoveryDomCaptureEnabled = False
        m_discoveryManualPumpActive = False
        m_discoveryDomContextId = vbNullString
        m_nextDiscoveryDomPollMs = 0#
        Exit Sub
    End If

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
                      " | excludeResources=" & m_excludeResources & _
                      " | manualDomCapture=" & m_discoveryDomCaptureEnabled & vbCrLf
        ' Ring-buffer head eviction is reported so a truncated timeline is never mistaken for
        ' a complete capture.
        If m_droppedHeadCount > 0 Then
            stm.WriteText "TRUNCATED | dropped " & m_droppedHeadCount & _
                          " oldest timeline line(s); only the most recent " & MAX_LOG_SIZE & _
                          " event(s) are shown. Earlier events existed but were evicted (ring buffer)." & vbCrLf
        End If
        stm.WriteText "LEGEND | [RES] carries: id=<reqId; a redirect chain reuses the SAME id, pairing its hops> key=<host+path, volatile query stripped, rpcids kept> idle=<skip|track: was this URL in the idle probes ignore set>" & vbCrLf
        stm.WriteText "LEGEND | [REQ]=request start (network.beforeRequestSent; recording-only, not wired to consensus). [REQ] pairs with [RES] on the shared id. Per-request latency (dur=) is not computed." & vbCrLf
        stm.WriteText "LEGEND | [SPA-ACT]=DOM mutation burst (manual RecordEventsForSeconds samples this by default)  [SPA-STABLE]=consensus  [SPA-WAIT-BEGIN/END]=one wait (END has outcome + slackMs)  [SUPPRESSED]=filtered-noise tally" & vbCrLf
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
    m_discoveryDomCaptureEnabled = False
    m_discoveryManualPumpActive = False
    m_discoveryDomContextId = vbNullString
    m_nextDiscoveryDomPollMs = 0#
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

    ' Only this explicit manual pump is allowed to drain the SPA probe for Discovery DOM logs.
    ' Other ProcessRecorderQueue callers (download/navigation/file-dialog/etc.) therefore retain
    ' the pre-existing SPA diagnostic behavior even while Discovery recording is enabled.
    m_discoveryManualPumpActive = (m_isRecording And m_discoveryDomCaptureEnabled)

    Do While (socket_.nowMs() - startMs) < durationMs
        ProcessRecorderQueue

        If durationMs > 5000 Then
            socket_.PumpMessages 100
            socket_.YieldExecution 100, False
        Else
            socket_.YieldExecution 100, True
        End If
    Loop

Cleanup:
    ' Capture the tail while the manual-pump guard is still active, then release it even on an
    ' error. These calls are best-effort and must not replace the original recorder error.
    On Error Resume Next
    If m_discoveryManualPumpActive Then
        ProcessRecorderQueueCore False
        PollDiscoverySpaActivityIfDue forceSample:=True
    End If
    m_discoveryManualPumpActive = False
    On Error GoTo 0

    ' RecordEventsForSeconds does not own m_isProcessingQueue.
    ' ProcessRecorderQueueCore is the sole normal owner and releases the lock in its own Cleanup.
    Exit Sub

ErrorHandler:
    ' Use Resume before Cleanup to deactivate VBA's active error handler; otherwise an error
    ' inside Cleanup can terminate the call stack instead of being handled.
    Debug.Print "BiDi Warning in RecordEventsForSeconds: " & Err.description
    Resume Cleanup
End Sub

Public Sub ProcessRecorderQueue()
    ProcessRecorderQueueCore False

    ' Manual Discovery DOM sampling is intentionally outside ProcessRecorderQueueCore and is
    ' gated to RecordEventsForSeconds. Querying the SPA probe calls back through
    ' CallScript/ResolveTargetContext, which may non-blockingly drain transport events; keeping
    ' the probe poll outside m_isProcessingQueue avoids recursion and preserves one queue owner.
    PollDiscoverySpaActivityIfDue
End Sub

Private Sub DrainNonBlockingRecorderQueue()
    ' Internal drains must remain transport-only. In particular, QuerySpaProbe -> CallScript ->
    ' ResolveTargetContext can enter here; polling Discovery DOM again would recurse.
    ProcessRecorderQueueCore True
End Sub

Private Sub ScheduleNextDiscoveryDomPoll(ByVal probeAvailable As Boolean)
    If socket_ Is Nothing Then
        m_nextDiscoveryDomPollMs = 0#
        Exit Sub
    End If

    If probeAvailable Then
        m_nextDiscoveryDomPollMs = socket_.nowMs() + DISCOVERY_DOM_POLL_MS
    Else
        m_nextDiscoveryDomPollMs = socket_.nowMs() + DISCOVERY_PROBE_RETRY_MS
    End If
End Sub

Private Sub PollDiscoverySpaActivityIfDue(Optional ByVal forceSample As Boolean = False)
    If Not m_isRecording Then Exit Sub
    If Not m_discoveryDomCaptureEnabled Then Exit Sub
    If Not m_discoveryManualPumpActive Then Exit Sub

    ' Never drain the shared SPA probe changeList while MasterWaitEngine owns a wait. The normal
    ' SPA-wait path must remain the sole consumer in that case so consensus/render diagnostics
    ' are bit-for-bit unaffected by enabling Discovery recording.
    If m_waitStackDepth > 0 Then Exit Sub
    If socket_ Is Nothing Then Exit Sub

    Dim nowMs As Double: nowMs = socket_.nowMs()
    If Not forceSample Then
        If m_nextDiscoveryDomPollMs > 0# And nowMs < m_nextDiscoveryDomPollMs Then Exit Sub
    End If

    Dim probeReady As Boolean
    probeReady = CaptureDiscoverySpaActivity(m_discoveryDomContextId)
    ScheduleNextDiscoveryDomPoll probeReady
End Sub

' Poll the existing SPA probe for manual Discovery recording. QuerySpaProbe drains changeList
' even from an unhealthy probe, so changes are consumed before recovery. discardChanges is used
' only when StartDiscoveryLog establishes the baseline; it prevents pre-recording mutations from
' being attributed to the first manual action. This routine never evaluates SPA consensus.
Private Function CaptureDiscoverySpaActivity(ByRef targetCtx As String, _
                                             Optional ByVal discardChanges As Boolean = False) As Boolean
    On Error GoTo ProbeUnavailable

    If Len(targetCtx) = 0 Then targetCtx = ResolveTargetContext(vbNullString)
    If Len(targetCtx) = 0 Then GoTo ProbeUnavailable

    Dim probe As Dictionary
    Set probe = QuerySpaProbe(targetCtx)

    If Not discardChanges Then ProcessDiscoverySpaProbeChanges probe
    LogInvalidProbeSelectors probe, targetCtx

    If IsSpaProbeReady(probe) Then
        If Not IsProbeFilterCurrent(targetCtx) Then
            If Not PushIdleFiltersToProbe(targetCtx) Then
                If Not StartSpaIdleProbe(targetCtx) Then GoTo ProbeUnavailable
            End If
        End If

        CaptureDiscoverySpaActivity = True
        Exit Function
    End If

    ' Full navigation destroys the page-world observer. Re-inject against the same browsing
    ' context. If the context itself disappeared, the error path clears targetCtx so the next
    ' recorder pump resolves the current main context instead.
    If StartSpaIdleProbe(targetCtx) Then
        CaptureDiscoverySpaActivity = True
        Exit Function
    End If

ProbeUnavailable:
    If Err.Number <> 0 Then
        If m_debugMode Then
            Debug.Print "BiDi Discovery: SPA probe unavailable: " & Err.Number & " " & Err.description
        End If
        Err.Clear
    End If
    targetCtx = vbNullString
    CaptureDiscoverySpaActivity = False
End Function

' Manual Discovery has different semantics from ProcessSpaProbeChangesForDiagnostics: this is a
' pure recorder sink. It MUST NOT clear unresolved-json/carryover state or influence consensus
' merely because the user opened a manual recording window.
Private Sub ProcessDiscoverySpaProbeChanges(ByVal probe As Dictionary)
    On Error GoTo Done
    If Not m_isRecording Then Exit Sub
    If probe Is Nothing Then Exit Sub
    If Not probe.Exists("version") Then Exit Sub
    If CLng(probe("version")) <> SPA_PROBE_VERSION Then Exit Sub
    If Not probe.Exists("changes") Then Exit Sub
    If Not IsObject(probe("changes")) Then Exit Sub

    Dim changes As Collection
    Set changes = probe("changes")
    If changes Is Nothing Then Exit Sub
    If changes.count = 0 Then Exit Sub

    Dim v As Variant
    Dim logSummary As String
    For Each v In changes
        logSummary = logSummary & v & " "
    Next v
    logSummary = Trim$(logSummary)

    If Len(logSummary) > 500 Then logSummary = Left$(logSummary, 497) & "..."
    If m_debugMode Then Debug.Print "[BiDi Discovery SPA Activity] " & logSummary

    Dim relTime As String
    relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
    PushTimelineLine "[" & relTime & "] [SPA-ACT] " & logSummary

Done:
    Err.Clear
End Sub

Private Sub IncrementDispatchLossCounter(ByRef counter As Long)
    If counter < 2147483647 Then counter = counter + 1
End Sub

' The communicator owns transport/FIFO/parse loss. This records the distinct case where an
' event was parsed successfully but its upper-layer handler failed after FIFO consumption.
Private Sub RecordDispatchEventLoss(ByVal json As Object, ByRef detectedMethod As String)
    detectedMethod = vbNullString
    IncrementDispatchLossCounter m_dispatchEventLossCount

    If json Is Nothing Then
        IncrementDispatchLossCounter m_dispatchUnknownCriticalEventLossCount
        Exit Sub
    End If

    If TypeName(json) <> "Dictionary" Then
        IncrementDispatchLossCounter m_dispatchUnknownCriticalEventLossCount
        Exit Sub
    End If

    If Not json.Exists("method") Then
        IncrementDispatchLossCounter m_dispatchUnknownCriticalEventLossCount
        Exit Sub
    End If

    If IsObject(json("method")) Or IsNull(json("method")) Or IsEmpty(json("method")) Then
        IncrementDispatchLossCounter m_dispatchUnknownCriticalEventLossCount
        Exit Sub
    End If

    detectedMethod = CStr(json("method"))
    If Len(detectedMethod) = 0 Then
        IncrementDispatchLossCounter m_dispatchUnknownCriticalEventLossCount
        Exit Sub
    End If

    Select Case detectedMethod
        Case "browsingContext.contextCreated", _
             "browsingContext.contextDestroyed", _
             "browsingContext.navigationStarted"
            IncrementDispatchLossCounter m_dispatchLifecycleEventLossCount

        Case "browsingContext.downloadWillBegin", _
             "browsingContext.downloadEnd"
            IncrementDispatchLossCounter m_dispatchDownloadEventLossCount

        Case "input.fileDialogOpened"
            IncrementDispatchLossCounter m_dispatchFileDialogEventLossCount
    End Select
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
        messageConsumed = socket_.TryGetParsedEvent(json, dropReason, MIN_MSG_SIZE, socket_.ParsedAsyncEventMaxChars, nonBlocking)
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

            Dim dispatchErrNum As Long
            Dim dispatchErrDesc As String
            dispatchErrNum = Err.Number
            dispatchErrDesc = Err.description
            Err.Clear

            If dispatchErrNum <> 0 Then
                Dim dispatchMethod As String
                RecordDispatchEventLoss json, dispatchMethod

                consecutiveErrors = consecutiveErrors + 1
                If m_debugMode Then
                    Debug.Print "BiDi: Dispatch error #" & consecutiveErrors & _
                                " | method=" & IIf(Len(dispatchMethod) > 0, dispatchMethod, "(unknown)") & _
                                " | Err " & CStr(dispatchErrNum) & ": " & dispatchErrDesc
                End If
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
    If Not health.Exists("shadowObserverReady") Then Exit Function
    If Not health.Exists("inputWrapped") Then Exit Function

    IsSpaProbeReady = CBool(health("versionOk")) And _
                      CBool(health("fetchWrapped")) And _
                      CBool(health("xhrWrapped")) And _
                      CBool(health("observerActive")) And _
                      CBool(health("shadowObserverReady")) And _
                      CBool(health("inputWrapped"))
    Exit Function

NotReady:
    IsSpaProbeReady = False
End Function

Private Function ProbeHealthFlagText(ByVal health As Object, ByVal key As String) As String
    On Error GoTo Missing
    If health Is Nothing Then GoTo Missing
    If Not health.Exists(key) Then GoTo Missing
    ProbeHealthFlagText = LCase$(CStr(CBool(health(key))))
    Exit Function
Missing:
    ProbeHealthFlagText = "missing"
    Err.Clear
End Function

Private Function BuildSpaProbeFalseFlags(ByVal health As Object) As String
    Dim keys As Variant
    keys = Array("versionOk", "fetchWrapped", "xhrWrapped", "observerActive", "shadowObserverReady", "inputWrapped")

    Dim i As Long
    Dim flagText As String
    Dim s As String
    For i = LBound(keys) To UBound(keys)
        flagText = ProbeHealthFlagText(health, CStr(keys(i)))
        If flagText <> "true" Then
            If Len(s) > 0 Then s = s & ","
            s = s & CStr(keys(i))
        End If
    Next i

    BuildSpaProbeFalseFlags = s
End Function

Private Function CompactProbeDiagnosticText(ByVal value As String, _
                                            Optional ByVal maxChars As Long = 240) As String
    Dim s As String
    s = Replace(value, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    If Len(s) > maxChars Then s = Left$(s, maxChars - 3) & "..."
    CompactProbeDiagnosticText = s
End Function

Private Function BuildSpaProbeHealthSummary(ByVal probe As Dictionary) As String
    On Error GoTo Malformed

    If probe Is Nothing Then
        BuildSpaProbeHealthSummary = "probe=missing | falseFlags=probeMissing"
        Exit Function
    End If

    Dim health As Object
    If probe.Exists("health") Then
        If IsObject(probe("health")) Then Set health = probe("health")
    End If

    Dim falseFlags As String
    falseFlags = BuildSpaProbeFalseFlags(health)
    If Len(falseFlags) = 0 Then falseFlags = "(none)"

    Dim s As String
    s = "versionOk=" & ProbeHealthFlagText(health, "versionOk") & _
        " | fetchWrapped=" & ProbeHealthFlagText(health, "fetchWrapped") & _
        " | xhrWrapped=" & ProbeHealthFlagText(health, "xhrWrapped") & _
        " | observerActive=" & ProbeHealthFlagText(health, "observerActive") & _
        " | shadowObserverReady=" & ProbeHealthFlagText(health, "shadowObserverReady") & _
        " | inputWrapped=" & ProbeHealthFlagText(health, "inputWrapped") & _
        " | falseFlags=" & falseFlags

    Dim lastHealthError As String
    If probe.Exists("lastHealthError") Then
        If Not IsObject(probe("lastHealthError")) Then
            lastHealthError = CStr(probe("lastHealthError") & "")
        End If
    End If
    lastHealthError = CompactProbeDiagnosticText(lastHealthError)
    If Len(lastHealthError) > 0 Then s = s & " | lastHealthError=" & lastHealthError

    BuildSpaProbeHealthSummary = s
    Exit Function

Malformed:
    BuildSpaProbeHealthSummary = "probe=malformed | falseFlags=probeMalformed"
    Err.Clear
End Function

Private Sub TrackSpaProbeUnhealthy(ByVal probe As Dictionary, _
                                   ByRef ctx As Dictionary, _
                                   ByVal targetCtx As String, _
                                   Optional ByVal forceLog As Boolean = False)
    Dim summary As String
    summary = BuildSpaProbeHealthSummary(probe)

    If ctx.Exists("lastProbeHealthSummary") Then
        ctx("lastProbeHealthSummary") = summary
    Else
        ctx.Add "lastProbeHealthSummary", summary
    End If

    Dim previousSignature As String
    If ctx.Exists("lastProbeHealthSignature") Then
        previousSignature = CStr(ctx("lastProbeHealthSignature"))
    Else
        ctx.Add "lastProbeHealthSignature", vbNullString
    End If

    If forceLog Or StrComp(previousSignature, summary, vbBinaryCompare) <> 0 Then
        LogProbeDiagnosticLine "[SPA-PROBE-UNHEALTHY] " & summary & _
                               " | ctx=" & targetCtx & _
                               IIf(forceLog, " | final=true", "")
        ctx("lastProbeHealthSignature") = summary
    End If
End Sub

Private Sub ResetSpaProbeHealthBurst(ByRef ctx As Dictionary)
    If ctx.Exists("lastProbeHealthSignature") Then
        ctx("lastProbeHealthSignature") = vbNullString
    Else
        ctx.Add "lastProbeHealthSignature", vbNullString
    End If

    If ctx.Exists("lastProbeHealthSummary") Then
        ctx("lastProbeHealthSummary") = vbNullString
    Else
        ctx.Add "lastProbeHealthSummary", vbNullString
    End If
End Sub

' QuerySpaProbe drains p.changeList on every query, including an unhealthy one. Consume the
' current-version diagnostic batch before readiness decides whether the sample may influence
' consensus. This preserves [SPA-ACT] / render diagnostics without treating unhealthy metrics
' as stability evidence.
Private Sub ProcessSpaProbeChangesForDiagnostics(ByVal probe As Dictionary, _
                                                  ByRef ctx As Dictionary)
    On Error GoTo Done
    If probe Is Nothing Then Exit Sub
    If Not probe.Exists("version") Then Exit Sub
    If CLng(probe("version")) <> SPA_PROBE_VERSION Then Exit Sub
    If Not probe.Exists("changes") Then Exit Sub
    If Not IsObject(probe("changes")) Then Exit Sub

    Dim changes As Collection
    Set changes = probe("changes")
    If changes Is Nothing Then Exit Sub
    If changes.count = 0 Then Exit Sub

    Dim v As Variant
    Dim logSummary As String
    For Each v In changes
        logSummary = logSummary & v & " "
    Next v
    logSummary = Trim$(logSummary)

    If Len(logSummary) > 500 Then logSummary = Left$(logSummary, 497) & "..."
    If m_debugMode Then Debug.Print "[BiDi SPA Activity] " & logSummary

    If m_isRecording Then
        Dim relTime As String
        relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
        PushTimelineLine "[" & relTime & "] [SPA-ACT] " & logSummary
    End If

    ' A render observed during an unhealthy sample is diagnostic evidence only. It may clear
    ' the unresolved-json advisory state, but no timing/idle metric from this sample is used.
    Dim ce As Variant
    Dim renderXPath As String
    renderXPath = vbNullString
    For Each ce In changes
        If Left$(CStr(ce), 4) = "DOM:" Then
            If IsContentRenderPath(Mid$(CStr(ce), 5)) Then
                renderXPath = CStr(ce)
                Exit For
            End If
        End If
    Next ce

    If Len(renderXPath) > 0 Then
        m_unresolvedJsonThisWait = False
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

Done:
    Err.Clear
End Sub

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

' Arms a rewrite signal for an existing element inside nested Shadow DOM.
Public Sub ArmShadowContentSignal(ByVal selectorPath As Variant, _
                                  Optional ByVal includeAttributes As Boolean = False, _
                                  Optional ByVal contextId As String = "", _
                                  Optional ByVal requireUnlocker As Boolean = False)
    ValidateShadowUnlockerRequirement requireUnlocker, "ArmShadowContentSignal"

    Dim targetCtx As String
    targetCtx = ResolveSignalContextForArm(contextId, "ArmShadowContentSignal")

    Dim selectorJson As String
    selectorJson = BuildShadowSelectorJson(selectorPath, "ArmShadowContentSignal")

    Dim js As String
    js = "var p=window.__vbaIdleProbe;if(!p||!p.armShadowDomSignal)return -9;" & _
         "return p.armShadowDomSignal(" & selectorJson & "," & _
         LCase$(CStr(includeAttributes)) & ");"

    Dim rc As Long
    rc = InvokeProbeArm(js, targetCtx)

    Select Case rc
        Case 0
            ClaimSignalOwner targetCtx
            If m_sigDomArmed Then LogSignalLine "[SPA-SIGNAL-ARM] type=dom | overwrite=true (previous dom arm replaced)"
            m_sigDomArmed = True
            m_sigDomSeen = False
            m_sigDomXPath = "ShadowPath=" & selectorJson
            m_sigLostWarned = False
            LogSignalLine "[SPA-SIGNAL-ARM] type=dom | shadowPath=" & selectorJson & " | ctx=" & targetCtx
        Case -1
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=not-found | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowContentSignal", _
                "ArmShadowContentSignal: target element not found at arm time: " & selectorJson & vbCrLf & _
                "This method watches an EXISTING element for rewrites. To wait for an element " & _
                "that appears only after the action, use ArmShadowVisibilitySignal instead."
        Case -2
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=invalid-selector | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowContentSignal", _
                "ArmShadowContentSignal: invalid Shadow selector path: " & selectorJson
        Case Else
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=dom | reason=probe-unavailable | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowContentSignal", _
                "ArmShadowContentSignal: SPA probe unavailable in the target context (rc=" & rc & ")."
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

' Arms a visibility signal for an element that may appear later inside nested Shadow DOM.
Public Sub ArmShadowVisibilitySignal(ByVal selectorPath As Variant, _
                                     Optional ByVal contextId As String = "", _
                                     Optional ByVal requireUnlocker As Boolean = False)
    ValidateShadowUnlockerRequirement requireUnlocker, "ArmShadowVisibilitySignal"

    Dim targetCtx As String
    targetCtx = ResolveSignalContextForArm(contextId, "ArmShadowVisibilitySignal")

    Dim selectorJson As String
    selectorJson = BuildShadowSelectorJson(selectorPath, "ArmShadowVisibilitySignal")

    Dim js As String
    js = "var p=window.__vbaIdleProbe;if(!p||!p.armShadowVisSignal)return -9;" & _
         "return p.armShadowVisSignal(" & selectorJson & ");"

    Dim rc As Long
    rc = InvokeProbeArm(js, targetCtx)

    Select Case rc
        Case 0
            ClaimSignalOwner targetCtx
            If m_sigVisArmed Then LogSignalLine "[SPA-SIGNAL-ARM] type=vis | overwrite=true (previous vis arm replaced)"
            m_sigVisArmed = True
            m_sigVisSeen = False
            m_sigVisXPath = "ShadowPath=" & selectorJson
            m_sigLostWarned = False
            LogSignalLine "[SPA-SIGNAL-ARM] type=vis | shadowPath=" & selectorJson & " | ctx=" & targetCtx
        Case -3
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=already-visible | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowVisibilitySignal", _
                "ArmShadowVisibilitySignal: target is ALREADY visible at arm time: " & selectorJson & vbCrLf & _
                "Waiting for the appearance of something already visible is meaningless and " & _
                "would open the gate instantly. To wait for this element's content to be " & _
                "REWRITTEN, use ArmShadowContentSignal instead."
        Case -2
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=invalid-selector | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowVisibilitySignal", _
                "ArmShadowVisibilitySignal: invalid Shadow selector path: " & selectorJson
        Case Else
            LogSignalLine "[SPA-SIGNAL-ARM-FAIL] type=vis | reason=probe-unavailable | shadowPath=" & selectorJson
            Err.Raise ERR_ACTION_FAILED, "BiDi:ArmShadowVisibilitySignal", _
                "ArmShadowVisibilitySignal: SPA probe unavailable in the target context (rc=" & rc & ")."
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

' Fatal Reset clears more than the ordinary SPA completion arms. Build one concise
' pre-reset snapshot so diagnostics do not falsely imply that only SPA state was lost.
Private Function BuildFatalResetDroppedArmsDesc() As String
    Dim s As String
    Dim spaDesc As String

    spaDesc = BuildSignalEndDesc()
    If Len(spaDesc) > 0 Then s = "spa(" & spaDesc & ")"

    If m_fileDlgArmed Then
        s = s & IIf(Len(s) > 0, ",", "") & "fileDialog:armed"
    ElseIf m_fileDlgSeen Then
        s = s & IIf(Len(s) > 0, ",", "") & "fileDialog:seen"
    End If

    If m_downloadArmed Then
        s = s & IIf(Len(s) > 0, ",", "") & "download:armed"
    ElseIf m_downloadBatchActive And m_downloadStartedCount > 0 Then
        s = s & IIf(Len(s) > 0, ",", "") & "download:started"
    End If

    If m_newContextArmed Then
        s = s & IIf(Len(s) > 0, ",", "") & "newContext:armed"
    ElseIf m_newContextSeen Then
        s = s & IIf(Len(s) > 0, ",", "") & "newContext:seen"
    End If

    If m_navHttpCaptureArmed Then
        s = s & IIf(Len(s) > 0, ",", "") & "navigationHttp:armed"
    ElseIf m_navHttpCaptureInProgress Then
        s = s & IIf(Len(s) > 0, ",", "") & "navigationHttp:active"
    End If

    If Len(s) = 0 Then s = "(none)"
    BuildFatalResetDroppedArmsDesc = s
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

' Critical incident selection is deliberately narrow so this ring cannot become a second
' high-rate Discovery Log. It retains one-shot capture/fatal evidence even when recording is OFF.
Private Function IsForensicIncidentSignal(ByVal line As String) As Boolean
    IsForensicIncidentSignal = _
        (InStr(1, line, "[FATAL-RESET]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[DOWNLOAD-", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[FILE-DIALOG-", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[NEW-CONTEXT-", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[TRANSPORT-OVERSIZE]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[CORRELATION-TIMEOUT]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[AUTOCLICKER-AMBIGUOUS]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[AUTOCLICKER-CONTAINMENT-FAILED]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[AUTOCLICKER-SUPPRESSOR-RESIDUAL]", vbBinaryCompare) = 1 Or _
         InStr(1, line, "[SHUTDOWN-CAPTURE-ABANDON]", vbBinaryCompare) = 1)
End Function

Private Sub EnsureForensicIncidentLog()
    If m_forensicIncidentLogs Is Nothing Then Set m_forensicIncidentLogs = New Collection
End Sub

' Forensic logging is best-effort and must never change the outcome of the operation being
' diagnosed. The ring is bounded by count and each line is independently capped.
Private Sub PushForensicIncidentLine(ByVal line As String)
    On Error Resume Next
    EnsureForensicIncidentLog

    Dim storedLine As String
    storedLine = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & line
    If Len(storedLine) > MAX_FORENSIC_INCIDENT_LINE_CHARS Then
        storedLine = Left$(storedLine, MAX_FORENSIC_INCIDENT_LINE_CHARS - 14) & "...(truncated)"
    End If

    If m_forensicIncidentLogs.count >= MAX_FORENSIC_INCIDENT_LOG_SIZE Then
        m_forensicIncidentLogs.Remove 1
        m_forensicIncidentDroppedHeadCount = m_forensicIncidentDroppedHeadCount + 1
    End If
    m_forensicIncidentLogs.Add storedLine
    On Error GoTo 0
End Sub

' LogSignalLine remains the Discovery Log/developer-trace path, but selected critical markers are
' also copied to the independent forensic ring regardless of m_isRecording or DebugMode.
Private Sub LogSignalLine(ByVal line As String)
    If IsForensicIncidentSignal(line) Then PushForensicIncidentLine line

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

' Action advisories share the recorder/developer-trace path but do not alter action success.
Private Sub LogActionDiagnosticLine(ByVal line As String)
    If m_isRecording Then
        Dim relTime As String
        relTime = Format((socket_.nowMs() / 1000) - m_recorderStartTime, "0.00") & "s"
        PushTimelineLine "[" & relTime & "] " & line
    End If
    If m_debugMode Then Debug.Print "BiDi: " & line
End Sub

' Advisory-only postcondition check for the fast bulk-input path. Never raises into the
' completed action: failure to inspect is less important than preserving the original result.
Private Sub LogInputValueMismatchBestEffort(ByVal locatorKind As ElementLocatorKind, _
                                             ByVal locator As Variant, _
                                             ByVal expectedValue As String, _
                                             ByVal searchTimeoutMs As Long, _
                                             ByVal contextId As String, _
                                             ByVal callerName As String)
    On Error Resume Next
    Err.Clear

    Dim postSharedId As String
    postSharedId = ResolveElementSharedId( _
        locatorKind, locator, contextId, searchTimeoutMs, callerName & ":postcheck")

    Dim readArgs As Collection
    Set readArgs = NewElementArgs(postSharedId, "value")

    Dim readRaw As String
    readRaw = CallScript(GetActionJs("get_property"), readArgs, contextId, False, False, 5000, _
                         Silent:=True, allowAmbiguousRetry:=True)

    Dim actualValue As Variant
    actualValue = GetCleanResult(readRaw)

    Dim inspectErr As Long: inspectErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If inspectErr <> 0 Then Exit Sub
    If IsEmpty(actualValue) Or IsNull(actualValue) Or IsObject(actualValue) Then Exit Sub

    Dim actualText As String: actualText = CStr(actualValue)
    If StrComp(actualText, expectedValue, vbBinaryCompare) <> 0 Then
        Dim locatorNote As String
        If locatorKind = ElementLocatorXPath Then
            locatorNote = " | xpath=" & CStr(locator)
        Else
            locatorNote = " | locator=" & ElementLocatorKindName(locatorKind)
        End If
        LogActionDiagnosticLine "[INPUT-VALUE-MISMATCH] exactMatch=false" & _
                                " | expectedLen=" & CStr(Len(expectedValue)) & _
                                " | actualLen=" & CStr(Len(actualText)) & locatorNote
    End If
End Sub

' Called only after lazy_scroll returned False. The secondary read is best-effort and cannot
' change the completed action outcome. The browser-side diagnostic object is reset on each run.
Private Sub LogLazyScrollUnstableBestEffort(ByVal contextId As String, _
                                             ByVal budgetMs As Long, _
                                             ByVal pollMs As Long)
    Dim line As String
    line = "[LAZY-SCROLL-UNSTABLE] convergence=false" & _
           " | budgetMs=" & CStr(budgetMs) & _
           " | pollMs=" & CStr(pollMs)

    On Error Resume Next
    Err.Clear

    Dim diagRaw As String
    diagRaw = CallScript( _
        "const d=window.__vbaLazyScrollDiag;return d?JSON.stringify(d):'';", _
        Nothing, contextId, False, False, 5000, Silent:=True, allowAmbiguousRetry:=True)

    Dim diagText As Variant
    diagText = GetCleanResult(diagRaw)

    If Err.Number = 0 And Not IsEmpty(diagText) And Not IsNull(diagText) And Not IsObject(diagText) Then
        If Len(CStr(diagText)) > 0 Then
            Dim diag As Object
            Set diag = WebJsonConverter.ParseJson(CStr(diagText))
            If Err.Number = 0 And Not diag Is Nothing Then
                If diag.Exists("lastHeight") Then line = line & " | lastHeight=" & CStr(diag("lastHeight"))
                If diag.Exists("elapsedMs") Then line = line & " | elapsedMs=" & CStr(diag("elapsedMs"))
            End If
        End If
    End If

    Err.Clear
    On Error GoTo 0

    LogActionDiagnosticLine line
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
    waitCtx.Add "lastProbeHealthSignature", vbNullString
    waitCtx.Add "lastProbeHealthSummary", vbNullString
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
        If m_fatalResetDeferred Then
            res("Error") = "Fatal: CDP instability. Context reset is deferred until the active " & _
                           "exclusive capture finishes cleanup."
        Else
            res("Error") = "Fatal: CDP instability. Context state has been cleared."
        End If
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
    ' Capture every volatile arm/capture state BEFORE any reset can clear it.
    Dim droppedArms As String
    droppedArms = BuildFatalResetDroppedArmsDesc()

    ' B4/N6: DoEvents may allow unrelated CDP work to trip the fatal threshold while download,
    ' file-dialog, new-context, or navigation-status capture owns class-level transaction state.
    ' Clearing that state here would corrupt the outer operation before its cleanup can inspect it.
    If Len(m_exclusiveCaptureOperationOwner) > 0 Then
        If Not m_fatalResetDeferred Then
            m_fatalResetDeferred = True
            m_fatalResetDeferredOwner = m_exclusiveCaptureOperationOwner
            m_fatalResetDeferredDroppedArms = droppedArms

            LogSignalLine "[FATAL-RESET] deferred=true" & _
                          " | owner=" & m_fatalResetDeferredOwner & _
                          " | droppedArms=" & droppedArms
            Debug.Print "BiDi Critical: Fatal Reset deferred until exclusive capture cleanup completes."
        End If
        Exit Sub
    End If

    LogSignalLine "[FATAL-RESET] deferred=false | droppedArms=" & droppedArms
    Debug.Print "BiDi Critical: Performing Fatal Reset due to protocol instability."
    ApplyFatalResetNow
End Sub

Private Sub ApplyDeferredFatalReset()
    If Not m_fatalResetDeferred Then Exit Sub
    If Len(m_exclusiveCaptureOperationOwner) > 0 Then Exit Sub

    Dim deferredOwner As String
    Dim deferredDroppedArms As String
    deferredOwner = m_fatalResetDeferredOwner
    deferredDroppedArms = m_fatalResetDeferredDroppedArms

    LogSignalLine "[FATAL-RESET] deferredApply=true" & _
                  " | owner=" & deferredOwner & _
                  " | droppedArms=" & deferredDroppedArms
    Debug.Print "BiDi Critical: Applying deferred Fatal Reset after exclusive capture cleanup."

    ApplyFatalResetNow
End Sub

Private Sub ApplyFatalResetNow()
    ResetVolatileContextState
    m_consecutiveCdpFailures = 0

    m_fatalResetDeferred = False
    m_fatalResetDeferredOwner = vbNullString
    m_fatalResetDeferredDroppedArms = vbNullString
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

    Dim injectHealth As Variant
    Assign injectHealth, GetPathValue(validated, "result.result.value")
    If VarType(injectHealth) <> vbBoolean Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:StartSpaIdleProbe", _
                  "Probe injection returned a non-Boolean health result."
    End If

    If Not CBool(injectHealth) Then
        StartSpaIdleProbe = False
        Exit Function
    End If

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
    ' Error formatting must never replace the JavaScript exception it is trying to describe.
    ' Treat every nested field as optional even when the remote end is expected to be spec-valid.
    On Error GoTo FormatFallback

    Dim msg As String
    msg = "JS Exception"

    If details Is Nothing Then
        FormatJavaScriptError = msg & ": (no exception details)"
        Exit Function
    End If

    If TypeName(details) <> "Dictionary" Then
        FormatJavaScriptError = msg & ": (unexpected details type: " & TypeName(details) & ")"
        Exit Function
    End If

    Dim textValue As String
    textValue = GetPathString(details, "text")
    If Len(textValue) > 0 Then
        msg = msg & ": " & textValue
    Else
        msg = msg & ": (text unavailable)"
    End If

    Dim lineValue As String
    Dim colValue As String
    lineValue = GetPathString(details, "lineNumber")
    colValue = GetPathString(details, "columnNumber")
    If Len(lineValue) > 0 Or Len(colValue) > 0 Then
        If Len(lineValue) = 0 Then lineValue = "?"
        If Len(colValue) = 0 Then colValue = "?"
        msg = msg & " [Line: " & lineValue & ", Col: " & colValue & "]"
    End If

    Dim exceptionObj As Object
    Set exceptionObj = GetPathObject(details, "exception")
    If Not exceptionObj Is Nothing Then
        If TypeName(exceptionObj) = "Dictionary" Then
            Dim descriptionValue As String
            descriptionValue = GetPathString(exceptionObj, "description")
            If Len(descriptionValue) > 0 Then
                msg = msg & vbCrLf & "Description: " & descriptionValue
            End If
        End If
    End If

    Dim stackTraceObj As Object
    Set stackTraceObj = GetPathObject(details, "stackTrace")
    If Not stackTraceObj Is Nothing Then
        If TypeName(stackTraceObj) = "Dictionary" Then
            Dim framesObj As Object
            Set framesObj = GetPathObject(stackTraceObj, "callFrames")

            If Not framesObj Is Nothing Then
                If TypeName(framesObj) = "Collection" Then
                    Dim stack As String
                    Dim i As Long
                    stack = "--- JS Stack Trace ---"

                    For i = 1 To framesObj.count
                        Dim frameObj As Object
                        Set frameObj = Nothing
                        If IsObject(framesObj(i)) Then Set frameObj = framesObj(i)

                        If Not frameObj Is Nothing Then
                            If TypeName(frameObj) = "Dictionary" Then
                                Dim functionName As String
                                Dim frameUrl As String
                                Dim frameLine As String
                                Dim frameCol As String

                                functionName = GetPathString(frameObj, "functionName")
                                frameUrl = GetPathString(frameObj, "url")
                                frameLine = GetPathString(frameObj, "lineNumber")
                                frameCol = GetPathString(frameObj, "columnNumber")

                                If Len(functionName) = 0 Then functionName = "(anonymous)"
                                If Len(frameUrl) = 0 Then frameUrl = "(unknown)"
                                If Len(frameLine) = 0 Then frameLine = "?"
                                If Len(frameCol) = 0 Then frameCol = "?"

                                stack = stack & vbCrLf & "  at " & functionName & _
                                        " (" & frameUrl & ":" & frameLine & ":" & frameCol & ")"
                            End If
                        End If
                    Next i

                    If Len(stack) > Len("--- JS Stack Trace ---") Then
                        msg = msg & vbCrLf & stack
                    End If
                End If
            End If
        End If
    End If

    FormatJavaScriptError = msg
    Exit Function

FormatFallback:
    ' Never propagate a secondary formatter failure over the original JS exception.
    FormatJavaScriptError = "JS Exception: (error details could not be formatted safely)"
    Err.Clear
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

' 608/604/600/601/517 mean dispatch outcome is unknown. 608 specifically means the
' browser response arrived but exceeded the receive cap and was discarded. Promise.race does
' not cancel the losing browser task, so state-changing actions must not replay.
Private Function IsAmbiguousOutcome(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case 608, 604, 601, 600, 517
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

' Automatic navigation retry covers selected Chromium network failures only.
' net::ERR_TIMED_OUT is intentionally excluded because the request may already have reached
' the server; automatically re-sending the navigation would create an ambiguous duplicate.
Private Function IsRetryableNavFailure(ByVal errDesc As String) As Boolean
    Dim d As String: d = LCase$(errDesc)
    Dim codes As Variant
    codes = Array("connection_reset", "connection_closed", "connection_aborted", _
                  "connection_refused", "connection_failed", "connection_timed_out", _
                  "empty_response", "socket_not_connected", _
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
        Case "click", "input", "input_keys", "select", "select_text"
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
   
        Case "input_keys"

            js = js & "try{let e=arguments[0],v=arguments[1];const da=()=>{let a=document.activeElement;while(a&&a.shadowRoot&&a.shadowRoot.activeElement){a=a.shadowRoot.activeElement;}return a;};if(!e||!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');"
            js = js & "e.scrollIntoView({block:'center',inline:'center'});if(!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');e.click();e.focus();"
            js = js & "if(!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');let prev=null,stb=0;for(let t=performance.now();performance.now()-t<3000;){hb();let a=da();"
            js = js & "if(a&&a.isConnected&&a.tagName==='INPUT'&&(a.type==='text'||a.type===''||!a.type)&&typeof a.selectionStart==='number'){if(a===prev){stb++;if(stb>=2){e=a;break;}}else{prev=a;stb=1;}}else{prev=null;stb=0;}await new Promise(r=>setTimeout(r,15));}"
            js = js & "if(!e||!e.isConnected)throw new Error('STALE_ELEMENT_PRE_INPUT');(e.select?e.select():document.execCommand('selectAll'));await new Promise(r=>setTimeout(r,80));document.execCommand('delete');await new Promise(r=>setTimeout(r,80));"
            js = js & "if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(e.value){if(e.select)e.select();document.execCommand('forwardDelete');await new Promise(r=>setTimeout(r,80));}if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');"
            js = js & "if(e.value){e.value='';e.dispatchEvent(new Event('input',{bubbles:true}));await new Promise(r=>setTimeout(r,80));}if(!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');" & rp & "await new Promise(r=>setTimeout(r,40));"
            js = js & "for(let c of v){hb();if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');let s=(typeof e.selectionStart==='number'?e.selectionStart:0);document.execCommand('insertText',false,c);for(let k=0;k<80;k++){let a=da();if(a&&a.isConnected&&a!==e&&a.tagName==='INPUT'&&(a.type==='text'||!a.type))e=a;if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(da()===e&&typeof e.selectionStart==='number'&&e.selectionStart>=s+1)break;await new Promise(r=>setTimeout(r,10));}}"
            js = js & "if(!e||!e.isConnected)throw new Error('INPUT_TARGET_CHANGED_DURING_KEYS');if(v.length>0&&(!e.value||e.value.length<v.length))throw new Error('Wiz Input Validation Failed: '+e.value);return true;}catch(err){throw err;}"

        Case "autoclicker"
            ' XPath and Shadow AutoClickers share one generation registry. Shadow mode observes
            ' every open shadow root and also polls briefly so late attachShadow calls are visible.
            js = js & "function(){const k={{KEY}},m={{MODE}},p={{LOCATOR}},t={{TIMEOUT}},g={{GENERATION}},en={{ENABLED}};"
            js = js & "const rk='__seleniumvba_bidi_autoclicker_v41',bag=globalThis[rk]||(globalThis[rk]=Object.create(null));"
            js = js & "const q=bag[k];if(q&&Number(q.g)>=g)return;"
            js = js & "if(q){q.cancelled=true;try{if(q.o)q.o.disconnect()}catch(_){ }try{for(const z of(q.os||[]))z.disconnect()}catch(_){ }try{if(q.tm)clearTimeout(q.tm)}catch(_){ }try{if(q.poll)clearInterval(q.poll)}catch(_){ }}"
            js = js & "const s={g:g,cancelled:!en,os:[],tm:0,poll:0};bag[k]=s;if(!en)return;const own=()=>bag[k]===s&&!s.cancelled;"
            js = js & "const stop=()=>{s.cancelled=true;try{for(const z of s.os)z.disconnect()}catch(_){ }try{if(s.tm)clearTimeout(s.tm)}catch(_){ }try{if(s.poll)clearInterval(s.poll)}catch(_){ }};"
            js = js & "const fd=(r,x)=>{let n=null;try{n=r.querySelector(String(x))}catch(_){return null}if(n)return n;let a=[];try{a=r.querySelectorAll('*')}catch(_){return null}for(const c of a){if(c.shadowRoot){n=fd(c.shadowRoot,x);if(n)return n}}return null};"
            js = js & "const find=()=>{if(m==='xpath'){try{return document.evaluate(p,document,null,9,null).singleNodeValue}catch(_){return null}}if(!Array.isArray(p)||p.length===0)return null;let e=document;for(const x of p){e=fd(e.shadowRoot||e,x);if(!e)return null}return e};"
            js = js & "const f=()=>{if(!own())return 0;const e=find();if(e&&e.isConnected&&(e.offsetWidth>0||e.offsetHeight>0||e.getClientRects().length>0)){e.click();console.log('BiDi-AutoClicker: Target clicked');return 1}return 0};"
            js = js & "const seen=new WeakSet();let scan=null;const watch=r=>{if(!own()||!r||seen.has(r))return;seen.add(r);const o=new MutationObserver(()=>{if(!own()){stop();return}if(m==='shadow')scan(document);if(f())stop()});s.os.push(o);try{o.observe(r,{childList:1,subtree:1,attributes:1})}catch(_){ }};"
            js = js & "scan=r=>{if(!own()||!r)return;watch(r);let a=[];try{a=r.querySelectorAll('*')}catch(_){return}for(const c of a){if(c.shadowRoot)scan(c.shadowRoot)}};"
            js = js & "setTimeout(()=>{if(!own())return;if(f()){stop();return}watch(document);if(m==='shadow'){scan(document);s.poll=setInterval(()=>{if(!own()){stop();return}scan(document);if(f())stop()},100)}s.tm=setTimeout(()=>{if(own())stop()},t)},0);}"

        Case "lazy_scroll"
            js = js & "try{const tout=arguments[0],poll=arguments[1],start=performance.now();"
            js = js & "const se=document.scrollingElement||document.documentElement||document.body;"
            js = js & "let lastH=se.scrollHeight,stable=0;"
            js = js & "while(performance.now()-start<tout){hb();"
            js = js & "se.scrollTo(0,se.scrollHeight);"      ' cross the lazy-load trigger line
            js = js & "await new Promise(r=>setTimeout(r,poll));"  ' let fetch -> DOM settle
            js = js & "let curH=se.scrollHeight;"
            js = js & "if(curH===lastH){stable++;if(stable>=2)break;}else{lastH=curH;stable=0;}}"
            js = js & "const elapsed=Math.max(0,Math.round(performance.now()-start)),ok=(stable>=2);"
            js = js & "window.__vbaLazyScrollDiag={stable:ok,lastHeight:lastH,elapsedMs:elapsed};"
            js = js & "se.scrollTo(0,0);return ok;}catch(err){throw err;}"  ' reset geometry

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
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:CallScript", _
                  "Internal context-resolution invariant failed: ResolveTargetContext returned an empty context id."
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

' Browser-side operation deadlines and VBA-side response-correlation ownership are different
' budgets. The operation may legitimately be configured up to 600s, while SendAndReceive uses
' short receive slices plus guarded Office yields. The correlation cap still bounds command-ID
' ownership independently. A correlation timeout is ambiguous and the browser task may survive.
Private Function ResolveTransportTimeout(ByVal operationTimeoutMs As Long) As Long
    Dim candidate As Double
    candidate = CDbl(operationTimeoutMs) + CDbl(COMMAND_CORRELATION_TIMEOUT_MARGIN_MS)

    If candidate < DEFAULT_COMMAND_TIMEOUT_MS Then candidate = DEFAULT_COMMAND_TIMEOUT_MS
    If candidate > socket_.MaxCommandCorrelationWaitMs Then candidate = socket_.MaxCommandCorrelationWaitMs

    ResolveTransportTimeout = CLng(candidate)
End Function

Private Function ResolveNavigationTransportTimeout(ByVal maxTimeoutMs As Long) As Long
    Dim candidate As Long
    candidate = ResolveTransportTimeout(maxTimeoutMs)
    If candidate < NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS Then
        candidate = NAVIGATION_INTERACTIVE_TRANSPORT_FLOOR_MS
    End If
    If candidate > socket_.MaxCommandCorrelationWaitMs Then
        candidate = socket_.MaxCommandCorrelationWaitMs
    End If
    ResolveNavigationTransportTimeout = candidate
End Function

Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeoutMs As Double) As Boolean
    If timeoutMs <= 0 Then Exit Function
    IsTimeout = socket_.elapsedMs(startTicks) > timeoutMs
End Function

Private Sub WaitRecoveryBackoffNoOfficePump(ByVal sleepMs As Long)
    Dim vSleep As Long
    vSleep = sleepMs
    If vSleep < 0 Then vSleep = 0

    Dim startedAt As Currency
    startedAt = socket_.NowTicks()

    Do
        DrainNonBlockingRecorderQueue
        If socket_.elapsedMs(startedAt) >= vSleep Then Exit Do
        socket_.YieldExecution 25, False
    Loop
End Sub

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
            js = js & "var probeHealth=()=>{var d=getInputDesc();return{versionOk:!!p&&p.version===V,fetchWrapped:(typeof window.fetch!=='function')||(!!p&&!!p.fetchWrapper&&window.fetch===p.fetchWrapper),xhrWrapped:(!proto)||(!!p&&!!p.xhrOpenWrapper&&!!p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper),observerActive:!!p&&!!p.baseObserver,shadowObserverReady:!!p&&typeof p.observeShadowRoot==='function'&&Array.isArray(p.shadowObservers),inputWrapped:!!d&&(!!p&&!!p.inputSetterWrapper&&d.set===p.inputSetterWr"
            js = js & "apper)}};"
            js = js & "if(p){p.selectors=selectors;p.ignoreNetwork=ignoreNetwork;p.lastHealthError='';var h0=probeHealth();if(h0.versionOk&&h0.fetchWrapped&&h0.xhrWrapped&&h0.observerActive&&h0.shadowObserverReady&&h0.inputWrapped){window.__vbaSpaProbeInstalled=true;return true}}"
            js = js & "var legacy=!!(window.__vbaSpaProbeInstalled&&p&&(p.version===undefined||p.version===null));"
            js = js & "p=p||{inflightFetchCount:0,inflightXhrCount:0,lastMutationTs:performance.now(),changeList:[],inflightUrlCounts:new Map()};window.__vbaIdleProbe=p;window.__vbaSpaProbeInstalled=true;p.version=V;p.selectors=selectors;p.ignoreNetwork=ignoreNetwork;p.lastHealthError='';"
            js = js & "if(typeof p.inflightFetchCount!=='number')p.inflightFetchCount=0;if(typeof p.inflightXhrCount!=='number')p.inflightXhrCount=0;if(typeof p.lastMutationTs!=='number')p.lastMutationTs=performance.now();if(!Array.isArray(p.changeList))p.changeList=[];"
            js = js & "if(!(p.inflightUrlCounts instanceof Map)){p.inflightUrlCounts=new Map();if(p.inflightUrls&&p.inflightUrls.forEach)p.inflightUrls.forEach((u)=>p.inflightUrlCounts.set(u,1))}"
            js = js & "p.invalidSelectors=new Set();"
            js = js & "if(p.sigCount===undefined){p.sigCount=0;p.sigArmed=false;p.sigObs=null;p.visArmed=false;p.visSeen=false;p.visXPath='';p.visKind='';p.visPath=null}"
            js = js & "var gX=function(e){if(!e||e.nodeType!==1)return '';var s=[];while(e&&e.nodeType===1){if(e.id){s.unshift('//*[@id=\''+e.id+'\']');return s.join('/')}var i=1,b=e.previousSibling;while(b){if(b.nodeType===1&&b.tagName===e.tagName)i++;b=b.previousSibling}s.unshift(e.tagName.toLowerCase()+'['+i+']');e=e.parentNode}return '/'+s.join('/')};"
            js = js & "p.vChk=function(e){if(!e||!e.isConnected)return false;var s=window.getComputedStyle(e),r=e.getBoundingClientRect();return(s.display!=='none'&&s.visibility!=='hidden'&&parseFloat(s.opacity||'1')>0&&(r.width>0||r.height>0||e.getClientRects().length>0))};"
            js = js & "p.shadowResolve=function(a){if(!Array.isArray(a)||a.length===0)return null;var fd=function(r,s){var n=r.querySelector(String(s));if(n)return n;var z=r.querySelectorAll('*');for(var j=0;j<z.length;j++){var c=z[j];if(c.shadowRoot){n=fd(c.shadowRoot,s);if(n)return n}}return null};var e=document;for(var i=0;i<a.length;i++){e=fd(e.shadowRoot||e,a[i]);if(!e)return null}return e};"
            js = js & "p.armDomSignal=function(x,a){var n;try{n=document.evaluate(x,document,null,9,null).singleNodeValue}catch(e){return -2}if(!n)return -1;if(p.sigObs){try{p.sigObs.disconnect()}catch(e){}}p.sigCount=0;p.sigArmed=true;p.sigObs=new MutationObserver(function(ms){p.sigCount+=ms.length});p.sigObs.observe(n,{childList:true,characterData:true,subtree:true,attributes:!!a});return 0};"
            js = js & "p.armShadowDomSignal=function(a,b){var n;try{n=p.shadowResolve(a)}catch(e){return -2}if(!n)return -1;if(p.sigObs){try{p.sigObs.disconnect()}catch(e){}}p.sigCount=0;p.sigArmed=true;p.sigObs=new MutationObserver(function(ms){p.sigCount+=ms.length});p.sigObs.observe(n,{childList:true,characterData:true,subtree:true,attributes:!!b});return 0};"
            js = js & "p.armVisSignal=function(x){var n=null;try{n=document.evaluate(x,document,null,9,null).singleNodeValue}catch(e){return -2}if(n&&p.vChk(n))return -3;p.visKind='xpath';p.visXPath=x;p.visPath=null;p.visArmed=true;p.visSeen=false;return 0};"
            js = js & "p.armShadowVisSignal=function(a){var n=null;try{n=p.shadowResolve(a)}catch(e){return -2}if(n&&p.vChk(n))return -3;p.visKind='shadow';p.visXPath='';p.visPath=a.slice();p.visArmed=true;p.visSeen=false;return 0};"
            js = js & "p.disarmSignal=function(){if(p.sigObs){try{p.sigObs.disconnect()}catch(e){}p.sigObs=null}p.sigArmed=false;p.sigCount=0;p.visArmed=false;p.visSeen=false;p.visXPath='';p.visKind='';p.visPath=null};"
            js = js & "var uT=()=>{p.lastMutationTs=performance.now()},lg=(e)=>{if(e){uT();if(p.changeList.indexOf(e)===-1&&p.changeList.length<50)p.changeList.push(e)}};p.touch=uT;p.logChange=lg;"
            js = js & "var iN=(n,s)=>{if(!n)return 1;var t=(n.nodeType===1)?n:n.parentElement;if(!t)return 0;for(var i=0;i<s.length;i++){var x=s[i],k=String(x);if(p.invalidSelectors.has(k))continue;var q=t;while(q){try{if(q.closest&&q.closest(x)!==null)return 1}catch(e){p.invalidSelectors.add(k);break}var r=null;try{r=q.getRootNode?q.getRootNode():null}catch(e){}q=(r&&r.host)?r.host:null}}return 0};"
            js = js & "try{var d=getInputDesc();if(legacy&&!p.inputSetterWrapper&&d&&typeof d.set==='function'){p.inputSetterWrapper=d.set;p.inputGetterWrapper=d.get||null}else if(!(p.inputSetterWrapper&&d&&d.set===p.inputSetterWrapper&&!p.inputOriginalSet)){var bS=(p.inputSetterWrapper&&d&&d.set===p.inputSetterWrapper&&p.inputOriginalSet)?p.inputOriginalSet:(d&&d.set),bG=(p.inputGetterWrapper&&d&&d.get===p.inputGetterWrapper&&p.inputOrigi"
            js = js & "nalGet)?p.inputOriginalGet:(d&&d.get);if(bS&&bG){var sW=function(v){if(!iN(this,p.selectors||[]))lg('PROP:'+gX(this)+'[value]');return bS.call(this,v)},gW=function(){return bG.call(this)};Object.defineProperty(HTMLInputElement.prototype,'value',{set:sW,get:gW,configurable:true});p.inputOriginalSet=bS;p.inputOriginalGet=bG;p.inputSetterWrapper=sW;p.inputGetterWrapper=gW}}}catch(e){p.lastHealthError='input:'+String(e&&e.message||e)}"
            js = js & "try{if(p.baseObserver&&p.baseObserver.disconnect)p.baseObserver.disconnect()}catch(e){}p.baseObserver=null;"
            js = js & "try{if(Array.isArray(p.shadowObservers)){for(const o of p.shadowObservers){try{o.disconnect()}catch(e){}}}}catch(e){}p.shadowObservers=[];p.shadowSeen=new WeakSet();"
            js = js & "p.handleMutations=(ms)=>{ms.forEach((m)=>{if(!iN(m.target,p.selectors||[])){var t=(m.type==='characterData')?m.target.parentElement:m.target,path=gX(t),entry=(m.type==='attributes')?'ATTR:'+path+'['+m.attributeName+']':'DOM:'+path;lg(entry)}if(m.type==='childList'&&m.addedNodes){for(const n of m.addedNodes){if(n&&n.nodeType===1){try{if(n.shadowRoot)p.observeShadowRoot(n.shadowRoot)}catch(e){}try{p.scanShadowRoots(n)}catch(e){}}}}})};"
            js = js & "p.scanShadowRoots=function(r){if(!r||!r.querySelectorAll)return;var a=[];try{a=r.querySelectorAll('*')}catch(e){return}for(const c of a){try{if(c.shadowRoot)p.observeShadowRoot(c.shadowRoot)}catch(e){}}};"
            js = js & "p.observeShadowRoot=function(r){if(!r||p.shadowSeen.has(r))return false;p.shadowSeen.add(r);try{var o=new MutationObserver(p.handleMutations);o.observe(r,{childList:true,characterData:true,subtree:true,attributes:true});p.shadowObservers.push(o);p.scanShadowRoots(r);return true}catch(e){p.lastHealthError='shadow-observer:'+String(e&&e.message||e);return false}};"
            js = js & "try{p.baseObserver=new MutationObserver(p.handleMutations);p.baseObserver.observe(document,{childList:true,characterData:true,subtree:true,attributes:true});p.scanShadowRoots(document)}catch(e){p.baseObserver="
            js = js & "null;p.lastHealthError='observer:'+String(e&&e.message||e)}"
            js = js & "var __vbaMatch=function(u,pat){u=String(u||'').toLowerCase();pat=String(pat||'').toLowerCase();if(!pat)return false;if(pat.indexOf('*')<0)return u.indexOf(pat)!==-1;var a=pat.split('*'),pos=0;for(var i=0;i<a.length;i++){var q=a[i];if(!q)continue;var n=u.indexOf(q,pos);if(n<0)return false;pos=n+q.length}return true};"
            js = js & "if(typeof window.fetch==='function'){if(legacy&&!p.fetchWrapper&&window.fetch.__vbaWrapped){p.fetchWrapper=window.fetch}else if(!(p.fetchWrapper&&window.fetch===p.fetchWrapper&&!p.fetchOriginal)){var oF=(p.fetchWrapper&&window.fetch===p.fetchWrapper&&p.fetchOriginal)?p.fetchOriginal:window.fetch;p.fetchGen=(p.fetchGen||0)+1;var myFetchGen=p.fetchGen;var fW=function(u,o){if(myFetchGen!==p.fetchGen)return oF.apply(this,arguments);var r=(typeof u==='string')?u:(u&&u.url)||'';if((p.ignoreNetwork||[]).some((pat)=>__vbaMatch(r,pat)))return oF"
            js = js & ".apply(this,arguments);p.inflightFetchCount++;p.inflightUrlCounts.set(r,(p.inflightUrlCounts.get(r)||0)+1);var dec=()=>{p.inflightFetchCount=Math.max(0,p.inflightFetchCount-1);var n=(p.inflightUrlCounts.get(r)||0)-1;if(n<=0)p.inflightUrlCounts.delete(r);else p.inflightUrlCounts.set(r,n);uT()};try{return oF.apply(this,arguments).then((r2)=>{dec();return r2}).catch((e)=>{dec();throw e})}catch(e2){dec();throw e2}};fW.__"
            js = js & "vbaWrapped=true;fW.__vbaProbeVersion=V;fW.__vbaProbeGeneration=myFetchGen;p.fetchOriginal=oF;p.fetchWrapper=fW;window.fetch=fW}}else{p.fetchOriginal=null;p.fetchWrapper=null}"
            js = js & "if(proto){if(legacy&&!p.xhrOpenWrapper&&proto.__vbaWrapped){p.xhrOpenWrapper=proto.open;p.xhrSendWrapper=proto.send}else if(!(p.xhrOpenWrapper&&p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper&&(!p.xhrOpenOriginal||!p.xhrSendOriginal))){var oO=(p.xhrOpenWrapper&&proto.open===p.xhrOpenWrapper&&p.xhrOpenOriginal)?p.xhrOpenOriginal:proto.open,oS=(p.xhrSendWrapper&&proto.send===p.xhrSendWrap"
            js = js & "per&&p.xhrSendOriginal)?p.xhrSendOriginal:proto.send;p.xhrGen=(p.xhrGen||0)+1;var myXhrGen=p.xhrGen;var xO=function(m,u){if(myXhrGen!==p.xhrGen)return oO.apply(this,arguments);this.__vbaUrl=u;this.__vbaHooked=true;return oO.apply(this,arguments)},xS=function(){if(myXhrGen!==p.xhrGen)return oS.apply(this,arguments);var s=this,u=this.__vbaUrl||'';if((p.ignoreNetwork||[]).some((pat)=>__vbaMatch(u,pat))){s.__vbaHooked=false;return oS.apply(this,arguments)}p.inflightXhrCount++;p.inflightUrlCounts.set(u,(p.inflightUrlCounts.get(u)||0)+1);var dn=()=>{if(s.__vbaHooked){p.inflightXhrCoun"
            js = js & "t=Math.max(0,p.inflightXhrCount-1);var n=(p.inflightUrlCounts.get(u)||0)-1;if(n<=0)p.inflightUrlCounts.delete(u);else p.inflightUrlCounts.set(u,n);s.__vbaHooked=false;uT()}s.removeEventListener('loadend',dn)};this.addEventListener('loadend',dn);try{return oS.apply(this,arguments)}catch(e2){dn();throw e2}};xO.__vbaProbeGeneration=myXhrGen;xS.__vbaProbeGeneration=myXhrGen;p.xhrOpenOriginal=oO;p.xhrSendOriginal=oS;p.xhrOpenWrapper=xO;p.xhrSendWrapper=xS;proto.open=xO;proto.send=xS;pro"
            js = js & "to.__vbaWrapped=true;proto.__vbaProbeVersion=V}}"
            js = js & "var h=probeHealth();return h.versionOk&&h.fetchWrapped&&h.xhrWrapped&&h.observerActive&&h.shadowObserverReady&&h.inputWrapped})();"

        Case "query"
            ' Query reports hook-by-hook health. VBA treats any missing/mismatched field as
            ' inactive, resets the stable window, and routes through bounded probe recovery.
            js = js & "var V=" & SPA_PROBE_VERSION & ";"
            js = js & "var p=window.__vbaIdleProbe;if(!p)return JSON.stringify({active:false,version:0,health:{versionOk:false,fetchWrapped:false,xhrWrapped:false,observerActive:false,shadowObserverReady:false,inputWrapped:false},lastHealthError:''});"
            js = js & "var proto=(typeof XMLHttpRequest==='function')?XMLHttpRequest.prototype:null,d=null;try{d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')}catch(e){}"
            js = js & "var h={versionOk:p.version===V,fetchWrapped:(typeof window.fetch!=='function')||(!!p.fetchWrapper&&window.fetch===p.fetchWrapper),xhrWrapped:(!proto)||(!!p.xhrOpenWrapper&&!!p.xhrSendWrapper&&proto.open===p.xhrOpenWrapper&&proto.send===p.xhrSendWrapper),observerActive:!!p.baseObserver,shadowObserverReady:typeof p.observeShadowRoot==='function'&&Array.isArray(p.shadowObservers),inputWrapped:!!d&&(!!p.inputSetterWrapper&&d.set===p.inputSetterWrapper)};"
            js = js & "var healthy=h.versionOk&&h.fetchWrapped&&h.xhrWrapped&&h.observerActive&&h.shadowObserverReady&&h.inputWrapped;"
            js = js & "if(p.visArmed&&!p.visSeen){try{var vn=null;if(p.visKind==='shadow'&&p.visPath)vn=p.shadowResolve(p.visPath);else if(p.visXPath)vn=document.evaluate(p.visXPath,document,null,9,null).singleNodeValue;if(vn&&p.vChk&&p.vChk(vn))p.visSeen=true}catch(e){}}"
            js = js & "var res={active:healthy,version:p.version||0,health:h,lastHealthError:String(p.lastHealthError||''),now:performance.now(),lastMutationTs:p.lastMutationTs,inflightFetchCount:p.inflightFetchCount,inflightXhrCount:p.inflightXhrCount,changes:p.changeList,inflightUrls:Array.from((p.inflightUrlCounts||new Map()).keys()),invalidSelectors:Array.from(p.invalidSelectors||[]),sigArmed:!!p.sigArmed,sigCount:p.sigCount||0,visArmed:!!p.visArmed,visSeen:!!p.visSeen};p.changeList=[];return JSON.stringify(res);"
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

            ' QuerySpaProbe drains changeList even when health is false. Preserve current-version
            ' activity/render diagnostics before readiness decides whether this sample may affect
            ' consensus.
            ProcessSpaProbeChangesForDiagnostics probe, ctx

            Dim isProbeActive As Boolean: isProbeActive = IsSpaProbeReady(probe)

            If isProbeActive Then
                ctx("missingProbeCount") = 0
                ResetSpaProbeHealthBurst ctx

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

                ' Expected navigation-driven disappearance is handled above without an unhealthy
                ' warning. All other missing/unhealthy samples are explicit diagnostics.
                TrackSpaProbeUnhealthy probe, ctx, targetCtx

                Dim bCount As Long: bCount = CLng(ctx("missingProbeCount")) + 1
                ctx("missingProbeCount") = bCount

                ctx("stableStart") = socket_.nowMs()

                ' Repeated inability to establish the probe raises logical panic 505. Include the
                ' latest explicit false/missing health flags so the caller can diagnose the failed hook.
                If bCount > MAX_RECOVERY_ATTEMPTS Then
                    TrackSpaProbeUnhealthy probe, ctx, targetCtx, True

                    Dim lastProbeHealth As String
                    If ctx.Exists("lastProbeHealthSummary") Then
                        lastProbeHealth = CStr(ctx("lastProbeHealthSummary"))
                    End If
                    If Len(lastProbeHealth) = 0 Then lastProbeHealth = BuildSpaProbeHealthSummary(probe)

                    Err.Raise 505, "BiDi:ConsensusError", _
                        "SPA Consensus Panic: Stability probe failed to initialize after " & _
                        MAX_RECOVERY_ATTEMPTS & " attempts. The page state is too volatile. " & _
                        "Last unhealthy probe: " & lastProbeHealth
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
                                    Optional ByVal retryLocatorKind As ElementLocatorKind = ElementLocatorNone, _
                                    Optional ByVal retryLocator As Variant, _
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

            If staleFailure And retryLocatorKind <> ElementLocatorNone Then
                Dim freshSharedId As String
                Dim resolverName As String

                If IsMissing(retryLocator) Then
                    Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteBaseAction", _
                              "Retry locator is missing for stale-element recovery."
                End If

                resolverName = retryCallerName
                If Len(resolverName) = 0 Then resolverName = "ExecuteBaseAction"

                freshSharedId = ResolveElementSharedId( _
                    retryLocatorKind, retryLocator, targetCtx, retrySearchTimeoutMs, resolverName)

                Set currentArgs = ReplaceFirstElementReference(currentArgs, freshSharedId)

                If m_debugMode Then
                    Debug.Print "BiDi: Stale element re-resolved via " & _
                                ElementLocatorKindName(retryLocatorKind) & "."
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

                    Dim shellStatus As String
                    shellStatus = CStr(innerJson("status"))

                    Select Case shellStatus
                        Case "timeout"
                            Err.Raise 517, "BiDi:SafetyTimeout:" & methodName, _
                                "Safety Shell timeout: the wrapper stopped waiting, but the browser-side task may still be running."

                        Case "error"
                            Dim shellMsg As String
                            If innerJson.Exists("message") Then
                                shellMsg = CStr(innerJson("message"))
                            Else
                                shellMsg = "JavaScript execution failed without an error message."
                            End If
                            Err.Raise ERR_ACTION_FAILED, "BiDi:JSError:" & methodName, shellMsg

                        Case "ok"
                            If Not innerJson.Exists("value") Then
                                Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                                    "Safety Shell protocol violation: status='ok' without value."
                            End If
                            Call AssignValue(resultObj, "value", innerJson("value"))

                        Case Else
                            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                                "Safety Shell protocol violation: unknown status '" & shellStatus & "'."
                    End Select
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

Private Function ElementLocatorKindName(ByVal locatorKind As ElementLocatorKind) As String
    Select Case locatorKind
        Case ElementLocatorXPath
            ElementLocatorKindName = "XPath"
        Case ElementLocatorShadowPath
            ElementLocatorKindName = "ShadowPath"
        Case Else
            ElementLocatorKindName = "Unknown"
    End Select
End Function


' Centralizes element lookup while keeping locator-specific logic isolated.
Private Function ResolveElementSharedId(ByVal locatorKind As ElementLocatorKind, _
                                        ByVal locator As Variant, _
                                        ByVal contextId As String, _
                                        ByVal searchTimeoutMs As Long, _
                                        ByVal callerName As String, _
                                        Optional ByVal required As Boolean = True) As String
    Select Case locatorKind
        Case ElementLocatorXPath
            If VarType(locator) <> vbString Then
                Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                          "XPath locator must be a String."
            End If

            ResolveElementSharedId = ResolveXPathSharedId( _
                CStr(locator), contextId, searchTimeoutMs, callerName, required)

        Case ElementLocatorShadowPath
            ResolveElementSharedId = ResolveShadowElementSharedId( _
                locator, contextId, searchTimeoutMs, callerName, required)

        Case Else
            Err.Raise ERR_ACTION_FAILED, "BiDi:" & callerName, _
                      "Unsupported element locator kind: " & CStr(locatorKind)
    End Select
End Function

Private Function ResolveXPathSharedId(ByVal xpath As String, _
                                      ByVal contextId As String, _
                                      ByVal timeoutMs As Long, _
                                      ByVal methodName As String, _
                                      Optional ByVal required As Boolean = True) As String
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

    If Len(sid) = 0 And required Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "Element not found within " & IIf(vTimeout = 0, "0 (immediate)", vTimeout) & "ms for XPath: " & xpath
    End If

    ResolveXPathSharedId = sid
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

    ' Maintenance (A10-OBS): each poll below passes through CallScript -> ResolveTargetContext,
    ' which drains the async event queue non-blockingly. Re-evaluate A10-OBS if a future
    ' URL-stability refactor removes that per-poll call chain; do not add a drain only for symmetry.

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
    
    ' The watchdog rejects with a private marker object rather than an Error message token.
    ' User JavaScript may therefore throw Error("VBA_TIMEOUT") without being misclassified.
    js = js & "let _wdTimer; const _wd = new Promise((_,r)=>{ _wdTimer=setInterval(()=>{if(performance.now()-_la>_t){clearInterval(_wdTimer);r({__svbaSafetyTimeout:1})}},100);}); "

    js = js & "try{ const res = await Promise.race([_tk(), _wd]); clearInterval(_wdTimer); return JSON.stringify({status:'ok', value:(typeof res === 'undefined' ? null : res)}); "
    js = js & "}catch(e){ clearInterval(_wdTimer); if(e&&e.__svbaSafetyTimeout===1)return JSON.stringify({status:'timeout'}); "
    js = js & "return JSON.stringify({status:'error', message:(e&&typeof e.message!=='undefined'?String(e.message):String(e))}); } }"

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

    ' Shutdown is teardown, not a competing capture operation. If a host interruption or a
    ' re-entrant teardown leaves an exclusive owner active, abandon that transaction and continue
    ' cleanup rather than blocking the only path that can release browser-side residual state.
    ' ClearState remains guarded because it is not teardown and must not erase a live capture.
    On Error Resume Next
    AbandonExclusiveCaptureForShutdownBestEffort

    ' Shutdown is best-effort from this point; cleanup must continue even after detach or
    ' close failure.

    ' AutoClicker preload scripts are session-side state. Remove every tracked active registration
    ' and leave a higher-generation inert suppressor before the socket is detached/closed. This also
    ' contains any older preload whose addPreloadScript response was lost and whose ID is unknown.
    Dim autoClickerCleanupFailures As Long
    autoClickerCleanupFailures = ReleaseAllAutoClickersBestEffort()
    If autoClickerCleanupFailures > 0 And m_debugMode Then
        Debug.Print "BiDi CRITICAL: Shutdown could not confirm cleanup of " & _
                    CStr(autoClickerCleanupFailures) & " AutoClicker cleanup/containment operation(s)."
    End If

    ' Shadow Unlocker is also preload/browser-realm state. Remove the tracked preload, restore
    ' Element.prototype.attachShadow in current realms, and contain any ambiguous orphan.
    Dim shadowCleanupFailures As Long
    shadowCleanupFailures = ReleaseGlobalShadowUnlockerBestEffort()
    If shadowCleanupFailures > 0 And m_debugMode Then
        Debug.Print "BiDi CRITICAL: Shutdown could not confirm " & _
                    CStr(shadowCleanupFailures) & " Shadow Unlocker cleanup/containment operation(s)."
    End If

    ' Download behavior is browser-side state and may survive wrapper/WebSocket teardown. Clear
    ' only when this wrapper attempted to change it; unrelated browser defaults remain untouched.
    ' Shutdown owns the result check so cleanup failure is not silently discarded immediately
    ' before ResetSessionBoundaryState forgets the pending/unknown flags.
    If m_downloadBehaviorCleanupRequired Then
        If Not ClearDownloadBehaviorBestEffort(False) Then
            If m_debugMode Then
                Debug.Print "BiDi CRITICAL: Shutdown could not confirm browser.setDownloadBehavior cleanup. " & _
                            "Reconnect before further download use; ConnectTo will perform the residual-state cleanup sweep."
            End If
        End If
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

    ' Subscriptions belong to the BiDi session, not the WebSocket transport. Use one bounded
    ' unsubscribe while the socket is still live. If acknowledgement is lost, retain only the
    ' issued IDs and session identity so a same-session ConnectTo must reconcile them first.
    If Not ReleaseTrackedSubscriptionsBeforeClearBestEffort(2000) Then
        CaptureTrackedSubscriptionsForDeferredCleanup
        If m_debugMode Then
            Debug.Print "BiDi Warning: Shutdown deferred subscription cleanup to the next same-session ConnectTo."
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
    ' path can observe m_isRecording=True with m_recorderLogs already Nothing. Subscription
    ' ownership has already been released or copied into the deferred-ID registry above.
    ResetSessionBoundaryState
    ClearStateCore False

    If ForceClose Then
        socket_.ForceClose
    Else
        ' Graceful shutdown must explicitly end the transport even when another object
        ' still holds a reference to the communicator.
        socket_.Disconnect
    End If

    ' Retain the disconnected communicator so this wrapper remains reusable, as ConnectTo's
    ' own error message promises. ForceClose and Disconnect are both idempotent. Deferred
    ' subscription cleanup carries its own session identity, so the active identity can be dropped.
    m_currentSessionIdentity = vbNullString
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

Private Function NextShadowUnlockerGuardGeneration() As Double
    ' QPC is system-wide and monotonic, so a later wrapper instance naturally supersedes any
    ' orphaned preload from an earlier instance even after local generation state was lost.
    Dim candidate As Double
    candidate = Fix(socket_.nowMs() * 1000#)

    If candidate <= m_lastShadowUnlockerGuardGeneration Then
        candidate = m_lastShadowUnlockerGuardGeneration + 1#
    End If

    m_lastShadowUnlockerGuardGeneration = candidate
    NextShadowUnlockerGuardGeneration = candidate
End Function

Private Function ShadowUnlockerGenerationLiteral(ByVal generation As Double) As String
    ShadowUnlockerGenerationLiteral = Format$(generation, "0")
End Function

Private Function BuildShadowUnlockerPreloadSource(ByVal generation As Double, _
                                                  ByVal enabled As Boolean) As String
    Dim enabledLiteral As String
    If enabled Then
        enabledLiteral = "true"
    Else
        enabledLiteral = "false"
    End If

    Dim js As String
    js = "function(){const rk=Symbol.for('seleniumvba.bidi.shadow_unlocker.v44')," & _
         "g=" & ShadowUnlockerGenerationLiteral(generation) & ",en=" & enabledLiteral & ";"
    js = js & "const prior=globalThis[rk];if(prior&&Number(prior.g)>=g)return;if(prior)prior.enabled=false;"
    js = js & "if(prior&&prior.wrapper&&prior.original&&Element.prototype.attachShadow===prior.wrapper){" & _
              "try{Element.prototype.attachShadow=prior.original}catch(_){ }}"
    js = js & "const s={g:g,enabled:en,original:null,wrapper:null};globalThis[rk]=s;if(!en)return;"
    js = js & "const original=Element.prototype.attachShadow;if(typeof original!=='function')return;s.original=original;"
    js = js & "const wrapper=function(options){const next=s.enabled?Object.assign({},options||{},{mode:'open'}):options;" & _
              "const r=Reflect.apply(original,this,[next]);" & _
              "if(s.enabled){try{const p=globalThis.__vbaIdleProbe;if(p&&typeof p.observeShadowRoot==='function')p.observeShadowRoot(r)}catch(_){ }}" & _
              "return r};s.wrapper=wrapper;Element.prototype.attachShadow=wrapper;}"

    BuildShadowUnlockerPreloadSource = js
End Function

Private Function BuildShadowUnlockerRuntimeCleanupSource() As String
    BuildShadowUnlockerRuntimeCleanupSource = _
        "function(){const rk=Symbol.for('seleniumvba.bidi.shadow_unlocker.v44'),s=globalThis[rk];" & _
        "if(!s)return false;s.enabled=false;if(s.wrapper&&s.original&&Element.prototype.attachShadow===s.wrapper){" & _
        "try{Element.prototype.attachShadow=s.original}catch(_){return false}}return true;}"
End Function

Private Function TryAddShadowUnlockerPreload(ByVal generation As Double, _
                                             ByVal enabled As Boolean, _
                                             ByRef scriptId As String, _
                                             ByRef failureNumber As Long, _
                                             ByRef failureSource As String, _
                                             ByRef failureDescription As String) As Boolean
    scriptId = vbNullString
    failureNumber = 0
    failureSource = vbNullString
    failureDescription = vbNullString

    On Error GoTo Failed

    Dim params As New Dictionary
    params.Add "functionDeclaration", BuildShadowUnlockerPreloadSource(generation, enabled)

    Dim raw As String
    raw = ExecuteBiDiCommand("script.addPreloadScript", params, True, _
                             allowAmbiguousRetry:=False, _
                             allowContextRetarget:=False)

    Dim json As Object
    Set json = ValidateProtocolResponse(raw, "script.addPreloadScript")

    scriptId = GetPathString(json, "result.script")
    If Len(scriptId) = 0 Then
        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:script.addPreloadScript", _
                  "Shadow Unlocker preload registration succeeded without returning result.script."
    End If

    TryAddShadowUnlockerPreload = True
    Exit Function

Failed:
    failureNumber = Err.Number
    failureSource = Err.source
    failureDescription = Err.description
    Err.Clear
End Function

Private Function TryInstallShadowUnlockerSuppressor(ByRef scriptId As String, _
                                                    ByRef failureNumber As Long, _
                                                    ByRef failureSource As String, _
                                                    ByRef failureDescription As String) As Boolean
    Dim previousSuppressorId As String
    previousSuppressorId = m_shadowUnlockerSuppressorScriptId

    Dim generation As Double
    generation = NextShadowUnlockerGuardGeneration()

    If Not TryAddShadowUnlockerPreload(generation, False, scriptId, _
                                       failureNumber, failureSource, failureDescription) Then
        m_shadowUnlockerContainmentRisk = True
        Exit Function
    End If

    m_shadowUnlockerSuppressorScriptId = scriptId
    m_shadowUnlockerContainmentRisk = False

    If Len(previousSuppressorId) > 0 And _
       StrComp(previousSuppressorId, scriptId, vbBinaryCompare) <> 0 Then
        Dim oldErrNum As Long
        Dim oldErrSource As String
        Dim oldErrDesc As String
        Call TryRemovePreloadScript(previousSuppressorId, False, oldErrNum, oldErrSource, oldErrDesc)
    End If

    If m_debugMode Then
        Debug.Print "BiDi: Shadow DOM Unlocker containment preload registered (ID: " & _
                    scriptId & ", generation=" & ShadowUnlockerGenerationLiteral(generation) & ")"
    End If

    TryInstallShadowUnlockerSuppressor = True
End Function

Private Sub BestEffortRemoveKnownShadowUnlockerSuppressor()
    If Len(m_shadowUnlockerSuppressorScriptId) = 0 Then Exit Sub

    Dim scriptId As String
    scriptId = m_shadowUnlockerSuppressorScriptId

    Dim errNum As Long
    Dim errSource As String
    Dim errDesc As String
    If TryRemovePreloadScript(scriptId, False, errNum, errSource, errDesc) Then
        m_shadowUnlockerSuppressorScriptId = vbNullString
    ElseIf m_debugMode Then
        Debug.Print "BiDi Warning: inert Shadow Unlocker containment preload could not be removed. " & _
                    "ID=" & scriptId & AppendErrorDescription(errDesc)
    End If
End Sub

Private Function TryRestoreShadowUnlockerInContext(ByVal contextId As String, _
                                                   ByRef failureDescription As String) As Boolean
    failureDescription = vbNullString
    On Error GoTo Failed

    Dim target As New Dictionary
    target.Add "context", contextId

    Dim args As New Collection
    Dim params As New Dictionary
    params.Add "functionDeclaration", BuildShadowUnlockerRuntimeCleanupSource()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim raw As String
    raw = ExecuteBiDiCommand("script.callFunction", params, True, _
                             allowAmbiguousRetry:=True, _
                             allowRecoveryRetry:=False, _
                             allowContextRetarget:=False, _
                             commandTimeoutMs:=2000)
    Call ValidateProtocolResponse(raw, "script.callFunction")

    TryRestoreShadowUnlockerInContext = True
    Exit Function

Failed:
    failureDescription = "Err " & CStr(Err.Number) & ": " & Err.description
    Err.Clear
End Function

Private Function RestoreShadowUnlockerInCurrentRealmsBestEffort() As Long
    If socket_ Is Nothing Then Exit Function
    If socket_.state <> StateConnected Then Exit Function

    On Error GoTo TreeFailed

    Dim params As New Dictionary
    Dim raw As String
    raw = ExecuteBiDiCommand("browsingContext.getTree", params, True, _
                             allowAmbiguousRetry:=True, _
                             allowRecoveryRetry:=False, _
                             allowContextRetarget:=False, _
                             commandTimeoutMs:=2000)

    Dim tree As Object
    Set tree = ValidateProtocolResponse(raw, "browsingContext.getTree")

    Dim roots As Collection
    Set roots = GetPathObject(tree, "result.contexts")
    If roots Is Nothing Then Exit Function

    Dim contexts As New Collection
    CollectContextIdsForCleanup roots, contexts

    Dim i As Long
    Dim failure As String
    For i = 1 To contexts.count
        If Not TryRestoreShadowUnlockerInContext(CStr(contexts(i)), failure) Then
            RestoreShadowUnlockerInCurrentRealmsBestEffort = _
                RestoreShadowUnlockerInCurrentRealmsBestEffort + 1
            If m_debugMode Then
                Debug.Print "BiDi Warning: Shadow Unlocker runtime restoration failed for context " & _
                            Left$(CStr(contexts(i)), 16) & "... | " & failure
            End If
        End If
    Next i
    Exit Function

TreeFailed:
    RestoreShadowUnlockerInCurrentRealmsBestEffort = 1
    If m_debugMode Then
        Debug.Print "BiDi Warning: Shadow Unlocker runtime restoration could not enumerate contexts. " & _
                    "Err " & CStr(Err.Number) & ": " & Err.description
    End If
    Err.Clear
End Function

' Removes the tracked preload when possible, restores attachShadow in every current browsing
' context, then leaves a higher-generation inert preload at the session boundary. This mirrors
' the AutoClicker containment rule and covers older registrations whose response ID was lost.
Private Function ReleaseGlobalShadowUnlockerBestEffort() As Long
    Dim hadShadowState As Boolean
    hadShadowState = m_shadowUnlockerRegistered Or _
                     Len(m_shadowUnlockerScriptId) > 0 Or _
                     Len(m_shadowUnlockerSuppressorScriptId) > 0 Or _
                     m_shadowUnlockerContainmentRisk
    If Not hadShadowState Then Exit Function

    Dim containmentRequired As Boolean
    containmentRequired = m_shadowUnlockerContainmentRisk

    If Len(m_shadowUnlockerScriptId) > 0 Then
        Dim removeErrNum As Long
        Dim removeErrSource As String
        Dim removeErrDesc As String
        If TryRemovePreloadScript(m_shadowUnlockerScriptId, False, _
                                  removeErrNum, removeErrSource, removeErrDesc) Then
            m_shadowUnlockerRegistered = False
            m_shadowUnlockerScriptId = vbNullString
        Else
            containmentRequired = True
            ReleaseGlobalShadowUnlockerBestEffort = ReleaseGlobalShadowUnlockerBestEffort + 1
            If m_debugMode Then
                Debug.Print "BiDi CRITICAL: Shadow Unlocker tracked preload removal was not confirmed. " & _
                            "ID=" & m_shadowUnlockerScriptId & AppendErrorDescription(removeErrDesc)
            End If
        End If
    End If

    Dim restoreFailures As Long
    restoreFailures = RestoreShadowUnlockerInCurrentRealmsBestEffort()
    ReleaseGlobalShadowUnlockerBestEffort = ReleaseGlobalShadowUnlockerBestEffort + restoreFailures

    ' Always leave one newer inert generation at the session boundary, even when the tracked
    ' active ID was removed successfully. This is deliberate hard-reset insurance for an older
    ' addPreloadScript response that may have been lost by a previous wrapper instance.
    Dim suppressorId As String
    Dim suppressErrNum As Long
    Dim suppressErrSource As String
    Dim suppressErrDesc As String
    If Not TryInstallShadowUnlockerSuppressor(suppressorId, suppressErrNum, _
                                              suppressErrSource, suppressErrDesc) Then
        ReleaseGlobalShadowUnlockerBestEffort = ReleaseGlobalShadowUnlockerBestEffort + 1
        If m_debugMode Then
            Debug.Print "BiDi CRITICAL: Shadow Unlocker session-boundary containment could not be confirmed." & _
                        AppendErrorDescription(suppressErrDesc)
        End If
    ElseIf containmentRequired And m_debugMode Then
        Debug.Print "BiDi: Shadow Unlocker ambiguous residual state contained for future realms."
    End If
End Function

' Returns the registered preload script ID. Raises on protocol or malformed-response failure.
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

    ' A prior ambiguous add must first be covered by a newer inert generation. If containment
    ' cannot be established, do not stack another state-changing preload on top of uncertainty.
    If m_shadowUnlockerContainmentRisk Then
        Dim riskSuppressorId As String
        Dim riskErrNum As Long
        Dim riskErrSource As String
        Dim riskErrDesc As String
        If Not TryInstallShadowUnlockerSuppressor(riskSuppressorId, riskErrNum, _
                                                  riskErrSource, riskErrDesc) Then
            Dim riskRaiseNum As Long
            riskRaiseNum = riskErrNum
            If riskRaiseNum = 0 Then riskRaiseNum = ERR_BIDI_PROTOCOL
            Err.Raise riskRaiseNum, "BiDi:ExecuteEnableGlobalShadowUnlocker", _
                      "Shadow Unlocker registration blocked because a prior ambiguous preload " & _
                      "could not be contained." & AppendErrorDescription(riskErrDesc)
        End If
    End If

    Dim generation As Double
    generation = NextShadowUnlockerGuardGeneration()

    Dim scriptId As String
    Dim addErrNum As Long
    Dim addErrSource As String
    Dim addErrDesc As String
    If Not TryAddShadowUnlockerPreload(generation, True, scriptId, _
                                       addErrNum, addErrSource, addErrDesc) Then
        Dim finalDesc As String
        finalDesc = addErrDesc

        If IsAmbiguousDispatchFailure(addErrNum, addErrSource, addErrDesc) Then
            m_shadowUnlockerContainmentRisk = True

            Dim containmentId As String
            Dim containmentErrNum As Long
            Dim containmentErrSource As String
            Dim containmentErrDesc As String
            If TryInstallShadowUnlockerSuppressor(containmentId, containmentErrNum, _
                                                  containmentErrSource, containmentErrDesc) Then
                finalDesc = finalDesc & vbCrLf & _
                            "ShadowUnlockerAmbiguousContained: a higher-generation inert preload " & _
                            "was installed, so any untracked registration cannot alter attachShadow " & _
                            "in future realms."
            Else
                finalDesc = finalDesc & vbCrLf & _
                            "ShadowUnlockerAmbiguousUncontained: browser-side preload state may " & _
                            "remain active." & AppendErrorDescription(containmentErrDesc)
            End If
        End If

        Dim addRaiseNum As Long
        addRaiseNum = addErrNum
        If addRaiseNum = 0 Then addRaiseNum = ERR_BIDI_PROTOCOL

        Dim addRaiseSource As String
        addRaiseSource = addErrSource
        If Len(addRaiseSource) = 0 Then addRaiseSource = "BiDi:ExecuteEnableGlobalShadowUnlocker"

        Err.Raise addRaiseNum, addRaiseSource, finalDesc
    End If

    m_shadowUnlockerRegistered = True
    m_shadowUnlockerScriptId = scriptId
    m_shadowUnlockerContainmentRisk = False
    ExecuteEnableGlobalShadowUnlocker = scriptId

    ' A successful higher generation supersedes a previous containment preload in every future
    ' realm, so removing the known inert preload is only resource cleanup, not a correctness gate.
    BestEffortRemoveKnownShadowUnlockerSuppressor
    
    If m_debugMode Then
        Debug.Print "BiDi: Global Shadow DOM Unlocker registered (ID: " & scriptId & _
                    ", generation=" & ShadowUnlockerGenerationLiteral(generation) & ")"
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

' The forensic incident ring is independent of Discovery Log recording and DebugMode.
' GetForensicIncidentLog returns a snapshot so callers cannot mutate the internal ring.
Public Property Get ForensicIncidentCount() As Long
    EnsureForensicIncidentLog
    ForensicIncidentCount = m_forensicIncidentLogs.count
End Property

Public Property Get ForensicIncidentDroppedHeadCount() As Long
    ForensicIncidentDroppedHeadCount = m_forensicIncidentDroppedHeadCount
End Property

Public Function GetForensicIncidentLog() As Collection
    EnsureForensicIncidentLog

    Dim snapshot As New Collection
    Dim item As Variant
    For Each item In m_forensicIncidentLogs
        snapshot.Add CStr(item)
    Next item

    Set GetForensicIncidentLog = snapshot
End Function

Public Sub ClearForensicIncidentLog()
    RejectIfExclusiveCaptureActive "ClearForensicIncidentLog"
    Set m_forensicIncidentLogs = New Collection
    m_forensicIncidentDroppedHeadCount = 0
End Sub

' DebugMode controls wrapper trace only; it does not affect Discovery Log or communicator
' wire tracing.
Public Property Get DebugMode() As Boolean
    DebugMode = m_debugMode
End Property
Public Property Let DebugMode(ByVal value As Boolean)
    m_debugMode = value
End Property

' Controls guarded Office DoEvents inside communicator connection/command waits.
' Default False preserves the non-reentrant wait behavior unless the caller explicitly opts in.
Public Property Get AllowUiPump() As Boolean
    AllowUiPump = socket_.AllowUiPump
End Property
Public Property Let AllowUiPump(ByVal value As Boolean)
    socket_.AllowUiPump = value
End Property

Public Property Get ShadowUnlockerRegistered() As Boolean
    ShadowUnlockerRegistered = m_shadowUnlockerRegistered
End Property
Public Property Get ShadowUnlockerScriptId() As String
    ShadowUnlockerScriptId = m_shadowUnlockerScriptId
End Property
Public Property Get AutoClickerRegistrationCount() As Long
    EnsureAutoClickerRegistry
    AutoClickerRegistrationCount = m_autoClickerScriptIds.count
End Property
Public Function GetAutoClickerScriptIdByXPath(ByVal xpath As String) As String
    EnsureAutoClickerRegistry
    Dim registrationKey As String
    registrationKey = BuildAutoClickerXPathKey(xpath)
    If m_autoClickerScriptIds.Exists(registrationKey) Then
        GetAutoClickerScriptIdByXPath = CStr(m_autoClickerScriptIds(registrationKey))
    End If
End Function
Public Function GetShadowAutoClickerScriptId(ByVal selectorPath As Variant) As String
    EnsureAutoClickerRegistry
    Dim registrationKey As String
    registrationKey = BuildAutoClickerShadowKey(selectorPath, "GetShadowAutoClickerScriptId")
    If m_autoClickerScriptIds.Exists(registrationKey) Then
        GetShadowAutoClickerScriptId = CStr(m_autoClickerScriptIds(registrationKey))
    End If
End Function
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

    ' Missing paths return Empty; explicit JSON null is represented as VBA Null.
    ' Both mean "no string value" to callers. Never pass Null to CStr.
    If IsObject(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    GetPathString = CStr(v)
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

    Call Me.ExecuteElementActionByXPath( _
        triggerXPath, ElementActionClick, searchTimeoutMs:=searchTimeoutMs, _
        waitNetworkIdle:=False, maxTimeoutMs:=maxTimeoutMs, contextId:=targetCtx)

    Call ExecuteUniqueVisibleRoleOptionClick( _
        targetCtx, optionText, searchTimeoutMs, True, maxTimeoutMs, _
        "ExecuteWizSelectDropdown")
End Sub
