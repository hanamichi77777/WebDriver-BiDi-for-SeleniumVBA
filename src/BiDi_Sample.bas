Attribute VB_Name = "BiDi_Sample"
Option Explicit
' ==========================================================================
' WebDriver BiDi for SeleniumVBA(https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA)
' Version: 2.6
' MIT License Copyright (c) hanamichi77777
' ==========================================================================
' SAMPLE MODULE GUIDE
' --------------------------------------------------------------------------
' This module is a practical tour of WebDriver BiDi for SeleniumVBA. Each
' MainXX procedure demonstrates one synchronization or browser-control pattern.
'
' What sample users usually need to know:
'   1. PURPOSE      - what problem the sample solves.
'   2. PREREQUISITE - browser, page state, extension, or manual action required.
'   3. CUSTOMIZE    - URL, XPath, signal pattern, timeout, or verification point.
'   4. EXPECTED     - the observable result that proves the wait worked.
'   5. DIAGNOSIS    - what to inspect when a live site changes or a wait is early.
'
' General usage notes:
'   - Run one MainXX procedure at a time from the VBA editor.
'   - Chrome and Edge both require caps.EnableBiDiMode before OpenBrowser.
'   - Live-site HTML, ARIA labels, URLs, and network endpoints may change. Treat
'     selectors and signal patterns as examples to re-discover, not permanent APIs.
'   - Discovery Log samples show how to collect evidence before changing waits.
'   - Keep bidi.Shutdown before browser shutdown so the WebSocket and injected
'     browser-side helpers are released in a predictable order.
'   - Production macros should add an error handler that performs the same cleanup.
'
' Sample index:
'   Main01 - install the Google Translate extension through WebDriver BiDi
'   Main02 - lazy-load scrolling
'   Main03 - form input, resource filtering, and Discovery Log
'   Main04 - manual login/navigation wait
'   Main05 - choosing when to disable or retain an action wait
'   Main06 - iframe browsing-context targeting
'   Main07 - consent auto-clicker, Shadow DOM, and network gate
'   Main08 - heavy Google Flights SPA stress scenario
'   Main09 - manual Discovery Log recorder
'   Main10 - content-signal gate for settle-to-render gaps
'
' Wait-selection guide:
'   - Ordinary action with observable SPA activity -> Execute* action default wait.
'   - Known response marks completion              -> ArmNetworkSignal.
'   - Existing DOM subtree must be rewritten       -> ArmContentSignal.
'   - Existing element must become visible         -> ArmVisibilitySignal.
'   - Unknown third-party SPA                      -> record a Discovery Log first.
' ==========================================================================
' Foreground message box: declared with hWnd=0 (no owner) so the dialog is an
' independent WS_EX_TOPMOST window and floats above other processes (e.g. the
' browser) during automation. VBA's own MsgBox is owned by Excel and cannot do
' this when Excel sits behind the browser, so MessageBoxA is used here on purpose.
Public Declare PtrSafe Function MESSAGEbox Lib "user32.dll" Alias "MessageBoxA" _
                                (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
Public Const MB_OK = &H0                          ' OK button flag
Public Const MB_ForeFront = &H40000              ' Topmost flag (MB_TOPMOST)

' ==========================================================================
' Main01 - INSTALL GOOGLE TRANSLATE THROUGH WEBDRIVER BiDi
' --------------------------------------------------------------------------
' USER GOAL:
'   Install one unpacked Chrome extension into the automation browser and learn
'   why current Chrome requires a runtime WebDriver BiDi command instead of the
'   traditional extension-loading capability used at session startup.
'
' WHY THIS SAMPLE USES BiDi INSTEAD OF CLASSIC CAPABILITIES:
'   Older Selenium examples commonly loaded an extension before browser startup
'   through ChromeOptions / capabilities, such as AddExtensions or a
'   --load-extension argument.
'
'   In branded Chrome 137 and later, the old --load-extension startup path was
'   removed. Therefore, for the current normal Chrome build used by this project,
'   the classic capability route is not a practical substitute for loading this
'   unpacked extension.
'
'   This sample first creates a normal WebDriver session with BiDi enabled, then
'   sends the W3C WebDriver BiDi command webExtension.install. Installation is a
'   runtime browser command, not a startup capability.
'
' REQUIRED CHROME STARTUP ARGUMENTS:
'   --remote-debugging-pipe
'       Uses Chrome's pipe-based debugging transport required by the current
'       extension-debugging path. Do not replace it with --remote-debugging-port.
'
'   --enable-unsafe-extension-debugging
'       Allows the automation session to install/debug an unpacked extension.
'       Use this only in a controlled automation browser/profile.
'
'   caps.EnableBiDiMode
'       Requests the WebDriver BiDi WebSocket endpoint. Without it,
'       ExecuteWebExtensionInstall cannot send webExtension.install.
'
' WHAT extensionPath MUST POINT TO:
'   - An absolute local folder containing manifest.json at its top level.
'   - For an extension copied from Chrome's profile, use the version directory,
'     not merely the extension-ID directory.
'   - The Google Translate extension ID used below is:
'         aapbdbdomjkkjkaonfhkkikfgjllcleb
'   - Chrome may update the version directory automatically. If 2.0.17_0 no
'     longer exists, inspect the Extensions folder and replace that segment.
'
' IMPORTANT DISTINCTIONS:
'   - The source folder is read from the user's normal Chrome profile, but the
'     extension is installed into the browser session opened by this macro.
'   - This does not use caps.AddExtensions and does not depend on
'     --load-extension.
'   - No page navigation or SPA wait is involved. A failure here is therefore
'     normally related to the path, manifest, browser startup arguments,
'     Chrome/ChromeDriver support, or an organizational browser policy.
'   - The wrapper intentionally does not retry webExtension.install after an
'     ambiguous transport failure, because repeating an install command could
'     duplicate a state-changing operation whose first result is unknown.
'
' EXPECTED RESULT:
'   ExecuteWebExtensionInstall returns a successful BiDi response and Chrome
'   shows the Google Translate extension as installed in the automation session.
'
' SECURITY / ENTERPRISE NOTE:
'   A managed Chrome environment may block extension installation or restrict
'   the required debugging switches. Such a policy restriction is external to
'   SeleniumVBA and cannot be bypassed by changing the wait settings.
' ==========================================================================
Public Sub Main01()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

    ' This demo targets Chrome because the source path and startup requirements
    ' below are Chrome-specific.
    .StartChrome

    ' Capabilities still create the browser session and enable BiDi, but they do
    ' NOT carry the extension itself. The extension is installed after startup by
    ' ExecuteWebExtensionInstall.
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities

    ' Optional user-interface settings.
    caps.AddArguments "--start-maximized"
    caps.AddArguments "--propagate-iph-for-testing"

    ' Required for runtime installation of an unpacked extension in current Chrome.
    caps.AddArguments "--remote-debugging-pipe"
    caps.AddArguments "--enable-unsafe-extension-debugging"

    ' Mandatory for obtaining the WebDriver BiDi WebSocket endpoint.
    caps.EnableBiDiMode

    ' Start Chrome first. webExtension.install is a post-startup BiDi command.
    .OpenBrowser caps

    ' Connect the VBA wrapper to the BiDi endpoint exposed by ChromeDriver.
    Dim bidi As New BiDiCommandWrapper
    bidi.ConnectTo .GetWebSocketUrl
    'bidi.DebugMode = False  'Mute developer trace in the Immediate window (default: On)

    ' Point to the unpacked Google Translate version directory that directly
    ' contains manifest.json. Update the version segment after Chrome updates it.
    Dim extensionPath As String
    extensionPath = Environ("LOCALAPPDATA") & _
        "\Google\Chrome\User Data\Default\Extensions\" & _
        "aapbdbdomjkkjkaonfhkkikfgjllcleb\2.0.17_0"

    ' Send W3C WebDriver BiDi webExtension.install with extensionData.type="path".
    ' This is the actual installation step; no navigation or wait command follows.
    Dim installResponse As String
    installResponse = bidi.ExecuteWebExtensionInstall(extensionPath)

    ' Keep the raw response available for troubleshooting in the Immediate window.
    Debug.Print installResponse

    Dim msgText As String, msgCaption As String
    msgText = "Google Translate was installed through WebDriver BiDi." & vbCrLf & _
              "The raw command response was written to the Immediate window."
    msgCaption = "Extension Installation Complete"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront

    ' Cleanup order: close the BiDi connection before shutting down the browser.
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

  End With
End Sub

' ==========================================================================
' Main02 - LAZY-LOAD SCROLL
' --------------------------------------------------------------------------
' USER GOAL:
'   Load content that appears only after repeated scrolling, then continue only
'   after scrolling and SPA activity have reached a practical quiescent state.
'
' CUSTOMIZE:
'   Replace the URL and final article XPath with selectors from the target site.
'   Because this is a live page, the article selector and amount of lazy-loaded
'   content may change without notice.
'
' EXPECTED RESULT:
'   A message shows the number of article links found after ExecuteLazyLoadScroll.
'   The count is evidence of loaded content, not a fixed expected number.
' ==========================================================================
Public Sub Main02()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
   
  ' Start
  .StartEdge
   
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
  ' /Open maximized
  caps.AddArguments "--start-maximized"
  ' ==========================================
  ' Enable BiDi (True is mandatory for this program)
  caps.EnableBiDiMode
  ' ==========================================
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
    ' Navigate with the BiDi wrapper so navigation and initial page activity are
    ' synchronized before the lazy-load routine starts.
    Dim url As String: url = "https://note.com/topic/novel"
    bidi.ExecuteNavigateAndGetStatus url, True
    ' Scroll in controlled steps. The method performs a best-effort SPA-quiescence
    ' assessment after scrolling; it does not claim that an infinite feed is exhausted.
    bidi.ExecuteLazyLoadScroll
    
    ' Business-level verification: take one snapshot of the DOM after the wait.
    Dim elms_title As WebElements ' List of article elements
    
    ' FindElements returns an empty collection when no matching article is present,
    ' which makes the resulting count useful for quick manual verification.
    Set elms_title = .FindElements(By.xpath, "//a[contains(@href, '/n/') and @aria-label and @title]")
    
    ' Display result
    Dim msgText As String, msgCaption As String
    msgText = "Article count: " & elms_title.count
    msgCaption = "Wait Complete"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
End With

End Sub

' ==========================================================================
' Main03 - FORM INPUT + RESOURCE FILTERING + DISCOVERY LOG
' --------------------------------------------------------------------------
' USER GOAL:
'   Automate a real route-search form while recording enough evidence to tune
'   waits on a third-party SPA.
'
' WHY TWO FILTER TYPES EXIST:
'   - ExecuteEnableResourceBlocking prevents selected requests from being sent.
'   - AddIdleIgnoreNetworkPattern still allows a request, but excludes matching
'     background traffic from SPA-idle judgment.
'   Blocking is stronger and can break a site; ignoring is safer for required but
'   continuously noisy telemetry.
'
' CUSTOMIZE / DIAGNOSIS:
'   Update URLs, field XPaths, and filters for the target site. If the action ends
'   too early or never stabilizes, inspect discovery_log.txt before adding delays.
' ==========================================================================
Public Sub Main03()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    ' Start
    .StartEdge
      
    ' Browser startup settings (for both Chrome and Edge)
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    ' Open maximized
    caps.AddArguments "--start-maximized"
    ' Enable BiDi (True is mandatory for this program)
    caps.EnableBiDiMode
    .OpenBrowser caps
     Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
    ' ==========================================
    ' Resource blocking reduces traffic before it reaches the page. Block only
    ' resources proven non-essential; an over-broad wildcard can change site behavior.
    ' ==========================================
    Dim blockList As Variant
    blockList = Array( _
        "*criteo.net*", "*criteo.com*", "*googlesyndication.com*", "*doubleclick.net*", "*adtrafficquality.google*")
    bidi.ExecuteEnableResourceBlocking blockList
    ' ==========================================
    ' Idle-ignore patterns leave requests functional but prevent known background
    ' activity from keeping the SPA consensus permanently busy.
    ' ==========================================
    bidi.AddIdleIgnoreNetworkPattern "*generate_204*"
    bidi.AddIdleIgnoreNetworkPattern "*/api/compat/suggest/*"
    bidi.AddIdleIgnoreNetworkPattern "*client-side-metrics*"
    bidi.AddIdleIgnoreNetworkPattern "*fundingchoicesmessages.google.com*"
    
    ' Page transition
    Dim url As String: url = "https://world.jorudan.co.jp/mln/en/"
    bidi.ExecuteNavigateAndGetStatus url
    ' ==========================================
    ' Start recording after initial navigation so the log focuses on the form
    ' interaction being investigated. The saved log is intended for human or AI
    ' analysis of requests, mutations, suppressed noise, and stability margins.
    bidi.StartDiscoveryLog
    ' ==========================================
    ' Enter values as a user would. The default action wait is retained because
    ' these fields may trigger suggestions or other asynchronous page work.
    ' Departure: Tokyo
    bidi.ExecuteInputValueByXPath "//input[@id='from_value']", "Tokyo"
    ' Arrival: Shinjuku
    bidi.ExecuteInputValueByXPath "//input[@id='to_value']", "Shinjuku"
    ' Two selector styles are demonstrated below. Because an exact id also matches
    ' the starts-with expression, confirm whether the live page truly requires two
    ' separate clicks. If both resolve to the same element, keep only one to avoid
    ' duplicate submission when adapting this sample.
    bidi.ExecuteClickByXPath "//button[starts-with(@id, 'search_button_main')]"
    ' Exact-id selector example.
    bidi.ExecuteClickByXPath "//button[@id='search_button_main']"
    ' ==========================================
    ' Stop only after the final action wait completes; stopping earlier would omit
    ' the response/mutation tail needed to diagnose settle-to-render gaps.
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    ' ==========================================
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
    
    ' Completion
    MsgBox "Completed"
    
End With
End Sub

' ==========================================================================
' Main04 - MANUAL LOGIN NAVIGATION WAIT
' --------------------------------------------------------------------------
' USER GOAL:
'   Let the user complete login manually, while VBA waits for navigation to the
'   authenticated page instead of polling the UI with arbitrary Sleep calls.
'
' PREREQUISITE:
'   After the login page opens, enter the sample credentials in the browser and
'   submit the form within 30 seconds. Credentials are intentionally not automated.
'
' CUSTOMIZE:
'   Replace loginUrl, the expected URL fragment, and the timeout. For identity
'   providers that stay on one URL, wait for an authenticated DOM signal instead.
' ==========================================================================
Public Sub Main04()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
     
    .StartEdge
    
    ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    ' ==========================================
    ' Enable BiDi (Mandatory)
    caps.EnableBiDiMode
    ' ==========================================
      
    ' Open
    .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
        
    ' Open the login page and wait for its initial navigation/network activity.
    Dim loginUrl As String: loginUrl = "https://hotel-example-site.takeyaqa.dev/ja/login.html"
    'userName = "ichiro@example.com"
    'pw = "password"
    bidi.ExecuteNavigateAndGetStatus loginUrl, True
      
    ' waitNetworkIdle=True means a matching URL alone is not enough; the method
    ' also waits for the configured post-navigation activity to settle. A zero
    ' timeout would be normalized because immediate checking and idle waiting are
    ' intentionally mutually exclusive.
    Dim isLoginSuccess As Boolean
    isLoginSuccess = bidi.ExecuteIsUrlContains("https://hotel-example-site.takeyaqa.dev/ja/mypage.html", True, , 30000)
      
    ' Verification
    Dim msgText As String, msgCaption As String
      
    If isLoginSuccess Then
        msgText = "BiDi Event Received!" & vbCrLf & "Login (Navigation) Confirmed."
        msgCaption = "Success"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
        msgText = "Timed out while waiting for login event."
        msgCaption = "Failed"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
      
  End With
End Sub

' ==========================================================================
' Main05 - CHOOSING WHEN AN ACTION NEEDS A POST-ACTION WAIT
' --------------------------------------------------------------------------
' USER GOAL:
'   Avoid unnecessary waiting for actions known to be synchronous, while keeping
'   a real SPA wait for the action that triggers an asynchronous DOM update.
'
' EXPECTED RESULT:
'   After clicking Add Label, #update_butter contains "Done!". Debug.Assert
'   breaks in the VBA editor if the expected update was not observed.
'
' ADAPTATION RULE:
'   Pass waitForCompletion=False only when the action is known not to start
'   relevant asynchronous work. Disabling the wait merely for speed can reintroduce
'   race conditions.
' ==========================================================================
Public Sub Main05()
    Dim driver As WebDriver: Set driver = New WebDriver
    With driver

    .StartEdge
    
   ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = driver.CreateCapabilities
    caps.EnableBiDiMode
    
    ' Open
    .OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl

    .NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"

    ' These two actions do not need a post-action SPA wait on this test page.
    ' The fourth argument False disables only that post-action wait.
    bidi.ExecuteInputValueByXPath "//input[@name='typer']", "aaa", , False
    bidi.ExecuteClickByXPath "//input[@id='red']", , False
    
    ' Add Label starts asynchronous work. Keep waiting enabled and require a
    ' 1000ms stable window so the final DOM update is present before verification.
    bidi.ExecuteClickByXPath "//input[@value='Add Label']", , , 1000

    Debug.Assert driver.FindElement(By.xpath, "//div[@id='update_butter']").GetText = "Done!"

    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
   End With
End Sub

' ==========================================================================
' Main06 - IFRAME CONTEXT DISCOVERY AND TARGETED ACTION
' --------------------------------------------------------------------------
' USER GOAL:
'   Click an element inside an iframe without switching SeleniumVBA's active frame.
'   WebDriver BiDi identifies the child browsing context, and the action receives
'   that context id explicitly.
'
' CUSTOMIZE:
'   - Replace the page URL.
'   - Replace the iframe URL fragment passed to GetIframeContextIdByUrl.
'   - Replace the XPath inside the frame.
'
' DIAGNOSIS:
'   If no context is found, inspect the actual iframe URL after navigation; it may
'   be redirected, generated dynamically, or represented by a nested frame.
' ==========================================================================
Public Sub Main06()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartEdge
    
  ' Browser startup settings
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
  caps.AddArguments "--start-maximized"
  ' ==========================================
  ' Enable BiDi (Mandatory)
  caps.EnableBiDiMode
  ' ==========================================
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' =========================================
   bidi.ExecuteNavigateAndGetStatus "https://www.customs.go.jp/toukei/srch/index.htm?M=01&P=0", False
   
   ' Resolve the child browsing context by a stable fragment of its current URL.
   ' Passing conID to ExecuteClickByXPath scopes both lookup and execution to that frame.
   Dim conID As String
   conID = bidi.GetIframeContextIdByUrl("jccht00d")
   bidi.ExecuteClickByXPath "//input[@id='la_imp']", , , , , conID
   
   ' Cleanup
   bidi.Shutdown: Set bidi = Nothing
   .CloseBrowser: .Shutdown
   
  End With
End Sub

' ==========================================================================
' Main07 - SHADOW DOM LOGIN FLOW + PRE-NAVIGATION AUTO-CLICKER + NETWORK GATE
' --------------------------------------------------------------------------
' USER GOAL:
'   Automate a complex third-party site that combines a consent banner, Shadow DOM,
'   SPA navigation, and an input page protected by framework/WAF timing behavior.
'
' WHY THE ORDER MATTERS:
'   - StartDiscoveryLog before the interaction to capture its complete causal tail.
'   - RegisterAutoClicker before navigation so the browser-side helper can dismiss
'     the consent button as soon as it appears.
'   - ArmNetworkSignal immediately before the Shadow click; armed signals are
'     one-shot and belong to the next action wait only.
'
' CUSTOMIZE / DIAGNOSIS:
'   Re-discover the consent XPath, shadow selector, username XPath, and network
'   pattern if ServiceNow changes. The string metadata/application was selected
'   from observed traffic and is not a universal login-completion endpoint.
' ==========================================================================
Public Sub Main07()
    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim targetUrl As String: targetUrl = "https://developer.servicenow.com/"
        
    With driver
    .StartEdge
    Set caps = .CreateCapabilities
    caps.EnableBiDiMode
    .OpenBrowser caps
    Set bidi = New BiDiCommandWrapper
    bidi.ConnectTo driver.GetWebSocketUrl
    
    ' ==========================================
    ' Start before auto-clicker registration/navigation so the log captures consent,
    ' navigation, the armed network response, and the final input action.
    bidi.StartDiscoveryLog
    ' ==========================================

    ' Pre-navigation registration is essential: registering after navigation can
    ' miss a short-lived banner or allow it to block the first target action.
    bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
    
    ' NavigateTo Page
    bidi.ExecuteNavigateAndGetStatus targetUrl
        
    ' Gate the next action on a response known to accompany creation of the sign-in
    ' experience. Arm immediately before Act to avoid attaching it to another wait.
    bidi.ArmNetworkSignal "metadata/application"
    bidi.ExecuteShadowClick "#utility-sign-in button"
            
    ' Dummy input used only to prove the newly rendered page is interactive.
    ' Replace "aaa" with a non-sensitive test value when adapting the sample.
    bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa"
    
    ' ==========================================
    ' Stop and Save AFTER the wait is finished
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    ' ==========================================
    
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
           
    End With
End Sub

' ========================================================================================
' Main08 - GOOGLE FLIGHTS: HEAVY REACTIVE SPA STRESS SAMPLE
' ----------------------------------------------------------------------------------------
' USER GOAL:
'   See how selector strategy, noise filtering, trusted keyboard input, and explicit
'   completion signals work together on a high-traffic Google SPA.
'
' SCENARIO:
'   Select One way, enter Sapporo -> Paris, choose a departure date, and search.
'
' WHY THIS SAMPLE IS ADVANCED:
'   Google Flights can replace active inputs, emit continuous telemetry, mutate large
'   DOM subtrees, and settle network responses before the corresponding UI is painted.
'   The sample therefore combines:
'     - ARIA-based selectors instead of obfuscated CSS classes
'     - [last()] for duplicated/replaced combobox inputs
'     - trusted per-character input with active-element tracking
'     - blocked non-essential resources and ignored required background noise
'     - one-shot network/visibility gates around known render boundaries
'
' WHAT USERS SHOULD CUSTOMIZE:
'   - City names, date selection, and final search behavior.
'   - XPath/ARIA labels when Google changes the UI or locale. --lang=en is set so
'     the English aria-label selectors below remain meaningful.
'   - Signal patterns after inspecting a fresh Discovery Log. GetCalendarPicker and
'     GetShoppingResults are observed implementation details, not stable public APIs.
'
' EXPECTED RESULT:
'   The search action completes after the shopping-results request is observed and the
'   wrapper's SPA wait reaches its configured conclusion. The sample demonstrates a
'   robust strategy; it cannot guarantee immunity from future live-site changes.
'
' TROUBLESHOOTING ORDER:
'   1. Confirm the target XPath still resolves to the visible control.
'   2. Check whether a consent dialog, locale change, or account state altered the UI.
'   3. Review discovery_log.txt for new noise or renamed completion requests.
'   4. Prefer updating selectors/signals over adding arbitrary fixed delays.
' ========================================================================================

Public Sub Main08()
    Dim driver As WebDriver: Set driver = New WebDriver
    With driver
        .StartEdge
        
        Dim caps As WebCapabilities: Set caps = .CreateCapabilities
        caps.AddArguments "--start-maximized"
        caps.AddArguments "--lang=en"
        caps.EnableBiDiMode
        
        .OpenBrowser caps
        Dim bidi As New BiDiCommandWrapper
        bidi.ConnectTo .GetWebSocketUrl
        ' ==========================================
        ' Block only traffic judged non-essential to the tested workflow. This
        ' reduces event volume but should be revalidated after major site changes.
        ' ==========================================
        Dim blockList As Variant
        blockList = Array( _
            "*googletagmanager*", "*doubleclick*", "*googlesyndication*", _
            "*google-analytics*", _
            "*/collect*", "*/beacon*", "*pagead*")
        bidi.ExecuteEnableResourceBlocking blockList
        ' ==========================================
        ' Ignore required background requests in idle judgment rather than blocking
        ' them. Continuous telemetry would otherwise prevent a stable consensus.
        ' ==========================================
        bidi.AddIdleIgnoreNetworkPattern "/log?"
        bidi.AddIdleIgnoreNetworkPattern "*generate_204*"
        bidi.AddIdleIgnoreNetworkPattern "GetAsyncData"
        
        ' ==========================================
        ' Record the entire scenario. The output is the evidence source for selector,
        ' noise, signal, and stability-window adjustments after a live-site change.
        bidi.StartDiscoveryLog
        ' ==========================================
        
        ' Navigation
        Dim url As String: url = "https://www.google.com/travel/flights"
        bidi.ExecuteNavigateAndGetStatus url
        
        ' STEP 0: Set ticket type through ARIA semantics rather than volatile classes.
        ' ExecuteSelectValueByXPath performs the interaction and the normal SPA wait.
        ' Reference selector retained to show the more specific listbox-trigger form.
        ' The executable line below intentionally uses the simpler first-combobox XPath.
        Dim ticketTypeTrigger As String
        ticketTypeTrigger = "(//div[@role='combobox' and @aria-haspopup='listbox'])[1]"
        
        ' Switch from "Round trip" to "One way"
        bidi.ExecuteSelectValueByXPath "(//div[@role='combobox'])[1]", "One way"
                
        ' STEP 1: Set Departure City - "Sapporo"
        ' [last()] ensures we resolve the final (real) input when Wiz spawns
        ' multiple copies.  Phase 0's stability check provides a second layer
        ' of protection inside the JS execution.
        
        Dim depXPath As String
        depXPath = "(//input[contains(@aria-label, 'Where from')])[last()]"
        
        ' The wrapper clicks, tracks the final active input, clears it, types with
        ' trusted events, and validates the resulting value before returning.
        bidi.ExecuteInputValueByXPath depXPath, "Sapporo"
        
        ' Limit the selector to the visible listbox and choose by accessible text, not
        ' by list index. minStableMs=1000 adds margin for framework replacement/mutation.
        Dim depSuggestXPath As String
        depSuggestXPath = "//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Sapporo')][1]"
        bidi.ExecuteClickByXPath depSuggestXPath, minStableMs:=1000
        
        ' STEP 2: Set Destination City - "Paris"
        ' After selecting departure, Google Flights auto-activates
        ' the destination combobox.
        
        Dim destXPath As String
        destXPath = "(//input[contains(@aria-label, 'Where to')])[last()]"
        
        ' Reuse the same active-element-aware input path for the destination.
        bidi.ExecuteInputValueByXPath destXPath, "Paris"
        
        ' Click matching suggestion
        Dim destSuggestXPath As String
        destSuggestXPath = "//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Paris')][1]"
        bidi.ExecuteClickByXPath destSuggestXPath, minStableMs:=1000
        
        ' STEP 3: Select Departure Date (one way -> no return date)
        ' The calendar may open automatically, but the explicit Departure click below
        ' keeps the sample deterministic when account/UI state differs.
        ' Strategy:
        '   - Pick a single departure date (the 8th available date cell)
        
        ' Arm completion signals for the fare calendar:
        '   - Wait for the GetCalendarPicker request to complete
        '   - Wait until fare cells (data-gs) become visible
        ' Together these gates reduce the risk of treating network quiet as completion
        ' before the fare cells become observable. They remain one-shot diagnostics,
        ' not permanent guarantees against every future Google UI implementation.
        bidi.ArmNetworkSignal "GetCalendarPicker"
        bidi.ArmVisibilitySignal "//div[@data-gs]"
        bidi.ExecuteClickByXPath "//input[@aria-label='Departure']"
        
        bidi.ExecuteClickByXPath "(//div[@role='gridcell' and @aria-hidden='false'])[8]//div[@role='button']"
        bidi.ExecuteClickByXPath "//button[contains(., 'Done')]"
        
        ' STEP 4: Click Search Button
        Dim searchXPath As String
        searchXPath = "//button[@aria-label='Search']"
        
        ' One-shot gate: the next action wait cannot conclude as stable until a
        ' matching shopping-results response is observed. Re-discover this pattern
        ' when the log no longer contains the expected request.
        bidi.ArmNetworkSignal "GetShoppingResults"
        bidi.ExecuteClickByXPath searchXPath
        
        ' ==========================================
        ' Stop and Save AFTER the wait is finished
        Dim logPath As String
        logPath = .ResolvePath(".\") & "\discovery_log.txt"
        bidi.StopAndSaveDiscoveryLog logPath
        ' ==========================================
        
        ' Cleanup
        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown
        
        ' Completion
        MsgBox "Google Flights Test Completed"
        
    End With
End Sub

' ==========================================================================
' Main09 - MANUAL DISCOVERY RECORDER
' --------------------------------------------------------------------------
' USER GOAL:
'   Discover which requests, DOM bursts, navigation events, and background noise
'   occur during a manual action before writing its automation code.
'
' RECOMMENDED METHOD:
'   Record one meaningful user action per run. For example: open a calendar, select
'   a suggestion, or press Search. A narrow recording is easier for a human or AI
'   to map from [REQ] to response, mutation tail, and stability margin.
'
' CUSTOMIZE:
'   Change the URL and RECORDING_SECONDS. Longer is not always better because it
'   captures more unrelated background traffic. excludeImagesAndCss=True removes
'   common static-resource noise but does not alter the page itself.
'
' OUTPUT:
'   discovery_log.txt in the workbook-resolved current directory. Use it to decide
'   what to block, ignore, arm, or verify; do not select a signal from its name alone.
' ==========================================================================
Sub Main09()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    .StartEdge
    
    ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    ' Enable BiDi (Mandatory)
    caps.EnableBiDiMode
    ' Open
    .OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper
    bidi.ConnectTo .GetWebSocketUrl

    ' Navigate to Page
    Dim url As String: url = "https://note.com/"
    bidi.ExecuteNavigateAndGetStatus url
    
    ' ==========================================================
    ' Keep the window short enough to isolate the intended manual interaction.
    Const RECORDING_SECONDS As Long = 20
    ' Show Message (Blocks execution until OK is clicked)
    Dim msgText As String, msgCaption As String
    msgText = "Please prepare the browser for recording." & vbCrLf & vbCrLf & _
              "Click [OK] to start recording." & vbCrLf & _
              "Duration: " & RECORDING_SECONDS & " seconds." & vbCrLf & _
              "Please manually interact with the page immediately after clicking OK."
    msgCaption = "Ready to Record"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    
    ' Start immediately before the manual action. True suppresses image/CSS entries
    ' from the log; it does not block those resources in the browser.
    bidi.StartDiscoveryLog excludeImagesAndCss:=True
    ' Process incoming BiDi events for the observation window while the user acts.
    bidi.RecordEventsForSeconds RECORDING_SECONDS
    ' ==========================================================
    
    ' Manual interaction occurs during RecordEventsForSeconds above. Avoid unrelated
    ' browsing during the window or the causal sequence will be harder to interpret.
    
    ' ==========================================
    ' Stop and Save AFTER the wait is finished
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    ' ==========================================
    
    MsgBox "Discovery Log Saved!" & vbCrLf & logPath
    
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
    
End With
End Sub

' ==========================================================================
' Main10 - CONTENT-SIGNAL GATE FOR THE SETTLE-TO-RENDER GAP
' --------------------------------------------------------------------------
' USER GOAL:
'   Prevent a click wait from concluding on apparent quiet after XHR completion but
'   before the existing results subtree has actually been rewritten.
'
' ARM-THEN-ACT CONTRACT:
'   - The target XPath must already exist when ArmContentSignal is called.
'   - The signal is one-shot and is consumed by the next relevant wait.
'   - Re-arm immediately before every action that requires the same completion gate.
'
' CUSTOMIZE:
'   Replace #table-body with a stable existing container whose rewrite genuinely
'   marks completion for the target application. Do not arm a broad ancestor that
'   changes for unrelated animations or telemetry.
'
' EXPECTED RESULT:
'   Each year click returns only after the table-body rewrite is observed and the
'   remaining SPA-idle conditions are satisfied.
' ==========================================================================
Public Sub Main10()
    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim targetUrl As String: targetUrl = "https://www.scrapethissite.com/pages/ajax-javascript/#2010"

    With driver
    .StartEdge
    Set caps = .CreateCapabilities
    caps.EnableBiDiMode
    .OpenBrowser caps
    Set bidi = New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl
    
    ' These requests remain allowed but are excluded from idle judgment because
    ' they are unrelated background telemetry on this page.
    bidi.AddIdleIgnoreNetworkPattern "www.scrapethissite.com/cdn-cgi/rum"
    bidi.AddIdleIgnoreNetworkPattern "www.facebook.com/tr/"

    ' ==========================================
    ' Record both year selections so the log shows the armed content signal, the
    ' response/mutation tail, and why the gate must be re-armed for the second click.
    bidi.StartDiscoveryLog
    ' ==========================================
    
    ' NavigateTo Page
    bidi.ExecuteNavigateAndGetStatus targetUrl
    
    ' Arm immediately before the action. The existing #table-body is the semantic
    ' completion marker, bridging the response-settle-to-DOM-render gap.
    bidi.ArmContentSignal "//*[@id='table-body']"
    bidi.ExecuteClickByXPath "//section[@id='oscars']//a[@id='2015']"
    
    ' The previous signal was consumed by the 2015 click, so the 2014 action must
    ' explicitly arm a new one-shot signal.
    bidi.ArmContentSignal "//*[@id='table-body']"
    bidi.ExecuteClickByXPath "//section[@id='oscars']//a[@id='2014']"
    
    ' ==========================================
    ' Stop and Save AFTER the wait is finished
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    ' ==========================================

    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

    End With
End Sub
