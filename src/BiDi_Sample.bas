Attribute VB_Name = "BiDi_Sample"
Option Explicit
' WebDriver BiDi for SeleniumVBA
' https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA
' Version 8.1 / MIT License / Copyright (c) hanamichi77777
'
' Run one MainXX procedure at a time. Live-site selectors and network signals may
' change; rediscover them with the Discovery Log instead of adding fixed delays.
' Use the default action wait for ordinary SPA work and Arm* for known completion
' evidence. Shut down BiDi before the browser.
'
' MessageBoxA with hWnd=0 keeps sample result dialogs above the browser.
Public Declare PtrSafe Function MESSAGEbox Lib "user32.dll" Alias "MessageBoxA" _
                                (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
Public Const MB_OK = &H0                          ' OK button flag
Public Const MB_ForeFront = &H40000              ' Topmost flag (MB_TOPMOST)

' Main01 - Install Google Translate through WebDriver BiDi
' Current Chrome installs this unpacked extension at runtime with
' webExtension.install; classic capabilities and --load-extension are not used.
' Google Translate must already be installed in your everyday Chrome profile
' (Default); the newest version folder containing manifest.json is used, so
' extension updates need no code change. Older Chrome/ChromeDriver builds may
' also need --remote-debugging-pipe and --enable-unsafe-extension-debugging.
' Managed browser policy may still block installation.
Public Sub Main01()
  Const EXTENSION_ID As String = "aapbdbdomjkkjkaonfhkkikfgjllcleb"

  ' Resolve the extension folder before starting the browser.
  Dim extensionRoot As String
  extensionRoot = Environ("LOCALAPPDATA") & _
      "\Google\Chrome\User Data\Default\Extensions\" & EXTENSION_ID

  Dim extensionPath As String
  extensionPath = FindLatestExtensionVersionFolder(extensionRoot)
  If Len(extensionPath) = 0 Then
    MESSAGEbox 0, "No version folder with manifest.json was found under:" & vbCrLf & _
               extensionRoot & vbCrLf & vbCrLf & _
               "Install Google Translate in Chrome (Default profile) first.", _
               "Extension Not Found", MB_OK Or MB_ForeFront
    Exit Sub
  End If

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

    .StartChrome

    Dim caps As WebCapabilities: Set caps = .CreateCapabilities

    caps.AddArguments "--start-maximized"
    caps.AddArguments "--propagate-iph-for-testing"

    caps.EnableBiDiMode

    .OpenBrowser caps

    Dim bidi As New BiDiCommandWrapper
    bidi.ConnectTo .GetWebSocketUrl
    'bidi.DebugMode = False  'Mute developer trace in the Immediate window (default: On)

    Debug.Print "Extension folder: " & extensionPath

    Dim installResponse As String
    installResponse = bidi.ExecuteWebExtensionInstall(extensionPath)

    Debug.Print installResponse

    Dim msgText As String, msgCaption As String
    msgText = "Google Translate was installed through WebDriver BiDi." & vbCrLf & _
              "The raw command response was written to the Immediate window."
    msgCaption = "Extension Installation Complete"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

  End With
End Sub

' Returns the newest version subfolder (for example 2.0.17_0) that directly
' contains manifest.json, or "" when none exists.
Private Function FindLatestExtensionVersionFolder(ByVal extensionRoot As String) As String
  Dim fso As New FileSystemObject
  If Not fso.FolderExists(extensionRoot) Then Exit Function

  Dim versionFolder As Scripting.Folder
  Dim sortKey As String
  Dim bestKey As String
  For Each versionFolder In fso.GetFolder(extensionRoot).SubFolders
    If fso.FileExists(fso.BuildPath(versionFolder.path, "manifest.json")) Then
      sortKey = ExtensionVersionSortKey(versionFolder.name)
      If sortKey > bestKey Then
        bestKey = sortKey
        FindLatestExtensionVersionFolder = versionFolder.path
      End If
    End If
  Next versionFolder
End Function

' "2.0.17_0" -> "0000000002.0000000000.0000000017.0000000000.0000000000"
' Numeric parts are zero-padded so a plain string comparison orders versions.
Private Function ExtensionVersionSortKey(ByVal folderName As String) As String
  Const PART_COUNT As Long = 5
  Dim parts() As String
  parts = Split(Replace(folderName, "_", "."), ".")

  Dim i As Long
  Dim part As String
  For i = 0 To PART_COUNT - 1
    part = "0"
    If i <= UBound(parts) Then
      If parts(i) Like "#*" And Not parts(i) Like "*[!0-9]*" Then part = parts(i)
    End If
    If i > 0 Then ExtensionVersionSortKey = ExtensionVersionSortKey & "."
    ExtensionVersionSortKey = ExtensionVersionSortKey & Right$(String$(10, "0") & part, 10)
  Next i
End Function

' Main02 - Lazy-load scroll
' Replace the URL and result XPath for the target site. The final count is a
' manual verification value, not a fixed expected result.
Public Sub Main02()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

  .StartEdge

  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
  caps.AddArguments "--start-maximized"
  caps.EnableBiDiMode
  .OpenBrowser caps
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
    Dim url As String: url = "https://note.com/topic/novel"
    bidi.ExecuteNavigateAndGetStatus url, True
    ' The scroll wait seeks practical quiescence; it does not exhaust an infinite feed.
    bidi.ExecuteLazyLoadScroll

    Dim elms_title As WebElements ' List of article elements

    Set elms_title = .FindElements(By.xpath, "//a[contains(@href, '/n/') and @aria-label and @title]")

    Dim msgText As String, msgCaption As String
    msgText = "Article count: " & elms_title.count
    msgCaption = "Wait Complete"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
End With

End Sub

' Main03 - Form input, resource filtering and Discovery Log
' Blocking prevents requests; idle-ignore keeps requests functional but excludes
' them from SPA-idle judgment. Tune selectors and filters from discovery_log.txt.
' Stop recording only after the final action wait has completed.
Public Sub Main03()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

    .StartEdge

    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    caps.EnableBiDiMode
    .OpenBrowser caps
     Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
    Dim blockList As Variant
    blockList = Array( _
        "*criteo.net*", "*criteo.com*", "*googlesyndication.com*", "*doubleclick.net*", "*adtrafficquality.google*")
    bidi.ExecuteEnableResourceBlocking blockList
    bidi.AddIdleIgnoreNetworkPattern "*generate_204*"
    bidi.AddIdleIgnoreNetworkPattern "*/api/compat/suggest/*"
    bidi.AddIdleIgnoreNetworkPattern "*client-side-metrics*"
    bidi.AddIdleIgnoreNetworkPattern "*fundingchoicesmessages.google.com*"

    Dim url As String: url = "https://world.jorudan.co.jp/mln/en/"
    bidi.ExecuteNavigateAndGetStatus url
    bidi.StartDiscoveryLog
    bidi.ExecuteInputValue "//input[@id='from_value']", "Tokyo", useKeyEvents:=False
    bidi.ExecuteInputValue "//input[@id='to_value']", "Shinjuku", useKeyEvents:=False
    bidi.ExecuteClick "//button[starts-with(@id, 'search_button_main')]"
    bidi.ExecuteClick "//button[@id='search_button_main']"
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

    MsgBox "Completed"

End With
End Sub

' Main04 - Manual login navigation wait
' Complete login in the browser within 30 seconds. Replace the login URL and
' expected authenticated URL; use a DOM signal instead when login stays on one URL.
Public Sub Main04()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

    .StartEdge

    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    caps.EnableBiDiMode

    .OpenBrowser caps
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl

    Dim loginUrl As String: loginUrl = "https://hotel-example-site.takeyaqa.dev/ja/login.html"
    'userName = "ichiro@example.com"
    'pw = "password"
    bidi.ExecuteNavigateAndGetStatus loginUrl, True

    ' waitNetworkIdle=True requires both the URL match and post-navigation settling.
    Dim isLoginSuccess As Boolean
    isLoginSuccess = bidi.ExecuteIsUrlContains("https://hotel-example-site.takeyaqa.dev/ja/mypage.html", True, , 30000)

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

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

  End With
End Sub

' Main05 - Choosing whether an action needs a post-action wait
' Use waitForCompletion=False only for actions known to be synchronous. Keep the
' wait for Add Label because it triggers an asynchronous DOM update.
Public Sub Main05()
    Dim driver As WebDriver: Set driver = New WebDriver
    With driver

    .StartEdge

    Dim caps As WebCapabilities: Set caps = driver.CreateCapabilities
    caps.EnableBiDiMode

    .OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl

    .NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"

    bidi.ExecuteInputValue "//input[@name='typer']", "aaa", , False
    bidi.ExecuteClick "//input[@id='red']", , False

    bidi.ExecuteClick "//input[@value='Add Label']", , , 1000

    Debug.Assert driver.FindElement(By.xpath, "//div[@id='update_butter']").GetText = "Done!"

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
   End With
End Sub

' Main06 - Iframe context discovery and targeted action
' Resolve the child browsing-context ID from a stable iframe URL fragment, then
' pass that ID to scope lookup and execution to the frame.
Public Sub Main06()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

  .StartEdge

  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
  caps.AddArguments "--start-maximized"
  caps.EnableBiDiMode
  .OpenBrowser caps
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
   bidi.ExecuteNavigateAndGetStatus "https://www.customs.go.jp/toukei/srch/index.htm?M=01&P=0", False

   Dim conID As String
   conID = bidi.GetIframeContextIdByUrl("jccht00d")
   bidi.ExecuteClick "//input[@id='la_imp']", , , , , conID

   bidi.Shutdown: Set bidi = Nothing
   .CloseBrowser: .Shutdown

  End With
End Sub

' Main07 - Shadow DOM login, pre-navigation auto-clicker and network gate
' Register the consent auto-clicker before navigation. Arm the observed
' metadata/application response immediately before the Shadow click; the signal
' is one-shot and site-specific.
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

    bidi.StartDiscoveryLog

    bidi.ExecuteRegisterAutoClicker "//button[@id='truste-consent-button']"

    bidi.ExecuteNavigateAndGetStatus targetUrl

    bidi.ArmNetworkSignal "metadata/application"
    ' searchShadow:=True evaluates the XPath in the document and in every nested ShadowRoot.
    ' The match is searched root by root, so //*[@id=...]//button stays inside one tree.
    bidi.ExecuteClick "//*[@id='utility-sign-in']//button", searchShadow:=True

    ' Use a non-sensitive test value when adapting this sample.
    bidi.ExecuteInputValue "//input[@id='username']", "aaa"

    ' Saves discovery_log.txt in the same folder as the current VBA host file.
    bidi.StopAndSaveDiscoveryLog

    bidi.Shutdown: Set bidi = Nothing

    ' Modern SPA pages may still have background activity when VBA-side work has completed.
    ' CloseBrowser/Shutdown can therefore take a few seconds while the browser exits cleanly.
    ' Do not add an artificial wait just to make shutdown appear faster.
    .CloseBrowser: .Shutdown

    End With
End Sub

' Main08 - Google Flights reactive-SPA stress sample
' Scenario: One way, Sapporo to Paris, select a date, then Search.
' The live-site selectors and GetCalendarPicker/GetShoppingResults signals may
' change; use a fresh Discovery Log. --lang=en supports the English ARIA labels.
' Blocking removes non-essential traffic; idle-ignore retains required telemetry.
' Arm* completion signals are one-shot.
' Observed: with navigator.webdriver=true, Google Flights silently returns a degraded
' list for both Best and Cheapest; AutomationControlled keeps it false.

Public Sub Main08()
    Dim driver As WebDriver: Set driver = New WebDriver
    With driver
        .StartEdge

        Dim caps As WebCapabilities: Set caps = .CreateCapabilities
        caps.AddArguments "--start-maximized"
        caps.AddArguments "--lang=en"
        caps.AddArguments "--disable-blink-features=AutomationControlled"
        caps.EnableBiDiMode

        .OpenBrowser caps
        Dim bidi As New BiDiCommandWrapper
        bidi.ConnectTo .GetWebSocketUrl
        Dim blockList As Variant
        blockList = Array("*googletagmanager*", "*doubleclick*", "*googlesyndication*", "*google-analytics*")
        bidi.ExecuteEnableResourceBlocking blockList
        bidi.AddIdleIgnoreNetworkPattern "/log?"
        bidi.AddIdleIgnoreNetworkPattern "*generate_204*"
        bidi.AddIdleIgnoreNetworkPattern "GetAsyncData"

        bidi.StartDiscoveryLog

        Dim url As String: url = "https://www.google.com/travel/flights"
        bidi.ExecuteNavigateAndGetStatus url

        bidi.ExecuteSelectValue "(//div[@role='combobox'])[1]", "One way"

        ' Actions use the first visible match among duplicates, so plain XPaths suffice.
        Dim depXPath As String
        depXPath = "//input[contains(@aria-label, 'Where from')]"

        Dim depSuggestXPath As String
        depSuggestXPath = "//li[@role='option' and contains(@aria-label, 'Sapporo')]"
        bidi.ExecuteInputValue depXPath, "Sapporo", useKeyEvents:=False

        bidi.ArmNetworkSignal "rpcids=tDoGIe"
        bidi.ExecuteClick depSuggestXPath, minStableMs:=1000


        Dim destXPath As String
        destXPath = "//input[contains(@aria-label, 'Where to')]"

        Dim destSuggestXPath As String
        destSuggestXPath = "//li[@role='option' and contains(@aria-label, 'Paris')]"
        bidi.ExecuteInputValue destXPath, "Paris", useKeyEvents:=False

        bidi.ArmNetworkSignal "rpcids=BVAT3"
        bidi.ExecuteClick destSuggestXPath, minStableMs:=1000


        bidi.ArmNetworkSignal "GetCalendarPicker"
        bidi.ArmVisibilitySignal "//div[@data-gs]"
        bidi.ExecuteClick "//input[@aria-label='Departure']"

        ' Departure date: one week from today. Day cells carry data-iso="yyyy-mm-dd".
        Dim departDate As String
        departDate = Format$(Date + 7, "yyyy-mm-dd")
        bidi.ExecuteClick "//div[@role='gridcell' and @data-iso='" & departDate & "']//div[@role='button']"
        bidi.ExecuteClick "//button[contains(., 'Done')]"


        Dim searchXPath As String
        searchXPath = "//button[@aria-label='Search']"


        bidi.ArmNetworkSignal "GetShoppingResults"
        bidi.ExecuteClick searchXPath, maxTimeoutMs:=30000

        ' Saves discovery_log.txt in the same folder as the current VBA host file.
        bidi.StopAndSaveDiscoveryLog

        bidi.Shutdown: Set bidi = Nothing

        ' Modern SPA pages may still have background activity when VBA-side work has completed.
        ' CloseBrowser/Shutdown can therefore take a few seconds while the browser exits cleanly.
        ' Do not add an artificial wait just to make shutdown appear faster.
        .CloseBrowser: .Shutdown

        MsgBox "Google Flights Test Completed"

    End With
End Sub

' Main09 - Manual Discovery Log recorder with a separate STOP tab.
' The target tab is unchanged. Click STOP RECORDING in the dedicated control tab,
' or close that tab using X to end recording; 300 seconds is a safety maximum.
' Wait for the page to settle before each action; typed input values are not logged.
' Popups and new tabs opened from the target are recorded too (capturePopupActions
' defaults to True); [USER-ACTION-SCOPE] status=added lines list those contexts.
' alert/confirm/prompt dialogs stay open for you to answer (unhandledPromptBehavior=ignore);
' [USER-PROMPT-OPENED]/[USER-PROMPT-CLOSED] lines record the type and your OK/Cancel choice.
' No worksheet button or UserForm is needed.
Public Sub Main09()
    Const MAX_RECORDING_SECONDS As Long = 300

    Dim url As String
    url = Trim$(InputBox("Paste the URL to record.", "Main09 - Target URL", "https://note.com/"))
    If Len(url) = 0 Then Exit Sub
    If Not (LCase$(url) Like "http://*" Or LCase$(url) Like "https://*") Then
        MsgBox "Not a valid http(s) URL: " & url, vbExclamation
        Exit Sub
    End If

    Dim driver As WebDriver
    Dim bidi As BiDiCommandWrapper
    Dim caps As WebCapabilities
    Dim controlCtx As String
    Dim recordingStarted As Boolean
    Dim stoppedFromControl As Boolean
    Dim failure As String
    Dim closeFailure As String

    On Error GoTo Failed
    Set driver = New WebDriver
    driver.StartEdge
    Set caps = driver.CreateCapabilities
    caps.AddArguments "--start-maximized"
    caps.SetUnhandledPromptBehavior svbaIgnore   ' leave JS dialogs open for the person recording
    caps.EnableBiDiMode
    driver.OpenBrowser caps

    Set bidi = New BiDiCommandWrapper
    bidi.ConnectTo driver.GetWebSocketUrl

    MESSAGEbox 0, "Target: " & url & vbCrLf & vbCrLf & _
               "Click OK to start the recording and open the page." & vbCrLf & _
               "Operate the target tab. Then use STOP RECORDING in the separate tab " & _
               "(or close the STOP tab with X)." & vbCrLf & _
               "Safety maximum: " & MAX_RECORDING_SECONDS & " seconds after page load." & vbCrLf & _
               "Allow each action to settle before the next one.", _
               "Ready to Record", MB_OK Or MB_ForeFront

    ' Capture initial navigation and record actions in the target tab, its frames,
    ' and popups/new tabs it opens. The STOP tab and unrelated tabs are excluded.
    bidi.StartDiscoveryLog excludeImagesAndCss:=True, captureUserActions:=True
    recordingStarted = True
    bidi.ExecuteNavigateAndGetStatus url
    controlCtx = bidi.OpenDiscoveryStopTab()
    Debug.Print "Recorder control tab: " & controlCtx
    bidi.RecordEventsForSeconds MAX_RECORDING_SECONDS
    stoppedFromControl = bidi.DiscoveryStopRequested

    ' Close the control tab before saving; this does not discard the already logged stop event.
    ' A dialog left open in the target tab is cancelled first so cleanup scripts can run.
    On Error Resume Next
    bidi.ExecuteHandleUserPrompt False
    Err.Clear
    bidi.CloseDiscoveryStopTab
    If Err.Number <> 0 Then closeFailure = Err.description
    Err.Clear
    On Error GoTo Failed
    bidi.StopAndSaveDiscoveryLog
    recordingStarted = False
    GoTo CleanUp

Failed:
    failure = "Err " & CStr(Err.Number) & ": " & Err.description
    Resume CleanUp

CleanUp:
    On Error Resume Next
    If Not bidi Is Nothing Then
        bidi.CloseDiscoveryStopTab
        If recordingStarted Then bidi.StopAndSaveDiscoveryLog
        bidi.Shutdown
    End If
    If Not driver Is Nothing Then
        driver.CloseBrowser
        driver.Shutdown
    End If
    On Error GoTo 0

    If Len(failure) > 0 Then
        MsgBox "Recording failed: " & failure & vbCrLf & _
               "A best-effort Discovery Log save was attempted if recording began.", vbExclamation
    ElseIf stoppedFromControl Then
        MsgBox "Recording stopped from the control tab (STOP or X). Discovery Log saved.", vbInformation
    Else
        MsgBox "Safety time limit reached. Discovery Log saved.", vbInformation
    End If
    If Len(closeFailure) > 0 Then
        MsgBox "Control tab cleanup warning: " & closeFailure, vbExclamation
    End If
End Sub

' Main10 - Content-signal gate for a response-to-render gap
' ArmContentSignal requires an existing target, is consumed by the next action
' wait, and must be re-armed before each later action. Choose a stable subtree
' whose rewrite genuinely marks completion; idle-ignore patterns still allow the
' requests.
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

    bidi.AddIdleIgnoreNetworkPattern "www.scrapethissite.com/cdn-cgi/rum"
    bidi.AddIdleIgnoreNetworkPattern "www.facebook.com/tr/"

    bidi.StartDiscoveryLog

    bidi.ExecuteNavigateAndGetStatus targetUrl

    bidi.ArmContentSignal "//*[@id='table-body']"
    bidi.ExecuteClick "//section[@id='oscars']//a[@id='2015']"

    bidi.ArmContentSignal "//*[@id='table-body']"
    bidi.ExecuteClick "//section[@id='oscars']//a[@id='2014']"

    ' Saves discovery_log.txt in the same folder as the current VBA host file.
    bidi.StopAndSaveDiscoveryLog

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

    End With
End Sub

' Main11: short-lived file input (created on click, removed immediately after).
' Run as-is first: with no completion signal the wait falls back to idle only, and
' discovery_log.txt reports the DOM paths this page produced in its [SPA-ACT] burst.
' Pick one from that burst, uncomment the Arm line below, and run again.
Public Sub Main11()

    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper

    Dim targetUrl As String
    Dim filePath As String

    targetUrl = "https://hanamichi77777.github.io/WebDriver-BiDi-for-SeleniumVBA/file-dialog-probe/"

    With driver
        .StartEdge

        ' Create the file used by this example.
        .SaveStringToFile "Hello World", ".\sample.txt"

        ' input.setFiles needs the actual local file path.
        filePath = .ResolvePath(".\sample.txt", True)

        Set caps = .CreateCapabilities
        caps.EnableBiDiMode
        .OpenBrowser caps
        Set bidi = New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl

        bidi.StartDiscoveryLog
        bidi.ExecuteNavigateAndGetStatus targetUrl
        ' ------------------------------------------------------------
        ' Completion signal.
        ' On the test page, #done is initially hidden and becomes
        ' visible only after the short-lived file input fires change.
        '
        ' The waiting strategy can be replaced as appropriate when
        ' adapting this example to an actual application.
        ' ------------------------------------------------------------

'        bidi.ArmVisibilitySignal "//*[@id='done']"
        ' ------------------------------------------------------------
        ' IMPORTANT:
        ' This XPath identifies the visible trigger button.
        ' It does NOT identify input[type=file].
        '
        ' The input does not exist until this button is clicked and
        ' is removed immediately after input.click().
        ' ------------------------------------------------------------
        bidi.ExecuteSetFileSelectionViaDialog "//*[@id='short-lived']", filePath

        ' Saves discovery_log.txt in the same folder as the current VBA host file.
        bidi.StopAndSaveDiscoveryLog

        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown

    End With

End Sub

' Main12 - Unified multiple-download lifecycle and destination-folder sample
' ExecuteDownloads uses the same engine for one or many trigger XPaths.
' This sample dispatches three independent download triggers and waits until all
' accepted transactions reach complete/canceled.
'
' A single String XPath is also valid:
'   Set result = bidi.ExecuteDownloads("//*[@id='download-a']")
'
' StartDiscoveryLog preserves the full download signal chain even when DebugMode
' is disabled. filePath is browser-reported only; this sample does not assert
' filesystem existence, size, hash, or rename stability.

Public Sub Main12()

    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim result As Dictionary
    Dim downloads As Dictionary
    Dim tx As Dictionary
    Dim key As Variant
    Dim fso As New FileSystemObject

    Dim targetUrl As String
    Dim downloadFolder As String
    Dim msgText As String
    Dim msgCaption As String

    ' The published download-probe page should contain the same three triggers as
    '   download-a -> bidi-batch-A.bin
    '   download-b -> bidi-batch-B.bin
    '   download-c -> bidi-batch-C.bin
    targetUrl = "https://hanamichi77777.github.io/WebDriver-BiDi-for-SeleniumVBA/download-probe/"

    With driver
        .StartEdge

        downloadFolder = .ResolvePath(".\download-sample", False)
        If Not fso.FolderExists(downloadFolder) Then fso.CreateFolder downloadFolder

        Set caps = .CreateCapabilities
        caps.AddArguments "--start-maximized"
        caps.EnableBiDiMode

        .OpenBrowser caps

        Set bidi = New BiDiCommandWrapper
        bidi.ConnectTo .GetWebSocketUrl

        bidi.SetDownloadFolder downloadFolder
        bidi.StartDiscoveryLog

        bidi.ExecuteNavigateAndGetStatus targetUrl

        Set result = bidi.ExecuteDownloads( _
                        Array( _
                            "//*[@id='download-a']", _
                            "//*[@id='download-b']", _
                            "//*[@id='download-c']"), _
                        searchTimeoutMs:=5000, _
                        timeoutMs:=30000)

        Set downloads = result("downloads")

        Debug.Print "Batch status: " & CStr(result("status"))
        Debug.Print "Expected: " & CStr(result("expectedCount"))
        Debug.Print "Dispatched: " & CStr(result("dispatchCount"))
        Debug.Print "Started: " & CStr(result("startedCount"))
        Debug.Print "Terminal: " & CStr(result("terminalCount"))
        Debug.Print "Complete: " & CStr(result("completeCount"))
        Debug.Print "Canceled: " & CStr(result("canceledCount"))

        Debug.Print String$(72, "-")

        For Each key In downloads.keys
            Set tx = downloads(key)

            Debug.Print "Correlation key: " & CStr(key)
            Debug.Print "  Status: " & CStr(tx("status"))
            Debug.Print "  Suggested filename: " & CStr(tx("suggestedFilename"))
            Debug.Print "  Correlation: " & CStr(tx("correlationMode"))
            Debug.Print "  Browser-reported filepath: " & CStr(tx("filePath"))
        Next key

        ' The log contains DOWNLOAD-BATCH-ARM / DOWNLOAD-BATCH-START-ACCEPTED /
        ' DOWNLOAD-BATCH-END-MATCHED and any structural diagnostics.
        bidi.StopAndSaveDiscoveryLog

        msgText = "Batch status: " & CStr(result("status")) & vbCrLf & _
                  "Expected: " & CStr(result("expectedCount")) & vbCrLf & _
                  "Dispatched: " & CStr(result("dispatchCount")) & vbCrLf & _
                  "Started: " & CStr(result("startedCount")) & vbCrLf & _
                  "Terminal: " & CStr(result("terminalCount")) & vbCrLf & _
                  "Complete: " & CStr(result("completeCount")) & vbCrLf & _
                  "Canceled: " & CStr(result("canceledCount"))

        msgCaption = "BiDi Downloads - " & CStr(result("status"))
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront

        bidi.ClearDownloadBehavior

        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown

    End With

End Sub

' Main13 - New tab / new window capture through browsingContext.contextCreated
' ExecuteOpenNewContext clicks the trigger exactly once and returns the
' newly created top-level browsing context. originalOpener correlation accepts the
' owner context or another frame in the owner's tab (openerScope=owner|same-tab),
' excludes foreign contexts, and multiple candidates fail instead of being guessed.
Public Sub Main13()

    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim tabResult As Dictionary
    Dim windowResult As Dictionary

    Dim targetUrl As String
    Dim ownerContext As String
    Dim tabContext As String
    Dim windowContext As String
    Dim msgText As String
    Dim msgCaption As String

    targetUrl = "https://hanamichi77777.github.io/WebDriver-BiDi-for-SeleniumVBA/new-context-probe/"

    With driver
        .StartEdge

        Set caps = .CreateCapabilities
        caps.AddArguments "--start-maximized"
        caps.EnableBiDiMode
        .OpenBrowser caps

        Set bidi = New BiDiCommandWrapper
        bidi.ConnectTo .GetWebSocketUrl

        bidi.ExecuteNavigateAndGetStatus targetUrl
        ownerContext = bidi.GetMainContextId()

        Set tabResult = bidi.ExecuteOpenNewContext( _
            "//*[@id='open-tab']", _
            searchTimeoutMs:=5000, _
            timeoutMs:=10000)

        tabContext = CStr(tabResult("context"))

        Debug.Print "New tab context: " & tabContext
        Debug.Print "New tab kind: " & CStr(tabResult("kind"))
        Debug.Print "New tab correlation: " & CStr(tabResult("correlation"))
        Debug.Print "New tab originalOpener: " & CStr(tabResult("originalOpener"))
        Debug.Print "New tab openerScope: " & CStr(tabResult("openerScope"))

        bidi.ExecuteCloseContext tabContext

        Set windowResult = bidi.ExecuteOpenNewContext( _
            "//*[@id='open-window']", _
            searchTimeoutMs:=5000, _
            timeoutMs:=10000)

        windowContext = CStr(windowResult("context"))

        Debug.Print "New window context: " & windowContext
        Debug.Print "New window kind: " & CStr(windowResult("kind"))
        Debug.Print "New window correlation: " & CStr(windowResult("correlation"))
        Debug.Print "New window originalOpener: " & CStr(windowResult("originalOpener"))
        Debug.Print "New window title: " & bidi.ExecuteGetTitleByContextId(windowContext)

        ' Pin the new window explicitly, then return to the original owner.
        bidi.SetMainContextId windowContext
        Debug.Print "Pinned main context: " & bidi.GetMainContextId()

        bidi.SetMainContextId ownerContext
        bidi.ExecuteCloseContext windowContext

        msgText = "New Tab: " & CStr(tabResult("kind")) & vbCrLf
        msgText = msgText & "New Window: " & CStr(windowResult("kind")) & vbCrLf
        msgText = msgText & "Correlation: " & CStr(windowResult("correlation"))
        msgCaption = "BiDi New Context Capture"

        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront

        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown

    End With

End Sub
