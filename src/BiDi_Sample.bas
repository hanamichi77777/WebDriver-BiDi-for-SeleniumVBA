Attribute VB_Name = "BiDi_Sample"
Option Explicit
' WebDriver BiDi for SeleniumVBA
' https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA
' Version 2.8 / MIT License / Copyright (c) hanamichi77777
'
' Run one MainXX procedure at a time. Live-site selectors and network signals may
' change; rediscover them with the Discovery Log instead of adding fixed delays.
' Use the default action wait for ordinary SPA work and Arm* for known completion
' evidence. Shut down BiDi before the browser.
'
' MessageBoxA with hWnd=0 keeps sample result dialogs above the browser.
Public Declare PtrSafe Function MESSAGEbox Lib "user32.dll" Alias "MessageBoxA" _
                                (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
Public Const MB_OK = &H0                          ' OK button flag
Public Const MB_ForeFront = &H40000              ' Topmost flag (MB_TOPMOST)

' Main01 - Install Google Translate through WebDriver BiDi
' Current Chrome installs this unpacked extension at runtime with
' webExtension.install; classic capabilities and --load-extension are not used.
' Requires --remote-debugging-pipe, --enable-unsafe-extension-debugging and BiDi.
' extensionPath must be the version folder containing manifest.json; managed
' browser policy may still block installation.
Public Sub Main01()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver

    .StartChrome

    Dim caps As WebCapabilities: Set caps = .CreateCapabilities

    caps.AddArguments "--start-maximized"
    caps.AddArguments "--propagate-iph-for-testing"

    caps.AddArguments "--remote-debugging-pipe"
    caps.AddArguments "--enable-unsafe-extension-debugging"

    caps.EnableBiDiMode

    .OpenBrowser caps

    Dim bidi As New BiDiCommandWrapper
    bidi.ConnectTo .GetWebSocketUrl
    'bidi.DebugMode = False  'Mute developer trace in the Immediate window (default: On)

    ' Use the extension version directory that directly contains manifest.json.
    Dim extensionPath As String
    extensionPath = Environ("LOCALAPPDATA") & _
        "\Google\Chrome\User Data\Default\Extensions\" & _
        "aapbdbdomjkkjkaonfhkkikfgjllcleb\2.0.17_0"

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
    bidi.ExecuteInputValueByXPath "//input[@id='from_value']", "Tokyo"
    bidi.ExecuteInputValueByXPath "//input[@id='to_value']", "Shinjuku"
    ' These two selectors may resolve to the same button; keep only one when adapting.
    bidi.ExecuteClickByXPath "//button[starts-with(@id, 'search_button_main')]"
    bidi.ExecuteClickByXPath "//button[@id='search_button_main']"
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

    bidi.ExecuteInputValueByXPath "//input[@name='typer']", "aaa", , False
    bidi.ExecuteClickByXPath "//input[@id='red']", , False
    
    bidi.ExecuteClickByXPath "//input[@value='Add Label']", , , 1000

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
   bidi.ExecuteClickByXPath "//input[@id='la_imp']", , , , , conID
   
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

    bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
    
    bidi.ExecuteNavigateAndGetStatus targetUrl
        
    bidi.ArmNetworkSignal "metadata/application"
    bidi.ExecuteShadowClick "#utility-sign-in button"
            
    ' Use a non-sensitive test value when adapting this sample.
    bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa"
    
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
           
    End With
End Sub

' Main08 - Google Flights reactive-SPA stress sample
' Scenario: One way, Sapporo to Paris, select a date, then Search.
' The live-site selectors and GetCalendarPicker/GetShoppingResults signals may
' change; use a fresh Discovery Log. --lang=en supports the English ARIA labels.
' Blocking removes non-essential traffic; idle-ignore retains required telemetry.
' Arm* completion signals are one-shot.

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
        Dim blockList As Variant
        blockList = Array( _
            "*googletagmanager*", "*doubleclick*", "*googlesyndication*", _
            "*google-analytics*", _
            "*/collect*", "*/beacon*", "*pagead*")
        bidi.ExecuteEnableResourceBlocking blockList
        bidi.AddIdleIgnoreNetworkPattern "/log?"
        bidi.AddIdleIgnoreNetworkPattern "*generate_204*"
        bidi.AddIdleIgnoreNetworkPattern "GetAsyncData"
        
        bidi.StartDiscoveryLog
        
        Dim url As String: url = "https://www.google.com/travel/flights"
        bidi.ExecuteNavigateAndGetStatus url
        
        ' The specific trigger XPath is retained as a reference; the executable call uses
        ' the simpler first-combobox XPath.
        Dim ticketTypeTrigger As String
        ticketTypeTrigger = "(//div[@role='combobox' and @aria-haspopup='listbox'])[1]"
        
        bidi.ExecuteSelectValueByXPath "(//div[@role='combobox'])[1]", "One way"
                
        Dim depXPath As String
        depXPath = "(//input[contains(@aria-label, 'Where from')])[last()]"
        
        Dim depSuggestXPath As String
        depSuggestXPath = _
            "//*[@role='listbox' and not(@aria-hidden='true')]" & _
            "//li[@role='option' and contains(@aria-label, 'Sapporo')][1]"
        bidi.ExecuteInputValueByXPath depXPath, "Sapporo"
        
        bidi.ArmNetworkSignal "rpcids=tDoGIe"
        bidi.ExecuteClickByXPath depSuggestXPath, minStableMs:=1000
        
        
        Dim destXPath As String
        destXPath = "(//input[contains(@aria-label, 'Where to')])[last()]"
        
        Dim destSuggestXPath As String
        destSuggestXPath = _
            "//*[@role='listbox' and not(@aria-hidden='true')]" & _
            "//li[@role='option' and contains(@aria-label, 'Paris')][1]"
        bidi.ExecuteInputValueByXPath destXPath, "Paris"
        
        bidi.ArmNetworkSignal "rpcids=BVAT3"
        bidi.ExecuteClickByXPath destSuggestXPath, minStableMs:=1000
        
        
        bidi.ArmNetworkSignal "GetCalendarPicker"
        bidi.ArmVisibilitySignal "//div[@data-gs]"
        bidi.ExecuteClickByXPath "//input[@aria-label='Departure']"
        
        bidi.ExecuteClickByXPath _
            "(//div[@role='gridcell' and @aria-hidden='false'])[8]//div[@role='button']"
        bidi.ExecuteClickByXPath "//button[contains(., 'Done')]"
        
        
        Dim searchXPath As String
        searchXPath = "//button[@aria-label='Search']"


        bidi.ArmNetworkSignal "GetShoppingResults"
        bidi.ExecuteClickByXPath searchXPath
        
        Dim logPath As String
        logPath = .ResolvePath(".\") & "\discovery_log.txt"
        bidi.StopAndSaveDiscoveryLog logPath
        
        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown
        
        MsgBox "Google Flights Test Completed"
        
    End With
End Sub

' Main09 - Manual Discovery Log recorder
' Record one meaningful manual action per run. Change the URL and recording
' window; a narrow window makes causal analysis easier. excludeImagesAndCss
' filters log entries only and does not block browser resources.
Sub Main09()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    .StartEdge
    
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    caps.EnableBiDiMode
    .OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper
    bidi.ConnectTo .GetWebSocketUrl

    Dim url As String: url = "https://note.com/"
    bidi.ExecuteNavigateAndGetStatus url
    
    Const RECORDING_SECONDS As Long = 20
    Dim msgText As String, msgCaption As String
    msgText = "Please prepare the browser for recording." & vbCrLf & vbCrLf & _
              "Click [OK] to start recording." & vbCrLf & _
              "Duration: " & RECORDING_SECONDS & " seconds." & vbCrLf & _
              "Please manually interact with the page immediately after clicking OK."
    msgCaption = "Ready to Record"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    
    ' Start immediately before the manual action; True filters image/CSS log entries.
    bidi.StartDiscoveryLog excludeImagesAndCss:=True
    bidi.RecordEventsForSeconds RECORDING_SECONDS
    
    
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    
    MsgBox "Discovery Log Saved!" & vbCrLf & logPath
    
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
    
End With
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
    bidi.ExecuteClickByXPath "//section[@id='oscars']//a[@id='2015']"
    
    bidi.ArmContentSignal "//*[@id='table-body']"
    bidi.ExecuteClickByXPath "//section[@id='oscars']//a[@id='2014']"
    
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath

    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown

    End With
End Sub
