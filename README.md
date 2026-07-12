# WebDriver BiDi for SeleniumVBA

![WebDriver BiDi for SeleniumVBA](image/pr_image.png)

This project is a WebDriver BiDi extension for **[SeleniumVBA](https://github.com/GCuser99/SeleniumVBA)** by @GCuser99.

It is not intended to replace SeleniumVBA, Playwright, Puppeteer, or Selenium itself.  
Instead, it focuses on extending SeleniumVBA with WebDriver BiDi capabilities, especially for modern third-party SPA sites where no explicit completion signal is available.

Modern SPAs such as React, Vue.js, and enterprise web applications often update the DOM asynchronously after network responses have completed. This makes traditional VBA automation fragile, because "request completed" does not always mean "page is ready."

To address this, the tool observes network activity, Fetch/XHR activity, DOM mutations, and quiet windows to infer when the page has become stable enough for the next automation step.

Since this project assumes concurrent use of classic SeleniumVBA methods and BiDi-based observation, the BiDi API surface is intentionally kept minimal.

## Discovery Log

The Discovery Log is designed for third-party SPA sites where the automation code cannot access an internal “ready” or “completed” flag.

It records network responses, DOM mutation bursts, suppressed background noise, and stability margins such as slackMs.

This makes it easier to determine which requests should be tracked, which requests should be ignored, which resources may be safely blocked, and whether the current wait thresholds are appropriate for the target site.

The structured log can also be provided to an AI assistant for analysis. By examining the sequence and timing of network activity, DOM updates, and stability decisions, the AI can help identify likely completion signals, recommend suitable network or content signals to arm, suggest noise-filtering or resource-blocking rules, and propose adjustments to the wait strategy.

In other words, the Discovery Log is not just an execution log.
It is a diagnostic tool for discovering what should be waited for—and for helping AI devise practical synchronization solutions—when automating unknown third-party SPAs.

## Validation Benchmark

For validation, this project uses a challenging benchmark: entering text into the ServiceNow login form.

ServiceNow is frequently regarded as one of the more difficult Single Page Applications to automate due to its complex SPA behavior, Shadow DOM usage, and asynchronous UI updates.

The validation code is contained in the `Main07` procedure.

---
## [Supported OS]
* **Windows11**
## [Supported Browsers]
* **Edge / Chrome**
* *Firefox is not supported due to functional limitations.*

## Installation

Setup and import instructions are available in the **[Wiki](https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA/wiki)**.

## Scope and Limitations

* **SPA completion is inferred, not guaranteed.** The default idle consensus is based on observed network activity, Fetch/XHR counters, DOM mutations, and a quiet window. It cannot prove the target application's internal logical completion or rule out future delayed work.
* **Use explicit completion signals for important actions.** When a specific DOM rewrite or network response marks completion, arm `ArmContentSignal` and/or `ArmNetworkSignal` immediately before the action. A targeted signal is safer than relying on quietness alone.
* **The WebSocket transport is synchronous.** In the rare case that the TCP connection remains open while the peer returns no frames, the underlying WinHTTP receive call may not return control to VBA, so VBA-side timeout logic cannot intervene.
* **This project is not intended for large-scale parallel browser execution.** It is optimized for precise control and observation of one browser, or a small number of sessions, rather than dozens or hundreds of concurrent browsers.
* **Resource blocking changes page behavior.** `ExecuteEnableResourceBlocking` should be limited to resources known to be unnecessary. Blocking application code, authentication endpoints, or content APIs can break the target page.
* **Idle-ignore patterns require careful selection.** An overly broad `AddIdleIgnoreNetworkPattern` rule can exclude meaningful requests and cause an early `STABLE` result.

---

## 📂 Procedure Overview (Sample Module: `BiDi_Sample`)

### 1. Main01: Enhanced Select Box & Extension Injection
This procedure focuses on handling elements that trigger complex JavaScript state changes.
* **Dynamic Extension Injection:** Utilizes the WebDriver BiDi `ExecuteWebExtensionInstall` command to load extensions directly into the browser session from a local path. This enables session-scoped extension installation through WebDriver BiDi without permanently registering the extension in the browser profile or system registry. *(Note: Please ensure that the Google Translate Chrome extension is installed on your PC in advance.)*
* **Smart Selection:** Utilizes `ExecuteSelectValueByXPath`. This command can be configured to wait for the browser's "Idle" state immediately after selection, allowing the monitored activity to reach a stable state before proceeding.

### 2. Main02: Auto-Scrolling for Lazy Load & Dynamic SPA Synchronization
Designed for Single-Page Application (SPA) environments that utilize infinite scrolling (lazy loading) like note.com, this procedure ensures reliable interaction with elements dynamically added to the DOM.
* **Inducing Dynamic Loads via Auto-Scrolling:** By using the ExecuteLazyLoadScroll method, the script repeatedly scrolls to the bottom of the page to forcefully trigger the loading of additional content (e.g., article lists).
* **Full-Stack Idleness Monitoring:** After navigation and during scrolling, the script injects window.__vbaIdleProbe to monitor the browser's internal state.
* **Real-Time Traffic Tracking & Synchronization:** The probe continuously tracks inflightXhrCount (active XHR requests) and inflightFetchCount (active Fetch requests). The VBA code waits for these counts to return to zero and for lastMutationTs (the timestamp of the final observed DOM mutation) to stabilize, providing a practical indication that the currently observed loading activity has settled.

### 3. Main03: Performance Optimization via CDP-over-BiDi Bridge
This procedure demonstrates how resource blocking can significantly reduce page-loading overhead by controlling the network layer through a hybrid protocol approach.
* **Hybrid Protocol Bridge:** Utilizes `ExecuteEnableResourceBlocking` to filter out heavy resources like ad scripts, analytics, and tracking beacons before navigation.
* **Post-Navigation Idleness Probe:** Injects `window.__vbaIdleProbe` to ensure the environment is quiescent before entering data, maximizing execution speed and reliability.

### 4. Main04: Event-Driven URL Monitoring
Bypasses the "flaky" nature of login redirections by moving away from polling.
* **Event vs. Polling:** Uses `ExecuteIsUrlContains` to hook into the browser's internal navigation events. The script reacts to matching navigation events without relying on a fixed sleep interval, avoiding unnecessary delay from coarse polling.

### 5. Main05: Asynchronous DOM Mutation & State Validation
Focuses on synchronizing with elements that are delayed or generated via AJAX, ensuring the script does not outpace the UI updates.
* **Smart Async Interaction:** Utilizes `ExecuteClickByXPath` to interact with AJAX-driven content. The command internally monitors BiDi events to ensure the action is processed during a stable browser state.
* **Instant State Verification:** Demonstrates how to validate dynamic DOM insertions (e.g., the "Done!" label) immediately after an action, eliminating the need for manual polling loops.

### 6. Main06: Iframe Context Piercing & Hierarchical Mapping
Solves the "nested frame" problem found in legacy portals.
* **Context ID Retrieval:** Executes `GetIframeContextIdByUrl` to reliably map and target deeply nested sub-frames.
* **Direct Context Targeting:** Instead of using traditional context switching, the script retrieves a unique context ID for the specific frame and passes it directly into interaction commands like `ExecuteClickByXPath`.

### 7. Main07: SPA Idleness Detection & Shadow DOM Traversal
Targeting heavy JavaScript platforms (e.g., ServiceNow), this procedure implements a sophisticated "BiDi Probe" system.
* **Shadow DOM Interaction:** Uses `ExecuteShadowClick` to pierce shadow boundaries and interact with encapsulated web components.
* **Auto-Clicker Registration:** Utilizes `ExecuteRegisterAutoClickerByXPath` to silently and automatically handle intrusive overlays (like cookie banners) without polluting the main automation logic.

### 8. Main08: Heavy SPA Stress Test & Advanced Combobox Handling
Designed as a stress test targeting highly reactive SPAs (e.g., Google Flights) to manage complex React/Wiz-controlled comboboxes and heavy background network traffic.
* **Multi-Phase Input Synchronization:** Implements a robust, per-character input routine (`ExecuteInputValueByXPath`) that waits for field activation, detects React/Wiz double DOM replacements, and safely clears values to ensure dynamic suggestion dropdowns trigger correctly.
* **Telemetry Noise Filtering:** Utilizes `AddIdleIgnoreNetworkPattern` to continuously ignore background tracking and telemetry requests (like `/log?` or `ogs.google.com`), allowing the internal idleness probe to more accurately estimate when relevant page activity has settled.
* **Semantic ARIA Targeting:** Bypasses obfuscated class names and layout shifts by relying on W3C ARIA attributes (e.g., `@role='combobox'`, `@aria-label`) to reliably locate and interact with changing UI elements.

### 9. Main09: Discovery Log & Diagnostic Recording
A specialized tool for reverse-engineering and debugging complex automation scenarios.
* **Event Stream:** Uses `StartDiscoveryLog` to capture a raw feed of every browser event, including network requests, console logs, and DOM changes (with an option to exclude image/css noise).
* **Analysis:** Records activity using `RecordEventsForSeconds` for a specified duration and saves it via `StopAndSaveDiscoveryLog` for post-mortem analysis.

### 10. Main10: Completion-Signal Gate & Bridging the Settle-to-Render Gap
This procedure demonstrates the completion-signal gate (`ArmContentSignal`), which injects the operator's knowledge of "what marks done" into the idle consensus — closing the class of failures where an XHR settles quickly but the corresponding DOM paint lands noticeably later.
* **Arm-then-Act Pattern:** Uses `ArmContentSignal` immediately before each click to declare the exact DOM subtree (`#table-body`) whose rewrite signals completion. The subsequent SPA wait cannot conclude as STABLE on quiescence alone while the declared signal is still outstanding — the wait is gated on the actual content rewrite, not on apparent quiet.
* **One-Shot Consumption:** Each armed signal is consumed exactly once at wait end, so it never leaks into an unrelated subsequent wait. This is why the sample re-arms before every click: one signal, one wait.
* **Fail-Fast Arming:** `ArmContentSignal` raises immediately if the declared target does not exist at arm time, so a mistyped XPath surfaces as an instant error instead of a silently meaningless wait.
* **Noise Filtering as a Prerequisite:** Registers background telemetry (`cdn-cgi/rum`, Facebook tracking) with `AddIdleIgnoreNetworkPattern` first, so the consensus engine judges idleness only from traffic that actually matters.

---
### 🔗 External Links
* [Medium - Article by hanamichi77777](https://medium.com/@hanamichi77777/webdriver-bidi-for-seleniumvba-ee4687887d03)