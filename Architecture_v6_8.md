# SeleniumVBA_BiDi_Architecture v6.7 → v6.8 changes

Apply each edit to `SeleniumVBA_BiDi_Architecture_v6_7.md` and save as `SeleniumVBA_BiDi_Architecture_v6_8.md`.

---

## Edit 1 — Title

Replace:

```
# SeleniumVBA WebDriver BiDi Architecture — v6.7
```

with:

```
# SeleniumVBA WebDriver BiDi Architecture — v6.8
```

---

## Edit 2 — New release scope (insert as the first paragraph of the opening quote block, before the v6.7 paragraph)

```
> Release scope: v6.8 replaces the JSON parser used for asynchronous events. The communicator gains `ParseJsonText`, a port of the parse half of VBA-JSON as shipped in SeleniumVBA's `WebJsonConverter`, and both event parse paths use it: `ParseEventText` and the Recorder's reduced-copy parse (`TryParseReducedJson`). The result objects are the same `Scripting.Dictionary` / `Collection` structure, so dispatch is unchanged; input the new parser rejects is handed to `WebJsonConverter.ParseJson`, so errors and loss accounting are unchanged (Chapter 7, "Event JSON Parser"). Command responses and all other parses still use `WebJsonConverter`, which is not modified. `[RECORDER-QUEUE-STATS]` now shows how many lost events were dropped at the 5 MiB transport receive cap, as `asyncLossSinceStart=<n> (transportOversize=<m>)` when `n` is nonzero. Verification: on 416 raw ServiceNow events (3.26 M characters) the new parser took 0.28 of `WebJsonConverter`'s time and all results matched recursively, including dictionary key order and the `VarType` of every scalar; strings with 5,000–20,000 escapes of four kinds grew linearly and matched; on Main07 (Edge) the per-event parse time fell from 3,486 µs (v6.7, 442 parsed) to 1,286 µs (557 parsed) on the release build, all waits ended STABLE, `fallbackOriginalParse=0`, and the two lost events were reported as `transportOversize=2`. VirusTotal reported no detection for the release build. Public signatures are unchanged. All three modules carry version 6.8.
>
```

---

## Edit 3 — Chapter 2, "Parsed-Event Admission and Queue Continuation"

Replace:

```
`ParseEventText(rawMsg, parsed, dropReason)` performs the JSON parse and the malformed-event accounting.
```

with:

```
`ParseEventText(rawMsg, parsed, dropReason)` performs the JSON parse through the communicator's event parser `ParseJsonText` (since v6.8; Chapter 7, "Event JSON Parser") and the malformed-event accounting.
```

Append at the end of the section (after the paragraph that begins "Recorder-side drops are counted separately"):

```
The "no root method or id" admission check is a literal search for `"method":` and `"id":`. A message that writes whitespace between the key and the colon is valid JSON but is rejected here and counted in `DroppedParsedFilteredCount`, which is not part of `TotalAsyncEventLossCount`. Current ChromeDriver and msedgedriver send compact JSON, so this has not been observed; the check is left unchanged in v6.8. A future fix should fall back to `TryExtractTopLevelEventMethod`, which already accepts that whitespace, only when the literal search fails.
```

---

## Edit 4 — Chapter 7, new section after "Recorder Throughput and Observation Semantics (v6.2)"

Insert after the paragraph that begins "**Health line.**":

```
## Event JSON Parser (v6.8)

**Why.** After the v6.2 shortcuts, parsing was still the largest share of Recorder work. `WebJsonConverter.ParseJson` reads one character at a time with `Mid$` and appends each string character through a `Variant` buffer helper, checks for spaces several times per value, and runs three `Replace` passes over the whole message. `WebJsonConverter` belongs to SeleniumVBA, which this extension does not modify, so the faster parser lives in the communicator.

**Design.** `ParseJsonText` keeps the control flow and the result objects of `WebJsonConverter.ParseJson`: the same `Scripting.Dictionary` / `Collection` nesting, the same handling of duplicate keys, `\n` as `vbNewLine`, `\uXXXX`, single-quoted strings, the `<` case, large-number strings, and the `OptionsUseDoubleForLargeNumbers` / `OptionsAllowUnquotedKeys` settings read from `WebJsonConverter`. Only character access changed: characters are read from a byte view of the string, a string without escapes is copied with one `Mid$`, a number is copied as one slice, and the CR/LF/TAB removal runs only when those characters are present. When the parser rejects an input it calls `WebJsonConverter.ParseJson` on the same text, so a malformed event produces the same error number and the same loss accounting as before. The code carries an attribution comment for VBA-JSON (MIT, Tim Hall) and vba-json (BSD, Ryo Yokoyama); the full license texts are in `WebJsonConverter`.

**Escapes stay linear.** A string with escapes is processed segment by segment. The closing-quote position is searched again only when an escape consumed it, and the next backslash position is cached for the whole parse (`JpNextBackslash`); the read position only moves forward, so scanning is linear in the message length. An earlier development build searched again and copied the rest of the string at every escape, which made escape-dense strings quadratic; an external review found it before release.

**Measured.** Same 416 raw ServiceNow events (3.26 M characters, about 6,000 backslashes), five passes after a warm-up, in one session:

| Parser | ms per pass | µs per event | Ratio |
|---|---|---|---|
| `WebJsonConverter.ParseJson` | 1,001.7 | 2,408 | 1.00 |
| `ParseJsonText` | 278.3 | 669 | 0.28 |

All 416 results matched recursively, including key order and `VarType`. With escape-dense strings (one escape every three to four characters; `\n`, `\"`, `\\`, `あ`; 5,000, 10,000, and 20,000 escapes) the time roughly doubled with the count, and all twelve cases matched; in this extreme density `ParseJsonText` was up to about twice as slow as `WebJsonConverter` in the twelve measured cases, because each escape costs several calls. Real event streams are far sparser, which is where the overall gain comes from. On Main07 (Edge) the per-event Recorder parse time fell from 3,486 µs (v6.7) to 1,286 µs on the release build. The peak FIFO depth did not change, because it is set by events that arrive while `browsingContext.navigate` is waiting for `interactive` (Chapter 16); in two runs per parser on a development build, the backlog cleared about 6 s after recording started with the new parser and about 7 s and 10 s with `WebJsonConverter`.

**Not adopted.** VBA-FastJSON with `Scripting.Dictionary` took 1.60 of `WebJsonConverter`'s time on the same corpus. With VBA-FastDictionary it was comparable to or slightly faster than `ParseJsonText` (measured in a separate workbook), but that dictionary's class is named `Dictionary` and would replace `Scripting.Dictionary` for the whole user project, it adds two modules to the three-file distribution, and VBA-FastJSON uses pointer-based memory access. The experiment also showed that most remaining parse time is spent creating `Scripting.Dictionary` objects, so further gains would require fewer objects rather than a faster scanner.
```

---

## Edit 5 — Chapter 7, "Health line" paragraph

Replace:

```
and `asyncLossSinceStart`. The header LEGEND defines each field for an AI reader.
```

with:

```
and `asyncLossSinceStart`. Since v6.8, a nonzero loss carries its transport-cap share in parentheses, `asyncLossSinceStart=<n> (transportOversize=<m>)`: `m` events were dropped at the 5 MiB receive cap (`MaxIncomingSize`), never reached the timeline, and are absent from `[REQ-UNPAIRED]`. The value is the change of the communicator's existing `TransportOversizedEventCount` since `StartDiscoveryLog`; each such event is already counted once in the total. The header LEGEND defines each field for an AI reader.
```

---

## Edit 6 — Chapter 17, "Recorder health" bullet

Replace:

```
`fallbackOriginalParse` should be 0.
```

with:

```
`fallbackOriginalParse` should be 0. `transportOversize=` inside `asyncLossSinceStart` identifies loss caused by oversized messages, such as an image request whose URL is a multi-megabyte `data:` URI; that loss is expected and is not a Recorder failure.
```
