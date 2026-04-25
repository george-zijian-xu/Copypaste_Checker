# IsItYours — Product Requirements Document

**Internal name:** Copy-paste checker
**Version:** 0.5
**Date:** 2026-04-23

---

## 1. Problem

Students submit written work they didn't write. Existing plagiarism tools (Turnitin, etc.) detect matching text but not the act of pasting — a student can paste from an unindexed source, paraphrase slightly, and pass. IsItYours detects the paste event itself, regardless of source, by monitoring the document as it's written.

---

## 2. Users

- **Student** — installs the add-in, writes their assignment, submits the .docx
- **Teacher** — receives the .docx, uploads it to the web app, reads the report
---

## 3. Products

### 3.1 VSTO Word Add-in (Windows) — Canon

The reference implementation. All other platforms are measured against this.

**Status: Working. Encryption, adjacency fusion, 50ms two-lane polling, session wrap-around, and test suite all implemented and passing (38/38 tests).**

---

#### 3.1.1 Polling & Diffing

- Two-lane timer architecture (replaced 1Hz single timer):
  - `_captureTimer` fires every **50 ms** on the UI thread — reads `doc.Content.Text`, diffs against `PrevText` using LCP + LCS, appends ticks to in-memory list
  - `_flushTimer` fires every **2 s** — builds, encrypts, and writes the XML payload to the custom XML part; skips if `PasteTraceState.Dirty == false`
- Emits one `ins` tick or one `del` tick per capture cycle (never both — the larger change wins)
- Tick ID format: `{SessionId:X3}{PollCounter:X5}` — 3 hex session ID + 5 hex poll counter (e.g. `001000A3` = session 1, poll 163). Counter increments every 50 ms capture tick.
- `TickRow.CreatedElapsedMs`: real milliseconds since session start via `Stopwatch` (monotonic); stored as `ms` attribute on each tick in XML (`tickfmt="2"`)
- Session IDs: 0–FFF (0–4095), wrap at 4096 using `(maxId + 1) % 4096`
- Del ticks are intentional — they provide a full audit trail and are a signal against fabricated "clean" traces (real writing has corrections)
- `PasteTraceState.Dirty`: set on any new tick; flush timer skips build/encrypt/write when false (avoids unnecessary writes)
- Incremental HMAC: `LastComputedChainHmac` + `LastHmacTickIndex` — each flush only recomputes HMAC over new ticks (O(new) not O(all))

#### 3.1.2 Paste Detection

- Any single insert of ≥ 20 chars (`DefaultPasteThreshold`) is flagged `paste=1`
- Threshold is a named constant (`PasteTraceEngine.DefaultPasteThreshold = 20`); tunable at construction time

#### 3.1.3 Adjacency Fusion (slow-injection defense)

Defends against AutoHotkey-style slow injection (~19 chars/sec, each tick just under threshold).

Algorithm (runs after every insert tick):
1. Walk backwards through recent `ins` ticks
2. Accumulate a run if: each tick is `ins`, not already `paste=1`, ≥ 3 chars (`FusionMinCharsEach`), contiguous in document position, and gap ≤ 1000 ms between consecutive ticks (`FusionMaxGapMs = 1000`)
3. If combined length of the run ≥ `PasteThreshold`, retroactively set all ticks in the run to `paste=1`

#### 3.1.4 Session Attribution

Every tick ID encodes its session (`sss` prefix). The `Sessions` list records each session's start UTC. The report can therefore show: "this paste happened in session 3, which started at 14:32 UTC on 2026-03-10."

#### 3.1.5 Clipboard Provenance

`ClipboardProbe` listens for `WM_CLIPBOARDUPDATE` on a hidden Win32 window (STA thread). On each clipboard change:
- Identifies the clipboard owner process by PID (`GetClipboardOwner` → `GetWindowThreadProcessId` → `Process.GetProcessById`)
- Reads `CF_UNICODETEXT` (full pasted text)
- Reads `CF_HTML` / `HTML Format` → extracts `SourceURL` header (file path or http URL)
- If Chromium: reads `Chromium internal source URL` format
- If Firefox: reads foreground window title as fallback URL
- Minimum copy length: 20 chars (same as paste threshold) — ignores trivial clipboard events
- De-duplicates within 800ms window to suppress double-fire

Captured per paste event: `{ Utc, Process, Text, SourceUrl, ChromiumUrl, FirefoxTitle }`

#### 3.1.6 Word-to-Word Provenance

When a paste is detected and the clipboard source is another `.docx`:

1. Resolve the file path from the `file://` SourceURL
2. If the source doc is already open in Word, use the live instance; otherwise open read-only, invisible
3. Look for `urn:paste-monitor` custom XML in the source doc
4. If found: hydrate a `PasteTraceState` from the source, flatten its visible text via the B+ tree piece table
5. Locate the pasted substring in the flattened text
6. Map the match back to `(tickId, offset, length)` segments — these are the origin ticks in the source doc
7. Record: `SrcDocGuid`, `SrcFile`, `SrcAuthor`, `SrcTitle`, `SrcTotalEditMin`, `WasPaste`, `Origins[]`
8. If the source doc has no trace: record `origin = "word-plain"` (Word doc, no trace)
9. If the source doc is not a `.docx` with a file URL: scan all open Word docs for a text match

#### 3.1.7 Recursive Provenance Chain — IMPLEMENTED

When a student copies from file 2, which itself copied from file 1, the teacher sees the full chain back to the canon source — not just "copied from file 2."

**How it works:**

After resolving the immediate Word-to-Word source (§3.1.6), `ResolveChain()` walks backwards through each hop's own paste evidence recursively:

1. For each contributing tick in the immediate source, check if that tick itself has `WasPaste="yes"` and a `SrcDocGuid`
2. If so, open that grandparent doc (same open-or-already-open logic as §3.1.6), hydrate its state, and record a `ProvenanceHop`
3. Recurse into the grandparent's own paste evidence
4. Stop when: a tick was typed originally (no paste evidence), a doc is unavailable, or a cycle is detected

**`ProvenanceHop` fields:** `DocGuid`, `SrcFile`, `SrcAuthor`, `SrcTitle`, `SrcTotalEditMin`, `Status`, `Origins[]`

**Status values:**
- `"resolved"` — doc found, trace hydrated, chain continues
- `"source-unavailable"` — file URL not accessible on disk
- `"no-trace"` — doc found but has no `urn:paste-monitor` XML part
- `"cycle-detected"` — DocGuid already seen in this chain walk

**Limits:** max 10 hops; cycle detection via `HashSet<string>` of DocGuids visited during the walk.

**Serialized in XML** as `<chain><hop g="..." file="..." status="..."><origins>...</origins></hop></chain>` inside each `<pid>` in the encrypted payload. Round-trips through `TryHydrate`.

**Report output:** "This text was originally typed in [file 1] by [author], then copied into [file 2], then copied into the submitted document."

#### 3.1.8 B+ Tree Piece Table

`BTreeSeq` tracks every character's origin tick. Each `Piece` records `{ TickId, OffsetInTick, Len, Visible }`. Tombstoning marks deleted characters invisible without removing them (preserves history). Used to:
- Flatten visible document text for Word-to-Word matching
- Map character ranges back to origin ticks for the report

Leaf capacity: 64 pieces. Splits automatically on overflow.

#### 3.1.9 Encryption & Integrity

**AES session key lifecycle:**
- Generated once on first document open (`RNGCryptoServiceProvider`, 32 bytes)
- RSA-2048-OAEP wrapped with server public key → stored as `ek` in XML header (only server can unwrap)
- DPAPI-wrapped (CurrentUser scope, entropy = DocGuid bytes) → stored as `lk` in XML header (add-in re-hydrates on document reopen without server)
- Never persisted in plaintext

**Payload encryption:**
- Inner XML (sessions + ticks + B+ tree snapshot + paste evidence) encrypted with AES-256-GCM
- IV: 12 random bytes per write; tag: 16 bytes
- Implemented via Windows BCrypt P/Invoke (`AesGcmBCrypt`) — required because `System.Security.Cryptography.AesGcm` is .NET 5+ only

**HMAC chain:**
- Chain root: `HMAC-SHA256(aesKey, docGuid | sessionId | sessionStartUtc)`
- Each tick: `HMAC[n] = HMAC-SHA256(aesKey, HMAC[n-1] | tickId | op | loc | len | paste)`
- Stored as `hmac` attribute on each tick in the inner XML
- Server verifies the full chain on upload — any deleted, reordered, or forged tick breaks it

**XML envelope format:**
```xml
<pasteTrace xmlns="urn:paste-monitor">
  <doc g="GUID" a="0.1.0"/>
  <header kv="1" ek="BASE64_RSA_WRAPPED" lk="BASE64_DPAPI_WRAPPED"/>
  <payload iv="BASE64_IV" ct="BASE64_CIPHERTEXT" tag="BASE64_TAG"/>
</pasteTrace>
```

**RSA key import on .NET 4.8:** Use `RSACryptoServiceProvider.FromXmlString()` with `<RSAKeyValue>` format. `ImportSubjectPublicKeyInfo` is .NET 5+ only and must not be used.

#### 3.1.10 Persistence

- Custom XML part (`urn:paste-monitor`) embedded in the `.docx` file itself
- Written on every poll cycle that produces a change, and on document open/activate/close
- Read-only documents are never written to (source docs opened for provenance lookup)
- On document reopen: `TryHydrate` decrypts using DPAPI local key, restores full state, increments session ID

#### 3.1.11 Multi-Document Support

- One `PasteTraceEngine` per document, keyed by `DocGuid`
- `ThisAddIn` maintains a `Dictionary<string, PasteTraceEngine>` — engines are created on `DocumentOpen` / `WindowActivate`
- Clipboard candidate is held in `ThisAddIn._pendingClipboard` and bound to whichever document emits the next paste tick

#### 3.1.12 Distribution

- ClickOnce self-install — student runs a URL, no IT admin required
- `.pfx` signing certificate is for ClickOnce only, not related to trace encryption

---

### 3.2 Office JS Add-in (macOS + Windows) — Tier 2

**What it does (same as VSTO):**
- 1Hz polling via `Word.run` + `setInterval`
- B+ tree piece table (TypeScript port)
- Paste threshold + adjacency fusion
- Del ticks
- Session/tick IDs (same 3+5 hex scheme)
- AES-256-GCM encryption + HMAC chain (Web Crypto API)

#### Capability Gaps vs VSTO Canon

| Capability | VSTO | Office JS | Notes |
|---|---|---|---|
| Clipboard source URL (CF_HTML) | ✅ | ❌ | No Win32 clipboard API from task pane |
| Browser process name | ✅ | ❌ | No OS process access |
| Chromium internal URL | ✅ | ❌ | Win32 only |
| Firefox window title fallback | ✅ | ❌ | Win32 only |
| Word-to-Word provenance | ✅ | ❌ | No file system access from Office JS |
| Recursive provenance chain | ✅ IMPLEMENTED | ❌ | Requires file system access |
| Persist trace in .docx | ✅ | ❌ on Mac | CustomXMLParts API blocked on Mac Word entirely |
| Offline operation | ✅ | ❌ on Mac | Mac requires network to save trace |

**Confidence tier:** Tier 2 — paste detected (velocity-based), no origin. Report clearly labels the tier.

**Persistence on Mac:** Trace stored in Supabase keyed by DocGuid. DocGuid stored in document properties (plaintext — just an ID). When teacher uploads .docx, web app fetches trace from Supabase by DocGuid.

**Status:** Not yet built.

---

### 3.3 Chrome Extension — Two Roles

#### Role A: Google Docs Monitor (build now)

- Content script injects into `docs.google.com`
- Listens to `paste` event on the editor's `contenteditable` div
- Captures: pasted text, `clipboardData.getData('text/html')` (contains SourceURL), UTC timestamp, document ID (from URL)
- MutationObserver on editor DOM for delta verification
- Adjacency fusion equivalent: merges consecutive paste events within 1s
- Encrypts payload (AES-256-GCM, same scheme), POSTs to Supabase keyed by Google Doc ID
- Session model: extension session = tab open. Session ID: same 3-hex scheme.

#### Capability Gaps vs VSTO Canon

| Capability | VSTO | Google Docs Extension | Notes |
|---|---|---|---|
| Trace travels with document | ✅ in .docx | ❌ | Stored in Supabase |
| Student can tamper with trace | ❌ (encrypted) | ❌ | Trace goes server-side directly |
| Word-to-Word provenance | ✅ | ❌ | No equivalent in Google Docs |
| Recursive provenance chain | ✅ (planned) | ❌ | No equivalent |
| Del ticks (precise) | ✅ | ⚠️ | MutationObserver less precise than LCP/LCS |
| B+ tree piece table | ✅ | ⚠️ | Harder to map offsets in Google Docs DOM |
| Offline operation | ✅ | ❌ | Requires network |

**Teacher flow:** Teacher gives the web app the Google Doc URL. Web app fetches trace from Supabase by document ID.

**Status:** Not yet built.

#### Role B: Mac Office JS Clipboard Bridge (future)

Chrome extension acts as native messaging host to supply clipboard provenance to Office JS on Mac. Partially closes the Tier 2 gap.

**Status:** Deferred.

---

### 3.4 Web App (IsItYours.com) — Report Viewer

**Teacher flow:**
1. Upload .docx (Word) or paste Google Doc URL
2. Web app decrypts trace using server RSA private key
3. Renders report:
   - Timeline of all sessions with start times
   - Highlighted document text (paste-suspect spans colored)
   - Per-paste evidence panel: timestamp, source URL, browser, Word-to-Word chain if available
   - Recursive provenance chain if available: "originally typed in [file 1] → copied to [file 2] → submitted"
   - Confidence tier label (Tier 1 = VSTO Windows, Tier 2 = Office JS Mac, Tier 3 = Google Docs)
4. If no trace found: warning banner — *"No IsItYours trace detected. The document may not have been written with the add-in installed."* — no further analysis shown

**Status:** Current site is RSID prototype. Needs full revamp.

---

## 4. Encryption & Integrity Model

### Per-document AES session key
- Generated on first document open
- RSA-2048-OAEP wrapped with server public key → `ek` in XML header
- DPAPI-wrapped (CurrentUser, entropy = DocGuid) → `lk` in XML header for local re-hydration
- Never persisted in plaintext

### Payload
- AES-256-GCM encrypted — student cannot read or modify

### HMAC chain
- `HMAC[n] = HMAC-SHA256(key, HMAC[n-1] || tickData[n])`
- Chain root seeded with `DocGuid + SessionId + SessionStartUtc`
- Server verifies chain on upload — any deleted, reordered, or forged tick breaks it

### Server
- Holds RSA private key only
- Decrypts AES key on upload, decrypts payload, verifies HMAC chain
- Rejects: broken chain, unknown key version, DocGuid already submitted with a different chain

**Key rotation:** Deferred until first customers.

---

## 5. Attack Resistance

| Attack | Mitigation |
|---|---|
| Don't install add-in | No trace = teacher warned |
| Delete custom XML part from .docx | No trace = teacher warned |
| Edit XML to remove paste flags | AES-GCM — payload is opaque to student |
| Forge a clean trace | HMAC chain — server rejects forged/modified chains |
| Replay old clean trace | DocGuid deduplication on server |
| Slow injection (AutoHotkey ~19 chars/sec) | Adjacency fusion closes this gap |
| Paste via macro / VBA | Poll delta catches it regardless of input method |
| Drag-and-drop text | Poll delta catches it |
| Suspend Word process during paste | Session gap in tick counter flags anomaly |
| Autocorrect false positive | Threshold is 20 chars; autocorrect is 1–5 chars |
| Dictation false positive | Dictation mode flag (student-toggleable, recorded in trace) |

---

## 6. Build Order

1. ~~**VSTO cleanup + encryption + adjacency fusion**~~ — **DONE**
2. ~~**Recursive provenance chain**~~ — **DONE**
3. **Web app revamp** — report viewer, server-side decryption, HMAC verification, "no trace" warning UI, chain display
4. **Office JS Mac add-in** — TypeScript port of engine, Supabase persistence for Mac
5. **Chrome extension Role A** — Google Docs full monitor
6. **Chrome extension Role B** — Mac clipboard bridge (future)

---

## 7. Infrastructure

| Service | Role |
|---|---|
| Vercel | Frontend (Next.js web app) |
| Railway | Python/FastAPI backend |
| Supabase | Trace storage for Office JS Mac + Google Docs |
| Cloudflare | Available (CDN, Workers if needed) |

---

## 8. Open Questions (not blocking)

- **RSID fallback analysis:** Deferred until after first customers
- **Key rotation policy:** Deferred until first customers
- **AppSource / Mac add-in store listing:** Deferred
- **IME / CJK input handling:** Out of scope for now (English-first)
- **Velocity histogram UI in report viewer:** Design when building web app revamp
- **Recursive chain depth limit:** Currently proposed at 10 hops — confirm when implementing
