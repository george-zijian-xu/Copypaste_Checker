using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace IsItYoursWordAddIn
{
    public sealed class ProvenanceHop
    {
        public string DocGuid;
        public string SrcFile;
        public string SrcAuthor;
        public string SrcTitle;
        public int? SrcTotalEditMin;
        // "resolved" | "unverified" | "source-unavailable" | "no-trace" | "cycle-detected"
        public string ChainStatus;
        public List<(string t, int off, int n)> Origins;
    }

    // Recursive provenance tree. A ProvenanceSource either bottoms out as a leaf
    // (Kind = "browser" / "doc-local" / "unknown" / "word-plain") or, when Kind = "word",
    // carries a Segments list that decomposes Text into ranges with their own sources.
    //
    // A "word" node with non-null Segments means we successfully walked into the source
    // doc's live state (via DocStateRegistry) and mapped the contributing tick subranges.
    // A "word" node with null Segments and a non-null MappingFailure means we knew the
    // source doc GUID but couldn't resolve its pieces (source doc closed, cycle, etc.).
    public sealed class ProvenanceSource
    {
        public string Kind;            // "browser" | "word" | "doc-local" | "unknown" | "word-plain"
        public string DocGuid;
        public string SrcFile;
        public string Url;
        public string Process;
        public string Text;            // the bytes contributed at this node
        public bool Live;            // true when we resolved against a live in-memory state
        public string MappingFailure;  // e.g. "source-doc-unavailable", "cycle-detected", "no-origin-mapping"

        // Leaf/hop locator — the tick and sub-range in the source doc that produced Text.
        public string TickId;
        public int OffsetInTick;
        public int Len;

        // Non-leaf: breakdown of Text into child sources. Starts are within this node's Text.
        public List<ProvenanceSegment> Segments;
    }

    public sealed class ProvenanceSegment
    {
        public int Start;   // offset within the parent node's Text
        public int Len;
        public string Text;    // convenience duplicate of Source?.Text
        public ProvenanceSource Source;
    }

    public sealed class PasteEvidence
    {
        public string Origin;          // "browser" | "word" | "word-plain" | "unknown"
        public DateTime ClipboardUtc;
        public string ClipboardProcess;
        public string Url;             // CF_HTML SourceURL (file:/// or http(s)://)
        public string ChromiumUrl;
        public string FirefoxTitle;
        public string FullText;
        public string Sha256;

        public string SrcDocGuid;
        public string SrcFile;
        public string SrcAuthor;
        public string SrcTitle;
        public int? SrcTotalEditMin;

        // MappingFailure: null = success | "no-substring-hit" | "empty-needle" |
        //                 "source-doc-unavailable" | "no-trace"
        public string MappingFailure;

        // Exact mapped source segments (set when exact substring match succeeded)
        public List<(string t, int off, int n)> Origins;
        // Fallback candidate paste ticks from source doc (set when exact match failed but source has a trace)
        public List<(string t, int off, int n)> OriginCandidates;

        public string OriginalPidXml;

        // Recursive provenance chain: hop[0] = immediate source, hop[n] = canon origin.
        // null = not yet resolved; empty list = typed originally at the immediate source.
        public List<ProvenanceHop> Chain;

        // Recursive provenance tree — nested source decomposition. When non-null this
        // is the authoritative, nested view that the analyzer should render. Chain is
        // retained alongside it for backward compatibility with the flat analyzer UI.
        public ProvenanceSource ProvenanceTree;
    }

    public static class Provenance
    {
        // The clipboard candidate is written to every live doc's state. The user may
        // copy while one doc is focused and then activate a different doc to paste; if
        // the candidate only lived on the source state the destination wouldn't see it
        // when ApplyInsert fires and AttachForPasteTick runs. Broadcasting keeps the
        // last-copy-wins semantics while making the candidate reachable everywhere.
        public static void SetCandidate(PasteTraceState s, ClipboardCandidate c)
        {
            if (c == null) return;
            foreach (var st in DocStateRegistry.AllLiveStates())
                st._clipCandidate = c;
            if (s != null) s._clipCandidate = c;     // covers states not yet registered
        }

        public static ClipboardCandidate ConsumeCandidate(PasteTraceState s)
        {
            var c = s?._clipCandidate;
            ClearAllCandidates();
            return c;
        }

        private static void ClearAllCandidates()
        {
            foreach (var st in DocStateRegistry.AllLiveStates())
                st._clipCandidate = null;
        }

        public static void AttachForPasteTick(Word.Application app, PasteTraceState s, TickRow tick)
        {
            var cand = s._clipCandidate;
            if (cand == null) return;

            string Norm(string x) => (x ?? "")
                .Replace("\r\n", "\r")
                .Replace("\n", "\r")
                .TrimEnd();

            var tickText = Norm(tick.Text);
            var clipText = Norm(cand.Text);

            bool nonEmpty = !string.IsNullOrEmpty(tickText) && !string.IsNullOrEmpty(clipText);

            bool subsetOk = nonEmpty &&
                            (clipText.IndexOf(tickText, StringComparison.Ordinal) >= 0 ||
                             tickText.IndexOf(clipText, StringComparison.Ordinal) >= 0);

            var proc = (cand.Process ?? "").ToLowerInvariant();
            var url = cand.SourceUrl ?? "";
            bool looksBrowser = proc.Contains("chrome") || proc.Contains("edge") || proc.Contains("opera")
                                || proc.Contains("brave") || proc.Contains("chromium") || proc.StartsWith("firefox");
            bool looksWord = proc.Contains("winword") ||
                             (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
                              url.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));

            var origin = looksBrowser ? "browser" : looksWord ? "word" : "unknown";

            // Browser and Word pastes always proceed — browser needs no text-match gate
            // (clipboard text may differ from tick text due to Word reformatting), and
            // Word pastes are verified by TryWordMapping directly.
            // Only truly unknown sources are gated by subsetOk.
            bool mayProceed = looksBrowser || looksWord || subsetOk;

            if (!mayProceed)
            {
                ClearAllCandidates();
                s._pasteEvidence[tick.TickId] = new PasteEvidence
                {
                    Origin = "unknown",
                    ClipboardUtc = cand.Utc,
                    ClipboardProcess = cand.Process,
                    Url = url,
                    ChromiumUrl = cand.ChromiumUrl ?? "",
                    FirefoxTitle = cand.FirefoxTitle ?? "",
                    FullText = tick.Text
                };
                return;
            }

            var ev = new PasteEvidence
            {
                Origin = origin,
                ClipboardUtc = cand.Utc,
                ClipboardProcess = cand.Process,
                Url = url,
                ChromiumUrl = cand.ChromiumUrl ?? "",
                FirefoxTitle = cand.FirefoxTitle ?? "",
                FullText = tick.Text,
                Sha256 = null,
                Origins = null
            };

            if (origin == "word")
                TryWordMapping(app, s, cand, ref ev);

            s._pasteEvidence[tick.TickId] = ev;
            ClearAllCandidates();
        }

        // Check open docs first; only if SourceURL is a .docx file URL, open it read-only.
        private static void TryWordMapping(Word.Application app, PasteTraceState s, ClipboardCandidate cand, ref PasteEvidence ev)
        {
            var url = cand.SourceUrl ?? "";
            if (!(url.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
                  url.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
            {
                Word.Document openSrc;
                PasteTraceState openState;
                List<(string t, int off, int n)> openOrigins;

                if (TryFindOpenWordSourceByText(app, (!string.IsNullOrEmpty(ev.FullText) ? ev.FullText : cand.Text), out openSrc, out openState, out openOrigins))
                {
                    ev.Origin = "word";
                    ev.SrcDocGuid = (openState != null && !string.IsNullOrEmpty(openState.DocGuid)) ? openState.DocGuid : null;
                    ev.SrcFile = TryFileUrl(openSrc);
                    ev.SrcAuthor = SafeGetBuiltinProperty(openSrc, "Author");
                    ev.SrcTitle = SafeGetBuiltinProperty(openSrc, "Title");
                    ev.SrcTotalEditMin = openState != null ? ComputeTotalEditMinutesApprox(openState) : null;
                    ev.Origins = openOrigins;

                    if (openOrigins != null && openOrigins.Count > 0 && openState != null)
                    {
                        // Full resolution: state registered, origins mapped. Build both
                        // the authoritative recursive tree and the flat legacy chain.
                        try
                        {
                            var seed = new HashSet<string>(StringComparer.Ordinal);
                            if (!string.IsNullOrEmpty(s.DocGuid)) seed.Add(s.DocGuid);
                            ev.ProvenanceTree = BuildProvenanceTree(app, openState, openOrigins, ev.SrcFile, seed);
                        }
                        catch { }

                        try
                        {
                            var visited = new HashSet<string>(StringComparer.Ordinal);
                            if (!string.IsNullOrEmpty(s.DocGuid)) visited.Add(s.DocGuid);
                            if (!string.IsNullOrEmpty(openState.DocGuid)) visited.Add(openState.DocGuid);
                            ev.Chain = ResolveChain(app, openState, openOrigins, visited, depth: 0);
                        }
                        catch { }
                    }
                    else
                    {
                        // Source doc identified (needle found in its live text) but we
                        // couldn't map the paste to specific ticks. Happens when the
                        // source doc has no registered trace state — e.g. an unsaved
                        // new doc that hasn't been registered with DocStateRegistry, or
                        // a state that exists but hasn't polled recently enough to cover
                        // the matching text. Emit a leaf ProvenanceTree noting the
                        // identified source so downstream tools can still show the hop.
                        ev.MappingFailure = (openState == null) ? "source-doc-no-trace" : "no-state-mapping";
                        ev.ProvenanceTree = new ProvenanceSource
                        {
                            Kind = "word",
                            DocGuid = ev.SrcDocGuid,
                            SrcFile = ev.SrcFile,
                            Text = ev.FullText,
                            Len = (ev.FullText ?? "").Length,
                            Live = (openState != null),
                            MappingFailure = ev.MappingFailure
                        };
                    }
                    return;
                }

                ev.MappingFailure = "source-doc-unavailable";
                return;
            }

            // 1) Resolve Windows path from file:// URL
            string srcPath;
            try
            {
                var uri = new Uri(url, UriKind.Absolute);
                srcPath = uri.LocalPath;
            }
            catch
            {
                ev.MappingFailure = "source-doc-unavailable";
                return;
            }

            Word.Document srcDoc = null;
            bool openedByUs = false;
            Word.Document priorActive = null;

            try
            {
                // 2) Prefer already-open docs
                foreach (Word.Document d in app.Documents)
                {
                    string full = null;
                    try { full = d.FullName; } catch { full = null; }
                    if (!string.IsNullOrEmpty(full) && string.Equals(full, srcPath, StringComparison.OrdinalIgnoreCase))
                    { srcDoc = d; break; }
                }

                // 3) If not open, open read-only and invisible
                if (srcDoc == null)
                {
                    try { priorActive = app.ActiveDocument; } catch { }
                    srcDoc = app.Documents.Open(FileName: srcPath, ReadOnly: true,
                                                AddToRecentFiles: false, Visible: false);
                    openedByUs = true;
                }

                // 4) Look for our custom XML part in the source
                var parts = srcDoc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1)
                {
                    ev.Origin = "word-plain";
                    ev.MappingFailure = "no-trace";
                    return;
                }

                // 5) Hydrate a temp state from the source doc's XML
                var srcState = new PasteTraceState { DocGuid = "", AppVersion = PasteTraceEngine.AppVersion, PasteThreshold = PasteTraceEngine.DefaultPasteThreshold };
                PasteTraceXml.TryHydrate(srcDoc, srcState);

                // 6) Flatten visible text from the source state
                var flat = BuildFlatVisibleText(srcState, out var pieceIndex);

                // 7) Locate the pasted text in the source
                string Norm(string x) => (x ?? "").Replace("\r\n", "\r").Replace("\n", "\r").TrimEnd();
                var needle = Norm(ev.FullText).Length >= Norm(cand.Text).Length ? Norm(ev.FullText) : Norm(cand.Text);

                if (string.IsNullOrEmpty(needle))
                {
                    ev.MappingFailure = "empty-needle";
                }
                else
                {
                    int hit = flat.IndexOf(needle, StringComparison.Ordinal);
                    if (hit < 0)
                    {
                        // Exact match failed. Collect all paste ticks as candidates so
                        // ResolveChain can still walk the chain; exact Origins stays null.
                        ev.MappingFailure = "no-substring-hit";
                        var candidates = new List<(string t, int off, int n)>();
                        foreach (var t in srcState.Ticks)
                            if (t.Op == "ins" && t.Paste == 1)
                                candidates.Add((t.TickId, 0, t.Len));
                        if (candidates.Count > 0)
                            ev.OriginCandidates = candidates;
                    }
                    else
                    {
                        // 8) Map the matched range back to (tickId, off, n) segments.
                        // The origin offset reported to the outer layer is the offset
                        // WITHIN the source tick's text, not within the piece. A tick may
                        // span multiple pieces (if a later edit split it) so we add the
                        // piece's own OffsetInTick to its internal overlap start.
                        int end = hit + needle.Length;
                        var origins = new List<(string t, int off, int n)>();
                        foreach (var p in pieceIndex)
                        {
                            int ovStart = Math.Max(p.Start, hit);
                            int ovEnd = Math.Min(p.Start + p.Len, end);
                            int ovLen = ovEnd - ovStart;
                            if (ovLen > 0) origins.Add((p.TickId, p.OffsetInTick + (ovStart - p.Start), ovLen));
                            if (ovEnd >= end) break;
                        }
                        if (origins.Count > 0) ev.Origins = origins;
                        else ev.MappingFailure = "no-substring-hit";
                    }
                }

                // 9) Populate source doc metadata
                ev.SrcDocGuid = string.IsNullOrEmpty(srcState.DocGuid) ? null : srcState.DocGuid;
                ev.SrcFile = url;
                ev.SrcAuthor = SafeGetBuiltinProperty(srcDoc, "Author");
                ev.SrcTitle = SafeGetBuiltinProperty(srcDoc, "Title");
                ev.SrcTotalEditMin = ComputeTotalEditMinutesApprox(srcState);

                // Build nested pid XML from exact origins (not candidates)
                try
                {
                    var contrib = ev.Origins;
                    if (contrib != null && contrib.Count > 0)
                    {
                        var tickMap = srcState.Ticks.ToDictionary(t => t.TickId, t => t, StringComparer.Ordinal);
                        var sbNested = new StringBuilder();
                        foreach (var tid in contrib.Select(o => o.t).Where(tid => !string.IsNullOrEmpty(tid)).Distinct(StringComparer.Ordinal))
                        {
                            if (!srcState._pasteEvidence.TryGetValue(tid, out var pe)) continue;
                            int len = tickMap.TryGetValue(tid, out var tr) ? tr.Len : (pe.FullText?.Length ?? 0);
                            if (len <= 0) continue;
                            sbNested.Append(
                                "<pid t=\"" + SecurityElement.Escape(tid) +
                                "\" len=\"" + len +
                                "\" origin=\"" + SecurityElement.Escape(pe.Origin ?? "unknown") + "\">" + Environment.NewLine +
                                "  <evidence>" + Environment.NewLine +
                                "    <clipboard ts=\"" + pe.ClipboardUtc.ToString("o") +
                                "\" process=\"" + SecurityElement.Escape(pe.ClipboardProcess ?? "") + "\"/>" + Environment.NewLine +
                                (string.IsNullOrEmpty(pe.Url) ? "" :
                                    "    <url>" + SecurityElement.Escape(pe.Url) + "</url>" + Environment.NewLine) +
                                "    <text>" + SecurityElement.Escape(pe.FullText ?? "") + "</text>" + Environment.NewLine +
                                "  </evidence>" + Environment.NewLine +
                                "</pid>" + Environment.NewLine);
                        }
                        if (sbNested.Length > 0)
                            ev.OriginalPidXml = sbNested.ToString().TrimEnd();
                    }
                }
                catch { }

                // 10) Recursively resolve the provenance chain.
                // Use exact Origins when available; fall back to OriginCandidates.
                var chainOrigins = ev.Origins ?? ev.OriginCandidates;
                if (chainOrigins != null && chainOrigins.Count > 0)
                {
                    // Build the nested tree. Only exact Origins produce a tree — candidate
                    // ticks (from no-substring-hit fallbacks) don't have verified byte-for-byte
                    // correspondence with the source, so stripe-slicing them would fabricate
                    // mappings. The flat Chain still honours candidates for UI continuity.
                    if (ev.Origins != null && ev.Origins.Count > 0)
                    {
                        try
                        {
                            var seed = new HashSet<string>(StringComparer.Ordinal);
                            if (!string.IsNullOrEmpty(s.DocGuid)) seed.Add(s.DocGuid);
                            ev.ProvenanceTree = BuildProvenanceTree(app, srcState, ev.Origins, ev.SrcFile, seed);
                        }
                        catch { }
                    }

                    try
                    {
                        var visited = new HashSet<string>(StringComparer.Ordinal);
                        if (!string.IsNullOrEmpty(s.DocGuid)) visited.Add(s.DocGuid);
                        if (!string.IsNullOrEmpty(srcState.DocGuid)) visited.Add(srcState.DocGuid);
                        ev.Chain = ResolveChain(app, srcState, chainOrigins, visited, depth: 0);
                    }
                    catch { }
                }
            }
            catch
            {
                ev.Origin = "word-plain";
                ev.SrcDocGuid = null;
                ev.SrcFile = null;
                ev.MappingFailure = "source-doc-unavailable";
            }
            finally
            {
                if (openedByUs && srcDoc != null)
                    try { srcDoc.Close(SaveChanges: false); } catch { }
                if (openedByUs && priorActive != null)
                    try { priorActive.Activate(); } catch { }
            }
        }

        private const int ChainMaxDepth = 10;

        private static List<ProvenanceHop> ResolveChain(
            Word.Application app,
            PasteTraceState srcState,
            List<(string t, int off, int n)> origins,
            HashSet<string> visited,
            int depth)
        {
            var chain = new List<ProvenanceHop>();
            if (origins == null || origins.Count == 0 || depth >= ChainMaxDepth) return chain;

            foreach (var tickId in origins.Select(o => o.t).Where(id => !string.IsNullOrEmpty(id)).Distinct(StringComparer.Ordinal))
            {
                PasteEvidence srcEv;
                if (!srcState._pasteEvidence.TryGetValue(tickId, out srcEv)) continue;

                // A tick with no paste evidence was typed originally — not a hop.
                // A tick with evidence (even failed mapping) is a hop; let the human judge.
                bool hasExactOrigins = srcEv.Origins != null && srcEv.Origins.Count > 0;
                bool hasCandidates = srcEv.OriginCandidates != null && srcEv.OriginCandidates.Count > 0;
                bool mappingFailed = srcEv.MappingFailure != null;
                bool isHop = hasExactOrigins || hasCandidates || mappingFailed || srcEv.SrcDocGuid != null;
                if (!isHop) continue;

                // No confirmed source doc — emit an unverified hop and stop recursing.
                if (string.IsNullOrEmpty(srcEv.SrcDocGuid))
                {
                    chain.Add(new ProvenanceHop
                    {
                        DocGuid = null,
                        SrcFile = srcEv.SrcFile,
                        SrcAuthor = srcEv.SrcAuthor,
                        SrcTitle = srcEv.SrcTitle,
                        SrcTotalEditMin = srcEv.SrcTotalEditMin,
                        Origins = srcEv.Origins ?? srcEv.OriginCandidates,
                        ChainStatus = "unverified"
                    });
                    continue;
                }

                if (visited.Contains(srcEv.SrcDocGuid))
                {
                    chain.Add(new ProvenanceHop { DocGuid = srcEv.SrcDocGuid, SrcFile = srcEv.SrcFile, ChainStatus = "cycle-detected" });
                    continue;
                }

                visited.Add(srcEv.SrcDocGuid);

                var hop = new ProvenanceHop
                {
                    DocGuid = srcEv.SrcDocGuid,
                    SrcFile = srcEv.SrcFile,
                    SrcAuthor = srcEv.SrcAuthor,
                    SrcTitle = srcEv.SrcTitle,
                    SrcTotalEditMin = srcEv.SrcTotalEditMin,
                    Origins = srcEv.Origins ?? srcEv.OriginCandidates
                };

                var grandState = TryLoadState(app, srcEv.SrcFile);
                if (grandState == null) { hop.ChainStatus = "source-unavailable"; chain.Add(hop); continue; }
                if (string.IsNullOrEmpty(grandState.DocGuid)) { hop.ChainStatus = "no-trace"; chain.Add(hop); continue; }

                hop.ChainStatus = hasExactOrigins ? "resolved" : "unverified";
                chain.Add(hop);
                var nextOrigins = srcEv.Origins ?? srcEv.OriginCandidates;
                chain.AddRange(ResolveChain(app, grandState, nextOrigins, visited, depth + 1));
            }

            return chain;
        }

        private static PasteTraceState TryLoadState(Word.Application app, string fileUrl)
        {
            if (string.IsNullOrEmpty(fileUrl)) return null;
            if (!(fileUrl.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
                  fileUrl.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))) return null;

            string path;
            try { path = new Uri(fileUrl, UriKind.Absolute).LocalPath; } catch { return null; }

            // Fast path: live state is authoritative. If the source doc is open in
            // this Word instance, its registered PasteTraceState has fresher tick,
            // session, and evidence data than the CustomXMLParts snapshot.
            var live = DocStateRegistry.GetByKey(path);
            if (live != null) return live;

            Word.Document doc = null;
            bool openedByUs = false;
            Word.Document priorActive = null;
            try
            {
                foreach (Word.Document d in app.Documents)
                {
                    string full = null;
                    try { full = d.FullName; } catch { }
                    if (!string.IsNullOrEmpty(full) && string.Equals(full, path, StringComparison.OrdinalIgnoreCase))
                    { doc = d; break; }
                }
                if (doc == null)
                {
                    try { priorActive = app.ActiveDocument; } catch { }
                    doc = app.Documents.Open(FileName: path, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                    openedByUs = true;
                }

                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1) return null;

                var st = new PasteTraceState { AppVersion = PasteTraceEngine.AppVersion, PasteThreshold = PasteTraceEngine.DefaultPasteThreshold };
                PasteTraceXml.TryHydrate(doc, st);
                return st;
            }
            catch { return null; }
            finally
            {
                if (openedByUs && doc != null) try { doc.Close(SaveChanges: false); } catch { }
                if (openedByUs && priorActive != null) try { priorActive.Activate(); } catch { }
            }
        }

        // Scan open docs for a CustomXMLPart whose pasteTrace g="…" attribute matches
        // the given GUID, hydrate it, and return the state. Used as a last-resort
        // fallback during recursive resolution when DocStateRegistry doesn't have the
        // state — typically because the host's per-doc engine-creation path didn't
        // call DocStateRegistry.Register. Because ThisAddIn.ForceFlush writes to the
        // in-memory CustomXMLParts of unsaved docs, we can still recover per-doc
        // trace data as long as at least one flush cycle has run for that doc.
        //
        // Fast pre-filter: a plain substring scan of the raw XML for the GUID marker
        // avoids decrypting the payload of every non-matching doc. A full hydrate
        // only runs once we've confirmed the GUID lives in this doc's part.
        private static PasteTraceState TryFindLiveDocStateByGuid(Word.Application app, string guid)
        {
            if (app == null || string.IsNullOrEmpty(guid)) return null;
            foreach (Word.Document d in app.Documents)
            {
                try
                {
                    var parts = d.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                    if (parts == null || parts.Count < 1) continue;

                    string rawXml = null;
                    try { rawXml = parts[1].XML; } catch { rawXml = null; }
                    if (rawXml == null || rawXml.IndexOf("g=\"" + guid + "\"", StringComparison.Ordinal) < 0)
                        continue;

                    var tmp = new PasteTraceState
                    {
                        DocGuid = "",
                        AppVersion = PasteTraceEngine.AppVersion,
                        PasteThreshold = PasteTraceEngine.DefaultPasteThreshold
                    };
                    if (PasteTraceXml.TryHydrate(d, tmp) &&
                        string.Equals(tmp.DocGuid, guid, StringComparison.Ordinal))
                    {
                        return tmp;
                    }
                }
                catch { }
            }
            return null;
        }

        // ----- Recursive provenance tree builder ------------------------------------
        //
        // Given a paste's (srcState, Origins) pair where Origins covers the pasted
        // text as a sequence of (tickId, offsetInTick, length) contributions from the
        // immediate source doc, build a ProvenanceSource rooted at that source doc
        // whose Segments list decomposes the paste into per-origin subtrees. Each
        // subtree recurses into the grand-source if the contributing tick was itself
        // a paste (Origin = "word") and we can find the grand-source's live state.
        //
        // Why this is structured around sub-ranges, not whole ticks: a single tick
        // in an intermediate doc may itself be a paste whose text was assembled from
        // several upstream ticks. When DOC2 pastes a 12-char snippet that originated
        // in DOC3, and the user later copies 6 chars of that snippet into DOC1, the
        // DOC1 paste's origin in DOC2 is {tick X, off Y, len 6} — and we need to
        // find which DOC3 ticks contributed to *those six bytes*, not the whole
        // original DOC2→DOC3 paste. SliceOrigins performs that sub-range slicing.

        private static ProvenanceSource BuildProvenanceTree(
            Word.Application app,
            PasteTraceState immediateSrc,
            List<(string t, int off, int n)> origins,
            string immediateSrcFile,
            HashSet<string> visitedSeed)
        {
            if (immediateSrc == null || origins == null || origins.Count == 0) return null;

            var visited = new HashSet<string>(visitedSeed ?? new HashSet<string>(StringComparer.Ordinal),
                                              StringComparer.Ordinal);
            if (!string.IsNullOrEmpty(immediateSrc.DocGuid))
                visited.Add(immediateSrc.DocGuid);

            var segments = new List<ProvenanceSegment>();
            var fullTextSb = new StringBuilder();
            int cursor = 0;
            foreach (var o in origins)
            {
                if (o.n <= 0) continue;
                var child = ResolveTickSubRange(app, immediateSrc, o.t, o.off, o.n, visited, depth: 0);
                string segText = child?.Text ?? "";
                segments.Add(new ProvenanceSegment
                {
                    Start = cursor,
                    Len = o.n,
                    Text = segText,
                    Source = child
                });
                fullTextSb.Append(segText);
                cursor += o.n;
            }

            return new ProvenanceSource
            {
                Kind = "word",
                DocGuid = immediateSrc.DocGuid,
                SrcFile = immediateSrcFile,
                Text = fullTextSb.ToString(),
                Live = true,
                Segments = segments
            };
        }

        // Resolve where the sub-range [takeOff, takeOff + takeLen) inside srcState's
        // tick `tickId` ultimately came from. Terminates at a leaf (browser / doc-local /
        // unknown / unresolvable) or recurses further via DocStateRegistry.
        private static ProvenanceSource ResolveTickSubRange(
            Word.Application app,
            PasteTraceState srcState,
            string tickId,
            int takeOff,
            int takeLen,
            HashSet<string> visitedGuids,
            int depth)
        {
            if (srcState == null || string.IsNullOrEmpty(tickId) || takeLen <= 0 || depth >= ChainMaxDepth)
                return null;

            TickRow tr = null;
            foreach (var t in srcState.Ticks)
                if (string.Equals(t.TickId, tickId, StringComparison.Ordinal)) { tr = t; break; }

            string subText = null;
            if (tr != null && tr.Text != null)
            {
                int off = Math.Max(0, Math.Min(takeOff, tr.Text.Length));
                int len = Math.Max(0, Math.Min(takeLen, tr.Text.Length - off));
                if (len > 0) subText = tr.Text.Substring(off, len);
            }

            bool isPasteTick = tr != null && tr.Paste == 1;
            PasteEvidence pe = null;
            if (isPasteTick) srcState._pasteEvidence.TryGetValue(tickId, out pe);

            // Typed-in-place (or a paste we failed to record) — this is the canon origin.
            if (!isPasteTick || pe == null)
            {
                return new ProvenanceSource
                {
                    Kind = "doc-local",
                    DocGuid = srcState.DocGuid,
                    Text = subText,
                    TickId = tickId,
                    OffsetInTick = takeOff,
                    Len = takeLen
                };
            }

            // Browser leaf — no further recursion possible.
            if (string.Equals(pe.Origin, "browser", StringComparison.Ordinal))
            {
                return new ProvenanceSource
                {
                    Kind = "browser",
                    Url = !string.IsNullOrEmpty(pe.Url) ? pe.Url : pe.ChromiumUrl,
                    Process = pe.ClipboardProcess,
                    Text = subText,
                    TickId = tickId,
                    OffsetInTick = takeOff,
                    Len = takeLen
                };
            }

            // Word-sourced tick. Try to descend into the grand-source.
            if (string.Equals(pe.Origin, "word", StringComparison.Ordinal))
            {
                var grandGuid = pe.SrcDocGuid;

                var outNode = new ProvenanceSource
                {
                    Kind = "word",
                    DocGuid = grandGuid,
                    SrcFile = pe.SrcFile,
                    Text = subText,
                    TickId = tickId,
                    OffsetInTick = takeOff,
                    Len = takeLen
                };

                // Cycle guard — don't recurse into a doc already on the stack.
                if (!string.IsNullOrEmpty(grandGuid) && visitedGuids.Contains(grandGuid))
                {
                    outNode.MappingFailure = "cycle-detected";
                    return outNode;
                }

                PasteTraceState grandState = null;
                if (!string.IsNullOrEmpty(grandGuid))
                    grandState = DocStateRegistry.GetByDocGuid(grandGuid);
                if (grandState == null && !string.IsNullOrEmpty(pe.SrcFile))
                    grandState = TryLoadState(app, pe.SrcFile);
                // Last resort: scan open docs' in-memory CustomXMLParts for a matching
                // DocGuid. ThisAddIn's ForceFlush writes to in-memory parts even for
                // unsaved docs, so this finds unsaved intermediaries in the chain that
                // were never registered in DocStateRegistry (e.g. because the host's
                // OnDocumentOpened integration was incomplete).
                if (grandState == null && !string.IsNullOrEmpty(grandGuid))
                    grandState = TryFindLiveDocStateByGuid(app, grandGuid);

                outNode.Live = (grandState != null);

                if (grandState == null)
                {
                    outNode.MappingFailure = "source-doc-unavailable";
                    return outNode;
                }

                // pe.Origins covers the full tick; slice it down to the sub-range we care about.
                var sliced = SliceOrigins(pe.Origins, takeOff, takeLen);
                if (sliced.Count == 0)
                {
                    outNode.MappingFailure = "no-origin-mapping";
                    return outNode;
                }

                if (!string.IsNullOrEmpty(grandGuid)) visitedGuids.Add(grandGuid);
                try
                {
                    var segments = new List<ProvenanceSegment>();
                    int cursor = 0;
                    foreach (var o in sliced)
                    {
                        var child = ResolveTickSubRange(app, grandState, o.t, o.off, o.n, visitedGuids, depth + 1);
                        string segText = child?.Text ?? "";
                        segments.Add(new ProvenanceSegment
                        {
                            Start = cursor,
                            Len = o.n,
                            Text = segText,
                            Source = child
                        });
                        cursor += o.n;
                    }
                    outNode.Segments = segments;
                }
                finally
                {
                    if (!string.IsNullOrEmpty(grandGuid)) visitedGuids.Remove(grandGuid);
                }
                return outNode;
            }

            // Word-plain / unknown — emit a best-effort leaf.
            return new ProvenanceSource
            {
                Kind = pe.Origin ?? "unknown",
                DocGuid = pe.SrcDocGuid,
                SrcFile = pe.SrcFile,
                Url = pe.Url,
                Process = pe.ClipboardProcess,
                Text = subText,
                TickId = tickId,
                OffsetInTick = takeOff,
                Len = takeLen,
                MappingFailure = pe.MappingFailure
            };
        }

        // Slice a consecutive origin list (each entry spans a range in the parent tick
        // whose total length sums to the tick's length) down to a flat sub-range.
        // Example: origins = [(A, 10, 3), (B, 0, 4)], subOff = 2, subLen = 4
        //          → byte 2 of A(10,3) = A offset 12 (len 1), bytes 3..5 of B(0,4) = B offset 0 (len 3)
        //          → result = [(A, 12, 1), (B, 0, 3)]
        private static List<(string t, int off, int n)> SliceOrigins(
            List<(string t, int off, int n)> origins, int subOff, int subLen)
        {
            var result = new List<(string t, int off, int n)>();
            if (origins == null || subLen <= 0) return result;
            int cum = 0;
            foreach (var o in origins)
            {
                if (o.n <= 0) continue;
                int segStart = cum;
                int segEnd = cum + o.n;
                cum = segEnd;

                int takeStart = Math.Max(subOff, segStart);
                int takeEnd = Math.Min(subOff + subLen, segEnd);
                if (takeEnd <= takeStart) continue;

                int localOff = takeStart - segStart;
                int takeLen = takeEnd - takeStart;
                result.Add((o.t, o.off + localOff, takeLen));
                if (takeEnd >= subOff + subLen) break;
            }
            return result;
        }
        // ----- end recursive tree builder -------------------------------------------

        static bool TryFindOpenWordSourceByText(Word.Application app0, string needle0,
            out Word.Document src0, out PasteTraceState st0, out List<(string t, int off, int n)> origins0)
        {
            src0 = null; st0 = null; origins0 = null;
            if (string.IsNullOrEmpty(needle0)) return false;

            // Line-ending normalisation. Word's Content.Text uses \r for paragraph breaks,
            // clipboard CF_UNICODETEXT typically uses \r\n, and intra-cell line breaks
            // can be \v. Without normalisation IndexOf silently misses matches that are
            // actually present in the doc — this was one of the two bugs causing
            // source-doc-unavailable for same-process Word→Word pastes.
            string Norm(string x) => (x ?? "")
                .Replace("\r\n", "\r")
                .Replace("\n", "\r")
                .Replace("\v", "\r")
                .TrimEnd();

            string needleN = Norm(needle0);
            if (needleN.Length == 0) return false;

            foreach (Word.Document d in app0.Documents)
            {
                bool isActive = false;
                try { isActive = (d == app0.ActiveDocument); } catch { }
                if (isActive) continue;

                try
                {
                    // Step 1: identify the source using Word's LIVE text. This works
                    // for any open doc regardless of whether it's saved or whether a
                    // PasteTraceState is registered for it — the addin's state registry
                    // may not cover every open doc (e.g. docs created before the addin
                    // loaded, or under a one-engine-many-docs design where only one
                    // State buffer exists at a time). Live text is always authoritative.
                    string liveText;
                    try { liveText = d.Content?.Text ?? ""; } catch { liveText = ""; }
                    string liveN = Norm(liveText);
                    if (liveN.Length == 0 || liveN.IndexOf(needleN, StringComparison.Ordinal) < 0)
                        continue;

                    // Source identified. Now look for a registered state so we can
                    // produce byte-level origin mapping. If no state is registered the
                    // caller still gets src + null origins and can emit a source-
                    // identified leaf with MappingFailure = "source-doc-no-trace".
                    PasteTraceState tmp = null;
                    string full = null;
                    try { full = d.FullName; } catch { }
                    if (!string.IsNullOrEmpty(full)) tmp = DocStateRegistry.GetByKey(full);

                    if (tmp == null)
                    {
                        try
                        {
                            var parts0 = d.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                            if (parts0 != null && parts0.Count >= 1)
                            {
                                tmp = new PasteTraceState
                                {
                                    DocGuid = "",
                                    AppVersion = PasteTraceEngine.AppVersion,
                                    PasteThreshold = PasteTraceEngine.DefaultPasteThreshold
                                };
                                PasteTraceXml.TryHydrate(d, tmp);
                            }
                        }
                        catch { }
                    }

                    // Step 2: build origins from the registered state's piece tree,
                    // if available and consistent with the needle. We search in the
                    // state's own flat text — that text is derived from the same
                    // piece tree as idx0, so offsets line up. We also normalise the
                    // state's flat text before searching.
                    List<(string t, int off, int n)> olist = null;
                    if (tmp != null)
                    {
                        try
                        {
                            var flat0 = BuildFlatVisibleText(tmp, out var idx0);
                            string flatN = Norm(flat0);
                            int hit0 = flatN.IndexOf(needleN, StringComparison.Ordinal);
                            if (hit0 >= 0)
                            {
                                // hit0 is an offset into flatN. Piece positions in idx0
                                // are into flat0. The normalised-vs-original offsets may
                                // differ when internal \r\n sequences collapse to \r.
                                // For correctness on well-formed inputs we do a raw
                                // IndexOf on flat0 as well; prefer raw hit when it exists.
                                int rawHit = flat0.IndexOf(needle0, StringComparison.Ordinal);
                                if (rawHit < 0) rawHit = hit0;

                                int end0 = rawHit + Math.Min(needle0.Length, flat0.Length - rawHit);
                                olist = new List<(string t, int off, int n)>();
                                foreach (var p in idx0)
                                {
                                    int ovStart = Math.Max(p.Start, rawHit);
                                    int ovEnd = Math.Min(p.Start + p.Len, end0);
                                    int ovLen = ovEnd - ovStart;
                                    if (ovLen > 0)
                                        olist.Add((p.TickId, p.OffsetInTick + (ovStart - p.Start), ovLen));
                                    if (ovEnd >= end0) break;
                                }
                                if (olist.Count == 0) olist = null;
                            }
                        }
                        catch { olist = null; }
                    }

                    src0 = d;
                    st0 = tmp;       // may be null — caller handles
                    origins0 = olist;     // may be null — caller handles
                    return true;
                }
                catch { }
            }
            return false;
        }

        static string TryFileUrl(Word.Document d0)
        {
            try
            {
                var full = d0?.FullName;
                if (string.IsNullOrEmpty(full)) return null;
                return new Uri(full, UriKind.Absolute).AbsoluteUri;
            }
            catch { return null; }
        }

        // PieceRow records a visible slice of a tick's text at (Start, Len) in the flat
        // view, plus OffsetInTick — the offset INTO the tick's text where that slice
        // begins. OffsetInTick is necessary because a single tick can be split into
        // multiple non-contiguous pieces (e.g. by an edit in the middle of a paste).
        // Callers mapping a flat substring back onto source ticks must add OffsetInTick
        // to the piece-local offset; earlier code used `ovStart - p.Start` on its own,
        // which silently reported wrong positions the moment any split occurred.
        private struct PieceRow { public int Start; public int Len; public string TickId; public int OffsetInTick; }

        private static string BuildFlatVisibleText(PasteTraceState st, out List<PieceRow> index)
        {
            var tickMap = new Dictionary<string, TickRow>(StringComparer.Ordinal);
            foreach (var t in st.Ticks) tickMap[t.TickId] = t;

            var sb = new StringBuilder(Math.Max(1024, st.Ticks.Sum(x => x.Len)));
            index = new List<PieceRow>();
            int visCursor = 0;

            foreach (var piece in st.EnumeratePiecesInOrder())
            {
                if (!piece.Visible || piece.Len <= 0) continue;
                TickRow tr;
                if (!tickMap.TryGetValue(piece.TickId, out tr) || tr.Text == null) continue;

                int off = Math.Max(0, piece.OffsetInTick);
                int n = Math.Max(0, Math.Min(piece.Len, tr.Text.Length - off));
                if (n <= 0) continue;

                sb.Append(tr.Text, off, n);
                index.Add(new PieceRow { Start = visCursor, Len = n, TickId = piece.TickId, OffsetInTick = off });
                visCursor += n;
            }

            return sb.ToString();
        }

        private static string SafeGetBuiltinProperty(Word.Document doc, string name)
        {
            try
            {
                var props = doc?.BuiltInDocumentProperties;
                if (props == null) return null;
                object prop = props[name];
                if (prop == null) return null;
                var v = prop.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, prop, null);
                return v?.ToString();
            }
            catch { return null; }
        }

        // Groups ticks by session (first 3 hex chars of TickId), takes the max CreatedElapsedMs
        // per session (≈ that session's duration), and sums across sessions.
        // Falls back to a tick-ID-based approximation for documents created before the 2026-04
        // refactor (those have CreatedElapsedMs == 0 throughout).
        private static int? ComputeTotalEditMinutesApprox(PasteTraceState st)
        {
            if (st.Ticks.Count == 0) return null;

            bool hasRealTimestamps = st.Ticks.Any(t => t.CreatedElapsedMs > 0);
            if (hasRealTimestamps)
            {
                long totalMs = st.Ticks
                    .GroupBy(t => t.TickId?.Length >= 3 ? t.TickId.Substring(0, 3) : "000")
                    .Sum(g => g.Max(t => t.CreatedElapsedMs));
                int mins = (int)(totalMs / 60_000L);
                return mins > 0 ? (int?)mins : null;
            }

            // Fallback: approximate from tick IDs. Values are rough (decimal-encoded IDs
            // are parsed as hex), but the path is display-only.
            int totalSeconds = st.Ticks.Max(t => ParseSessionSeconds(t.TickId));
            int fallbackMins = totalSeconds / 60;
            return fallbackMins > 0 ? (int?)fallbackMins : null;
        }

        // Parses chars [3..7] of the tick ID as base-16 (the poll counter field).
        private static int ParseSessionSeconds(string tickId)
        {
            if (tickId == null || tickId.Length < 8) return 0;
            try { return Convert.ToInt32(tickId.Substring(3, 5), 16); } catch { return 0; }
        }
    }
}