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
        // "resolved" | "source-unavailable" | "no-trace" | "cycle-detected"
        public string Status;
        public List<(string t, int off, int n)> Origins;
    }

    public sealed class PasteEvidence
    {
        public string Origin; // "browser" | "word" | "word-plain" | "unknown"
        public DateTime ClipboardUtc;
        public string ClipboardProcess;
        public string Url;           // CF_HTML SourceURL (file:/// or http(s)://)
        public string ChromiumUrl;
        public string FirefoxTitle;
        public string FullText;
        public string Sha256;

        public string SrcDocGuid;
        public string SrcFile;
        public string SrcAuthor;
        public string SrcTitle;
        public int? SrcTotalEditMin;
        public string WasPaste;     // "yes" | "no" | "unknown"
        public List<(string t, int off, int n)> Origins;
        public string OriginalPidXml;

        // Recursive provenance chain: hop[0] = immediate source, hop[n] = canon origin.
        // null = not yet resolved; empty list = typed originally at the immediate source.
        public List<ProvenanceHop> Chain;
    }

    public static class Provenance
    {
        public static void SetCandidate(PasteTraceState s, ClipboardCandidate c) { s._clipCandidate = c; }
        public static ClipboardCandidate ConsumeCandidate(PasteTraceState s)
        {
            var c = s._clipCandidate; s._clipCandidate = null; return c;
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

            // If it looks like Word, proceed regardless of subset match —
            // TryWordMapping will inspect the source doc directly.
            bool mayProceed = subsetOk || looksWord;

            if (!mayProceed)
            {
                s._clipCandidate = null;
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
            s._clipCandidate = null;
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
                    ev.SrcDocGuid = string.IsNullOrEmpty(openState.DocGuid) ? null : openState.DocGuid;
                    ev.SrcFile = TryFileUrl(openSrc);
                    ev.SrcAuthor = SafeGetBuiltinProperty(openSrc, "Author");
                    ev.SrcTitle = SafeGetBuiltinProperty(openSrc, "Title");
                    ev.SrcTotalEditMin = ComputeTotalEditMinutesApprox(openState);
                    ev.WasPaste = "yes";
                    ev.Origins = openOrigins;
                    return;
                }

                ev.WasPaste = "unknown";
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
                ev.WasPaste = "unknown";
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
                    {
                        srcDoc = d;
                        break;
                    }
                }

                // 3) If not open, open read-only and invisible
                if (srcDoc == null)
                {
                    try { priorActive = app.ActiveDocument; } catch { }
                    srcDoc = app.Documents.Open(
                        FileName: srcPath,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        Visible: false
                    );
                    openedByUs = true;
                }

                // 4) Look for our custom XML part in the source
                var parts = srcDoc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1)
                {
                    ev.Origin = "word-plain";
                    ev.SrcFile = null;
                    ev.SrcDocGuid = null;
                    ev.WasPaste = "unknown";
                    return;
                }

                // 5) Hydrate a temp state from the source doc's XML
                var srcState = new PasteTraceState { DocGuid = "", AppVersion = PasteTraceEngine.AppVersion, PasteThreshold = PasteTraceEngine.DefaultPasteThreshold };
                PasteTraceXml.TryHydrate(srcDoc, srcState);

                // 6) Flatten visible text from the source state
                var flat = BuildFlatVisibleText(srcState, out var pieceIndex);

                // 7) Locate the pasted text in the source (tolerant to line-ending/trailing-space differences)
                string Norm(string x) => (x ?? "").Replace("\r\n", "\r").Replace("\n", "\r").TrimEnd();
                var n1 = Norm(ev.FullText);
                var n2 = Norm(cand.Text);
                var needle = (n1.Length >= n2.Length ? n1 : n2);
                if (string.IsNullOrEmpty(needle))
                {
                    ev.WasPaste = "unknown";
                    return;
                }
                int hit = flat.IndexOf(needle, StringComparison.Ordinal);
                if (hit < 0)
                {
                    ev.WasPaste = "unknown";
                }
                else
                {
                    // 8) Map the matched range back to (tickId, off, n) segments
                    int end = hit + needle.Length;
                    var origins = new List<(string t, int off, int n)>();
                    foreach (var p in pieceIndex)
                    {
                        int ovStart = Math.Max(p.Start, hit);
                        int ovEnd   = Math.Min(p.Start + p.Len, end);
                        int ovLen   = ovEnd - ovStart;
                        if (ovLen > 0) origins.Add((p.TickId, ovStart, ovLen));
                        if (ovEnd >= end) break;
                    }

                    if (origins.Count > 0) { ev.WasPaste = "yes"; ev.Origins = origins; }
                    else                   { ev.WasPaste = "unknown"; }
                }

                // 9) Populate Word document metadata
                ev.SrcDocGuid = string.IsNullOrEmpty(srcState.DocGuid) ? null : srcState.DocGuid;
                ev.SrcFile = url;
                ev.SrcAuthor = SafeGetBuiltinProperty(srcDoc, "Author");
                ev.SrcTitle = SafeGetBuiltinProperty(srcDoc, "Title");
                ev.SrcTotalEditMin = ComputeTotalEditMinutesApprox(srcState);
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
                                "</pid>" + Environment.NewLine
                            );
                        }
                        if (sbNested.Length > 0)
                            ev.OriginalPidXml = sbNested.ToString().TrimEnd();
                    }
                }
                catch { }

                // 10) Recursively resolve the provenance chain
                if (ev.WasPaste == "yes" && ev.Origins != null && ev.Origins.Count > 0)
                {
                    try
                    {
                        var visited = new HashSet<string>(StringComparer.Ordinal);
                        if (!string.IsNullOrEmpty(s.DocGuid)) visited.Add(s.DocGuid);
                        if (!string.IsNullOrEmpty(srcState.DocGuid)) visited.Add(srcState.DocGuid);
                        ev.Chain = ResolveChain(app, srcState, ev.Origins, visited, depth: 0);
                    }
                    catch { }
                }
            }
            catch
            {
                ev.Origin = "word-plain";
                ev.SrcDocGuid = null;
                ev.SrcFile = null;
                ev.WasPaste = "unknown";
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
                if (srcEv.WasPaste != "yes" || string.IsNullOrEmpty(srcEv.SrcDocGuid)) continue;

                if (visited.Contains(srcEv.SrcDocGuid))
                {
                    chain.Add(new ProvenanceHop { DocGuid = srcEv.SrcDocGuid, SrcFile = srcEv.SrcFile, Status = "cycle-detected" });
                    continue;
                }

                visited.Add(srcEv.SrcDocGuid);

                var hop = new ProvenanceHop
                {
                    DocGuid         = srcEv.SrcDocGuid,
                    SrcFile         = srcEv.SrcFile,
                    SrcAuthor       = srcEv.SrcAuthor,
                    SrcTitle        = srcEv.SrcTitle,
                    SrcTotalEditMin = srcEv.SrcTotalEditMin,
                    Origins         = srcEv.Origins
                };

                var grandState = TryLoadState(app, srcEv.SrcFile);
                if (grandState == null)      { hop.Status = "source-unavailable"; chain.Add(hop); continue; }
                if (string.IsNullOrEmpty(grandState.DocGuid)) { hop.Status = "no-trace"; chain.Add(hop); continue; }

                hop.Status = "resolved";
                chain.Add(hop);
                chain.AddRange(ResolveChain(app, grandState, srcEv.Origins, visited, depth + 1));
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
                if (openedByUs && doc != null)      try { doc.Close(SaveChanges: false); } catch { }
                if (openedByUs && priorActive != null) try { priorActive.Activate(); } catch { }
            }
        }

        static bool TryFindOpenWordSourceByText(Word.Application app0, string needle0,
            out Word.Document src0, out PasteTraceState st0, out List<(string t, int off, int n)> origins0)
        {
            src0 = null; st0 = null; origins0 = null;
            if (string.IsNullOrEmpty(needle0)) return false;
            foreach (Word.Document d in app0.Documents)
            {
                bool isActive = false;
                try { isActive = (d == app0.ActiveDocument); } catch { }
                if (isActive) continue;
                try
                {
                    var parts0 = d.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                    if (parts0 == null || parts0.Count < 1) continue;

                    var tmp = new PasteTraceState { DocGuid = "", AppVersion = PasteTraceEngine.AppVersion, PasteThreshold = PasteTraceEngine.DefaultPasteThreshold };
                    PasteTraceXml.TryHydrate(d, tmp);

                    var flat0 = BuildFlatVisibleText(tmp, out var idx0);
                    int hit0 = flat0.IndexOf(needle0, StringComparison.Ordinal);
                    if (hit0 < 0) continue;

                    int end0 = hit0 + needle0.Length;
                    var olist = new List<(string t, int off, int n)>();
                    foreach (var p in idx0)
                    {
                        int ovStart = Math.Max(p.Start, hit0);
                        int ovEnd   = Math.Min(p.Start + p.Len, end0);
                        int ovLen   = ovEnd - ovStart;
                        if (ovLen > 0) olist.Add((p.TickId, ovStart, ovLen));
                        if (ovEnd >= end0) break;
                    }

                    if (olist.Count > 0) { src0 = d; st0 = tmp; origins0 = olist; return true; }
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

        private struct PieceRow { public int Start; public int Len; public string TickId; }

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
                int n   = Math.Max(0, Math.Min(piece.Len, tr.Text.Length - off));
                if (n <= 0) continue;

                sb.Append(tr.Text, off, n);
                index.Add(new PieceRow { Start = visCursor, Len = n, TickId = piece.TickId });
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
