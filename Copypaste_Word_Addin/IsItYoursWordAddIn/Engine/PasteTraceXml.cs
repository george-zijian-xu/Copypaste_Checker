using System;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;
using Office = Microsoft.Office.Core;
using Word   = Microsoft.Office.Interop.Word;

namespace IsItYoursWordAddIn
{
    public static class PasteTraceXml
    {
        private const string NS = "urn:paste-monitor";

        // tickfmt="2" = hex session IDs + real-elapsed-time ms on every tick.
        // Absent or other value → old decimal format (hydration falls back gracefully).
        private const string TickFormatVersion = "2";

        public static string Build(PasteTraceState s)
        {
            if (s.AesKey == null)
                s.AesKey = TraceEncryption.GenerateAesKey();
            if (s.WrappedAesKeyB64 == null)
                s.WrappedAesKeyB64 = TraceEncryption.WrapAesKey(s.AesKey);
            if (s.LocalWrappedAesKeyB64 == null)
                s.LocalWrappedAesKeyB64 = TraceEncryption.WrapAesKeyLocal(s.AesKey, s.DocGuid ?? "");

            string inner = BuildInner(s);
            var (ivB64, ctB64, tagB64) = TraceEncryption.Encrypt(s.AesKey, inner);

            var sb = new StringBuilder();
            sb.Append($@"<pasteTrace xmlns=""{NS}"">").AppendLine();
            sb.Append($@"  <doc g=""{E(s.DocGuid)}"" a=""{E(s.AppVersion)}"" tickfmt=""{TickFormatVersion}""/>").AppendLine();
            sb.Append($@"  <header kv=""{TraceEncryption.KeyVersion}"" ek=""{E(s.WrappedAesKeyB64)}"" lk=""{E(s.LocalWrappedAesKeyB64)}""/>").AppendLine();
            sb.Append($@"  <payload iv=""{E(ivB64)}"" ct=""{E(ctB64)}"" tag=""{E(tagB64)}""/>").AppendLine();
            sb.Append(@"</pasteTrace>");

            // Do NOT clear s.Dirty here. Build() only produces the XML string and cannot
            // know whether the caller's WriteCustomXml will succeed. The caller clears
            // Dirty only after a successful write so a failed write is retried.
            return sb.ToString();
        }

        private static string BuildInner(PasteTraceState s)
        {
            var sb = new StringBuilder();

            foreach (var srow in s.Sessions)
                sb.Append($@"  <session id=""{srow.Id:X3}"" startUtc=""{srow.StartUtc:yyyy-MM-ddTHH:mm:ssZ}""/>").AppendLine();

            // Incremental HMAC: resume from the cached value when possible.
            // Only new ticks (those without a stored Hmac) pay the HMAC cost.
            byte[] chainHmac;
            int    firstNewTick;

            if (s.LastComputedChainHmac != null && s.LastHmacTickIndex >= 0)
            {
                chainHmac    = s.LastComputedChainHmac;
                firstNewTick = s.LastHmacTickIndex + 1;
            }
            else
            {
                chainHmac    = TraceEncryption.ChainRoot(s.AesKey, s.DocGuid ?? "", s.SessionId, s.SessionStartUtc);
                firstNewTick = 0;
            }

            sb.Append(@"  <ticks>").AppendLine();
            for (int i = 0; i < s.Ticks.Count; i++)
            {
                var t = s.Ticks[i];
                if (i >= firstNewTick)
                {
                    chainHmac = TraceEncryption.ChainStep(
                        s.AesKey, chainHmac, t.TickId, t.Op ?? "ins", t.Loc, t.Len, t.Paste);
                    t.Hmac = TraceEncryption.HmacToString(chainHmac);
                }

                sb.Append($@"    <t id=""{E(t.TickId)}"" op=""{E(t.Op ?? "ins")}"" off=""{t.Loc}"" ms=""{t.CreatedElapsedMs}"" text=""{E(t.Text ?? string.Empty)}"" len=""{t.Len}""");
                if (t.Op == "ins") sb.Append($@" p=""{t.Paste}""");
                sb.Append($@" hmac=""{E(t.Hmac)}""");
                sb.Append(@"/>").AppendLine();
            }
            sb.Append(@"  </ticks>").AppendLine();

            if (s.Ticks.Count > 0)
            {
                s.LastHmacTickIndex     = s.Ticks.Count - 1;
                s.LastComputedChainHmac = chainHmac;
            }

            var snap = s.SnapshotTree();
            sb.Append(@"  <btree>").AppendLine();
            foreach (var n in snap.Nodes)
            {
                if (n.IsLeaf)
                {
                    sb.Append($@"    <leaf id=""{n.Id}"" parent=""{n.ParentId}"" vislen=""{n.VisibleLen}"">").AppendLine();
                    foreach (var lp in n.LeafPieces)
                        sb.Append($@"      <span t=""{E(lp.TickId)}"" off=""{lp.Off}"" n=""{lp.Len}"" vis=""{lp.Vis}""/>").AppendLine();
                    sb.Append(@"    </leaf>").AppendLine();
                }
                else
                {
                    sb.Append($@"    <node id=""{n.Id}"" parent=""{n.ParentId}"" vislen=""{n.VisibleLen}""/>").AppendLine();
                }
            }
            sb.Append(@"  </btree>").AppendLine();

            sb.Append(@"  <pastes>").AppendLine();
            foreach (var t in s.Ticks.Where(x => x.Op == "ins" && x.Paste == 1))
            {
                PasteEvidence ev;
                if (s._pasteEvidence.TryGetValue(t.TickId, out ev))
                {
                    sb.Append($@"    <pid t=""{E(t.TickId)}"" len=""{t.Len}"" origin=""{E(ev.Origin ?? "unknown")}"">").AppendLine();
                    sb.Append(@"      <evidence>").AppendLine();
                    sb.Append($@"        <clipboard ts=""{ev.ClipboardUtc:o}"" process=""{E(ev.ClipboardProcess ?? "")}""/>").AppendLine();

                    if (!string.IsNullOrEmpty(ev.Url))      sb.Append($@"        <url>{E(ev.Url)}</url>").AppendLine();
                    if (!string.IsNullOrEmpty(ev.FullText)) sb.Append($@"        <text>{E(ev.FullText)}</text>").AppendLine();
                    if (!string.IsNullOrEmpty(ev.Sha256))   sb.Append($@"        <sha256>{E(ev.Sha256)}</sha256>").AppendLine();

                    if (!string.IsNullOrEmpty(ev.ChromiumUrl) || !string.IsNullOrEmpty(ev.FirefoxTitle))
                    {
                        sb.Append(@"        <misc>").AppendLine();
                        if (!string.IsNullOrEmpty(ev.ChromiumUrl))  sb.Append($@"          <chromiumURL>{E(ev.ChromiumUrl)}</chromiumURL>").AppendLine();
                        if (!string.IsNullOrEmpty(ev.FirefoxTitle)) sb.Append($@"          <windowTitle>{E(ev.FirefoxTitle)}</windowTitle>").AppendLine();
                        sb.Append(@"        </misc>").AppendLine();
                    }

                    if (!string.IsNullOrEmpty(ev.OriginalPidXml))
                        foreach (var line in ev.OriginalPidXml.Split(new[] { Environment.NewLine }, StringSplitOptions.None))
                            sb.Append("        ").Append(line).AppendLine();

                    if (ev.Origin == "word")
                    {
                        bool hasAny = !string.IsNullOrEmpty(ev.SrcDocGuid) || !string.IsNullOrEmpty(ev.SrcFile) ||
                                      !string.IsNullOrEmpty(ev.SrcAuthor)  || !string.IsNullOrEmpty(ev.SrcTitle) ||
                                      ev.SrcTotalEditMin.HasValue || !string.IsNullOrEmpty(ev.MappingFailure) ||
                                      (ev.Origins != null && ev.Origins.Count > 0) || !string.IsNullOrEmpty(ev.OriginalPidXml);
                        if (hasAny)
                        {
                            sb.Append(@"        <doc");
                            if (!string.IsNullOrEmpty(ev.SrcDocGuid)) sb.Append($@" g=""{E(ev.SrcDocGuid)}""");
                            if (!string.IsNullOrEmpty(ev.SrcFile))    sb.Append($@" file=""{E(ev.SrcFile)}""");
                            sb.Append(@">").AppendLine();
                            if (!string.IsNullOrEmpty(ev.SrcAuthor))  sb.Append($@"          <author>{E(ev.SrcAuthor)}</author>").AppendLine();
                            if (!string.IsNullOrEmpty(ev.SrcTitle))   sb.Append($@"          <title>{E(ev.SrcTitle)}</title>").AppendLine();
                            if (ev.SrcTotalEditMin.HasValue)           sb.Append($@"          <totalEditMin>{ev.SrcTotalEditMin.Value}</totalEditMin>").AppendLine();
                            if (!string.IsNullOrEmpty(ev.MappingFailure)) sb.Append($@"          <mappingFailure>{E(ev.MappingFailure)}</mappingFailure>").AppendLine();
                            if (ev.Origins != null && ev.Origins.Count > 0)
                            {
                                sb.Append(@"          <origins>").AppendLine();
                                foreach (var o in ev.Origins)
                                    sb.Append($@"            <origin t=""{E(o.t)}"" off=""{o.off}"" n=""{o.n}""/>").AppendLine();
                                sb.Append(@"          </origins>").AppendLine();
                            }
                            if (ev.OriginCandidates != null && ev.OriginCandidates.Count > 0)
                            {
                                sb.Append(@"          <originCandidates>").AppendLine();
                                foreach (var o in ev.OriginCandidates)
                                    sb.Append($@"            <origin t=""{E(o.t)}"" off=""{o.off}"" n=""{o.n}""/>").AppendLine();
                                sb.Append(@"          </originCandidates>").AppendLine();
                            }
                            sb.Append(@"        </doc>").AppendLine();
                        }
                    }

                    if (ev.Chain != null && ev.Chain.Count > 0)
                    {
                        sb.Append(@"        <chain>").AppendLine();
                        foreach (var hop in ev.Chain)
                        {
                            sb.Append(@"          <hop");
                            if (!string.IsNullOrEmpty(hop.DocGuid))   sb.Append($@" g=""{E(hop.DocGuid)}""");
                            if (!string.IsNullOrEmpty(hop.SrcFile))   sb.Append($@" file=""{E(hop.SrcFile)}""");
                            if (!string.IsNullOrEmpty(hop.SrcAuthor)) sb.Append($@" author=""{E(hop.SrcAuthor)}""");
                            if (!string.IsNullOrEmpty(hop.SrcTitle))  sb.Append($@" title=""{E(hop.SrcTitle)}""");
                            if (hop.SrcTotalEditMin.HasValue)          sb.Append($@" editMin=""{hop.SrcTotalEditMin.Value}""");
                            sb.Append($@" status=""{E(hop.ChainStatus ?? "unknown")}""");
                            if (hop.Origins != null && hop.Origins.Count > 0)
                            {
                                sb.Append(@">").AppendLine();
                                sb.Append(@"            <origins>").AppendLine();
                                foreach (var o in hop.Origins)
                                    sb.Append($@"              <origin t=""{E(o.t)}"" off=""{o.off}"" n=""{o.n}""/>").AppendLine();
                                sb.Append(@"            </origins>").AppendLine();
                                sb.Append(@"          </hop>").AppendLine();
                            }
                            else sb.Append(@"/>").AppendLine();
                        }
                        sb.Append(@"        </chain>").AppendLine();
                    }

                    // Nested recursive provenance tree — authoritative view of this paste.
                    // Hydration doesn't parse this back yet (the browser analyzer reads it
                    // directly from the decrypted payload), so there's no round-trip
                    // obligation to preserve shape on save-close-reopen. The emitter
                    // tolerates empty branches and missing fields gracefully.
                    if (ev.ProvenanceTree != null)
                        WriteProvenanceSource(sb, ev.ProvenanceTree, "        ");

                    sb.Append(@"      </evidence>").AppendLine();
                    sb.Append(@"    </pid>").AppendLine();
                }
                else
                {
                    sb.Append($@"    <pid t=""{E(t.TickId)}"" len=""{t.Len}"" origin=""unknown""/>").AppendLine();
                }
            }
            sb.Append(@"  </pastes>").AppendLine();

            return sb.ToString();
        }

        public static bool TryHydrate(Word.Document doc, PasteTraceState state)
            => TryHydrate(doc, state, decryptKey: null);

        public static bool TryHydrate(Word.Document doc, PasteTraceState state, byte[] decryptKey)
        {
            try
            {
                Office.CustomXMLParts parts = doc?.CustomXMLParts;
                if (parts == null) return false;
                var existing = parts.SelectByNamespace(NS);
                if (existing == null || existing.Count < 1) return false;

                string xml  = existing[1].XML;
                var    xdoc = new XmlDocument();
                xdoc.LoadXml(xml);
                var nsmgr = new XmlNamespaceManager(xdoc.NameTable);
                nsmgr.AddNamespace("pm", NS);

                var docNode = xdoc.SelectSingleNode("/pm:pasteTrace/pm:doc", nsmgr);
                if (docNode?.Attributes?["g"] != null)
                    state.DocGuid = docNode.Attributes["g"].Value;

                bool hexFormat = (docNode?.Attributes?["tickfmt"]?.Value == TickFormatVersion);

                var headerNode = xdoc.SelectSingleNode("/pm:pasteTrace/pm:header", nsmgr) as XmlElement;
                if (headerNode != null)
                {
                    state.WrappedAesKeyB64      = headerNode.GetAttribute("ek");
                    state.LocalWrappedAesKeyB64 = headerNode.GetAttribute("lk");
                }

                var payloadNode = xdoc.SelectSingleNode("/pm:pasteTrace/pm:payload", nsmgr) as XmlElement;
                if (payloadNode == null) return false;

                string ivB64  = payloadNode.GetAttribute("iv");
                string ctB64  = payloadNode.GetAttribute("ct");
                string tagB64 = payloadNode.GetAttribute("tag");

                byte[] aesKey = decryptKey ?? state.AesKey;
                if (aesKey == null && !string.IsNullOrEmpty(state.LocalWrappedAesKeyB64))
                    aesKey = TraceEncryption.UnwrapAesKeyLocal(state.LocalWrappedAesKeyB64, state.DocGuid ?? "");
                if (aesKey == null) return false;

                string inner;
                try   { inner = TraceEncryption.Decrypt(aesKey, ivB64, ctB64, tagB64); }
                catch { return false; }

                state.AesKey = aesKey;

                var innerDoc = new XmlDocument();
                innerDoc.LoadXml("<root xmlns:pm=\"" + NS + "\">" + inner + "</root>");
                var insmgr = new XmlNamespaceManager(innerDoc.NameTable);
                insmgr.AddNamespace("pm", NS);

                return HydrateInner(innerDoc, insmgr, state, hexFormat);
            }
            catch { return false; }
        }

        private static bool HydrateInner(
            XmlDocument xdoc, XmlNamespaceManager nsmgr, PasteTraceState state, bool hexFormat = false)
        {
            state.Sessions.Clear();
            int maxId = -1;
            var sessionNodes = xdoc.SelectNodes("//pm:session", nsmgr) ?? xdoc.SelectNodes("//session");
            foreach (XmlNode n in sessionNodes)
            {
                string idStr = n.Attributes?["id"]?.Value ?? (hexFormat ? "000" : "0");
                int    id    = ParseSessionId(idStr, hexFormat);

                string startStr = n.Attributes?["startUtc"]?.Value ?? "";
                DateTime start;
                if (!DateTime.TryParseExact(startStr, "yyyy-MM-ddTHH:mm:ss'Z'",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out start))
                    DateTime.TryParse(startStr, CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out start);

                state.Sessions.Add(new PasteTraceState.SessionRow { Id = id, StartUtc = start });
                if (id > maxId) maxId = id;
            }
            state.SessionId = (maxId + 1) % 4096;

            state.Ticks.Clear();
            var tickNodes = xdoc.SelectNodes("//pm:t", nsmgr);
            if (tickNodes == null || tickNodes.Count == 0)
                tickNodes = xdoc.SelectNodes("//t");
            foreach (XmlNode t in tickNodes)
            {
                int.TryParse(t.Attributes?["off"]?.Value, out var off);
                int.TryParse(t.Attributes?["len"]?.Value, out var len);
                int.TryParse(t.Attributes?["p"]?.Value,   out var paste);
                long.TryParse(t.Attributes?["ms"]?.Value, out var elapsedMs);
                state.Ticks.Add(new TickRow
                {
                    TickId           = t.Attributes?["id"]?.Value   ?? "00000000",
                    Op               = t.Attributes?["op"]?.Value   ?? "ins",
                    Loc              = off, Text = t.Attributes?["text"]?.Value ?? "",
                    Len              = len, Paste = paste,
                    CreatedElapsedMs = elapsedMs,
                    Hmac             = t.Attributes?["hmac"]?.Value ?? ""
                });
            }

            // No HMAC cache priming on hydration. ChainRoot() takes the current SessionId and
            // SessionStartUtc as inputs. When StartSession() runs right after hydration those
            // values change, so any cached chain value would be inconsistent with what a
            // validator computes from the new session root. The first flush after re-open
            // recomputes the full chain from the current root (O(total ticks) once per
            // re-open); subsequent flushes in the same process lifetime are incremental.

            var btreeNode = xdoc.SelectSingleNode("//pm:btree", nsmgr) ?? xdoc.SelectSingleNode("//btree");
            if (btreeNode != null)
            {
                var snap = new BTreeSnapshot();
                foreach (XmlNode n in btreeNode.ChildNodes)
                {
                    if (n.NodeType != XmlNodeType.Element) continue;
                    var entry = new BTreeSnapshot.Node
                    {
                        Id         = int.Parse(n.Attributes?["id"]?.Value     ?? "0"),
                        ParentId   = int.Parse(n.Attributes?["parent"]?.Value ?? "-1"),
                        IsLeaf     = (n.LocalName == "leaf"),
                        VisibleLen = int.Parse(n.Attributes?["vislen"]?.Value ?? "0")
                    };
                    if (entry.IsLeaf)
                        foreach (XmlNode s in n.ChildNodes)
                        {
                            if (s.NodeType != XmlNodeType.Element || s.LocalName != "span") continue;
                            int.TryParse(s.Attributes?["off"]?.Value, out var soff);
                            int.TryParse(s.Attributes?["n"]?.Value,   out var slen);
                            int.TryParse(s.Attributes?["vis"]?.Value, out var vis);
                            entry.LeafPieces.Add(new BTreeSnapshot.LeafPiece
                            { TickId = s.Attributes?["t"]?.Value ?? "", Off = soff, Len = slen, Vis = vis });
                        }
                    snap.Nodes.Add(entry);
                }
                state.Seq.LoadSnapshot(snap);
            }

            state._pasteEvidence.Clear();
            var pidNodes = xdoc.SelectNodes("//pm:pid", nsmgr);
            if (pidNodes == null || pidNodes.Count == 0) pidNodes = xdoc.SelectNodes("//pid");
            if (pidNodes != null)
                foreach (XmlNode pid in pidNodes)
                {
                    var tickId = pid.Attributes?["t"]?.Value      ?? "";
                    var origin = pid.Attributes?["origin"]?.Value ?? "unknown";
                    var ev     = new PasteEvidence { Origin = origin };
                    var evNode = pid.SelectSingleNode("pm:evidence", nsmgr) ?? pid.SelectSingleNode("evidence");
                    if (evNode != null)
                    {
                        var clip = (evNode.SelectSingleNode("pm:clipboard", nsmgr) ?? evNode.SelectSingleNode("clipboard")) as XmlElement;
                        if (clip != null)
                        {
                            DateTime ts;
                            if (DateTime.TryParse(clip.GetAttribute("ts"), CultureInfo.InvariantCulture,
                                                  DateTimeStyles.RoundtripKind, out ts)) ev.ClipboardUtc = ts;
                            ev.ClipboardProcess = clip.GetAttribute("process") ?? "";
                        }
                        var url = evNode.SelectSingleNode("pm:url",    nsmgr) ?? evNode.SelectSingleNode("url");
                        if (url != null) ev.Url = url.InnerText ?? "";
                        var txt = evNode.SelectSingleNode("pm:text",   nsmgr) ?? evNode.SelectSingleNode("text");
                        if (txt != null) ev.FullText = txt.InnerText ?? "";
                        var sha = evNode.SelectSingleNode("pm:sha256", nsmgr) ?? evNode.SelectSingleNode("sha256");
                        if (sha != null) ev.Sha256 = sha.InnerText ?? "";
                        var misc = evNode.SelectSingleNode("pm:misc",  nsmgr) ?? evNode.SelectSingleNode("misc");
                        if (misc != null)
                        {
                            var chr = misc.SelectSingleNode("pm:chromiumURL", nsmgr) ?? misc.SelectSingleNode("chromiumURL");
                            if (chr != null) ev.ChromiumUrl = chr.InnerText ?? "";
                            var wt = misc.SelectSingleNode("pm:windowTitle", nsmgr) ?? misc.SelectSingleNode("windowTitle");
                            if (wt != null) ev.FirefoxTitle = wt.InnerText ?? "";
                        }
                        var docp = (evNode.SelectSingleNode("pm:doc",  nsmgr) ?? evNode.SelectSingleNode("doc")) as XmlElement;
                        if (docp != null)
                        {
                            ev.SrcDocGuid = docp.GetAttribute("g") ?? "";
                            ev.SrcFile    = docp.GetAttribute("file") ?? "";
                            var author  = docp.SelectSingleNode("pm:author",       nsmgr) ?? docp.SelectSingleNode("author");
                            if (author != null)  ev.SrcAuthor = author.InnerText ?? "";
                            var title   = docp.SelectSingleNode("pm:title",        nsmgr) ?? docp.SelectSingleNode("title");
                            if (title != null)   ev.SrcTitle  = title.InnerText ?? "";
                            var tem     = docp.SelectSingleNode("pm:totalEditMin", nsmgr) ?? docp.SelectSingleNode("totalEditMin");
                            if (tem != null)     { int v; if (int.TryParse(tem.InnerText ?? "", out v)) ev.SrcTotalEditMin = v; }
                            var mf      = docp.SelectSingleNode("pm:mappingFailure", nsmgr) ?? docp.SelectSingleNode("mappingFailure");
                            if (mf != null)      ev.MappingFailure = mf.InnerText ?? "";
                            var origins = docp.SelectSingleNode("pm:origins",      nsmgr) ?? docp.SelectSingleNode("origins");
                            if (origins != null)
                            {
                                ev.Origins = new System.Collections.Generic.List<(string t, int off, int n)>();
                                var oNodes = origins.SelectNodes("pm:origin", nsmgr) ?? origins.SelectNodes("origin");
                                foreach (XmlElement o in oNodes)
                                {
                                    int off = 0, n = 0;
                                    int.TryParse(o.GetAttribute("off"), out off);
                                    int.TryParse(o.GetAttribute("n"),   out n);
                                    ev.Origins.Add((o.GetAttribute("t") ?? "", off, n));
                                }
                            }
                            var candidates = docp.SelectSingleNode("pm:originCandidates", nsmgr) ?? docp.SelectSingleNode("originCandidates");
                            if (candidates != null)
                            {
                                ev.OriginCandidates = new System.Collections.Generic.List<(string t, int off, int n)>();
                                var cNodes = candidates.SelectNodes("pm:origin", nsmgr) ?? candidates.SelectNodes("origin");
                                foreach (XmlElement o in cNodes)
                                {
                                    int off = 0, n = 0;
                                    int.TryParse(o.GetAttribute("off"), out off);
                                    int.TryParse(o.GetAttribute("n"),   out n);
                                    ev.OriginCandidates.Add((o.GetAttribute("t") ?? "", off, n));
                                }
                            }
                        }
                        var chainNode = evNode.SelectSingleNode("pm:chain", nsmgr) ?? evNode.SelectSingleNode("chain");
                        if (chainNode != null)
                        {
                            ev.Chain = new System.Collections.Generic.List<ProvenanceHop>();
                            var hopNodes = chainNode.SelectNodes("pm:hop", nsmgr) ?? chainNode.SelectNodes("hop");
                            foreach (XmlElement hopEl in hopNodes)
                            {
                                var hop = new ProvenanceHop
                                {
                                    DocGuid     = hopEl.GetAttribute("g")      ?? "",
                                    SrcFile     = hopEl.GetAttribute("file")   ?? "",
                                    SrcAuthor   = hopEl.GetAttribute("author") ?? "",
                                    SrcTitle    = hopEl.GetAttribute("title")  ?? "",
                                    ChainStatus = hopEl.GetAttribute("status") ?? "unknown"
                                };
                                int em;
                                if (int.TryParse(hopEl.GetAttribute("editMin"), out em)) hop.SrcTotalEditMin = em;
                                var hopOrigins = hopEl.SelectSingleNode("pm:origins", nsmgr) ?? hopEl.SelectSingleNode("origins");
                                if (hopOrigins != null)
                                {
                                    hop.Origins = new System.Collections.Generic.List<(string t, int off, int n)>();
                                    var hoNodes = hopOrigins.SelectNodes("pm:origin", nsmgr) ?? hopOrigins.SelectNodes("origin");
                                    foreach (XmlElement ho in hoNodes)
                                    {
                                        int hoff = 0, hn = 0;
                                        int.TryParse(ho.GetAttribute("off"), out hoff);
                                        int.TryParse(ho.GetAttribute("n"),   out hn);
                                        hop.Origins.Add((ho.GetAttribute("t") ?? "", hoff, hn));
                                    }
                                }
                                ev.Chain.Add(hop);
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(tickId))
                        state._pasteEvidence[tickId] = ev;
                }

            return true;
        }

        private static int ParseSessionId(string s, bool hexFormat)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            if (hexFormat)
            {
                try   { return Convert.ToInt32(s, 16); }
                catch { int v; int.TryParse(s, out v); return v; }
            }
            else
            {
                int v;
                if (int.TryParse(s, out v)) return v;
                try   { return Convert.ToInt32(s, 16); } catch { return 0; }
            }
        }

        // Recursive emitter for ProvenanceSource trees. Callers pass the leading
        // whitespace the <source> element should be indented at; child elements add
        // two more spaces per level and <segment> children recurse with indent + 4.
        // Leaves (browser / doc-local / unknown) produce a <source> with a small
        // set of child elements. Word nodes with Segments recurse into <segments>.
        private static void WriteProvenanceSource(StringBuilder sb, ProvenanceSource src, string indent)
        {
            if (src == null) return;

            string childIndent = indent + "  ";
            string segIndent   = childIndent +"fuckkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk" +"  ";
            string innerIndent = segIndent + "  ";

            sb.Append(indent).Append(@"<source kind=""").Append(E(src.Kind ?? "unknown")).Append('"');
            if (!string.IsNullOrEmpty(src.DocGuid))      sb.Append(@" docGuid=""").Append(E(src.DocGuid)).Append('"');
            if (!string.IsNullOrEmpty(src.SrcFile))      sb.Append(@" file=""").Append(E(src.SrcFile)).Append('"');
            if (src.Len > 0)                              sb.Append(@" len=""").Append(src.Len).Append('"');
            if (!string.IsNullOrEmpty(src.TickId))       sb.Append(@" tick=""").Append(E(src.TickId)).Append(@""" off=""").Append(src.OffsetInTick).Append('"');
            sb.Append(@" live=""").Append(src.Live ? "true" : "false").Append('"');
            if (!string.IsNullOrEmpty(src.MappingFailure))
                sb.Append(@" mappingFailure=""").Append(E(src.MappingFailure)).Append('"');
            sb.Append('>').AppendLine();

            if (!string.IsNullOrEmpty(src.Url))
                sb.Append(childIndent).Append("<url>").Append(E(src.Url)).Append("</url>").AppendLine();
            if (!string.IsNullOrEmpty(src.Process))
                sb.Append(childIndent).Append("<process>").Append(E(src.Process)).Append("</process>").AppendLine();
            if (!string.IsNullOrEmpty(src.Text))
                sb.Append(childIndent).Append("<text>").Append(E(src.Text)).Append("</text>").AppendLine();

            if (src.Segments != null && src.Segments.Count > 0)
            {
                sb.Append(childIndent).Append("<segments>").AppendLine();
                foreach (var seg in src.Segments)
                {
                    sb.Append(segIndent).Append(@"<segment start=""").Append(seg.Start)
                      .Append(@""" len=""").Append(seg.Len).Append('"');
                    if (!string.IsNullOrEmpty(seg.Text))
                        sb.Append(@" text=""").Append(E(seg.Text)).Append('"');
                    sb.Append('>').AppendLine();
                    if (seg.Source != null)
                        WriteProvenanceSource(sb, seg.Source, innerIndent);
                    sb.Append(segIndent).Append("</segment>").AppendLine();
                }
                sb.Append(childIndent).Append("</segments>").AppendLine();
            }

            sb.Append(indent).Append("</source>").AppendLine();
        }

        private static string E(string s) => SecurityElement.Escape(s) ?? string.Empty;
    }
}