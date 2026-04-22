using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word   = Microsoft.Office.Interop.Word;

namespace IsItYoursWordAddIn
{
    public partial class ThisAddIn
    {
        // _captureTimer (50 ms): polls for text changes, sets State.Dirty.
        // _flushTimer (2 s): builds encrypted XML and writes CustomXML only when dirty.
        // Decoupling the two timers keeps XML/encrypt/write overhead (~12–30 ms) off
        // the hot capture path.
        private Timer _captureTimer;
        private Timer _flushTimer;

        private readonly Dictionary<string, PasteTraceEngine> _engines
            = new Dictionary<string, PasteTraceEngine>(StringComparer.Ordinal);

        // Tracks the Word.Document object for each engine so ForceFlush can target
        // the correct document without relying on Application.ActiveDocument.
        private readonly Dictionary<string, Word.Document> _engineDocs
            = new Dictionary<string, Word.Document>(StringComparer.Ordinal);

        private ClipboardCandidate _pendingClipboard;
        private IClipboardProbe    _clip;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _clip = new ClipboardProbe();
            _clip.CandidateAvailable += c => { _pendingClipboard = c; };
            _clip.Start();

            this.Application.DocumentOpen        += Application_DocumentOpen;
            this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.WindowActivate      += Application_WindowActivate;
            this.Application.DocumentBeforeSave  += Application_DocumentBeforeSave;

            _captureTimer = new Timer { Interval = 50 };
            _captureTimer.Tick += (s, _) =>
            {
                try
                {
                    if (Application?.ActiveDocument == null) return;

                    var doc    = Application.ActiveDocument;
                    var engine = EnsureEngineFor(doc);

                    if (_pendingClipboard != null)
                    {
                        try { Provenance.SetCandidate(engine.State, _pendingClipboard); }
                        catch { }
                        _pendingClipboard = null;
                    }

                    engine.PollOnce();
                }
                catch { }
            };

            // Flush ALL dirty engines, not just the active document. A user who edits
            // doc A, switches to doc B for a long time, then crashes would lose unflushed
            // ticks from doc A if only the active document were flushed periodically.
            _flushTimer = new Timer { Interval = 2000 };
            _flushTimer.Tick += (s, _) =>
            {
                foreach (var kv in _engines)
                {
                    var engine = kv.Value;
                    if (!engine.State.Dirty) continue;
                    if (!_engineDocs.TryGetValue(kv.Key, out var doc) || doc == null) continue;
                    try { ForceFlush(doc, engine); }
                    catch { }
                }
            };

            _captureTimer.Start();
            _flushTimer.Start();
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentOpened(Doc, DateTime.UtcNow);
                ForceFlush(Doc, engine);
            }
            catch { }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentActivated(Doc, DateTime.UtcNow);

                // Only flush if there is new data or the CustomXML part does not exist yet.
                // Flushing on every activate caused a ~12–30 ms WriteCustomXml stall
                // each time the user switched Word windows with no new ticks.
                if (engine.State.Dirty || !HasCustomXmlPart(Doc))
                    ForceFlush(Doc, engine);
            }
            catch { }
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                ForceFlush(Doc, engine);
                engine.OnDocumentClosing(Doc, DateTime.UtcNow);

                var key = engine.State.DocGuid;
                if (key != null)
                {
                    _engines.Remove(key);
                    _engineDocs.Remove(key);
                }
            }
            catch { }
        }

        private void Application_DocumentBeforeSave(
            Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try { ForceFlush(Doc, EnsureEngineFor(Doc)); }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _captureTimer?.Stop();
            _flushTimer?.Stop();

            foreach (var kv in _engines)
            {
                if (!kv.Value.State.Dirty) continue;
                if (_engineDocs.TryGetValue(kv.Key, out var doc))
                    try { ForceFlush(doc, kv.Value); } catch { }
            }

            _clip?.Stop();
            _clip?.Dispose();
            _clip = null;

            _captureTimer?.Dispose();
            _flushTimer?.Dispose();
            _captureTimer = null;
            _flushTimer   = null;
        }

        // Attaches pending provenance for all unbound paste ticks, builds encrypted XML,
        // writes the CustomXML part. Clears State.Dirty only on a successful write.
        private void ForceFlush(Word.Document doc, PasteTraceEngine engine)
        {
            var ticks = engine.State.Ticks;
            // Scan backwards so the most recent paste tick gets the clipboard candidate first.
            for (int i = ticks.Count - 1; i >= 0; i--)
            {
                var t = ticks[i];
                if (t.Op != "ins" || t.Paste != 1) continue;
#if !TEST_HARNESS
                if (engine.State._pasteEvidence.ContainsKey(t.TickId)) continue;
                Provenance.AttachForPasteTick(this.Application, engine.State, t);
                if (engine.State._clipCandidate == null) break;
#endif
            }

            string xml     = PasteTraceXml.Build(engine.State);
            bool   written = WriteCustomXml(doc, "urn:paste-monitor", xml);
            // Only clear Dirty on success. On failure the next flush cycle retries;
            // the HMAC cache is already advanced so the retry reuses stored t.Hmac values.
            if (written)
                engine.State.Dirty = false;
        }

        private string GetDocumentText(Word.Document doc)
        {
            try { return doc?.Content?.Text ?? string.Empty; }
            catch { return string.Empty; }
        }

        private int GetDocumentCharCount(Word.Document doc)
        {
            try { return doc?.Characters?.Count ?? -1; }
            catch { return -1; }
        }

        private int GetDocumentCaretPos(Word.Document doc)
        {
            try
            {
                var active = Application?.ActiveDocument;
                if (active == null || doc == null) return -1;
                if (!ComObjectsEqual(active, doc)) return -1;

                var sel = Application?.Selection;
                return sel == null ? -1 : Math.Max(0, sel.Start);
            }
            catch { return -1; }
        }

        // Returns true if two RCW wrappers point to the same underlying COM object.
        // Marshal.AreComObjectsEqual was removed from .NET 4; compare IUnknown pointers instead.
        private static bool ComObjectsEqual(object a, object b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            IntPtr p1 = IntPtr.Zero, p2 = IntPtr.Zero;
            try
            {
                p1 = Marshal.GetIUnknownForObject(a);
                p2 = Marshal.GetIUnknownForObject(b);
                return p1 == p2;
            }
            catch { return false; }
            finally
            {
                if (p1 != IntPtr.Zero) Marshal.Release(p1);
                if (p2 != IntPtr.Zero) Marshal.Release(p2);
            }
        }

        // Returns true on success, false on any failure (read-only doc, COM error, etc.).
        // Caller treats false as "still dirty; retry next flush."
        private bool WriteCustomXml(Word.Document targetDoc, string ns, string xml)
        {
            if (targetDoc == null) return false;

            bool isReadOnly = false;
            try { isReadOnly = targetDoc.ReadOnly; } catch { return false; }
            if (isReadOnly) return false;

            Office.CustomXMLParts parts = null;
            try { parts = targetDoc.CustomXMLParts; } catch { return false; }
            if (parts == null) return false;

            try
            {
                var existing = parts.SelectByNamespace(ns);
                if (existing != null)
                    for (int i = existing.Count; i >= 1; i--) existing[i].Delete();

                parts.Add(xml);
                return true;
            }
            catch { return false; }
        }

        private bool HasCustomXmlPart(Word.Document doc)
        {
            try
            {
                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                return parts != null && parts.Count >= 1;
            }
            catch { return false; }
        }

        // Two-phase lookup: COM identity first (handles new docs without an XML part),
        // then DocGuid from the XML part (handles re-opens after Word restart).
        // Creating a new engine on every poll for a brand-new document would silently
        // lose all data — each new engine's baseline = current text, so no diffs fire.
        private PasteTraceEngine EnsureEngineFor(Word.Document doc)
        {
            if (doc == null) return null;

            foreach (var kv in _engineDocs)
            {
                try
                {
                    if (kv.Value != null && ComObjectsEqual(kv.Value, doc)
                        && _engines.TryGetValue(kv.Key, out var existingByCom))
                        return existingByCom;
                }
                catch { }
            }

            var g = TryReadDocGuid(doc);
            if (g != null && _engines.TryGetValue(g, out var existingByGuid))
                return existingByGuid;

            Word.Document capturedDoc = doc;
            var engine = new PasteTraceEngine(
                () => GetDocumentText(capturedDoc),
                () => GetDocumentCaretPos(capturedDoc),
                () => GetDocumentCharCount(capturedDoc));

            try { engine.OnDocumentOpened(doc, DateTime.UtcNow); } catch { }

            var key = engine.State.DocGuid ?? g ?? Guid.NewGuid().ToString();
            _engines[key]    = engine;
            _engineDocs[key] = doc;
            return engine;
        }

        private static string TryReadDocGuid(Word.Document doc)
        {
            try
            {
                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1) return null;
                var xml    = parts[1].XML;
                var marker = " g=\"";
                var i      = xml.IndexOf(marker, StringComparison.Ordinal);
                if (i < 0) return null;
                i += marker.Length;
                var j = xml.IndexOf('"', i);
                if (j <= i) return null;
                return xml.Substring(i, j - i);
            }
            catch { return null; }
        }

        #region VSTO generated
        private void InternalStartup()
        {
            this.Startup  += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
