// =============================================================================
// ThisAddIn.cs
// CHANGE LOG (2026-04 refactor)
// =============================================================================
// [P0-5] Per-document engine closures (doc-specific, not ActiveDocument).
//   EnsureEngineFor() now captures the specific Word.Document reference in the
//   getDocText / getCharCount / getCaretPos lambdas, so each engine reads from
//   its own document rather than always reading Application.ActiveDocument.
//   New helper methods: GetDocumentText(doc), GetDocumentCharCount(doc),
//   GetDocumentCaretPos(doc) — take a Word.Document parameter instead of
//   using the Application.ActiveDocument global.
//
// [P1-1] Split single 1 Hz timer into two-lane architecture.
//   _captureTimer (50 ms): polls for text changes, appends ticks, sets
//     State.Dirty = true.  Does NOT build XML or write CustomXMLParts.
//   _flushTimer (2000 ms): fires only when the active engine is dirty;
//     builds encrypted XML and writes the CustomXML part.
//   Result: XML/encrypt/write overhead (~12–30 ms) is incurred at most
//   once every 2 seconds instead of on every changed capture tick.
//
// [P0-6] Flush on document save and close.
//   Added Application.DocumentBeforeSave hook → ForceFlush().
//   Fixed Application_DocumentBeforeClose to call ForceFlush() BEFORE
//   OnDocumentClosing(), preventing data loss on close-before-flush.
//
// [P1-2] ForceFlush(doc, engine) centralises the build+write path.
//   Called from the flush timer, DocumentBeforeSave, and
//   DocumentBeforeClose.  After a successful write it clears State.Dirty.
//
// [P1-3] Shutdown flushes all dirty engines.
//   ThisAddIn_Shutdown iterates _engines and flushes any that are dirty
//   before stopping timers, protecting against data loss on Word exit.
// =============================================================================

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word   = Microsoft.Office.Interop.Word;

namespace IsItYoursWordAddIn
{
    public partial class ThisAddIn
    {
        // [P1-1] Two timers replace the old single _timer.
        private Timer _captureTimer;  // 50 ms  — capture only, no XML write
        private Timer _flushTimer;    // 2000 ms — XML build + CustomXML write

        // Per-document engines keyed by DocGuid.
        private readonly Dictionary<string, PasteTraceEngine> _engines
            = new Dictionary<string, PasteTraceEngine>(StringComparer.Ordinal);

        // [P0-5] Track the Word.Document for each engine so ForceFlush can find
        // the right document object without relying on Application.ActiveDocument.
        private readonly Dictionary<string, Word.Document> _engineDocs
            = new Dictionary<string, Word.Document>(StringComparer.Ordinal);

        private ClipboardCandidate _pendingClipboard;
        private IClipboardProbe    _clip;

        // ── Startup ───────────────────────────────────────────────────────────
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _clip = new ClipboardProbe();
            _clip.CandidateAvailable += c => { _pendingClipboard = c; };
            _clip.Start();

            this.Application.DocumentOpen          += Application_DocumentOpen;
            this.Application.DocumentBeforeClose   -= Application_DocumentBeforeClose;
            this.Application.DocumentBeforeClose   += Application_DocumentBeforeClose;
            this.Application.WindowActivate        += Application_WindowActivate;
            // [P0-6] Flush on save so no ticks are lost between periodic flushes.
            this.Application.DocumentBeforeSave    += Application_DocumentBeforeSave;

            // ── [P1-1] Capture timer (50 ms) ──────────────────────────────────
            _captureTimer = new Timer { Interval = 50 };
            _captureTimer.Tick += (s, _) =>
            {
                try
                {
                    if (Application?.ActiveDocument == null) return;

                    var doc    = Application.ActiveDocument;
                    var engine = EnsureEngineFor(doc);

                    // Deliver any pending clipboard candidate before polling.
                    // If this is a Word copy and the source document is still active,
                    // capture the source's live provenance BEFORE the user pastes into
                    // another document. This is what makes unsaved doc3 -> doc2 -> doc1
                    // nesting possible without waiting for CustomXML flushes.
                    if (_pendingClipboard != null)
                    {
                        try { Provenance.TryEnrichWordCandidateFromActiveDocument(this.Application, engine.State, _pendingClipboard); }
                        catch (Exception ex) { engine.State.Log("prov.enrich.error", "TryEnrichWordCandidateFromActiveDocument threw", null, ex.GetType().Name + ": " + ex.Message); }
                        try { Provenance.SetCandidate(engine.State, _pendingClipboard); }
                        catch (Exception ex) { engine.State.Log("prov.setcandidate.error", "SetCandidate threw", null, ex.GetType().Name + ": " + ex.Message); }
                        _pendingClipboard = null;
                    }

                    // PollOnce() internally gates on the cheap sentinel; it sets
                    // State.Dirty only when a real change was detected.
                    bool changed = engine.PollOnce();

                    // Attach clipboard evidence immediately, not only at the 2s flush.
                    // Otherwise a second copy can overwrite the single clipboard slot
                    // before the first paste tick receives its evidence.
                    if (changed)
                        AttachRecentPasteEvidence(engine);
                }
                catch { /* swallow — must not kill the capture timer */ }
            };

            // ── [P1-1] Flush timer (2 s) ───────────────────────────────────────
            // [Fix-7] Flush ALL dirty engines, not just the active document.
            // Previously only the active document was flushed periodically.  If a
            // user edits doc A, switches to doc B for a long time, then Word crashes,
            // doc A's unflushed ticks are lost because the periodic timer never
            // reached it.  Save/close/shutdown hooks still flush on those events,
            // but the periodic flush is the durability backstop.
            _flushTimer = new Timer { Interval = 2000 };
            _flushTimer.Tick += (s, _) =>
            {
                foreach (var kv in _engines)
                {
                    var engine = kv.Value;
                    if (!engine.State.Dirty) continue;

                    if (!_engineDocs.TryGetValue(kv.Key, out var doc)) continue;
                    if (doc == null) continue;

                    try { ForceFlush(doc, engine); }
                    catch { /* one engine failing must not block others */ }
                }
            };

            _captureTimer.Start();
            _flushTimer.Start();
        }

        // ── Document event handlers ───────────────────────────────────────────
        private void Application_DocumentOpen(Word.Document Doc)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentOpened(Doc, DateTime.UtcNow);
                // Write initial XML immediately so the CustomXML part exists before
                // any typing or pasting occurs (important for hydration on re-open).
                ForceFlush(Doc, engine);
            }
            catch { /* log if you have logging */ }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentActivated(Doc, DateTime.UtcNow);

                // [P0-9] Only flush if there is something to write OR the XML part
                // doesn't exist yet.  Flushing on every activate regardless caused an
                // unnecessary ~12–30 ms WriteCustomXml stall every time the user
                // switched Word windows — even when no new ticks had been captured.
                if (engine.State.Dirty || !HasCustomXmlPart(Doc))
                    ForceFlush(Doc, engine);
            }
            catch { /* ignore */ }
        }

        // [P0-6] Flush BEFORE calling OnDocumentClosing, then let the engine
        // mark the session as closed.  Previously this only flipped a flag,
        // losing any ticks captured since the last flush timer fire.
        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                ForceFlush(Doc, engine);                         // [P0-6] persist last ticks
                engine.OnDocumentClosing(Doc, DateTime.UtcNow); // then mark session done

                // Remove the engine from the registry so it is not polled after close.
                var key = engine.State.DocGuid;
                if (key != null)
                {
                    _engines.Remove(key);
                    _engineDocs.Remove(key);
                }
            }
            catch { /* ignore */ }
        }

        // [P0-6] New: flush on save to avoid data loss between flush-timer fires.
        private void Application_DocumentBeforeSave(
            Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                var engine = EnsureEngineFor(Doc);
                ForceFlush(Doc, engine);
            }
            catch { /* ignore */ }
        }

        // ── Shutdown ──────────────────────────────────────────────────────────
        // [P1-3] Flush all dirty engines before stopping timers.
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _captureTimer?.Stop();
            _flushTimer?.Stop();

            // Flush any engines that have unsaved ticks.
            foreach (var kv in _engines)
            {
                var engine = kv.Value;
                if (!engine.State.Dirty) continue;
                if (_engineDocs.TryGetValue(kv.Key, out var doc))
                {
                    try { ForceFlush(doc, engine); } catch { /* best-effort */ }
                }
            }

            _clip?.Stop();
            _clip?.Dispose();
            _clip = null;

            _captureTimer?.Dispose();
            _flushTimer?.Dispose();
            _captureTimer = null;
            _flushTimer   = null;
        }

        private void AttachRecentPasteEvidence(PasteTraceEngine engine)
        {
            try
            {
                var ticks = engine.State.Ticks;
                int floor = Math.Max(0, ticks.Count - 64);
                for (int i = ticks.Count - 1; i >= floor; i--)
                {
                    var t = ticks[i];
                    if (t.Op != "ins" || t.Paste != 1) continue;
#if !TEST_HARNESS
                    if (engine.State._pasteEvidence.ContainsKey(t.TickId)) continue;
                    Provenance.AttachForPasteTick(this.Application, engine.State, t);
                    if (engine.State._clipCandidate == null) break;
#endif
                }
            }
            catch (Exception ex) { engine.State.Log("prov.attach_recent.error", "AttachRecentPasteEvidence threw", null, ex.GetType().Name + ": " + ex.Message); }
        }

        // ── ForceFlush ────────────────────────────────────────────────────────
        // Attaches pending provenance for ALL paste ticks that do not yet have
        // evidence, builds encrypted XML, writes CustomXML part.
        // [Fix-4] Previously only the last tick was checked; earlier paste ticks
        // that occurred between flush cycles were silently missed.  Clipboard
        // evidence is single-slot and overwrites, so the candidate is consumed on
        // the first matching tick found scanning backwards.
        private void ForceFlush(Word.Document doc, PasteTraceEngine engine)
        {
            var ticks = engine.State.Ticks;
            // Scan backwards so the most recent paste tick gets the current
            // clipboard candidate first (it is most likely to match).
            for (int i = ticks.Count - 1; i >= 0; i--)
            {
                var t = ticks[i];
                if (t.Op != "ins" || t.Paste != 1) continue;
#if !TEST_HARNESS
                // Skip ticks that already have evidence from a prior flush.
                if (engine.State._pasteEvidence.ContainsKey(t.TickId)) continue;
                Provenance.AttachForPasteTick(this.Application, engine.State, t);
                // AttachForPasteTick consumes _clipCandidate; stop scanning once
                // it is gone so we do not pass null to subsequent ticks.
                if (engine.State._clipCandidate == null) break;
#endif
            }

            string xml = PasteTraceXml.Build(engine.State);
            WriteCustomXml(doc, "urn:paste-monitor", xml);
            // State.Dirty is cleared inside PasteTraceXml.Build() after the encrypt
            // step succeeds (line that sets s.Dirty = false in Build()).
        }

        // ── Document text / caret / char-count helpers ────────────────────────
        // [P0-5] Each helper takes a specific Word.Document rather than reading
        // Application.ActiveDocument.  EnsureEngineFor captures these in closures.

        private string GetDocumentText(Word.Document doc)
        {
            try { return doc?.Content?.Text ?? string.Empty; }
            catch { return string.Empty; }
        }

        // [P0-4] Cheap sentinel: returns character count without fetching text.
        // Characters.Count is an integer COM property; much faster than Content.Text.
        private int GetDocumentCharCount(Word.Document doc)
        {
            try { return doc?.Characters?.Count ?? -1; }
            catch { return -1; }
        }

        private int GetDocumentCaretPos(Word.Document doc)
        {
            try
            {
                // [Fix-6] Compare COM object identity instead of FullName strings.
                // FullName requires a COM call and is empty for unsaved documents,
                // which makes the string comparison unreliable.
                // ComObjectsEqual() checks whether two RCW wrappers point
                // to the same underlying COM object — correct, cheap, and null-safe.
                var active = Application?.ActiveDocument;
                if (active == null || doc == null) return -1;
                if (!ComObjectsEqual(active, doc))
                    return -1;

                var sel = Application?.Selection;
                return sel == null ? -1 : Math.Max(0, sel.Start);
            }
            catch { return -1; }
        }

        // ── WriteCustomXml ────────────────────────────────────────────────────
        // Unchanged from original: delete any existing part, add the new one.
        // Called only from ForceFlush (flush cadence), not from the capture timer.
        private void WriteCustomXml(Word.Document targetDoc, string ns, string xml)
        {
            if (targetDoc == null) return;

            bool isReadOnly = false;
            try { isReadOnly = targetDoc.ReadOnly; } catch { }
            if (isReadOnly) return;

            Office.CustomXMLParts parts = null;
            try { parts = targetDoc.CustomXMLParts; } catch { }
            if (parts == null) return;

            try
            {
                var existing = parts.SelectByNamespace(ns);
                if (existing != null)
                    for (int i = existing.Count; i >= 1; i--) existing[i].Delete();

                parts.Add(xml);
            }
            catch { /* swallow; avoid killing the caller */ }
        }

        // [P0-9] Cheap existence check used by WindowActivate to skip unnecessary flushes.
        // Returns true if the document already has a urn:paste-monitor CustomXML part.
        // Does not decrypt or read the payload — just checks the namespace count.
        private bool HasCustomXmlPart(Word.Document doc)
        {
            try
            {
                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                return parts != null && parts.Count >= 1;
            }
            catch { return false; }
        }

        // ── EnsureEngineFor ───────────────────────────────────────────────────
        // [P0-5] Now captures the specific Word.Document in each delegate closure
        // so engines don't all read Application.ActiveDocument.
        private PasteTraceEngine EnsureEngineFor(Word.Document doc)
        {
            var g = TryReadDocGuid(doc);

            if (g != null && _engines.TryGetValue(g, out var existing))
                return existing;

            // Capture the document reference for the closures.
            // Word.Document COM objects remain valid as long as the document is open.
            Word.Document capturedDoc = doc;

            var engine = new PasteTraceEngine(
                () => GetDocumentText(capturedDoc),     // [P0-5]
                () => GetDocumentCaretPos(capturedDoc), // [P0-5]
                () => GetDocumentCharCount(capturedDoc) // [P0-4] new
            );

            try { engine.OnDocumentOpened(doc, DateTime.UtcNow); } catch { }

            var key = engine.State.DocGuid ?? g ?? Guid.NewGuid().ToString();
            _engines[key]    = engine;
            _engineDocs[key] = doc;    // [P0-5] track the document reference
            return engine;
        }

        // COM identity check: compares IUnknown pointers so two RCW wrappers for the
        // same underlying COM object compare equal. Marshal.AreComObjectsEqual does not
        // exist in .NET 4.8; this is the correct replacement.
        private static bool ComObjectsEqual(object a, object b)
        {
            if (a == null || b == null) return false;
            var pA = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(a);
            var pB = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(b);
            try   { return pA == pB; }
            finally
            {
                System.Runtime.InteropServices.Marshal.Release(pA);
                System.Runtime.InteropServices.Marshal.Release(pB);
            }
        }

        // Light parse of the existing XML part to extract the doc guid.
        // Unchanged from original.
        private static string TryReadDocGuid(Word.Document doc)
        {
            try
            {
                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1) return null;
                var xml = parts[1].XML;
                var marker = " g=\"";
                var i = xml.IndexOf(marker, StringComparison.Ordinal);
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