using System;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace IsItYoursWordAddIn
{
    public partial class ThisAddIn
    {
        private Timer _timer;
        // Per-document engines keyed by pasteTrace <doc g="...">
        private readonly System.Collections.Generic.Dictionary<string, PasteTraceEngine> _engines
            = new System.Collections.Generic.Dictionary<string, PasteTraceEngine>(StringComparer.Ordinal);
        // Defer clipboard candidate and bind it to whichever doc emits the next paste tick
        private ClipboardCandidate _pendingClipboard;
        private IClipboardProbe _clip;


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _clip = new ClipboardProbe();
            // Just stash the latest clipboard candidate; we’ll bind it to the active doc at tick time
            _clip.CandidateAvailable += c => { _pendingClipboard = c; };
            _clip.Start();

            this.Application.DocumentOpen += Application_DocumentOpen;
            this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose; // ensure no duplicates
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.WindowActivate += Application_WindowActivate;


            _timer = new Timer { Interval = 1000 };
            _timer.Tick += (s, _) =>
            {
                try
                {
                    if (Application?.ActiveDocument == null) return;

                    var doc = Application.ActiveDocument;
                    var engine = EnsureEngineFor(doc);                 // NEW: resolve per-<doc g="..."> engine
                    // If we have a pending clipboard event, give it to the active doc *before* we poll
                    if (_pendingClipboard != null)
                    {
                        try { Provenance.SetCandidate(engine.State, _pendingClipboard); }
                        catch { /* ignore */ }
                        _pendingClipboard = null;
                    }

                    bool changed = engine.PollOnce();
                    if (!changed) return;

                    // Attach provenance if last op is a paste-suspect insert
                    var ticks = engine.State.Ticks;
                    if (ticks.Count > 0)
                    {
                        var last = ticks[ticks.Count - 1];
                        if (last.Op == "ins" && last.Paste == 1)
                            Provenance.AttachForPasteTick(this.Application, engine.State, last);
                    }

                    // Write or update the custom XML part for THIS doc
                    string xml = PasteTraceXml.Build(engine.State);
                    WriteCustomXml(doc, "urn:paste-monitor", xml);


                }
                catch { /* swallow to avoid killing timer; consider logging */ }
            };
            _timer.Start();
        }
        private void Application_DocumentOpen(Word.Document Doc)
        {
            try { 
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentOpened(Doc, DateTime.UtcNow);
                // ——————— Write initial CustomXML immediately ————————
                // so that part.Exists == true before any typing/pasting
                var xml = PasteTraceXml.Build(engine.State);
                WriteCustomXml(Doc,"urn:paste-monitor", xml);
            }
            catch { /* log if you have logging */ }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            try {
                var engine = EnsureEngineFor(Doc);
                engine.OnDocumentActivated(Doc, DateTime.UtcNow);
                // ensure part is there even if no edits yet
                var xml = PasteTraceXml.Build(engine.State);
                WriteCustomXml(Doc,"urn:paste-monitor", xml);
            }
            catch { /* ignore */ }
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            try { EnsureEngineFor(Doc).OnDocumentClosing(Doc, DateTime.UtcNow); }
            catch { /* ignore */ }
        }



        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _timer?.Stop();
            _clip?.Stop();
            _clip?.Dispose();
            _clip = null;
            _timer?.Dispose();
            _timer = null;
        }

        private string GetActiveDocumentText()
        {
            try
            {
                var doc = Application?.ActiveDocument;
                if (doc == null) return string.Empty;
                // Word returns CRs; keep them (we’re linearizing text)
                return doc.Content?.Text ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        private int GetCaretPos()
        {
            try
            {
                var sel = Application?.Selection;
                if (sel == null) return -1;
                // Convert Word selection to 0-based char offset in document text.
                // Word characters are 1-based; Selection.Start is 0-based in content.
                // We keep it simple: use Start as an approx caret.
                return Math.Max(0, sel.Start);
            }
            catch { return -1; }
        }

        private void WriteCustomXml(Word.Document targetDoc, string ns, string xml)
        {
            if (targetDoc == null) return;

            // do not touch background read-only sources
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
            catch { /* swallow; avoid timer death */ }
        }

        // Return the pasteTrace <doc g="..."> if present; else null.
        private static string TryReadDocGuid(Word.Document doc)
        {
            try
            {
                var parts = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
                if (parts == null || parts.Count < 1) return null;
                var xml = parts[1].XML;
                // super-light parse: look for g="..."; we already trust our own writer’s shape
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

        // Create or fetch the per-document engine keyed by <doc g="...">
        private PasteTraceEngine EnsureEngineFor(Word.Document doc)
        {
            // If the doc already has pasteTrace XML, use that g; else we’ll mint one via the engine ctor
            var g = TryReadDocGuid(doc);

            if (g != null && _engines.TryGetValue(g, out var existing))
                return existing;

            var engine = new PasteTraceEngine(() => GetActiveDocumentText(), () => GetCaretPos());

            // Let the engine hydrate + start a session; this also sets State.DocGuid (either loaded or newly minted)
            try { engine.OnDocumentOpened(doc, DateTime.UtcNow); } catch { }

            var key = engine.State.DocGuid ?? g ?? Guid.NewGuid().ToString();
            _engines[key] = engine;
            return engine;
        }

        #region VSTO generated
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
