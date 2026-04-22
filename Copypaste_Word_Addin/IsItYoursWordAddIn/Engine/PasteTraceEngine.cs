using System;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace IsItYoursWordAddIn
{
    public sealed class PasteTraceEngine
    {
        public const string AppVersion            = "0.1.0";
        public const int    DefaultPasteThreshold = 20;

        // Delegates bound at construction — each engine owns lambdas over its specific document.
        private readonly Func<string> _getDocText;   // doc.Content.Text — expensive COM call
        private readonly Func<int>    _getCaretPos;  // Selection.Start  — cheap
        private readonly Func<int>    _getCharCount; // doc.Characters.Count — cheap sentinel

        public readonly PasteTraceState State;
        private string _activeDocId;
        private bool   _sessionStartedForActiveDoc;

        // Monotonic session clock. Stopwatch is immune to NTP jumps and VM resume steps
        // that can make DateTime.UtcNow go backwards, which would produce negative CreatedElapsedMs.
        private long _sessionStartTimestamp;

        // Sentinel state for the two-step poll optimisation.
        private int _prevCharCount       = -1;
        private int _prevCaretPos        = -1;
        private int _pollsSinceFullSweep = 0;

        // At 50 ms cadence: 20 polls × 50 ms = 1 s between mandatory full-text reads.
        // Full sweeps catch equal-length replacements that the char-count sentinel misses.
        private const int FullSweepEveryNPolls = 20;

        private const int FusionMinCharsEach = 3;     // ignore single-char autocorrect noise
        private const int FusionMaxGapMs     = 1000;

        // getCharCount is optional; pass null (or omit) to disable the sentinel.
        // When the sentinel is off every poll reads the full document text — safe, not optimal.
        public PasteTraceEngine(
            Func<string> getDocText,
            Func<int>    getCaretPos,
            Func<int>    getCharCount = null)
        {
            _getDocText   = getDocText;
            _getCaretPos  = getCaretPos;
            _getCharCount = getCharCount ?? (() => -1);

            State = new PasteTraceState
            {
                DocGuid         = Guid.NewGuid().ToString(),
                AppVersion      = AppVersion,
                SessionId       = 0,
                SessionStartUtc = DateTime.UtcNow,
                PasteThreshold  = DefaultPasteThreshold
            };
        }

        public bool PollOnce()
        {
            // Stage 1: cheap sentinel — skip the expensive Content.Text read when nothing changed.
            int charCount = -1;
            try { charCount = _getCharCount(); } catch { charCount = -1; }

            int caret = -1;
            try { caret = _getCaretPos(); } catch { caret = -1; }

            _pollsSinceFullSweep++;
            bool doFullSweep = _pollsSinceFullSweep >= FullSweepEveryNPolls;
            if (doFullSweep) _pollsSinceFullSweep = 0;

            bool skipFull = !doFullSweep
                            && charCount >= 0 && charCount == _prevCharCount
                            && caret    >= 0 && caret    == _prevCaretPos;

            if (skipFull) return false;

            // Stage 2: full text read and diff.
            string curr = _getDocText() ?? string.Empty;

            _prevCharCount = charCount >= 0 ? charCount : curr.Length;
            _prevCaretPos  = caret;

            if (State.PrevText == null)
            {
                State.PrevText = curr;
                return false;
            }

            // Inner caret-window fast exit: compare a 1 KB window around the caret before the LCS diff.
            if (caret >= 0 && curr.Length == State.PrevText.Length)
            {
                const int probe = 1024;
                int a = Math.Max(0, caret - probe);
                int b = Math.Min(curr.Length, caret + probe);
                int len = b - a;
                if (len > 0 &&
                    string.Compare(State.PrevText, a, curr, a, len,
                                   StringComparison.Ordinal) == 0)
                    return false;
            }

            if (curr.Length == State.PrevText.Length &&
                string.Equals(curr, State.PrevText, StringComparison.Ordinal))
                return false;

            string prev = State.PrevText;
            int lcp = LongestCommonPrefix(prev, curr);
            int lcs = LongestCommonSuffix(prev, curr, lcp);
            int prevMiddle = prev.Length - lcp - lcs;
            int currMiddle = curr.Length - lcp - lcs;

            if (prevMiddle > 0) ApplyDelete(lcp, prevMiddle);
            if (currMiddle > 0) ApplyInsert(lcp, curr.Substring(lcp, currMiddle));

            State.PrevText = curr;
            return true;
        }

        // Counter incremented here (not per-poll) so a replace in one poll emits two distinct IDs.
        private string MakeTickId()
        {
            State.SessionPollCounter++;
            return $"{State.SessionId:X3}{State.SessionPollCounter:X5}";
        }

        private long GetElapsedMs()
        {
            long ticks = Stopwatch.GetTimestamp() - _sessionStartTimestamp;
            return ticks * 1000L / Stopwatch.Frequency;
        }

        private static int LongestCommonPrefix(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length), i = 0;
            while (i < n && a[i] == b[i]) i++;
            return i;
        }

        private static int LongestCommonSuffix(string a, string b, int lcp)
        {
            int na = a.Length, nb = b.Length, n = Math.Min(na - lcp, nb - lcp), i = 0;
            while (i < n && a[na - 1 - i] == b[nb - 1 - i]) i++;
            return i;
        }

        private void ApplyDelete(int startVisible, int lengthVisible)
        {
            if (lengthVisible <= 0) return;

            string deleted = string.Empty;
            if (State.PrevText != null
                && startVisible >= 0
                && startVisible + lengthVisible <= State.PrevText.Length)
                deleted = State.PrevText.Substring(startVisible, lengthVisible);

            var del = new TickRow
            {
                TickId           = MakeTickId(),
                Op               = "del",
                Loc              = startVisible,
                Text             = deleted,
                Len              = lengthVisible,
                Paste            = 0,
                CreatedElapsedMs = GetElapsedMs(),
            };
            State.Ticks.Add(del);
            State.Dirty = true;

            State.TombstoneVisibleRange(startVisible, lengthVisible);
        }

        private void ApplyInsert(int atVisible, string text)
        {
            string tickId = MakeTickId();
            int paste = text.Length >= State.PasteThreshold ? 1 : 0;

            var ins = new TickRow
            {
                TickId           = tickId,
                Op               = "ins",
                Loc              = atVisible,
                Text             = text,
                Len              = text.Length,
                Paste            = paste,
                CreatedElapsedMs = GetElapsedMs(),
            };
            State.Ticks.Add(ins);
            State.Dirty = true;

            // Always update the B+ tree for every insert, even small ones.
            // Partial updates would corrupt visible-offset coordinates used by later
            // TombstoneVisibleRange and InsertSpanAtVisible calls.
            State.InsertSpanAtVisible(atVisible, tickId, offInTick: 0, len: text.Length);

            TryFuseAdjacentInserts();
        }

        // Defends against AutoHotkey-style slow injection (~19 chars/sec).
        // If a run of contiguous insert ticks combined >= PasteThreshold arrives
        // within FusionMaxGapMs of each other, all are retroactively flagged as paste.
        private void TryFuseAdjacentInserts()
        {
            var ticks = State.Ticks;
            int last = ticks.Count - 1;
            if (last < 1) return;

            int runStart    = last;
            int combinedLen = 0;
            int prevLoc     = -1;

            for (int i = last; i >= 0; i--)
            {
                var t = ticks[i];
                if (t.Op    != "ins")          break;
                if (t.Len   <  FusionMinCharsEach) break;
                if (t.Paste == 1)              break;

                if (prevLoc >= 0 && t.Loc + t.Len != prevLoc) break;

                if (i < last)
                {
                    long msThis = t.CreatedElapsedMs;
                    long msNext = ticks[i + 1].CreatedElapsedMs;
                    if (msNext - msThis > FusionMaxGapMs) break;
                }

                prevLoc      = t.Loc;
                combinedLen += t.Len;
                runStart     = i;

                if (combinedLen >= State.PasteThreshold)
                {
                    for (int j = runStart; j <= last; j++)
                        if (ticks[j].Op == "ins") ticks[j].Paste = 1;
                    return;
                }
            }
        }

        public void OnDocumentOpened(Word.Document doc, DateTime nowUtc)
        {
            var docId = SafeDocId(doc);
            if (_activeDocId == docId && _sessionStartedForActiveDoc) return;

            _activeDocId = docId;
            _sessionStartedForActiveDoc = false;

            State.ClearTransientCountersOnly();
#if !TEST_HARNESS
            PasteTraceXml.TryHydrate(doc, State);
#endif
            State.StartSession(nowUtc);
            _sessionStartedForActiveDoc = true;

            _sessionStartTimestamp = Stopwatch.GetTimestamp();

            _prevCharCount       = -1;
            _prevCaretPos        = -1;
            _pollsSinceFullSweep = 0;

            State.PrevText = _getDocText() ?? string.Empty;
        }

        public void OnDocumentActivated(Word.Document doc, DateTime nowUtc)
        {
            if (_activeDocId != SafeDocId(doc))
                OnDocumentOpened(doc, nowUtc);
        }

        public void OnDocumentClosing(Word.Document doc, DateTime nowUtc)
        {
            if (SafeDocId(doc) != _activeDocId) return;
            // Durability lives in ThisAddIn.Application_DocumentBeforeClose → ForceFlush,
            // which runs before this method. Do not set Dirty here.
            _sessionStartedForActiveDoc = false;
        }

        private static string SafeDocId(Word.Document doc)
        {
            try { return doc?.FullName ?? "(Untitled)"; } catch { return "(Unknown)"; }
        }
    }
}
