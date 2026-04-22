using System;
using Word = Microsoft.Office.Interop.Word;


namespace IsItYoursWordAddIn
{
    public sealed class PasteTraceEngine
    {
        public const string AppVersion = "0.1.0";
        public const int DefaultPasteThreshold = 20;

        private readonly Func<string> _getDocText;
        private readonly Func<int> _getCaretPos;
        public readonly PasteTraceState State;
        private string _activeDocId;
        private bool _sessionStartedForActiveDoc;


        public PasteTraceEngine(Func<string> getDocText, Func<int> getCaretPos)
        {
            _getDocText = getDocText;
            _getCaretPos = getCaretPos;
            State = new PasteTraceState
            {
                DocGuid = Guid.NewGuid().ToString(),
                AppVersion = AppVersion,
                SessionId = 0,
                SessionStartUtc = DateTime.UtcNow,
                PasteThreshold = DefaultPasteThreshold
            };
        }

        public bool PollOnce()
        {
            State.SessionSecondCounter++;
            string curr = _getDocText() ?? string.Empty;

            // FAST PATH: caret window probe
            int caret = -1;
            try { caret = _getCaretPos(); } catch { caret = -1; }

            if (State.PrevText == null)
            {
                // First observation of this document: establish baseline only.
                State.PrevText = (curr != null) ? curr : string.Empty;
                return false; // do not emit any ticks on baseline
            }

            // If lengths unchanged and caret is valid, compare a small window around the caret.
            // If that window is identical, assume no edit (fast common-case).
            if (caret >= 0 && curr.Length == State.PrevText.Length)
            {
                const int probe = 1024; // ��1KB
                int a = Math.Max(0, caret - probe);
                int b = Math.Min(curr.Length, caret + probe);
                int len = b - a;
                if (len > 0)
                {
                    // safe substring ranges guaranteed
                    if (string.Compare(State.PrevText, a, curr, a, len, StringComparison.Ordinal) == 0)
                    {
                        // quick no-change detection
                        return false;
                    }
                }
            }
            // Skip tick if nothing changed at all (cheap length + equality check)
            if (curr.Length == State.PrevText.Length)
            {
                if (string.Equals(curr, State.PrevText, StringComparison.Ordinal))
                    return false; 
            }
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

        private string MakeTickId() => $"{State.SessionId:000}{State.SessionSecondCounter:00000}";

        private static int LongestCommonPrefix(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length), i = 0; while (i < n && a[i] == b[i]) i++; return i;
        }
        private static int LongestCommonSuffix(string a, string b, int lcp)
        {
            int na = a.Length, nb = b.Length, n = Math.Min(na - lcp, nb - lcp), i = 0;
            while (i < n && a[na - 1 - i] == b[nb - 1 - i]) i++; return i;
        }

        private void ApplyDelete(int startVisible, int lengthVisible)
        {
            if (lengthVisible <= 0) return;

            string deleted = "";
            if (State.PrevText != null && startVisible >= 0 && startVisible + lengthVisible <= State.PrevText.Length)
                deleted = State.PrevText.Substring(startVisible, lengthVisible);

            var del = new TickRow { TickId = MakeTickId(), Op = "del", Loc = startVisible, Text = deleted, Len = lengthVisible, Paste = 0 };
            State.Ticks.Add(del);

            // Route through B-tree (tombstone; do not mutate offsets)
            State.TombstoneVisibleRange(startVisible, lengthVisible);
        }

        private void ApplyInsert(int atVisible, string text)
        {
            string tickId = MakeTickId();
            int paste = text.Length >= State.PasteThreshold ? 1 : 0;
            var ins = new TickRow { TickId = tickId, Op = "ins", Loc = atVisible, Text = text, Len = text.Length, Paste = paste };
            State.Ticks.Add(ins);

            State.InsertSpanAtVisible(atVisible, tickId, offInTick: 0, len: text.Length);

            // Adjacency fusion: defend against slow injection (e.g. AutoHotkey 19 chars/sec).
            // Walk back through recent insert ticks. If a run of contiguous inserts
            // (each individually below threshold, no intervening del, gap ≤1s between any two)
            // has a combined length >= PasteThreshold, flag all of them as paste.
            TryFuseAdjacentInserts();
        }

        // Adjacency fusion constants
        private const int FusionMinCharsEach = 3;   // ignore single-char autocorrect noise
        private const int FusionMaxGapSeconds = 1;  // ticks must be within 1 second of each other

        private void TryFuseAdjacentInserts()
        {
            var ticks = State.Ticks;
            int last = ticks.Count - 1;
            if (last < 1) return;

            // Walk backwards collecting a contiguous insert run ending at 'last'
            int runStart = last;
            int combinedLen = 0;
            int prevLoc = -1;

            for (int i = last; i >= 0; i--)
            {
                var t = ticks[i];
                if (t.Op != "ins") break;                          // del tick breaks the run
                if (t.Len < FusionMinCharsEach) break;             // noise tick breaks the run
                if (t.Paste == 1) break;                           // already flagged — no need to fuse further

                // Contiguity check: each tick must start exactly where the previous one ended
                if (prevLoc >= 0 && t.Loc + t.Len != prevLoc) break;

                // Timing check: gap between this tick and the next must be ≤ FusionMaxGapSeconds
                if (i < last)
                {
                    int secThis = ParseTickSeconds(t.TickId);
                    int secNext = ParseTickSeconds(ticks[i + 1].TickId);
                    if (secNext - secThis > FusionMaxGapSeconds) break;
                }

                prevLoc = t.Loc;
                combinedLen += t.Len;
                runStart = i;

                if (combinedLen >= State.PasteThreshold)
                {
                    // Flag every tick in the run as paste
                    for (int j = runStart; j <= last; j++)
                        if (ticks[j].Op == "ins") ticks[j].Paste = 1;
                    return;
                }
            }
        }

        // Parse the 5-hex second counter from a tick ID ("sssddddd")
        private static int ParseTickSeconds(string tickId)
        {
            if (tickId == null || tickId.Length < 8) return 0;
            try { return Convert.ToInt32(tickId.Substring(3, 5), 16); } catch { return 0; }
        }

        private void ApplyReplace(int at, int delLen, string insText)
        {
            if (delLen > 0) ApplyDelete(at, delLen);
            if (!string.IsNullOrEmpty(insText)) ApplyInsert(at, insText);
        }

        public void OnDocumentOpened(Word.Document doc, DateTime nowUtc)
        {
            var docId = SafeDocId(doc);
            if (_activeDocId == docId && _sessionStartedForActiveDoc) return;

            _activeDocId = docId;
            _sessionStartedForActiveDoc = false;

            // 1) Rehydrate existing sessions + ticks from the file
            State.ClearTransientCountersOnly();
#if !TEST_HARNESS
            PasteTraceXml.TryHydrate(doc, State);   // sets State.SessionId = max(existing)+1 when XML exists
#endif

            // 2) Start a new session using the SessionId decided by hydration (or 0 if none)
            State.StartSession(nowUtc);
            _sessionStartedForActiveDoc = true;

            // 3) Reset baseline text for diffing
            State.PrevText = _getDocText() ?? string.Empty;
        }

        public void OnDocumentActivated(Word.Document doc, DateTime nowUtc)
        {
            if (_activeDocId != SafeDocId(doc))
                OnDocumentOpened(doc, nowUtc);
        }

        public void OnDocumentClosing(Word.Document doc, DateTime nowUtc)
        {
            if (SafeDocId(doc) == _activeDocId)
                _sessionStartedForActiveDoc = false;
        }


        private static string SafeDocId(Word.Document doc)
        {
            try { return doc?.FullName ?? "(Untitled)"; } catch { return "(Unknown)"; }
        }

    }
}
