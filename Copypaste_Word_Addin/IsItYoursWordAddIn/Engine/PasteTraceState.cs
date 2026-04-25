using System;
using System.Collections.Generic;

namespace IsItYoursWordAddIn
{
    public sealed class PasteTraceState
    {
        public sealed class SessionRow { public int Id; public DateTime StartUtc; }
        public sealed class DebugRow
        {
            public DateTime Utc;
            public int SessionId;
            public string Stage;
            public string TickId;
            public string Message;
            public string Data;
        }

        public List<SessionRow> Sessions { get; } = new List<SessionRow>();
        public List<DebugRow> DebugLog { get; } = new List<DebugRow>();
        public const int MaxDebugRows = 1200;

        public void Log(string stage, string message, string tickId = null, string data = null)
        {
            try
            {
                string stg = stage ?? "";
                string msg = message ?? "";
                string dat = data ?? "";

                // Do not store raw timer heartbeat spam. It hides the useful provenance logs.
                // Keep actual text-change / tick.insert / tick.delete / flush / provenance events.
                if ((stg == "capture" && msg == "timer") ||
                    (stg == "capture" && dat.IndexOf("result=tick", StringComparison.OrdinalIgnoreCase) >= 0))
                    return;

                DebugLog.Add(new DebugRow
                {
                    Utc = DateTime.UtcNow,
                    SessionId = SessionId,
                    Stage = stg,
                    TickId = tickId ?? "",
                    Message = msg,
                    Data = dat
                });
                while (DebugLog.Count > MaxDebugRows) DebugLog.RemoveAt(0);
                Dirty = true;
            }
            catch { }
        }

        public string   DocGuid         { get; set; }
        public string   AppVersion      { get; set; }
        public int      SessionId       { get; set; }  // 0..4095 (3 hex digits), wraps at 4096
        public DateTime SessionStartUtc { get; set; }

        // Counts MakeTickId() calls per session. Transient — not serialised to XML.
        public int SessionPollCounter { get; set; } = 0;

        public int PasteThreshold { get; set; }

        // Per-document AES-256 key. Generated once on first open, never persisted in plaintext.
        public byte[] AesKey               { get; set; }
        // RSA-2048-OAEP wrapped — only the server can unwrap.
        public string WrappedAesKeyB64     { get; set; }
        // DPAPI-wrapped (CurrentUser scope, entropy = DocGuid bytes) for local re-hydration.
        public string LocalWrappedAesKeyB64 { get; set; }

#if !TEST_HARNESS
        // Single-slot clipboard candidate. Last copy wins; overwrite is intentional.
        public ClipboardCandidate _clipCandidate;
        public readonly Dictionary<string, PasteEvidence> _pasteEvidence
            = new Dictionary<string, PasteEvidence>(); // key: tickId
#endif

        public string PrevText { get; set; }

        public List<TickRow> Ticks { get; } = new List<TickRow>();

        public BTreeSeq Seq { get; } = new BTreeSeq();

        public int  SplitAtVisible(int offset) => Seq.EnsureBoundaryAt(offset);
        public void InsertSpanAtVisible(int at, string tickId, int offInTick, int len)
            => Seq.InsertSpan(at, new Piece
               { TickId = tickId, OffsetInTick = offInTick, Len = len, Visible = true });
        public void TombstoneVisibleRange(int start, int len) => Seq.TombstoneRange(start, len);

        public IEnumerable<Piece> EnumeratePiecesInOrder() => Seq.EnumeratePieces();
        public BTreeSnapshot SnapshotTree()                => Seq.Snapshot();

        // True when the capture path has appended a tick not yet persisted to the CustomXML part.
        // The flush timer skips the expensive Build/encrypt/write cycle when false.
        public bool Dirty { get; set; }

        // HMAC chain value after the last tick processed by BuildInner.
        // null = no flush yet this session; BuildInner recomputes from root.
        public byte[] LastComputedChainHmac { get; set; }

        // Count of ticks whose Hmac is already computed. BuildInner resumes from here.
        public int LastHmacTickIndex { get; set; }

        public void StartSession(DateTime utcNow)
        {
            SessionStartUtc       = utcNow;
            SessionPollCounter    = 0;
            LastComputedChainHmac = null;  // force full chain recompute from new root
            LastHmacTickIndex     = 0;

            // Always allocate a SessionId strictly greater than any previously used one
            // and always append a row. The previous behaviour ("skip if last row already
            // has this SessionId") caused two concrete regressions:
            //
            //   1. Session count stuck at 1. On a fresh doc SessionId defaults to 0 and is
            //      only bumped on re-open (HydrateInner). Re-entering StartSession within
            //      the same process (e.g. multi-doc activation cycles) left SessionId=0,
            //      so the guard short-circuited and no row was appended.
            //
            //   2. Duplicate tick IDs within one doc. TickId = {SessionId:X3}{PollCounter:X5}.
            //      Re-entering StartSession reset PollCounter to 0 without changing
            //      SessionId, so the next insert emitted "00000001" a second time, and
            //      later ticks overwrote earlier _pasteEvidence entries keyed on that ID.
            //
            // Computing nextId from max(existing) + 1 is also the behaviour HydrateInner
            // expects after a re-open, so the two paths agree.
            int nextId = 0;
            for (int i = 0; i < Sessions.Count; i++)
                if (Sessions[i].Id >= nextId) nextId = Sessions[i].Id + 1;
            SessionId = nextId % 4096;

            Sessions.Add(new SessionRow { Id = SessionId, StartUtc = utcNow });

            // Mark dirty so the new session row is persisted even if no edits follow.
            // Without this, a save-close-reopen with no edits drops the session row.
            Dirty = true;
            Log("session.start", "session row appended", null, "session=" + SessionId.ToString("X3") + ";sessions=" + Sessions.Count);
        }

        public void ClearTransientCountersOnly()
        {
            SessionPollCounter    = 0;
            Dirty                 = false;
            LastComputedChainHmac = null;
            LastHmacTickIndex     = 0;
        }
    }

    public sealed class TickRow
    {
        // Format: {SessionId:X3}{PollCounter:X5} (8 hex chars, e.g. "00000001")
        public string TickId { get; set; }

        public string Op    { get; set; }    // "ins" | "del"
        public int    Loc   { get; set; }    // visible-document character offset
        public string Text  { get; set; }    // inserted or deleted text
        public int    Len   { get; set; }
        public int    Paste { get; set; }    // 1 = paste-suspect

        // Milliseconds since session start. Used for fusion gap checks and XML "ms" attribute.
        public long CreatedElapsedMs { get; set; }

        // HMAC chain value (base64). Set during BuildInner, not at capture time.
        // Once non-null, BuildInner does not recompute this tick's chain step.
        public string Hmac { get; set; }
    }

    public sealed class Piece
    {
        public string TickId       { get; set; }
        public int    OffsetInTick { get; set; }
        public int    Len          { get; set; }
        public bool   Visible      { get; set; }  // false = tombstoned (deleted)
    }
}