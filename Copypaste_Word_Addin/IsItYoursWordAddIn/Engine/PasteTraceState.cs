using System;
using System.Collections.Generic;

namespace IsItYoursWordAddIn
{
    public sealed class PasteTraceState
    {
        public sealed class SessionRow { public int Id; public DateTime StartUtc; }
        public List<SessionRow> Sessions { get; } = new List<SessionRow>();

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
            Dirty                 = false;
            LastComputedChainHmac = null;  // force full chain recompute from new root
            LastHmacTickIndex     = 0;

            if (Sessions.Count == 0 || Sessions[Sessions.Count - 1].Id != SessionId)
                Sessions.Add(new SessionRow { Id = SessionId, StartUtc = utcNow });
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
