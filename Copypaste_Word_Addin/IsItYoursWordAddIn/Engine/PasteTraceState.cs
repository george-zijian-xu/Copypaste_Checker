using System;
using System.Collections.Generic;

namespace IsItYoursWordAddIn
{
    public sealed class PasteTraceState
    {
        public sealed class SessionRow { public int Id; public DateTime StartUtc; }
        public List<SessionRow> Sessions { get; } = new List<SessionRow>();

        // Doc/session meta
        public string DocGuid { get; set; }
        public string AppVersion { get; set; }
        public int SessionId { get; set; }              // 0..4095 (3 hex, wraps at 4096)
        public DateTime SessionStartUtc { get; set; }
        public int SessionSecondCounter { get; set; } = 0;
        public int PasteThreshold { get; set; }

        // Per-document AES-256 session key (32 bytes).
        // Generated once on first open, never persisted in plaintext.
        // Stored in the encrypted XML header as RSA-wrapped ciphertext.
        public byte[] AesKey { get; set; }

        // RSA-wrapped AES key (base64) as read from the XML header.
        // Held so we can round-trip it unchanged when re-writing the XML.
        public string WrappedAesKeyB64 { get; set; }

        // DPAPI-wrapped AES key (base64) for local re-hydration on document re-open.
        // Only valid on the same Windows user account that created it.
        public string LocalWrappedAesKeyB64 { get; set; }

#if !TEST_HARNESS
        public ClipboardCandidate _clipCandidate;                 // single-slot
        public readonly Dictionary<string, PasteEvidence> _pasteEvidence = new Dictionary<string, PasteEvidence>(); // key: tickId
#endif

        // Snapshot baseline
        public string PrevText { get; set; }

        // Ticks (insert + delete ops)
        public List<TickRow> Ticks { get; } = new List<TickRow>();

        // B+ tree piece table
        public BTreeSeq Seq { get; } = new BTreeSeq();

        // Facade used by the engine (visible-document offsets)
        public int SplitAtVisible(int offset) => Seq.EnsureBoundaryAt(offset);
        public void InsertSpanAtVisible(int at, string tickId, int offInTick, int len)
            => Seq.InsertSpan(at, new Piece { TickId = tickId, OffsetInTick = offInTick, Len = len, Visible = true });
        public void TombstoneVisibleRange(int start, int len) => Seq.TombstoneRange(start, len);

        // Snapshot enumerations for XML
        public IEnumerable<Piece> EnumeratePiecesInOrder() => Seq.EnumeratePieces();
        public BTreeSnapshot SnapshotTree() => Seq.Snapshot();

        public void StartSession(DateTime utcNow)
        {
            SessionStartUtc = utcNow;
            SessionSecondCounter = 0;

            if (Sessions.Count == 0 || Sessions[Sessions.Count - 1].Id != SessionId)
                Sessions.Add(new SessionRow { Id = SessionId, StartUtc = utcNow });
        }

        public void ClearTransientCountersOnly()
        {
            SessionSecondCounter = 0;
            // Keep Sessions and Ticks -- they will be reloaded from XML
        }
    }

    public sealed class TickRow
    {
        public string TickId { get; set; }   // "sssddddd"
        public string Op { get; set; }       // "ins" | "del"
        public int Loc { get; set; }         // visible doc offset where op applies
        public string Text { get; set; }
        public int Len { get; set; }
        public int Paste { get; set; }
        public string Hmac { get; set; }     // HMAC chain value (base64), set during XML build
    }

    public sealed class Piece
    {
        public string TickId { get; set; }    // immutable
        public int OffsetInTick { get; set; } // immutable tick-local offset of first char in this piece
        public int Len { get; set; }          // length of this span (chars)
        public bool Visible { get; set; }     // tombstone flag
    }
}
