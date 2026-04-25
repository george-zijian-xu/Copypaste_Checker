// DocStateRegistry.cs
// Process-wide registry of live per-document PasteTraceState instances.
//
// Why this exists:
//   Cross-document provenance resolution (e.g. "doc1 was pasted from doc2,
//   which was itself pasted from doc3") needs to read another doc's trace at
//   the moment the paste happens. Reading it from the other doc's
//   CustomXMLParts is unreliable — unsaved or in-progress docs may not have
//   flushed their trace yet, so the part is empty or stale. The live
//   in-memory state is always authoritative for an open doc.
//
// Two indexes are kept in parallel:
//   - byKey:  keyed on Word.Document.FullName (or a placeholder for unsaved
//             docs, e.g. "Document3"). Used by the paste-detection path,
//             which iterates app.Documents and already has a FullName.
//   - byGuid: keyed on PasteTraceState.DocGuid. Used during recursive chain
//             resolution, which only carries a DocGuid forward from the
//             previous hop's evidence.

using System;
using System.Collections.Generic;

namespace IsItYoursWordAddIn
{
    public static class DocStateRegistry
    {
        private static readonly object _lock = new object();

        private static readonly Dictionary<string, PasteTraceState> _byKey =
            new Dictionary<string, PasteTraceState>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, PasteTraceState> _byGuid =
            new Dictionary<string, PasteTraceState>(StringComparer.Ordinal);

        public static void Register(string docKey, PasteTraceState state)
        {
            if (state == null) return;
            lock (_lock)
            {
                if (!string.IsNullOrEmpty(docKey))
                    _byKey[docKey] = state;
                if (!string.IsNullOrEmpty(state.DocGuid))
                    _byGuid[state.DocGuid] = state;
            }
        }

        // Replace oldKey with newKey in-place — used when an unsaved doc
        // ("Document3") gets a real path ("C:\...\doc3.docx") on save.
        // DocGuid is stable across saves so byGuid is just refreshed.
        public static void Rekey(string oldKey, string newKey, PasteTraceState state)
        {
            if (state == null) return;
            lock (_lock)
            {
                if (!string.IsNullOrEmpty(oldKey) && !string.Equals(oldKey, newKey, StringComparison.OrdinalIgnoreCase))
                    _byKey.Remove(oldKey);
                if (!string.IsNullOrEmpty(newKey))
                    _byKey[newKey] = state;
                if (!string.IsNullOrEmpty(state.DocGuid))
                    _byGuid[state.DocGuid] = state;
            }
        }

        public static void Unregister(string docKey, string docGuid)
        {
            lock (_lock)
            {
                if (!string.IsNullOrEmpty(docKey))  _byKey.Remove(docKey);
                if (!string.IsNullOrEmpty(docGuid)) _byGuid.Remove(docGuid);
            }
        }

        public static PasteTraceState GetByKey(string docKey)
        {
            if (string.IsNullOrEmpty(docKey)) return null;
            lock (_lock)
                return _byKey.TryGetValue(docKey, out var s) ? s : null;
        }

        public static PasteTraceState GetByDocGuid(string docGuid)
        {
            if (string.IsNullOrEmpty(docGuid)) return null;
            lock (_lock)
                return _byGuid.TryGetValue(docGuid, out var s) ? s : null;
        }

        // Snapshot copy — callers iterate without holding the lock.
        public static List<PasteTraceState> AllLiveStates()
        {
            lock (_lock) return new List<PasteTraceState>(_byKey.Values);
        }
    }
}
