using System;

namespace IsItYoursWordAddIn
{
    // Compatibility wrapper. Older patches called TraceDebugLog.Add(...), while the
    // current state model stores logs as PasteTraceState.DebugRow via state.Log(...).
    public sealed class TraceDebugRow
    {
        public DateTime Utc;
        public string Phase;
        public string Event;
        public string DocGuid;
        public string DocKey;
        public string TickId;
        public string SourceGuid;
        public string SourceFile;
        public string Result;
        public string Detail;
    }

    public static class TraceDebugLog
    {
        public static void Add(
            PasteTraceState state,
            string phase,
            string ev,
            string result = null,
            string detail = null,
            string tickId = null,
            string docKey = null,
            string sourceGuid = null,
            string sourceFile = null)
        {
            if (state == null) return;
            string data = "docKey=" + (docKey ?? "") +
                          ";sourceGuid=" + (sourceGuid ?? "") +
                          ";sourceFile=" + (sourceFile ?? "") +
                          ";result=" + (result ?? "") +
                          ";detail=" + (detail ?? "");
            state.Log(phase ?? "", ev ?? "", tickId, data);
        }

        public static string E(string s) { return System.Security.SecurityElement.Escape(s) ?? string.Empty; }
    }
}
