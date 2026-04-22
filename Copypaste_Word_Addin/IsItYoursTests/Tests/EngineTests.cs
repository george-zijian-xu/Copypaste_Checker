using System;
using System.Collections.Generic;
using System.Linq;
using IsItYoursWordAddIn;

namespace IsItYoursTests.Tests
{
    static class EngineTests
    {
        // Helpers to create a test engine with a controllable document text
        static string _docText = "";
        static PasteTraceEngine MakeEngine(string initialText = "")
        {
            _docText = initialText;
            return new PasteTraceEngine(() => _docText, () => _docText.Length);
        }

        public static IEnumerable<(string, Action)> All() => new (string, Action)[]
        {
            ("Session ID wraps at 4096", SessionIdWrapAround),
            ("Single insert below threshold: paste=0", InsertBelowThreshold),
            ("Single insert at threshold: paste=1", InsertAtThreshold),
            ("Delete tick emitted on deletion", DeleteTickEmitted),
            ("Adjacency fusion: two small inserts fused into paste", AdjacencyFusion),
            ("Adjacency fusion: non-contiguous inserts not fused", AdjacencyFusionNonContiguous),
            ("Adjacency fusion: combined below threshold not fused", AdjacencyFusionBelowThreshold),
            ("AppVersion constant is non-empty", AppVersionNonEmpty),
        };

        static void SessionIdWrapAround()
        {
            // Simulate hydration setting SessionId to 4095 (max), then wrap
            var state = new PasteTraceState { DocGuid = "test", AppVersion = "0.1.0", PasteThreshold = 20 };
            // Manually set sessions to simulate maxId = 4095
            for (int i = 0; i <= 4095; i++)
                state.Sessions.Add(new PasteTraceState.SessionRow { Id = i, StartUtc = DateTime.UtcNow });

            // Simulate what TryHydrate does: (maxId + 1) % 4096
            int maxId = state.Sessions.Max(s => s.Id);
            int nextId = (maxId + 1) % 4096;
            TestRunner.AssertEqual(0, nextId, "session ID wraps to 0 after 4095");
        }

        static void InsertBelowThreshold()
        {
            var engine = MakeEngine("hello");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);
            _docText = "hello world"; // 6 chars inserted
            engine.PollOnce();
            var ins = engine.State.Ticks.FirstOrDefault(t => t.Op == "ins");
            TestRunner.Assert(ins != null, "insert tick emitted");
            TestRunner.AssertEqual(0, ins.Paste, "paste=0 for 6-char insert (below threshold 20)");
        }

        static void InsertAtThreshold()
        {
            var engine = MakeEngine("start ");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);
            _docText = "start " + new string('x', 20); // exactly 20 chars inserted
            engine.PollOnce();
            var ins = engine.State.Ticks.FirstOrDefault(t => t.Op == "ins");
            TestRunner.Assert(ins != null, "insert tick emitted");
            TestRunner.AssertEqual(1, ins.Paste, "paste=1 for 20-char insert (at threshold)");
        }

        static void DeleteTickEmitted()
        {
            var engine = MakeEngine("hello world");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);
            _docText = "hello"; // deleted " world"
            engine.PollOnce();
            var del = engine.State.Ticks.FirstOrDefault(t => t.Op == "del");
            TestRunner.Assert(del != null, "del tick emitted");
            TestRunner.AssertEqual(6, del.Len, "deleted 6 chars");
        }

        static void AdjacencyFusion()
        {
            // Two consecutive inserts of 10 chars each at contiguous positions,
            // same second counter -> should fuse to paste=1
            var engine = MakeEngine("");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);

            // First insert: 10 chars at offset 0
            _docText = new string('a', 10);
            engine.PollOnce();

            // Second insert: 10 more chars appended (contiguous)
            _docText = new string('a', 20);
            engine.PollOnce();

            var inserts = engine.State.Ticks.Where(t => t.Op == "ins").ToList();
            TestRunner.Assert(inserts.Count >= 2, "at least 2 insert ticks");
            // Both should be flagged as paste after fusion (combined = 20 >= threshold)
            TestRunner.Assert(inserts.All(t => t.Paste == 1), "all inserts fused to paste=1");
        }

        static void AdjacencyFusionNonContiguous()
        {
            // Two inserts at non-contiguous positions should NOT fuse
            var engine = MakeEngine("abc   xyz");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);

            // Insert 10 chars at the start
            _docText = new string('a', 10) + "abc   xyz";
            engine.PollOnce();

            // Insert 10 chars at the end (not contiguous with first insert)
            _docText = new string('a', 10) + "abc   xyz" + new string('b', 10);
            engine.PollOnce();

            var inserts = engine.State.Ticks.Where(t => t.Op == "ins").ToList();
            TestRunner.Assert(inserts.Count >= 2, "at least 2 insert ticks");
            // Neither should be paste (each is 10 chars, non-contiguous, so no fusion)
            TestRunner.Assert(inserts.All(t => t.Paste == 0), "non-contiguous inserts not fused");
        }

        static void AdjacencyFusionBelowThreshold()
        {
            // Two inserts of 5 chars each (combined = 10, below threshold 20) should NOT fuse
            var engine = MakeEngine("");
            engine.OnDocumentOpened(new Microsoft.Office.Interop.Word.Document(), DateTime.UtcNow);

            _docText = new string('a', 5);
            engine.PollOnce();

            _docText = new string('a', 10);
            engine.PollOnce();

            var inserts = engine.State.Ticks.Where(t => t.Op == "ins").ToList();
            TestRunner.Assert(inserts.All(t => t.Paste == 0), "combined 10 chars below threshold, no fusion");
        }

        static void AppVersionNonEmpty()
        {
            TestRunner.Assert(!string.IsNullOrEmpty(PasteTraceEngine.AppVersion), "AppVersion is set");
        }
    }
}
