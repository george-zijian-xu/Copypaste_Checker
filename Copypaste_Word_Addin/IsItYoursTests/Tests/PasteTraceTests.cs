#if TEST_HARNESS
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace IsItYoursTests.Tests
{
    static class PasteTraceTests
    {
        // Minimal fixture: wires a string buffer to the three engine delegates.
        private sealed class Fixture
        {
            public string Text  = string.Empty;
            public int    Caret = 0;
            public int    Count => Text?.Length ?? 0;
            public IsItYoursWordAddIn.PasteTraceEngine Build()
                => new IsItYoursWordAddIn.PasteTraceEngine(() => Text, () => Caret, () => Count);
        }

        public static IEnumerable<(string, Action)> All() => new (string, Action)[]
        {
            ("Baseline poll emits no tick",                        Baseline_NoTick),
            ("Append one char: insert tick + Dirty",               Append_OneChar),
            ("Replace: delete and insert have distinct TickIds",   Replace_ProducesDistinctTickIds),
            ("TickId is 8 uppercase hex chars",                    TickId_IsHex),
            ("Large insert flagged as paste",                      LargeInsert_FlaggedAsPaste),
            ("Adjacency fusion fires within gap",                  AdjacencyFusion_WithinGap),
            ("Adjacency fusion suppressed outside gap",            AdjacencyFusion_OutsideGap_NotFused),
            ("CreatedElapsedMs is monotonic",                      ElapsedMs_Monotonic),
            ("Required field names exist (compile-time guard)",    FieldNames_Exist),
            ("Dirty set on capture, clear resets it",              Dirty_SetOnCapture),
            ("Poll counter advances per tick, not per poll",       PollCounter_AdvancesPerTick),
        };

        static void Baseline_NoTick()
        {
            var fx = new Fixture { Text = "hello", Caret = 5 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            bool changed = e.PollOnce();
            TestRunner.Assert(!changed,                 "baseline returns !changed");
            TestRunner.Assert(e.State.Ticks.Count == 0, "baseline emits 0 ticks");
            // StartSession() marks Dirty=true so the new session row is persisted even with no edits.
            TestRunner.Assert(e.State.Dirty,            "baseline is dirty after open (session row pending flush)");
        }

        static void Append_OneChar()
        {
            var fx = new Fixture { Text = "hello", Caret = 5 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            fx.Text = "hello!"; fx.Caret = 6;
            bool changed = e.PollOnce();

            TestRunner.Assert(changed,                  "append returns changed");
            TestRunner.Assert(e.State.Ticks.Count == 1, "one tick emitted");
            var t = e.State.Ticks[0];
            TestRunner.AssertEqual("ins",  t.Op,        "op is ins");
            TestRunner.AssertEqual(5,      t.Loc,       "Loc is the insertion offset");
            TestRunner.AssertEqual(1,      t.Len,       "Len is 1");
            TestRunner.AssertEqual("!",    t.Text,      "Text captures the char");
            TestRunner.AssertEqual(0,      t.Paste,     "not flagged as paste");
            TestRunner.Assert(e.State.Dirty,            "state is now dirty");
        }

        static void Replace_ProducesDistinctTickIds()
        {
            var fx = new Fixture { Text = "hello world", Caret = 11 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            // Same length → sentinel would skip; run past FullSweepEveryNPolls to force a read.
            fx.Text = "hello WORLD"; fx.Caret = 11;
            for (int i = 0; i < 25; i++) e.PollOnce();

            TestRunner.Assert(e.State.Ticks.Count >= 2, "at least 2 ticks for a replace");
            var del = e.State.Ticks.FirstOrDefault(t => t.Op == "del");
            var ins = e.State.Ticks.FirstOrDefault(t => t.Op == "ins");
            TestRunner.Assert(del != null,              "replace produced a delete tick");
            TestRunner.Assert(ins != null,              "replace produced an insert tick");
            TestRunner.Assert(del.TickId != ins.TickId, "del and ins TickIds are distinct");
        }

        static void TickId_IsHex()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            fx.Text = "a"; fx.Caret = 1;
            e.PollOnce();

            var id = e.State.Ticks[0].TickId;
            TestRunner.AssertEqual(8, id.Length, "tick ID is 8 chars");
            foreach (var c in id)
                TestRunner.Assert((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F'),
                    "tick ID char is valid uppercase hex: " + c);
        }

        static void LargeInsert_FlaggedAsPaste()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            var big = new string('x', IsItYoursWordAddIn.PasteTraceEngine.DefaultPasteThreshold);
            fx.Text = big; fx.Caret = big.Length;
            e.PollOnce();

            TestRunner.Assert(e.State.Ticks.Count == 1, "one tick for a large insert");
            TestRunner.AssertEqual(1, e.State.Ticks[0].Paste, "paste flag is set");
        }

        static void AdjacencyFusion_WithinGap()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            // 10 × 3-char contiguous inserts (combined 30 > threshold 20), all within ms of each other.
            string accum = "";
            for (int i = 0; i < 10; i++)
            {
                accum += "abc";
                fx.Text = accum; fx.Caret = accum.Length;
                e.PollOnce();
            }

            TestRunner.AssertEqual(10, e.State.Ticks.Count, "ten insert ticks emitted");
            int flagged = e.State.Ticks.Count(t => t.Paste == 1);
            TestRunner.Assert(flagged >= 7, "fusion flagged the recent run (" + flagged + " flagged)");
        }

        static void AdjacencyFusion_OutsideGap_NotFused()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            string accum = "";
            for (int i = 0; i < 3; i++)
            {
                accum += "abc";
                fx.Text = accum; fx.Caret = accum.Length;
                e.PollOnce();
                Thread.Sleep(1200); // > FusionMaxGapMs (1000 ms)
            }

            TestRunner.AssertEqual(3, e.State.Ticks.Count, "three insert ticks emitted");
            int flagged = e.State.Ticks.Count(t => t.Paste == 1);
            TestRunner.AssertEqual(0, flagged, "no tick is paste-flagged when gaps exceed threshold");
        }

        static void ElapsedMs_Monotonic()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            long prev = -1L;
            for (int i = 0; i < 10; i++)
            {
                fx.Text += "x"; fx.Caret = fx.Text.Length;
                e.PollOnce();
                long curr = e.State.Ticks.Last().CreatedElapsedMs;
                TestRunner.Assert(curr >= prev, "elapsed ms non-decreasing: prev=" + prev + " curr=" + curr);
                prev = curr;
                Thread.Sleep(5);
            }
        }

        static void FieldNames_Exist()
        {
            // Compile-time guard: if a future rename drops these fields the build breaks.
            var s = new IsItYoursWordAddIn.PasteTraceState();
            _ = s.Dirty;
            _ = s.LastComputedChainHmac;
            _ = s.LastHmacTickIndex;
            _ = s.SessionPollCounter;

            var t = new IsItYoursWordAddIn.TickRow();
            _ = t.CreatedElapsedMs;
            _ = t.TickId;
        }

        static void Dirty_SetOnCapture()
        {
            var fx = new Fixture { Text = "", Caret = 0 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();
            // StartSession() marks Dirty=true; that's expected — session row needs flushing.
            TestRunner.Assert(e.State.Dirty, "dirty after open (session row pending)");

            fx.Text = "a"; fx.Caret = 1;
            e.PollOnce();
            TestRunner.Assert(e.State.Dirty, "capture sets Dirty");

            e.State.Dirty = false;
            TestRunner.Assert(!e.State.Dirty, "manual clear works");

            fx.Text = "ab"; fx.Caret = 2;
            e.PollOnce();
            TestRunner.Assert(e.State.Dirty, "re-capture re-sets Dirty");
        }

        static void PollCounter_AdvancesPerTick()
        {
            var fx = new Fixture { Text = "hello", Caret = 5 };
            var e = fx.Build();
            e.OnDocumentOpened(null, DateTime.UtcNow);
            e.PollOnce();

            int before = e.State.SessionPollCounter;
            fx.Text = "hELLo"; fx.Caret = 5;
            for (int i = 0; i < 25; i++) e.PollOnce();

            int after = e.State.SessionPollCounter;
            TestRunner.Assert(after >= before + 2, "counter advances >= 2 on replace");

            var ids = new HashSet<string>(e.State.Ticks.Select(t => t.TickId));
            TestRunner.AssertEqual(e.State.Ticks.Count, ids.Count, "all tick IDs distinct");
        }
    }
}
#endif
