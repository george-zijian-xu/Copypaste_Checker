using System;
using System.Collections.Generic;
using System.Linq;
using IsItYoursWordAddIn;

namespace IsItYoursTests.Tests
{
    static class BTreeTests
    {
        public static IEnumerable<(string, Action)> All() => new (string, Action)[]
        {
            ("Insert single span, enumerate returns it", InsertSingle),
            ("Insert two spans, enumerate in order", InsertTwo),
            ("Tombstone removes chars from visible count", TombstoneReducesVisible),
            ("Tombstone then enumerate skips invisible pieces", TombstoneEnumerate),
            ("Insert at middle splits correctly", InsertAtMiddle),
            ("Snapshot round-trips via LoadSnapshot", SnapshotRoundTrip),
            ("Large insert triggers leaf split", LargeInsertSplit),
        };

        static void InsertSingle()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 5, Visible = true });
            var pieces = tree.EnumeratePieces().ToList();
            TestRunner.AssertEqual(1, pieces.Count, "piece count");
            TestRunner.AssertEqual("t1", pieces[0].TickId, "tickId");
            TestRunner.AssertEqual(5, pieces[0].Len, "len");
        }

        static void InsertTwo()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 3, Visible = true });
            tree.InsertSpan(3, new Piece { TickId = "t2", OffsetInTick = 0, Len = 4, Visible = true });
            var pieces = tree.EnumeratePieces().Where(p => p.Visible).ToList();
            TestRunner.AssertEqual(2, pieces.Count, "piece count");
            TestRunner.AssertEqual("t1", pieces[0].TickId, "first tickId");
            TestRunner.AssertEqual("t2", pieces[1].TickId, "second tickId");
        }

        static void TombstoneReducesVisible()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 10, Visible = true });
            tree.TombstoneRange(2, 3); // remove 3 chars starting at offset 2
            int visible = tree.EnumeratePieces().Where(p => p.Visible).Sum(p => p.Len);
            TestRunner.AssertEqual(7, visible, "visible len after tombstone");
        }

        static void TombstoneEnumerate()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 5, Visible = true });
            tree.TombstoneRange(0, 5);
            var visible = tree.EnumeratePieces().Where(p => p.Visible).ToList();
            TestRunner.AssertEqual(0, visible.Count, "no visible pieces after full tombstone");
        }

        static void InsertAtMiddle()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 6, Visible = true });
            // Insert in the middle at visible offset 3
            tree.InsertSpan(3, new Piece { TickId = "t2", OffsetInTick = 0, Len = 2, Visible = true });
            var pieces = tree.EnumeratePieces().Where(p => p.Visible).ToList();
            // Should be: t1[0..3], t2[0..2], t1[3..6]
            TestRunner.Assert(pieces.Count >= 2, "at least 2 pieces after mid-insert");
            int totalVisible = pieces.Sum(p => p.Len);
            TestRunner.AssertEqual(8, totalVisible, "total visible len = 6 + 2");
        }

        static void SnapshotRoundTrip()
        {
            var tree = new BTreeSeq();
            tree.InsertSpan(0, new Piece { TickId = "t1", OffsetInTick = 0, Len = 4, Visible = true });
            tree.InsertSpan(4, new Piece { TickId = "t2", OffsetInTick = 0, Len = 3, Visible = true });
            tree.TombstoneRange(1, 2);

            var snap = tree.Snapshot();

            var tree2 = new BTreeSeq();
            tree2.LoadSnapshot(snap);

            var orig = tree.EnumeratePieces().ToList();
            var restored = tree2.EnumeratePieces().ToList();

            TestRunner.AssertEqual(orig.Count, restored.Count, "piece count after snapshot round-trip");
            for (int i = 0; i < orig.Count; i++)
            {
                TestRunner.AssertEqual(orig[i].TickId, restored[i].TickId, $"piece[{i}].TickId");
                TestRunner.AssertEqual(orig[i].Len, restored[i].Len, $"piece[{i}].Len");
                TestRunner.AssertEqual(orig[i].Visible, restored[i].Visible, $"piece[{i}].Visible");
            }
        }

        static void LargeInsertSplit()
        {
            // Insert 70 pieces (> MaxLeafItems=64) to trigger a leaf split
            var tree = new BTreeSeq();
            int offset = 0;
            for (int i = 0; i < 70; i++)
            {
                tree.InsertSpan(offset, new Piece { TickId = $"t{i:000}", OffsetInTick = 0, Len = 2, Visible = true });
                offset += 2;
            }
            var pieces = tree.EnumeratePieces().ToList();
            TestRunner.AssertEqual(70, pieces.Count, "all 70 pieces present after split");
            int totalVisible = pieces.Where(p => p.Visible).Sum(p => p.Len);
            TestRunner.AssertEqual(140, totalVisible, "total visible = 70 * 2");
        }
    }
}
