using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;


namespace IsItYoursWordAddIn
{
    // Minimal B+ tree sequence for piece spans.
    // - Leaves store Pieces in order; internals store child pointers + subtree visible lengths.
    // - All navigation is by VISIBLE document offsets (sum of piece.Len where piece.Visible == true).
    // - Splits preserve Piece.OffsetInTick (tick-local) invariants:
    //     right.OffsetInTick = left.OffsetInTick + left.Len_split
    //
    // Degree settings chosen to avoid frequent splits in typical essays.
    public sealed class BTreeSeq
    {
        private const int MaxLeafItems = 64;
        private const int MaxInnerArity = 16;

        private abstract class Node
        {
            public int VisibleLen; // total visible chars under this node
            public abstract bool IsLeaf { get; }
        }

        private sealed class Leaf : Node
        {
            public readonly List<Piece> Items = new List<Piece>(8);
            public Leaf Next, Prev;
            public override bool IsLeaf => true;
        }

        private sealed class Inner : Node
        {
            public readonly List<Node> Children = new List<Node>(4);
            public override bool IsLeaf => false;
        }

        private Node _root = new Leaf();

        // Public surface --------------

        // Ensure a boundary at visible offset "at": split piece if needed,
        // return the leaf+index position AFTER the boundary (as a linearized index)
        public int EnsureBoundaryAt(int at)
        {
            at = Math.Max(0, at);
            // Traverse to leaf by visible length
            var path = new Stack<(Inner parent, int childIndex)>();
            Leaf leaf = DescendToLeaf(at, path, out int atInLeaf);
            // Align inside leaf by splitting a piece if boundary falls inside it
            AlignLeafBoundary(leaf, atInLeaf, ref at);
            // Return a linearized index (for debugging or optional use)
            return -1; // not used by callers; tree is mutated in-place
        }

        public void InsertSpan(int at, Piece piece)
        {
            var path = new Stack<(Inner parent, int childIndex)>();
            Leaf leaf = DescendToLeaf(at, path, out int atInLeaf);
            AlignLeafBoundary(leaf, atInLeaf, ref at);

            // Insert piece at boundary (leaf index computed inside AlignLeafBoundary)
            int insertIdx = FindIndexInLeafByVisibleOffset(leaf, atInLeaf);
            leaf.Items.Insert(insertIdx, piece);

            // Update visible lengths upward
            int delta = piece.Visible ? piece.Len : 0;
            UpdateVisibleUpward(leaf, delta, path);

            // Split leaf if overflow
            if (leaf.Items.Count > MaxLeafItems) SplitLeaf(leaf, path);
        }

        public void TombstoneRange(int start, int len)
        {
            if (len <= 0) return;
            int end = start + len;

            var path = new Stack<(Inner parent, int childIndex)>();
            Leaf leaf = DescendToLeaf(start, path, out int atInLeafStart);
            AlignLeafBoundary(leaf, atInLeafStart, ref start);
            // tombstone until we reach end; we may cross leaves
            while (start < end)
            {
                int before = start;
                int tombstoned = TombstoneFromLeaf(leaf, start, end);
                if (tombstoned == 0)
                {
                    // Move to next leaf
                    if (leaf.Next == null) break;
                    leaf = leaf.Next;
                    continue;
                }
                // Adjust visible counts upward
                UpdateVisibleUpward(leaf, -tombstoned, path);
                start += tombstoned;
                if (start < end && leaf.Next != null)
                {
                    // descend again from current start into next leaf/path
                    path.Clear();
                    leaf = DescendToLeaf(start, path, out _);
                }
                if (start == before) break; // safety
            }
        }

        public IEnumerable<Piece> EnumeratePieces()
        {
            // in-order via leaf chain
            Leaf l = LeftmostLeaf();
            while (l != null)
            {
                foreach (var p in l.Items) yield return p;
                l = l.Next;
            }
        }

        // Snapshot structure for XML
        public BTreeSnapshot Snapshot()
        {
            var snap = new BTreeSnapshot();
            DumpSnapshot(_root, snap, parentId: -1, myId: 0);
            return snap;
        }

        public void LoadSnapshot(BTreeSnapshot snap)
        {
            // Build node objects for each snapshot node id
            var map = new Dictionary<int, Node>();

            // 1) Create nodes (leaves get their pieces now; inners get children later)
            foreach (var sn in snap.Nodes)
            {
                if (sn.IsLeaf)
                {
                    var leaf = new Leaf();
                    foreach (var lp in sn.LeafPieces)
                    {
                        leaf.Items.Add(new Piece
                        {
                            TickId = lp.TickId,
                            OffsetInTick = lp.Off,        // immutable tick-local off
                            Len = lp.Len,
                            Visible = lp.Vis == 1
                        });
                    }
                    leaf.VisibleLen = VisibleLenOf(leaf.Items);
                    map[sn.Id] = leaf;
                }
                else
                {
                    var inner = new Inner();
                    inner.VisibleLen = 0; // will compute after wiring children
                    map[sn.Id] = inner;
                }
            }

            // 2) Find root id (parent == -1)
            int rootId = -1;
            foreach (var sn in snap.Nodes)
                if (sn.ParentId == -1) { rootId = sn.Id; break; }
            if (rootId == -1)
            {
                // fallback to a single empty leaf if snapshot is malformed
                _root = new Leaf();
                return;
            }

            // 3) Wire children for inner nodes in id order (DumpSnapshot numbers left-to-right)
            foreach (var sn in snap.Nodes)
            {
                if (sn.IsLeaf) continue;
                var parent = (Inner)map[sn.Id];
                var children = snap.Nodes
                    .Where(ch => ch.ParentId == sn.Id)
                    .OrderBy(ch => ch.Id)
                    .Select(ch => map[ch.Id]);
                foreach (var ch in children) parent.Children.Add(ch);
            }

            // 4) Set root
            _root = map[rootId];

            // 5) Compute VisibleLen bottom-up
            ComputeVisibleLens(_root);

            // 6) Link leaf chain (Prev/Next) by in-order traversal
            var leaves = new List<Leaf>();
            CollectLeavesInOrder(_root, leaves);
            for (int i = 0; i < leaves.Count; i++)
            {
                leaves[i].Prev = i > 0 ? leaves[i - 1] : null;
                leaves[i].Next = i + 1 < leaves.Count ? leaves[i + 1] : null;
            }
        }

        private int ComputeVisibleLens(Node n)
        {
            if (n.IsLeaf) return n.VisibleLen; // already set on load
            var inner = (Inner)n;
            int sum = 0;
            foreach (var ch in inner.Children) sum += ComputeVisibleLens(ch);
            inner.VisibleLen = sum;
            return sum;
        }

        private void CollectLeavesInOrder(Node n, List<Leaf> acc)
        {
            if (n.IsLeaf) { acc.Add((Leaf)n); return; }
            var inner = (Inner)n;
            foreach (var ch in inner.Children) CollectLeavesInOrder(ch, acc);
        }


        // Internals --------------

        private Leaf LeftmostLeaf()
        {
            var n = _root;
            while (!n.IsLeaf) n = ((Inner)n).Children[0];
            return (Leaf)n;
        }

        private Leaf DescendToLeaf(int at, Stack<(Inner, int)> path, out int atInLeaf)
        {
            Node n = _root;
            int remaining = at;
            while (!n.IsLeaf)
            {
                var inner = (Inner)n;
                int i = 0, acc = 0;
                for (; i < inner.Children.Count; i++)
                {
                    int childVis = inner.Children[i].VisibleLen;
                    if (remaining <= acc + childVis) break;
                    acc += childVis;
                }
                if (i == inner.Children.Count) i = inner.Children.Count - 1;
                path.Push((inner, i));
                n = inner.Children[i];
                remaining -= Math.Min(remaining, inner.Children[i].VisibleLen);
            }
            atInLeaf = remaining;
            return (Leaf)n;
        }

        private void AlignLeafBoundary(Leaf leaf, int atInLeaf, ref int atGlobal)
        {
            // Walk items in leaf summing visible lengths to find boundary; if inside a visible piece, split it
            int cum = 0;
            for (int i = 0; i < leaf.Items.Count; i++)
            {
                var p = leaf.Items[i];
                int visLen = p.Visible ? p.Len : 0;
                if (cum + visLen == atInLeaf) return; // already aligned before item i
                if (cum + visLen > atInLeaf)
                {
                    if (!p.Visible)
                    {
                        // boundary falls ��inside�� an invisible piece �� we consider it aligned between pieces
                        return;
                    }
                    // Split p into left/right by visible count; since piece is fully visible,
                    // the split is simply by char count.
                    int leftLen = atInLeaf - cum;
                    if (leftLen <= 0 || leftLen >= p.Len) return;
                    var left = new Piece
                    {
                        TickId = p.TickId,
                        OffsetInTick = p.OffsetInTick,
                        Len = leftLen,
                        Visible = true
                    };
                    var right = new Piece
                    {
                        TickId = p.TickId,
                        OffsetInTick = p.OffsetInTick + leftLen, // *** immutable tick offset rule ***
                        Len = p.Len - leftLen,
                        Visible = true
                    };
                    leaf.Items[i] = left;
                    leaf.Items.Insert(i + 1, right);
                    // no visible count change overall; boundary now lies between i and i+1
                    return;
                }
                cum += visLen;
            }
            // If we reach here, boundary is at end of leaf; OK
        }

        private int FindIndexInLeafByVisibleOffset(Leaf leaf, int atInLeaf)
        {
            int cum = 0;
            for (int i = 0; i < leaf.Items.Count; i++)
            {
                var p = leaf.Items[i];
                int vis = p.Visible ? p.Len : 0;

                if (cum == atInLeaf) return i;     // insert before item i
                cum += vis;

                if (cum == atInLeaf) return i + 1; // insert after item i
                if (cum > atInLeaf) return i + 1; // boundary lies inside (AlignLeafBoundary has split)
            }
            return leaf.Items.Count;
        }

        private void UpdateVisibleUpward(Node start, int delta, Stack<(Inner, int)> path)
        {
            start.VisibleLen += delta;
            foreach (var (parent, idx) in path)
                parent.VisibleLen = SumVisible(parent.Children);
        }

        private static int SumVisible(List<Node> kids)
        {
            int s = 0;
            foreach (var c in kids) s += c.VisibleLen;
            return s;
        }

        private void SplitLeaf(Leaf leaf, Stack<(Inner parent, int childIndex)> path)
        {
            int mid = leaf.Items.Count / 2;
            var right = new Leaf();
            for (int i = mid; i < leaf.Items.Count; i++) right.Items.Add(leaf.Items[i]);
            leaf.Items.RemoveRange(mid, leaf.Items.Count - mid);

            right.Next = leaf.Next; if (right.Next != null) right.Next.Prev = right;
            leaf.Next = right; right.Prev = leaf;

            leaf.VisibleLen = VisibleLenOf(leaf.Items);
            right.VisibleLen = VisibleLenOf(right.Items);

            if (path.Count == 0)
            {
                // new root
                var root = new Inner();
                root.Children.Add(leaf);
                root.Children.Add(right);
                root.VisibleLen = leaf.VisibleLen + right.VisibleLen;
                _root = root;
                return;
            }

            var (parent, idx) = path.Pop();
            parent.Children.Insert(idx + 1, right);
            parent.VisibleLen = SumVisible(parent.Children);

            if (parent.Children.Count > MaxInnerArity) SplitInner(parent, path);
        }

        private void SplitInner(Inner inner, Stack<(Inner parent, int childIndex)> path)
        {
            int mid = inner.Children.Count / 2;
            var right = new Inner();
            for (int i = mid; i < inner.Children.Count; i++) right.Children.Add(inner.Children[i]);
            inner.Children.RemoveRange(mid, inner.Children.Count - mid);

            inner.VisibleLen = SumVisible(inner.Children);
            right.VisibleLen = SumVisible(right.Children);

            if (path.Count == 0)
            {
                var root = new Inner();
                root.Children.Add(inner);
                root.Children.Add(right);
                root.VisibleLen = inner.VisibleLen + right.VisibleLen;
                _root = root;
                return;
            }

            var (parent, idx) = path.Pop();
            parent.Children.Insert(idx + 1, right);
            parent.VisibleLen = SumVisible(parent.Children);

            if (parent.Children.Count > MaxInnerArity) SplitInner(parent, path);
        }

        private static int VisibleLenOf(List<Piece> items)
        {
            int s = 0; foreach (var p in items) if (p.Visible) s += p.Len; return s;
        }

        private int TombstoneFromLeaf(Leaf leaf, int startVisible, int endVisible)
        {
            // Tombstone as many visible chars as lie in [startVisible, endVisible) within this leaf.
            int cum = 0; int affected = 0;
            for (int i = 0; i < leaf.Items.Count; i++)
            {
                var p = leaf.Items[i];
                int vis = p.Visible ? p.Len : 0;
                if (vis == 0) continue;

                int next = cum + vis;
                if (next <= startVisible) { cum = next; continue; }
                if (cum >= endVisible) break;

                // We need exact alignment; split at leaf level so we hide whole pieces
                int localStart = Math.Max(0, startVisible - cum);
                int localEnd = Math.Min(vis, endVisible - cum);

                if (localStart > 0 && localStart < p.Len)
                {
                    // split left | right
                    var left = new Piece { TickId = p.TickId, OffsetInTick = p.OffsetInTick, Len = localStart, Visible = true };
                    var right = new Piece { TickId = p.TickId, OffsetInTick = p.OffsetInTick + localStart, Len = p.Len - localStart, Visible = true };
                    leaf.Items[i] = left;
                    leaf.Items.Insert(i + 1, right);
                    p = right; // continue on right
                    i++;       // advance to right
                    vis = p.Len;
                    cum += localStart;
                    // localEnd was computed relative to the original piece; adjust it
                    // so it is now relative to the start of the new right piece.
                    localEnd -= localStart;
                }

                if (localEnd < p.Len)
                {
                    // split p into mid (to tombstone) + tail (visible)
                    int midLen = localEnd - 0; // from start of p now
                    if (midLen > 0)
                    {
                        var mid = new Piece { TickId = p.TickId, OffsetInTick = p.OffsetInTick, Len = midLen, Visible = false };
                        var tail = new Piece { TickId = p.TickId, OffsetInTick = p.OffsetInTick + midLen, Len = p.Len - midLen, Visible = true };
                        leaf.Items[i] = mid;
                        leaf.Items.Insert(i + 1, tail);
                        affected += midLen;
                        cum += midLen;
                        continue;
                    }
                }
                else
                {
                    // whole piece is within range �� just tombstone it
                    p.Visible = false;
                    affected += p.Len;
                    cum = next;
                }
            }
            return affected;
        }

        // ----- Snapshot (structure) -----

        private void DumpSnapshot(Node n, BTreeSnapshot snap, int parentId, int myId)
        {
            if (n.IsLeaf)
            {
                var leaf = (Leaf)n;
                var entry = new BTreeSnapshot.Node
                {
                    Id = myId,
                    ParentId = parentId,
                    IsLeaf = true,
                    VisibleLen = leaf.VisibleLen
                };
                foreach (var p in leaf.Items)
                    entry.LeafPieces.Add(new BTreeSnapshot.LeafPiece { TickId = p.TickId, Off = p.OffsetInTick, Len = p.Len, Vis = p.Visible ? 1 : 0 });
                snap.Nodes.Add(entry);
                return;
            }

            var inner = (Inner)n;
            var node = new BTreeSnapshot.Node
            {
                Id = myId,
                ParentId = parentId,
                IsLeaf = false,
                VisibleLen = inner.VisibleLen
            };
            snap.Nodes.Add(node);
            for (int i = 0; i < inner.Children.Count; i++)
                DumpSnapshot(inner.Children[i], snap, myId, snap.NextId());
        }
    }

    // Lightweight serializable snapshot of the B-tree
    public sealed class BTreeSnapshot
    {
        public sealed class LeafPiece { public string TickId; public int Off; public int Len; public int Vis; }
        public sealed class Node
        {
            public int Id;
            public int ParentId;
            public bool IsLeaf;
            public int VisibleLen;
            public List<LeafPiece> LeafPieces = new List<LeafPiece>(); // only for leaves
        }


        public readonly List<Node> Nodes = new List<Node>();
        private int _nextId = 1; public int NextId() => _nextId++;
    }


}
