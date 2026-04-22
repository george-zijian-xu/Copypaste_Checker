# Recursive Provenance Chain Integration Tests

The recursive provenance chain (`ResolveChain()` in `Provenance.cs`) traces Word-to-Word paste sources through multiple hops. Testing requires **real Word** and **.docx files with embedded custom XML**.

## Test Scenarios

### 1. Linear Chain: A ← B ← C (source)
**Goal**: Verify that a chain resolves correctly through multiple hops.

- **Doc C** (canon source): Trace with no paste evidence (original typing)
- **Doc B** (intermediate): Trace showing paste from C's text
- **Doc A** (current): Trace showing paste from B's text
- **Expected**: When A's chain is resolved, it walks to B → C, marking both as "resolved"

### 2. Cycle Detection: A ← B ← A
**Goal**: Verify cycle detection prevents infinite loops.

- **Doc A**: Trace showing paste from B
- **Doc B**: Trace showing paste back to A
- **Expected**: Chain resolution detects the cycle, marks second hop to A with `status="cycle-detected"`

### 3. Depth Limit: Chain > 10 Hops
**Goal**: Verify truncation at depth 10 (see `ChainMaxDepth = 10`).

- Create 11 documents: D₀ ← D₁ ← D₂ ← ... ← D₁₀
- **Expected**: Chain walks from D₀ to D₁₀ but stops (≤10 hops)

### 4. Source Unavailable: File Path Invalid
**Goal**: Verify graceful handling when source file is missing.

- **Doc A**: Trace showing paste from `file:///Z:/nonexistent/missing.docx`
- **Expected**: Chain hop marked `status="source-unavailable"`, no crash

### 5. No Trace in Source: Doc Exists But No Custom XML
**Goal**: Verify handling when source is a plain Word doc (no IsItYours trace).

- **Doc A**: Trace showing paste from B
- **Doc B**: Plain .docx with no custom XML part
- **Expected**: Chain hop marked `status="no-trace"`

---

## Manual Testing Steps

### Prerequisites
- **Word 2016 or later** (16.0+ / Office 2016+)
- **C# project** with these assemblies:
  ```csharp
  using Microsoft.Office.Interop.Word;
  using Microsoft.Office.Core;
  using IsItYoursWordAddIn;
  ```

### Setup: Create Test Documents

```csharp
// 1. Create minimal traced documents
static void CreateTestDocs()
{
    var wordApp = new Application { Visible = false };
    string testDir = Path.Combine(Path.GetTempPath(), "IsItYourChainTests");
    Directory.CreateDirectory(testDir);

    // Doc C: Canon source (no paste, just trace)
    var stateC = new PasteTraceState
    {
        DocGuid = "guid-C",
        AppVersion = PasteTraceEngine.AppVersion,
        PasteThreshold = PasteTraceEngine.DefaultPasteThreshold
    };
    stateC.Ticks.Add(new TickRow
    {
        TickId = "000CCCCC",
        Op = "ins",
        Loc = 0,
        Len = 25,
        Text = "Original text from source",
        Paste = 0  // Not a paste
    });
    stateC.BTree.InsertAtEnd("000CCCCC", 0, 25, true);

    // Serialize stateC to Doc C
    WriteDocWithTrace(wordApp, Path.Combine(testDir, "C.docx"), stateC);

    // Doc B: Pastes from C
    var stateB = new PasteTraceState
    {
        DocGuid = "guid-B",
        AppVersion = PasteTraceEngine.AppVersion,
        PasteThreshold = PasteTraceEngine.DefaultPasteThreshold
    };
    stateB.Ticks.Add(new TickRow
    {
        TickId = "000BBBBB",
        Op = "ins",
        Loc = 0,
        Len = 25,
        Text = "Original text from source",
        Paste = 1  // This is a paste
    });
    stateB.BTree.InsertAtEnd("000BBBBB", 0, 25, true);

    // Add paste evidence pointing to C
    var evidenceB = new PasteEvidence
    {
        Origin = "word",
        ClipboardUtc = DateTime.UtcNow,
        ClipboardProcess = "winword.exe",
        Url = "file:///" + Path.Combine(testDir, "C.docx").Replace("\\", "/"),
        FullText = "Original text from source",
        SrcDocGuid = "guid-C",
        SrcFile = "file:///" + Path.Combine(testDir, "C.docx").Replace("\\", "/"),
        SrcAuthor = "TestUser",
        SrcTitle = "Source Doc",
        WasPaste = "yes",
        Origins = new List<(string t, int off, int n)> { ("000BBBBB", 0, 25) }
    };
    stateB._pasteEvidence["000BBBBB"] = evidenceB;

    WriteDocWithTrace(wordApp, Path.Combine(testDir, "B.docx"), stateB);

    // Doc A: Pastes from B (similar pattern)
    // ...

    wordApp.Quit(SaveChanges: false);
}
```

### Test: Resolve Chain

```csharp
static void TestChainResolution()
{
    var wordApp = new Application { Visible = false };
    string pathA = Path.Combine(Path.GetTempPath(), "IsItYourChainTests", "A.docx");

    Document docA = null;
    try
    {
        docA = wordApp.Documents.Open(pathA, ReadOnly: true, AddToRecentFiles: false, Visible: false);

        var state = new PasteTraceState
        {
            AppVersion = PasteTraceEngine.AppVersion,
            PasteThreshold = PasteTraceEngine.DefaultPasteThreshold
        };

        // Hydrate: deserialize custom XML from docA
        PasteTraceXml.TryHydrate(docA, state);

        // Check evidence
        var evidence = state._pasteEvidence.Values.FirstOrDefault(e => e.Origin == "word");
        if (evidence != null)
        {
            Console.WriteLine($"Paste evidence found:");
            Console.WriteLine($"  Origin: {evidence.Origin}");
            Console.WriteLine($"  Source: {evidence.SrcDocGuid}");
            Console.WriteLine($"  Chain hops: {evidence.Chain?.Count ?? 0}");

            if (evidence.Chain != null)
            {
                foreach (var hop in evidence.Chain)
                {
                    Console.WriteLine($"    → Hop: {hop.DocGuid} (status: {hop.Status})");
                }
            }
        }
    }
    finally
    {
        if (docA != null)
            docA.Close(SaveChanges: false);
        wordApp.Quit(SaveChanges: false);
    }
}
```

### Helper: Write State to Document

```csharp
static void WriteDocWithTrace(Application wordApp, string path, PasteTraceState state)
{
    Document doc = null;
    try
    {
        // Create or open
        if (File.Exists(path))
        {
            doc = wordApp.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: false);
        }
        else
        {
            doc = wordApp.Documents.Add();
        }

        // Remove old custom XML
        var existing = doc?.CustomXMLParts?.SelectByNamespace("urn:paste-monitor");
        if (existing?.Count > 0)
        {
            for (int i = existing.Count; i >= 1; i--)
                try { existing[i].Delete(); } catch { }
        }

        // Build and add encrypted XML
        string xml = PasteTraceXml.Build(state);
        doc?.CustomXMLParts?.Add(xml);

        // Set document text
        doc.Content.Text = state.Ticks.FirstOrDefault()?.Text ?? "Test document";

        // Save
        doc?.SaveAs2(path);
    }
    finally
    {
        if (doc != null)
            doc.Close(SaveChanges: true);
    }
}
```

---

## Verification Checklist

- [ ] **Linear chain**: Doc A's chain contains hops to B and C, both marked "resolved"
- [ ] **Cycle detection**: Hop to A from B is marked "cycle-detected"
- [ ] **Depth limit**: Chain truncates at 10 hops (not 11+)
- [ ] **Source unavailable**: Missing file marked "source-unavailable"
- [ ] **No trace**: Plain Word doc marked "no-trace"
- [ ] **No crashes**: All tests complete without exceptions

---

## Key Code Locations

- **Chain resolution**: `Provenance.ResolveChain()` (line 363)
- **Chain max depth**: `ChainMaxDepth = 10` (line 358)
- **Chain serialization**: `PasteTraceXml.Build()` → `<chain>` element (line 152–180)
- **Chain deserialization**: `HydrateInner()` (no chain parsing yet — may need implementation)

## Notes

- Chain is encrypted inside the custom XML part — PasteTraceXml handles encrypt/decrypt
- `ResolveChain()` is called from `Provenance.AttachForPasteTick()` line 327
- Visited DocGuids are tracked to detect cycles (line 324–326)
- Each hop `status` can be: `"resolved"`, `"source-unavailable"`, `"no-trace"`, `"cycle-detected"`
