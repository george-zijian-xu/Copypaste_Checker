// WordStubs.cs — minimal stubs for Word interop types, used only in the test harness.
// The real add-in uses the actual Microsoft.Office.Interop.Word assembly.
#if TEST_HARNESS

namespace Microsoft.Office.Interop.Word
{
    // Minimal stub — PasteTraceEngine only uses Document as an opaque handle (FullName property).
    public class Document
    {
        public string FullName { get; set; } = "(TestDoc)";
        public bool ReadOnly { get; set; } = false;
    }
}

namespace Microsoft.Office.Core
{
    public interface CustomXMLPart { string XML { get; } void Delete(); }
    public interface CustomXMLParts
    {
        int Count { get; }
        CustomXMLPart this[int index] { get; }
        CustomXMLParts SelectByNamespace(string ns);
        void Add(string xml);
    }
}

#endif
