/**
 * This hook orchestrates the entire document analysis process,
 * from file submission to state management.
 */
import { useDocumentStore } from "@/store/document-store";
import { analyzeDocxForCopying } from "@/lib/api";
import { applyHighlights } from "@/utils/highlight";

export function useDocumentAnalysis() {
  const { setStatus, setAnalysisResult, setHighlightedHtml, setError, setProgress } = useDocumentStore();

  const handleSubmit = async (file: File) => {
    if (!file) {
      setError("Please select a file first.");
      return;
    }

    try {
      setStatus("uploading");
      setHighlightedHtml(null); // Clear previous results
      setProgress(25);

      const analysisResult = await analyzeDocxForCopying(file);
      setProgress(75);
      setStatus("analyzing");

      // --- DEBUGGING: Log the raw API response ---
      console.log("Received analysis result from backend:", analysisResult);
      
      const html = applyHighlights(analysisResult.sourceText, analysisResult.highlights);
      setAnalysisResult(analysisResult);
      setHighlightedHtml(html);
      
      setProgress(100);
      setStatus("done");

    } catch (error: any) {
      console.error("Analysis failed:", error);
      setError(error.message || "An unknown error occurred.");
      setStatus("error");
      setProgress(0);
    }
  };

  return { handleSubmit };
}
