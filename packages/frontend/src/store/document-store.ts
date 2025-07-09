import { AnalysisResult } from "@/types";
import { create } from "zustand";

export type DocumentStatus = "idle" | "uploading" | "analyzing" | "done" | "error";

interface DocumentState {
  file: File | null;
  progress: number;
  status: DocumentStatus;
  error: string | null;
  analysisResult: AnalysisResult | null;
  highlightedHtml: string | null;
  setFile: (file: File | null) => void;
  setProgress: (progress: number) => void;
  setStatus: (status: DocumentStatus) => void;
  setError: (error: string | null) => void;
  setAnalysisResult: (result: AnalysisResult | null) => void;
  setHighlightedHtml: (html: string | null) => void;
  reset: () => void;
}

export const useDocumentStore = create<DocumentState>((set) => ({
  file: null,
  progress: 0,
  status: "idle",
  error: null,
  analysisResult: null,
  highlightedHtml: null,
  setFile: (file) => set({ file }),
  setProgress: (progress) => set({ progress }),
  setStatus: (status) => set({ status }),
  setError: (error) => set({ error }),
  setAnalysisResult: (result) => set({ analysisResult: result }),
  setHighlightedHtml: (html) => set({ highlightedHtml: html }),
  reset: () =>
    set({
      file: null,
      progress: 0,
      status: "idle",
      error: null,
      analysisResult: null,
      highlightedHtml: null,
    }),
})); 