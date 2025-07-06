import { create } from "zustand";

export type DocumentStatus = "idle" | "uploading" | "analyzing" | "done" | "error";

interface DocumentState {
  file: File | null;
  progress: number;
  status: DocumentStatus;
  error: string | null;
  setFile: (file: File | null) => void;
  setProgress: (progress: number) => void;
  setStatus: (status: DocumentStatus) => void;
  setError: (error: string | null) => void;
  reset: () => void;
}

export const useDocumentStore = create<DocumentState>((set) => ({
  file: null,
  progress: 0,
  status: "idle",
  error: null,
  setFile: (file) => set({ file }),
  setProgress: (progress) => set({ progress }),
  setStatus: (status) => set({ status }),
  setError: (error) => set({ error }),
  reset: () =>
    set({ file: null, progress: 0, status: "idle", error: null }),
})); 