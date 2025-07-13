/**
 * Defines the core data structures used throughout the frontend.
 */

export interface Highlight {
  start: number;
  end: number;
  category: string;
  note: string;
  rsid: string;
}

export interface AnalysisResult {
  documentId: string;
  sourceText: string;
  highlights: Highlight[];
} 