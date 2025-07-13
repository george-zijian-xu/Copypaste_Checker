/**
 * This file centralizes all API communication for the frontend.
 * It provides typed functions for interacting with the backend services.
 */
import { AnalysisResult } from "@/types";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "";

/**
 * Uploads a .docx file to get a copy-paste analysis.
 * @param file The .docx file to analyze.
 * @returns A promise that resolves to the analysis result.
 */
export async function analyzeDocxForCopying(file: File): Promise<AnalysisResult> {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(`${API_BASE_URL}/api/v1/analysis/copy-check`, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.detail || "Failed to analyze document.");
  }

  return response.json();
}

/**
 * Uploads a .docx file to extract its raw text content.
 * @param file The .docx file to process.
 * @returns A promise that resolves to an object containing the source text.
 */
export async function extractTextFromDocx(file: File): Promise<{ sourceText: string }> {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(`${API_BASE_URL}/api/v1/documents/extract-text`, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.detail || "Failed to extract text.");
  }

  return response.json();
} 