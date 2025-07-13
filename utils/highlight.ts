/**
 * This file contains utility functions for processing and displaying analysis results.
 */
import { Highlight } from "@/types";

/**
 * Escapes special HTML characters in a string to prevent XSS attacks.
 * @param str The string to escape.
 * @returns The escaped string.
 */
function escapeHtml(str: string) {
  return str.replace(/[&<>"']/g, (m) => {
    switch (m) {
      case '&': return '&amp;';
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '"': return '&quot;';
      default: return '&#039;';
    }
  });
}

/**
 * Applies highlight ranges to a plain text string, wrapping the specified
 * sections in <span. tags with a given CSS class.
 *
 * @param text The source text.
 * @param ranges An array of highlight objects with start/end offsets.
 * @param inlineStyle The inline style to apply to the highlight spans.
 * @returns A single HTML string with highlights applied.
 */
export function applyHighlights(
  text: string,
  ranges: Highlight[],
  // Use a hardcoded inline style to guarantee visibility.
  inlineStyle = "background-color: #F7E379; border-radius: 3px; padding: 0 2px;"
): string {
  if (!ranges || ranges.length === 0) {
    return escapeHtml(text);
  }

  // --- DEBUGGING: Log inputs ---
  console.log("Applying highlights. Text length:", text.length, "Number of ranges:", ranges.length);

  // Sort ranges to process them in order, preventing nesting issues.
  const sortedRanges = [...ranges].sort((a, b) => a.start - b.start);

  let lastIndex = 0;
  const parts = [];

  for (const range of sortedRanges) {
    // --- DEBUGGING: Log each range ---
    console.log(`Processing range: start=${range.start}, end=${range.end}, lastIndex=${lastIndex}`);

    // Add the text part before the current highlight
    if (range.start > lastIndex) {
      parts.push(escapeHtml(text.slice(lastIndex, range.start)));
    }
    // Add the highlighted part
    const highlightedText = escapeHtml(text.slice(range.start, range.end));
    
    // --- DEBUGGING: Log highlighted text ---
    console.log(`Highlighting text: "${highlightedText}"`);

    parts.push(
      `<span style="${inlineStyle}">${highlightedText}</span>`
    );
    lastIndex = range.end;
  }

  // Add any remaining text after the last highlight
  if (lastIndex < text.length) {
    parts.push(escapeHtml(text.slice(lastIndex)));
  }

  return parts.join("");
} 