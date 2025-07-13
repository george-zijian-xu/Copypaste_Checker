"""
Service layer responsible for detecting suspected copy-pasted sections
in a .docx file using run-level RSID analysis.
"""

import os
import sys
import tempfile
import shutil
import uuid
import logging
from collections import defaultdict

# Get a logger for this module
logger = logging.getLogger(__name__)

# Lazy-import the analysis and unzip utilities while supporting direct script execution
try:
    # Import the local analysis module (same package) so we get the correct return signature
    from .run_level_rsid_analysis import analyze_document_runs
    from ...utils.unzip_docx import unzip_docx
except (ImportError, ValueError):
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
    from services.copy_check.run_level_rsid_analysis import analyze_document_runs
from utils.unzip_docx import unzip_docx


def _filter_and_transform_highlights(raw_highlights):
    """
    Filters raw analysis highlights to identify suspected copy-pasted text.

    A text segment is considered a "suspected copy" if the analysis found either:
    1. Exactly one reason: "No property change detected".
    2. Exactly one reason: "Formatting: Font Hint Change (...)".
    
    Args:
        raw_highlights: A list of highlight dictionaries from the analysis script.

    Returns:
        A list of transformed highlight dictionaries that match the criteria.
    """
    # Group highlights by the text span they refer to (start, end, rsid)
    grouped_reasons = defaultdict(list)
    for h in raw_highlights:
        key = (h['start'], h['end'], h['rsid'])
        grouped_reasons[key].append(h['category'])
    
    suspected_highlights = []
    for (start, end, rsid), reasons in grouped_reasons.items():
        note = None
        # Criteria 1: Only "No property change detected"
        if len(reasons) == 1 and reasons[0] == "No property change detected":
            note = "No property change detected"
        
        # Criteria 2: Only a "Font Hint Change"
        elif len(reasons) == 1 and reasons[0].startswith("Formatting: Font Hint Change"):
            note = reasons[0]

        if note:
            suspected_highlights.append({
                "start": start,
                "end": end,
                "category": "suspected_copy",
                "note": note,
                "rsid": rsid
            })
            
    return suspected_highlights

def analyze_document_for_copying(docx_path: str):
    """
    Main service function to analyze a .docx file for copy-pasted content.

    It unzips the file, runs a detailed run-level analysis, filters the results
    to find likely copied text, and returns a structured JSON-like object.

    Args:
        docx_path: The absolute path to the .docx file.

    Returns:
        A dictionary containing the full source text and a list of highlights
        for suspected copied content, or None if the file cannot be processed.
    """
    if not os.path.isfile(docx_path):
        logger.error(f"File not found at '{docx_path}'")
        return None

    # Create a temporary directory to extract the .docx contents
    temp_dir = tempfile.mkdtemp()
    logger.info(f"Created temporary directory for analysis: {temp_dir}")
    try:
        logger.info(f"Unzipping '{docx_path}' to '{temp_dir}'")
        if not unzip_docx(docx_path, temp_dir):
            logger.error(f"Failed to unzip '{docx_path}'")
            return None # Unzip failed

        document_xml_path = os.path.join(temp_dir, "word", "document.xml")
        if not os.path.exists(document_xml_path):
            logger.error(f"Could not find 'word/document.xml' in '{temp_dir}'")
            return None

        # The analysis script now returns both the text and the highlights in one call.
        logger.info(f"Running run-level analysis on '{document_xml_path}'")
        source_text, raw_highlights = analyze_document_runs(document_xml_path)
        logger.info(f"Analysis completed. Found {len(raw_highlights)} raw highlights.")

        # Filter the raw highlights to find only the ones we suspect are copies.
        suspected_highlights = _filter_and_transform_highlights(raw_highlights)
        logger.info(f"Filtered highlights. Found {len(suspected_highlights)} suspected copies.")

        return {
            "documentId": str(uuid.uuid4()),
            "sourceText": source_text,
            "highlights": suspected_highlights
        }
    except Exception as e:
        logger.error(f"An error occurred during the analysis of '{docx_path}': {e}", exc_info=True)
        return None # Explicitly return None on failure
    finally:
        # Clean up the temporary directory
        logger.info(f"Removing temporary directory: {temp_dir}")
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # This block is for direct script execution testing.
    # It expects a single argument: the path to a .docx file.
    if len(sys.argv) != 2:
        print("Usage: python copy_check_service.py <path_to_docx_file>")
        sys.exit(1)

    docx_path = sys.argv[1]
    result = analyze_document_for_copying(docx_path)

    if result:
        print("\n--- Copy Check Results ---")
        print(f"Document ID: {result['documentId']}")
        print(f"Source Text Length: {len(result['sourceText'])} characters")
        print(f"Number of Suspected Copies: {len(result['highlights'])}")
        print("\n--- Suspected Copies ---")
        for h in result['highlights']:
            print(f"RSID: {h['rsid']}, Category: {h['category']}, Note: {h['note']}, Span: {h['start']}-{h['end']}")
    else:
        print("Copy check failed or file not found.")