"""
This service provides general document-related functionalities,
such as extracting raw text content from a .docx file.
"""
import os
import tempfile
import shutil
import sys

try:
    # This import works when the service is called from other parts of the package
    from ..utils.unzip_docx import unzip_docx
    from ..utils.docx_to_text import extract_text
except (ImportError, ValueError):
    # This is a fallback for direct script execution for testing
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
    from src.utils.unzip_docx import unzip_docx
    from src.utils.docx_to_text import extract_text


def get_document_text(docx_path: str):
    """
    Extracts the plain text content from a .docx file.

    It handles the process of unzipping the file to a temporary location
    and then running the text extraction utility.

    Args:
        docx_path: The absolute path to the .docx file.

    Returns:
        A string containing the full text of the document, or None if
        the file cannot be processed.
    """
    if not os.path.isfile(docx_path):
        print(f"Error: File not found at '{docx_path}'", file=sys.stderr)
        return None

    # Create a temporary directory to extract the .docx contents
    temp_dir = tempfile.mkdtemp()
    try:
        if not unzip_docx(docx_path, temp_dir):
            return None  # Unzip failed

        document_xml_path = os.path.join(temp_dir, "word", "document.xml")
        if not os.path.exists(document_xml_path):
            print("Error: Could not find 'word/document.xml' in the unzipped file.", file=sys.stderr)
            return None

        # Use the dedicated utility to get the clean text
        source_text = extract_text(document_xml_path)
        return source_text

    finally:
        # Clean up the temporary directory
        shutil.rmtree(temp_dir)

if __name__ == '__main__':
    # Example of how to run this service directly for testing
    if len(sys.argv) != 2:
        print("Usage: python document_service.py <path_to_docx>")
        sys.exit(1)
    
    test_docx_path = sys.argv[1]
    text_content = get_document_text(test_docx_path)

    if text_content is not None:
        print(text_content)
    else:
        print("Failed to extract text from the document.") 