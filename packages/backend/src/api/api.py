"""
This module defines the API endpoints for the document analysis services.
"""
import tempfile
import shutil
import os
import logging
from fastapi import APIRouter, UploadFile, File, HTTPException

try:
    from ..services.copy_check.copy_check_service import analyze_document_for_copying
    from ..services.document_service import get_document_text
except (ImportError, ValueError):
    # Fallback for direct testing
    import sys
    import os
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
    from src.services.copy_check.copy_check_service import analyze_document_for_copying
    from src.services.document_service import get_document_text

# Get a logger for this module
logger = logging.getLogger(__name__)

# Create a router to group the endpoints
router = APIRouter()

def _save_upload_file_tmp(upload_file: UploadFile):
    """
    Saves an uploaded file to a temporary file and returns the path.
    The caller is responsible for cleaning up the temporary file.
    """
    try:
        # Create a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=upload_file.filename) as tmp:
            # Write the uploaded file's content to the temporary file
            shutil.copyfileobj(upload_file.file, tmp)
            tmp_path = tmp.name
    finally:
        # Make sure to close the file object from the upload
        upload_file.file.close()
    return tmp_path

@router.post("/analysis/copy-check", tags=["Analysis"])
async def run_copy_check_analysis(file: UploadFile = File(...)):
    """
    Accepts a .docx file and returns an analysis of suspected copy-pasted text.
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .docx file.")

    logger.info(f"Starting copy-check analysis for file: {file.filename}")
    temp_file_path = None
    try:
        temp_file_path = _save_upload_file_tmp(file)
        logger.info(f"File saved to temporary path: {temp_file_path}")
        
        analysis_result = analyze_document_for_copying(temp_file_path)

        if analysis_result is None:
            logger.error("Analysis failed, service returned None.")
            raise HTTPException(status_code=500, detail="Failed to analyze the document.")
        
        logger.info(f"Analysis successful for file: {file.filename}")
        return analysis_result
    except Exception as e:
        logger.error(f"An unexpected error occurred during copy-check analysis for {file.filename}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"An internal server error occurred: {e}")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)
            logger.info(f"Cleaned up temporary file: {temp_file_path}")

@router.post("/documents/extract-text", tags=["Documents"])
async def extract_text_from_document(file: UploadFile = File(...)):
    """
    Accepts a .docx file and returns its raw text content.
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .docx file.")
    
    logger.info(f"Starting text extraction for file: {file.filename}")
    temp_file_path = None
    try:
        temp_file_path = _save_upload_file_tmp(file)
        logger.info(f"File saved to temporary path: {temp_file_path}")

        text_content = get_document_text(temp_file_path)

        if text_content is None:
            logger.error("Text extraction failed, service returned None.")
            raise HTTPException(status_code=500, detail="Failed to extract text from the document.")
        
        logger.info(f"Text extraction successful for file: {file.filename}")
        return {"sourceText": text_content}
    except Exception as e:
        logger.error(f"An unexpected error occurred during text extraction for {file.filename}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"An internal server error occurred: {e}")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)
            logger.info(f"Cleaned up temporary file: {temp_file_path}") 