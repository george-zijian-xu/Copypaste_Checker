"""
Vercel-compatible API entry point for the Copy-Paste Checker backend.
This file exposes a simple FastAPI application as required by Vercel's Python runtime.
"""
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import sys
import os
import tempfile
import shutil
import logging

# Add the backend directory to the Python path
backend_path = os.path.join(os.path.dirname(__file__), '..', 'backend')
sys.path.insert(0, backend_path)

# Import backend services
try:
    from services.copy_check.copy_check_service import analyze_document_for_copying
    from services.document_service import get_document_text
except ImportError as e:
    print(f"Import error: {e}")
    # Fallback - create dummy functions for now
    def analyze_document_for_copying(file_path):
        return {"analysis": "Copy-paste analysis not available", "error": str(e)}
    def get_document_text(file_path):
        return {"text": "Text extraction not available", "error": str(e)}

# Create FastAPI app
app = FastAPI(
    title="Copy-Paste Checker API",
    description="API for analyzing .docx files for copy-pasted content.",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def _save_upload_file_tmp(upload_file: UploadFile):
    """Save uploaded file to temporary location."""
    try:
        suffix = ".docx" if upload_file.filename.endswith('.docx') else ""
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            shutil.copyfileobj(upload_file.file, tmp_file)
            tmp_file_path = tmp_file.name
        return tmp_file_path
    except Exception as e:
        logger.error(f"Error saving uploaded file: {e}")
        raise HTTPException(status_code=500, detail="Error saving uploaded file")

@app.get("/")
async def root():
    """Root endpoint to confirm API is working."""
    return {"message": "Copy-Paste Checker API is running!"}

@app.post("/api/v1/analysis/copy-check")
async def copy_check_analysis(file: UploadFile = File(...)):
    """Analyze document for copy-paste detection."""
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only .docx files are supported")
    
    tmp_file_path = None
    try:
        tmp_file_path = _save_upload_file_tmp(file)
        result = analyze_document_for_copying(tmp_file_path)
        return result
    except Exception as e:
        logger.error(f"Error in copy-check analysis: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if tmp_file_path and os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)

@app.post("/api/v1/documents/extract-text")
async def extract_text(file: UploadFile = File(...)):
    """Extract text from document."""
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only .docx files are supported")
    
    tmp_file_path = None
    try:
        tmp_file_path = _save_upload_file_tmp(file)
        result = get_document_text(tmp_file_path)
        return result
    except Exception as e:
        logger.error(f"Error in text extraction: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if tmp_file_path and os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path) 