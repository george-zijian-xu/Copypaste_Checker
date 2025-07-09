import os
import sys

# Ensure backend directory is in PYTHONPATH
BACKEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'packages', 'backend'))
if BACKEND_DIR not in sys.path:
    sys.path.append(BACKEND_DIR)

# Import FastAPI and the router directly
from fastapi import FastAPI
from src.api.api import router as api_router
from src.middleware.cors_middleware import add_cors_middleware

# Create a new FastAPI app for Vercel without the /api/v1 prefix
app = FastAPI(
    title="Copy-Paste Checker API",
    description="API for analyzing .docx files for copy-pasted content.",
    version="1.0.0"
)

# Add CORS middleware
add_cors_middleware(app)

# Include the API router without the /api/v1 prefix since Vercel handles that
app.include_router(api_router, prefix="/v1")

@app.get("/", tags=["Root"])
async def read_root():
    return {"message": "Welcome to the Copy-Paste Checker API!"} 