import os
import sys

# Ensure backend directory is in PYTHONPATH
BACKEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'packages', 'backend'))
if BACKEND_DIR not in sys.path:
    sys.path.append(BACKEND_DIR)

# Import FastAPI and the router directly
from fastapi import FastAPI

try:
    from src.api.api import router as api_router
    from src.middleware.cors_middleware import add_cors_middleware
    IMPORTS_SUCCESSFUL = True
except ImportError as e:
    print(f"Import error: {e}")
    # Create a dummy router for testing
    from fastapi import APIRouter
    api_router = APIRouter()
    IMPORTS_SUCCESSFUL = False
    
    @api_router.get("/api/v1/analysis/copy-check")
    async def dummy_endpoint():
        return {"error": "Backend dependencies not available", "import_error": str(e)}

# Create a new FastAPI app for Vercel
app = FastAPI(
    title="Copy-Paste Checker API",
    description="API for analyzing .docx files for copy-pasted content.",
    version="1.0.0"
)

# Add CORS middleware
if IMPORTS_SUCCESSFUL:
    add_cors_middleware(app)
else:
    # Add basic CORS manually
    from fastapi.middleware.cors import CORSMiddleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

# Include the API router with /api/v1 prefix to match the frontend calls
app.include_router(api_router, prefix="/api/v1")

@app.get("/", tags=["Root"])
async def read_root():
    return {"message": "Welcome to the Copy-Paste Checker API!"}

@app.get("/api/test", tags=["Test"])
async def test_endpoint():
    return {
        "status": "API is working!", 
        "python_version": sys.version,
        "imports_successful": IMPORTS_SUCCESSFUL,
        "backend_dir": BACKEND_DIR,
        "sys_path": sys.path[:3]  # First 3 entries only
    } 