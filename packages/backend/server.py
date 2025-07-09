"""
Main backend server entry point.
This script initializes the FastAPI application, configures middleware,
and includes the API routers from other modules.
"""
from fastapi import FastAPI

try:
    from .src.api.api import router as api_router
    from .src.middleware.cors_middleware import add_cors_middleware
    from .src.config.logging_config import setup_logging
except ImportError:
    # This allows the server to be run from the root of the project for local dev
    from src.api.api import router as api_router
    from src.middleware.cors_middleware import add_cors_middleware
    from src.config.logging_config import setup_logging

# Configure logging before initializing the app
setup_logging()

# Initialize the main FastAPI application
app = FastAPI(
    title="Copy-Paste Checker API",
    description="API for analyzing .docx files for copy-pasted content.",
    version="1.0.0"
)

# Add CORS middleware to allow cross-origin requests from the frontend
add_cors_middleware(app)

# Include the main API router
# All routes from api.py will be prefixed with /api/v1
app.include_router(api_router, prefix="/api/v1")

@app.get("/", tags=["Root"])
async def read_root():
    """
    A simple root endpoint to confirm the API is running.
    """
    return {"message": "Welcome to the Copy-Paste Checker API!"} 