"""
Vercel-compatible API entry point for the Copy-Paste Checker backend.
This file exposes the FastAPI application as required by Vercel's Python runtime.
"""
import sys
import os

# Add the packages/backend directory to the Python path
backend_path = os.path.join(os.path.dirname(__file__), '..', 'packages', 'backend')
sys.path.insert(0, backend_path)

# Import the FastAPI application
from server import app

# Vercel expects the ASGI application to be available as 'app'
# The FastAPI app is already configured in server.py 