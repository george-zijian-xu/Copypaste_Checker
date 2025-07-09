import os
import sys

# Ensure backend directory is in PYTHONPATH
BACKEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'packages', 'backend'))
if BACKEND_DIR not in sys.path:
    sys.path.append(BACKEND_DIR)

# Import the FastAPI app from the backend server module
from server import app  # Vercel will look for this 'app' object 