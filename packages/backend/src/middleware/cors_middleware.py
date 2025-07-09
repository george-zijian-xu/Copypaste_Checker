"""
This module configures the Cross-Origin Resource Sharing (CORS) middleware.
"""
from fastapi.middleware.cors import CORSMiddleware

# Define the origins that are allowed to make requests to this API.
# In a production environment, you would restrict this to your frontend's domain.
# For local development, we allow the default Next.js port.
origins = [
    "http://localhost",
    "http://localhost:3000",
    "http://localhost:3001",
    '*'
]

def add_cors_middleware(app):
    """
    Adds CORS middleware to the FastAPI application.
    This allows the frontend to communicate with the backend.
    """
    app.add_middleware(
        CORSMiddleware,
        allow_origins=origins,
        allow_credentials=True,
        allow_methods=["*"],  # Allows all methods (GET, POST, etc.)
        allow_headers=["*"],  # Allows all headers
    ) 