"""Compatibility entry point for the FastAPI backend."""

import os

import uvicorn
from src.backend.app import app


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "5000")))
