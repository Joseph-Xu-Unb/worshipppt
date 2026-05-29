"""Compatibility entry point for the FastAPI backend."""

import uvicorn
from src.backend.app import app


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=5000)
