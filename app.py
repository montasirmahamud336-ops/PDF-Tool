from __future__ import annotations

"""ASGI entrypoint for Vercel, Render, Railway, Docker, and local Uvicorn."""

import importlib.util
from pathlib import Path


def _load_pdf_tool_app():
    module_path = Path(__file__).with_name("pdf-tool.py")
    spec = importlib.util.spec_from_file_location("pdf_tool_runtime", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load FastAPI app from {module_path}.")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.app


app = _load_pdf_tool_app()

