from pathlib import Path
from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    app_name: str = "GraphPilot AI"
    app_env: str = "development"
    gemini_model: str = "gemini-3.8-flash"
    gemini_fallback_model: str = "gemini-3.5-flash-lite"
    groq_model: str = "openai/gpt-oss-120b"
    groq_fallback_model: str = "openai/gpt-oss-20b"
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

settings = Settings()
app = FastAPI(title=settings.app_name, version="0.1.0")

@app.get("/api/health")
def health():
    return {"status": "ok", "service": "graphpilot-ai", "phase": 1}

@app.get("/api/config")
def public_config():
    # Never expose API keys or secret values through this endpoint.
    return {
        "app_name": settings.app_name,
        "environment": settings.app_env,
        "phase": 1,
        "ai_providers_configured": False,
        "message": "AI provider integration is scheduled for a later phase."
    }

STATIC_DIR = Path(__file__).resolve().parents[1] / "static"
if STATIC_DIR.exists():
    app.mount("/assets", StaticFiles(directory=STATIC_DIR / "assets"), name="assets")

    @app.get("/{full_path:path}")
    def serve_frontend(full_path: str):
        requested = STATIC_DIR / full_path
        if full_path and requested.is_file():
            return FileResponse(requested)
        return FileResponse(STATIC_DIR / "index.html")
