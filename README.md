# GraphPilot AI — Phase 1

A deployment-ready foundation for a premium AI document intelligence platform.

## Phase 1 scope
- Premium dark SaaS dashboard shell
- Responsive navigation and overview cards
- FastAPI backend with health and configuration endpoints
- React + TypeScript + Vite frontend
- Docker-based deployment for Render
- Basic backend tests
- Environment-variable-based configuration placeholders

## Not included yet
Document ingestion, RAG, embeddings, authentication, Google Drive OAuth, AI chat, report generation, PDF/PPTX export, and organization-level access controls are intentionally scheduled for later phases.

## Requirements
- Docker
- Python 3.11+ (for backend-only development)
- Node.js 20+ (for frontend-only development)

## Run with Docker
```bash
docker build -t graphpilot-ai .
docker run --rm -p 8000:8000 graphpilot-ai
```
Open `http://localhost:8000`.

## Run backend without Docker
```bash
cd backend
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload
```

## Run frontend without Docker
```bash
cd frontend
npm install
npm run dev
```
The Vite development server runs on `http://localhost:5173`. The production frontend is served by FastAPI in the Docker deployment.

## Render deployment
1. Push the contents of this phase to a GitHub repository.
2. In Render, create a new **Blueprint** and connect the repository, or create a new Web Service using the Docker runtime.
3. If using the Blueprint, Render reads `render.yaml`.
4. Confirm the service health check path is `/api/health`.
5. After deployment, open the Render URL.

No API keys are required in Phase 1. Do not add real secrets to GitHub.

## Environment variables
See `.env.example`. The model IDs are configuration placeholders for future provider integration and are not called in Phase 1.

## Health endpoints
- `GET /api/health`
- `GET /api/config`

## Testing
```bash
cd backend
pip install -r requirements.txt
pytest
```

## Security notes
- Never commit `.env` files or provider API keys.
- Phase 1 is a UI and deployment foundation, not a production-secure multi-tenant system.
- Do not upload confidential company documents until authentication, authorization, tenant isolation, and storage controls are implemented and reviewed.
