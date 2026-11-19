from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)

def test_health_endpoint():
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"
    assert response.json()["phase"] == 1

def test_public_config_does_not_expose_secrets():
    response = client.get("/api/config")
    assert response.status_code == 200
    payload = response.json()
    assert "GEMINI_API_KEY" not in payload
    assert "GROQ_API_KEY" not in payload
