from app.core.config import Settings
from app.models.database import SessionLocal
from app.models.event import N8nWebhookLog
from app.services.n8n_client import N8nClient


class FakeResponse:
    status_code = 200
    text = "ok"


def test_n8n_webhook_mock_logged(reset_state, monkeypatch):
    settings = Settings(
        n8n_enabled=True,
        n8n_webhook_lead_created="https://n8n.example.test/webhook",
        database_url="sqlite:///./test_optiquant.db",
    )
    client = N8nClient(settings=settings)
    monkeypatch.setattr(client, "_post", lambda url, payload: FakeResponse())

    client.send_event("lead_created", {"event": "lead_created"})

    db = SessionLocal()
    try:
        logs = db.query(N8nWebhookLog).all()
        assert len(logs) == 1
        assert logs[0].response_status == 200
    finally:
        db.close()

