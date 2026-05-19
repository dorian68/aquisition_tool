import logging
from typing import Any

import httpx

from app.core.config import Settings, get_settings
from app.models.database import SessionLocal
from app.models.event import N8nWebhookLog

logger = logging.getLogger(__name__)


class N8nClient:
    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()

    def send_event(self, event_type: str, payload: dict[str, Any]) -> None:
        url = self._url_for(event_type)
        if not self.settings.n8n_enabled or not url:
            self._log(event_type, payload, None, "n8n disabled or webhook missing")
            return

        try:
            response = self._post(url, payload)
            self._log(event_type, payload, response.status_code, response.text[:1000])
        except Exception as exc:  # noqa: BLE001 - webhook failures must not block users.
            logger.warning("n8n webhook failed event_type=%s error=%s", event_type, exc)
            self._log(event_type, payload, None, str(exc))

    def _post(self, url: str, payload: dict[str, Any]) -> httpx.Response:
        headers = {"Content-Type": "application/json"}
        if self.settings.n8n_api_key:
            headers["Authorization"] = f"Bearer {self.settings.n8n_api_key}"
        with httpx.Client(timeout=self.settings.n8n_timeout_seconds) as client:
            return client.post(url, json=payload, headers=headers)

    def _url_for(self, event_type: str) -> str | None:
        return {
            "lead_created": self.settings.n8n_webhook_lead_created,
            "dashboard_generated": self.settings.n8n_webhook_dashboard_generated,
            "file_downloaded": self.settings.n8n_webhook_file_downloaded,
        }.get(event_type)

    @staticmethod
    def _log(event_type: str, payload: dict[str, Any], status: int | None, body: str | None) -> None:
        db = SessionLocal()
        try:
            db.add(
                N8nWebhookLog(
                    event_type=event_type,
                    payload_json=payload,
                    response_status=status,
                    response_body=body,
                )
            )
            db.commit()
        finally:
            db.close()

