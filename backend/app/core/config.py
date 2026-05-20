from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


def parse_cors_origins(value: str) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    app_name: str = "OptiQuant IA Lead Tool Backend"
    environment: str = "local"
    api_v1_prefix: str = "/api/v1"
    cors_origins: str = "*"

    database_url: str = "sqlite:///./optiquant.db"

    local_storage_path: Path = Path("storage")
    max_upload_mb: int = 10
    csv_preview_row_limit: int = 10_000
    generated_file_ttl_days: int = 7

    jwt_secret_key: str = "change-me-in-production"
    jwt_algorithm: str = "HS256"
    jwt_expire_minutes: int = 60 * 24

    google_client_id: str | None = None
    allow_mock_google_auth: bool = False

    n8n_enabled: bool = False
    n8n_webhook_lead_created: str | None = None
    n8n_webhook_dashboard_generated: str | None = None
    n8n_webhook_file_downloaded: str | None = None
    n8n_api_key: str | None = None
    n8n_timeout_seconds: float = 3.0

    @property
    def cors_origin_list(self) -> list[str]:
        return parse_cors_origins(self.cors_origins)


@lru_cache
def get_settings() -> Settings:
    return Settings()
