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
    frontend_allowed_origins: str = "http://localhost:5173,https://*.lovable.app"

    database_url: str = "sqlite:///./optiquant.db"

    local_storage_path: Path = Path("storage")
    max_upload_mb: int = 25
    max_columns: int = 100
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

    ai_analyst_enabled: bool = True
    ai_analyst_provider: str = "langgraph"
    ai_analyst_model: str = "gpt-4.1-mini"
    ai_analyst_timeout_seconds: float = 25.0
    max_ai_context_chars: int = 20_000
    max_top_categories: int = 10
    max_trend_points: int = 12
    max_anomalies: int = 5
    max_cleaning_actions_for_ai: int = 20
    max_high_missing_columns_for_ai: int = 10
    ai_analyst_fail_open: bool = True
    openai_api_key: str | None = None

    @property
    def cors_origin_list(self) -> list[str]:
        if "*" in parse_cors_origins(self.cors_origins):
            return ["*"]
        exact = []
        for origin in [*parse_cors_origins(self.cors_origins), *parse_cors_origins(self.frontend_allowed_origins)]:
            if "*" not in origin and origin not in exact:
                exact.append(origin)
        return exact or ["*"]

    @property
    def cors_origin_regex(self) -> str | None:
        wildcard_origins = [origin for origin in parse_cors_origins(self.frontend_allowed_origins) if "*" in origin]
        patterns = []
        for origin in wildcard_origins:
            escaped = re_escape_origin(origin)
            patterns.append(escaped.replace(r"\*", r"[^/]+"))
        return "|".join(f"^{pattern}$" for pattern in patterns) if patterns else None


def re_escape_origin(origin: str) -> str:
    import re

    return re.escape(origin)


@lru_cache
def get_settings() -> Settings:
    return Settings()
