from dataclasses import dataclass

from app.core.config import Settings, get_settings


class GoogleAuthError(ValueError):
    pass


@dataclass(frozen=True)
class GoogleUserInfo:
    email: str
    name: str | None
    provider_user_id: str


class GoogleTokenVerifier:
    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()

    def verify(self, id_token_value: str) -> GoogleUserInfo:
        if self.settings.allow_mock_google_auth and id_token_value.startswith("mock:"):
            return self._mock_user(id_token_value)

        try:
            from google.auth.transport import requests
            from google.oauth2 import id_token
        except ImportError as exc:
            raise GoogleAuthError("google-auth dependency is not installed") from exc

        try:
            audience = self.settings.google_client_id
            payload = id_token.verify_oauth2_token(id_token_value, requests.Request(), audience)
        except Exception as exc:  # noqa: BLE001
            raise GoogleAuthError("Invalid Google ID token") from exc

        email = payload.get("email")
        if not email:
            raise GoogleAuthError("Google token does not contain an email")

        return GoogleUserInfo(
            email=email,
            name=payload.get("name"),
            provider_user_id=payload.get("sub", email),
        )

    @staticmethod
    def _mock_user(id_token_value: str) -> GoogleUserInfo:
        raw = id_token_value.removeprefix("mock:")
        parts = raw.split("|")
        email = parts[0] if parts and parts[0] else "mock.user@example.com"
        name = parts[1] if len(parts) > 1 and parts[1] else "Mock User"
        provider_user_id = parts[2] if len(parts) > 2 and parts[2] else f"mock-{email}"
        return GoogleUserInfo(email=email, name=name, provider_user_id=provider_user_id)

