from app.core.config import Settings


def test_cors_origins_accept_plain_star():
    settings = Settings(cors_origins="*")

    assert settings.cors_origin_list == ["*"]


def test_cors_origins_accept_comma_separated_domains():
    settings = Settings(cors_origins="https://app.example.com, https://admin.example.com", frontend_allowed_origins="")

    assert settings.cors_origin_list == ["https://app.example.com", "https://admin.example.com"]


def test_lovable_preview_origin_uses_regex():
    settings = Settings(cors_origins="", frontend_allowed_origins="http://localhost:5173,https://*.lovable.app")

    assert settings.cors_origin_list == ["http://localhost:5173"]
    assert settings.cors_origin_regex is not None
    assert "lovable\\.app" in settings.cors_origin_regex
