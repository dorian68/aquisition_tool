from app.core.config import Settings


def test_cors_origins_accept_plain_star():
    settings = Settings(cors_origins="*")

    assert settings.cors_origin_list == ["*"]


def test_cors_origins_accept_comma_separated_domains():
    settings = Settings(cors_origins="https://app.example.com, https://admin.example.com")

    assert settings.cors_origin_list == ["https://app.example.com", "https://admin.example.com"]
