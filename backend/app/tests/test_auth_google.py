def test_google_auth_mock_creates_user_and_token(client):
    response = client.post(
        "/api/v1/auth/google",
        json={"id_token": "mock:lead@example.com|Lead User", "utm_source": "test"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["user"]["email"] == "lead@example.com"
    assert body["access_token"]

