import os
import shutil
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

os.environ.setdefault("DATABASE_URL", "sqlite:///./test_optiquant.db")
os.environ.setdefault("LOCAL_STORAGE_PATH", "./test_storage")
os.environ.setdefault("ALLOW_MOCK_GOOGLE_AUTH", "true")
os.environ.setdefault("JWT_SECRET_KEY", "test-secret")
os.environ.setdefault("N8N_ENABLED", "false")

from app.main import app  # noqa: E402
from app.models.database import Base, engine, init_db  # noqa: E402


@pytest.fixture()
def reset_state():
    Base.metadata.drop_all(bind=engine)
    init_db()
    storage_path = Path("./test_storage")
    if storage_path.exists():
        shutil.rmtree(storage_path)
    storage_path.mkdir(parents=True, exist_ok=True)
    yield
    if storage_path.exists():
        shutil.rmtree(storage_path)


@pytest.fixture()
def client(reset_state):
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture()
def sample_csv_bytes() -> bytes:
    return (
        b"Date,Revenue,Product,Customer,Region,Status,Quantity\n"
        b"2026-01-05,12000,Alpha Console,Northwind,North America,Paid,12\n"
        b"2026-01-12,18000,Beta Suite,Contoso,Europe,Paid,18\n"
        b"2026-02-03,24000,Alpha Console,Fabrikam,North America,Paid,20\n"
        b"2026-03-02,31000,Beta Suite,Adventure Works,Asia,Paid,28\n"
    )


def upload_sample(client: TestClient, sample_csv_bytes: bytes) -> str:
    response = client.post(
        "/api/v1/uploads/csv",
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )
    assert response.status_code == 201, response.text
    return response.json()["upload_id"]


def create_dashboard(client: TestClient, sample_csv_bytes: bytes) -> str:
    upload_id = upload_sample(client, sample_csv_bytes)
    response = client.post(f"/api/v1/uploads/{upload_id}/dashboard-spec")
    assert response.status_code == 200, response.text
    return response.json()["dashboard_id"]


def auth_headers(client: TestClient) -> dict[str, str]:
    response = client.post("/api/v1/auth/google", json={"id_token": "mock:test@example.com|Test User"})
    assert response.status_code == 200, response.text
    return {"Authorization": f"Bearer {response.json()['access_token']}"}

