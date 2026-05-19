from app.tests.conftest import upload_sample


def test_upload_valid_csv(client, sample_csv_bytes):
    upload_id = upload_sample(client, sample_csv_bytes)
    assert upload_id


def test_empty_csv_refused(client):
    response = client.post(
        "/api/v1/uploads/csv",
        files={"file": ("empty.csv", b"", "text/csv")},
    )
    assert response.status_code == 400
    assert "empty" in response.json()["detail"].lower()


def test_bad_format_refused(client):
    response = client.post(
        "/api/v1/uploads/csv",
        files={"file": ("data.txt", b"a,b\n1,2\n", "text/plain")},
    )
    assert response.status_code == 400
    assert "csv" in response.json()["detail"].lower()

