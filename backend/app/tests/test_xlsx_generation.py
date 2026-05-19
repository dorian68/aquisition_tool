from app.tests.conftest import auth_headers, create_dashboard


def test_generate_xlsx_requires_auth(client, sample_csv_bytes):
    dashboard_id = create_dashboard(client, sample_csv_bytes)
    response = client.post(f"/api/v1/dashboards/{dashboard_id}/generate-xlsx")
    assert response.status_code == 401


def test_generate_and_download_xlsx(client, sample_csv_bytes):
    dashboard_id = create_dashboard(client, sample_csv_bytes)
    headers = auth_headers(client)

    response = client.post(f"/api/v1/dashboards/{dashboard_id}/generate-xlsx", headers=headers)
    assert response.status_code == 200, response.text
    file_id = response.json()["file_id"]

    download = client.get(f"/api/v1/files/{file_id}/download", headers=headers)
    assert download.status_code == 200, download.text
    assert download.content.startswith(b"PK")
    assert "spreadsheetml.sheet" in download.headers["content-type"]

