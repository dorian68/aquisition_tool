from app.tests.conftest import create_dashboard, upload_sample


def test_dashboard_spec_generated(client, sample_csv_bytes):
    upload_id = upload_sample(client, sample_csv_bytes)
    response = client.post(f"/api/v1/uploads/{upload_id}/dashboard-spec")
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["dashboard_type"] == "sales_performance"
    assert body["kpis"]
    assert body["charts"]


def test_preview_payload_valid(client, sample_csv_bytes):
    dashboard_id = create_dashboard(client, sample_csv_bytes)
    response = client.get(f"/api/v1/dashboards/{dashboard_id}/preview")
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["dashboard_spec"]["dashboard_id"] == dashboard_id
    assert body["preview_data"]["charts"]
    assert body["preview_data"]["sample_rows"]

