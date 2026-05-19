import pandas as pd

from app.services.dataset_profiler import DatasetProfiler
from app.tests.conftest import upload_sample


def test_profiler_detects_sales_columns():
    df = pd.DataFrame(
        {
            "Date": ["2026-01-01", "2026-02-01"],
            "Revenue": [1000, 2000],
            "Product": ["A", "B"],
            "Status": ["Paid", "Pending"],
        }
    )
    profile = DatasetProfiler().profile(df)
    semantics = {column["name"]: column["semantic_type"] for column in profile["columns"]}
    assert semantics["Revenue"] == "money"
    assert semantics["Product"] == "product"
    assert "Date" in profile["detected_date_columns"]
    assert profile["quality_score"] > 90


def test_analyze_endpoint_returns_dataset_profile(client, sample_csv_bytes):
    upload_id = upload_sample(client, sample_csv_bytes)
    response = client.post(f"/api/v1/uploads/{upload_id}/analyze")
    assert response.status_code == 200, response.text
    profile = response.json()["dataset_profile"]
    assert profile["row_count"] == 4
    assert "Revenue" in profile["detected_metrics"]

