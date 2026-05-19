def test_direct_premium_generator_returns_xlsx(client, sample_csv_bytes):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "auto", "output_format": "xlsx", "client_name": "Demo SaaS"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    assert response.content.startswith(b"PK")
    assert "spreadsheetml.sheet" in response.headers["content-type"]
    assert response.headers["x-dashboard-template"] == "dark-saas"
    assert "dashboard-dark-saas.xlsx" in response.headers["content-disposition"]


def test_direct_premium_generator_forced_template(client, sample_csv_bytes):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "light-consulting", "output_format": "xlsx"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    assert response.headers["x-dashboard-template"] == "light-consulting"


def test_direct_premium_generator_xlsm_requires_vba_project(client, sample_csv_bytes):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "dark-saas", "output_format": "xlsm"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 400
    assert "vbaProject.bin" in response.json()["detail"]
