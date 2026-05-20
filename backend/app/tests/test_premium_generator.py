import zipfile


def _workbook_xml(content: bytes, tmp_path):
    path = tmp_path / "dashboard.xlsx"
    path.write_bytes(content)
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/workbook.xml").decode("utf-8")


def _chart_xml_files(content: bytes, tmp_path):
    path = tmp_path / "dashboard.xlsx"
    path.write_bytes(content)
    with zipfile.ZipFile(path) as archive:
        return [
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/charts/chart") and name.endswith(".xml")
        ]


def _all_xml(content: bytes, tmp_path):
    path = tmp_path / "dashboard.xlsx"
    path.write_bytes(content)
    with zipfile.ZipFile(path) as archive:
        return "\n".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in archive.namelist()
            if name.endswith(".xml")
        )


def test_direct_premium_generator_returns_xlsx(client, sample_csv_bytes, tmp_path):
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
    assert 'name="AI Insights"' in _workbook_xml(response.content, tmp_path)


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


def test_direct_generator_analyze_returns_analysis_json(client, sample_csv_bytes):
    response = client.post(
        "/api/v1/generator/analyze",
        data={"template": "auto", "output_format": "xlsx"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    payload = response.json()
    assert payload["analysis_id"]
    assert payload["recommended_template"] == "dark-saas"
    assert payload["limits"]["raw_rows_sent_to_llm"] == 0
    assert payload["limits"]["ai_context_chars"] <= payload["limits"]["max_ai_context_chars"]
    assert payload["ai_metadata"]["provider"] == "langgraph"
    assert payload["ai_metadata"]["used_fallback"] is True
    assert payload["ai_report"]["dashboard_title"]


def test_dashboard_returns_excel_without_ai_when_disabled(client, sample_csv_bytes, tmp_path):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "dark-saas", "output_format": "xlsx", "include_ai_analysis": "false"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    assert response.content.startswith(b"PK")
    assert 'name="AI Insights"' not in _workbook_xml(response.content, tmp_path)


def test_dashboard_can_reuse_analysis_id(client, sample_csv_bytes, tmp_path):
    analyze = client.post(
        "/api/v1/generator/analyze",
        data={"template": "auto", "output_format": "xlsx"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )
    assert analyze.status_code == 200, analyze.text

    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "dark-saas", "output_format": "xlsx", "analysis_id": analyze.json()["analysis_id"]},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    assert 'name="AI Insights"' in _workbook_xml(response.content, tmp_path)


def test_generated_workbook_has_no_chart_gridlines(client, sample_csv_bytes, tmp_path):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "dark-saas", "output_format": "xlsx", "include_ai_analysis": "false"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    chart_xml_files = _chart_xml_files(response.content, tmp_path)
    assert chart_xml_files
    assert all("majorGridlines" not in chart_xml and "minorGridlines" not in chart_xml for chart_xml in chart_xml_files)


def test_sidebar_includes_ai_insights_nav_button(client, sample_csv_bytes, tmp_path):
    response = client.post(
        "/api/v1/generator/dashboard",
        data={"template": "dark-saas", "output_format": "xlsx", "include_ai_analysis": "true"},
        files={"file": ("sample_sales.csv", sample_csv_bytes, "text/csv")},
    )

    assert response.status_code == 200, response.text
    assert "AI Insights" in _all_xml(response.content, tmp_path)
