import json

from app.core.config import Settings
from app.services.ai_analyst import (
    AIContextLimits,
    build_ai_context_json,
    build_deterministic_ai_report,
    generate_ai_dashboard_report_result,
    prepare_python_analysis,
    serialized_context_length,
)


def test_build_ai_context_never_contains_raw_rows(tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)

    analysis = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")
    serialized = json.dumps(analysis["ai_context"], ensure_ascii=True)

    assert "2026-01-05,12000,Alpha Console,Northwind" not in serialized
    assert "raw_df" not in serialized
    assert "clean_df" not in serialized


def test_ai_context_respects_max_list_lengths(tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)
    analysis = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")

    context = build_ai_context_json(
        analysis["cleaning_result"],
        analysis["spec"],
        analysis["aggregations"],
        analysis["selected_template"],
        "xlsx",
        AIContextLimits(max_ai_context_chars=20_000, max_top_categories=2, max_trend_points=2, max_anomalies=1),
    )

    assert len(context["business_metrics"]["top_categories"]) <= 2
    assert len(context["business_metrics"]["secondary_distribution"]) <= 2
    assert len(context["business_metrics"]["trend"]) <= 2
    assert len(context["anomalies"]) <= 1
    assert len(context["cleaning_actions"]) <= 20


def test_fallback_report_is_valid(tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)
    context = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")["ai_context"]

    report = build_deterministic_ai_report(context)

    assert report["dashboard_title"]
    assert report["executive_summary"]
    assert isinstance(report["key_insights"], list)
    assert isinstance(report["recommended_actions"], list)
    assert isinstance(report["risk_notes"], list)


def test_langgraph_service_falls_back_when_openai_key_missing(tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)
    context = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")["ai_context"]

    result = generate_ai_dashboard_report_result(context, settings=Settings(openai_api_key=None))

    assert result["metadata"]["provider"] == "langgraph"
    assert result["metadata"]["used_fallback"] is True
    assert result["ai_report"]["dashboard_title"]


def test_langgraph_invalid_model_output_falls_back(monkeypatch, tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)
    context = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")["ai_context"]

    def fake_llm(*args, **kwargs):
        return {"invalid": "schema"}

    monkeypatch.setattr("app.services.ai_analyst.graph._call_llm_structured", fake_llm)
    settings = Settings(openai_api_key="test-key", ai_analyst_fail_open=True)
    result = generate_ai_dashboard_report_result(context, settings=settings)

    assert result["metadata"]["provider"] == "langgraph"
    assert result["metadata"]["used_fallback"] is True
    assert result["ai_report"]["dashboard_title"]


def test_serialized_context_is_capped(tmp_path, sample_csv_bytes):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_bytes(sample_csv_bytes)
    analysis = prepare_python_analysis(csv_path, template="auto", output_format="xlsx")

    context = build_ai_context_json(
        analysis["cleaning_result"],
        analysis["spec"],
        analysis["aggregations"],
        analysis["selected_template"],
        "xlsx",
        AIContextLimits(max_ai_context_chars=900, max_top_categories=10, max_trend_points=12, max_anomalies=5),
    )

    assert serialized_context_length(context) <= 900
