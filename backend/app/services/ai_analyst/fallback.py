from __future__ import annotations

from typing import Any

from .schemas import AIReport


def build_deterministic_ai_report(ai_context: dict[str, Any]) -> dict[str, Any]:
    overview = ai_context.get("dataset_overview", {})
    quality = ai_context.get("data_quality", {})
    metrics = ai_context.get("business_metrics", {})
    dashboard = ai_context.get("dashboard_context", {})
    anomalies = ai_context.get("anomalies", [])

    detected_type = str(overview.get("detected_type") or "business").replace("_", " ").title()
    metric = overview.get("main_metric") or "records"
    rows_after = int(overview.get("rows_after") or metrics.get("record_count") or 0)
    columns_after = int(overview.get("columns_after") or 0)
    quality_before = overview.get("quality_before", 0)
    quality_after = overview.get("quality_after", 0)
    duplicates = int(quality.get("duplicate_rows_removed") or 0)
    top_categories = metrics.get("top_categories") or []
    trend = metrics.get("trend") or []
    high_missing = quality.get("columns_with_high_missing_rate") or []

    key_insights = [
        f"The cleaned dataset contains {rows_after} records across {columns_after} columns and is classified as {detected_type}.",
        f"Data quality moved from {quality_before}% to {quality_after}% after Python cleaning and profiling.",
        _trend_note(trend, metric),
    ]
    if top_categories:
        key_insights.append(f"The leading segment is {top_categories[0]['name']} with {round(top_categories[0]['share_pct'], 1)}% of the summarized distribution.")
    if anomalies:
        key_insights.append(f"Python flagged {len(anomalies)} anomaly or risk signal(s) for review.")

    recommended_actions = [
        "Review the top segments and trend movement before sharing the dashboard with decision makers.",
        "Use the Data Quality and Cleaning Log sheets to trace source-data improvements.",
    ]
    if high_missing:
        recommended_actions.append("Prioritize fixing columns with high missing rates in the upstream data source.")
    if anomalies:
        recommended_actions.append("Investigate outliers, dominance warnings or large trend changes before operational decisions.")

    risk_notes = ["Raw CSV rows were not sent to the AI model; this report is based on a compact Python-generated profile."]
    for anomaly in anomalies[:3]:
        risk_notes.append(str(anomaly.get("details") or "Anomaly detected.")[:240])
    if float(quality_after or 0) < 80:
        risk_notes.append("The final quality score remains below 80%, so conclusions should be treated as directional.")

    report = AIReport.model_validate({
        "dashboard_title": f"{detected_type} Dashboard Analysis",
        "executive_summary": (
            f"Python cleaned and profiled the CSV, selected {dashboard.get('selected_template', 'a dashboard template')}, "
            f"and generated an Excel dashboard focused on {metric}. The dataset now has {rows_after} clean records and a "
            f"{quality_after}% quality score."
        )[:900],
        "data_quality_summary": (
            f"Missing cells changed from {quality.get('missing_cells_before_pct', 0)}% to "
            f"{quality.get('missing_cells_after_pct', 0)}%; {duplicates} duplicate rows were removed."
        ),
        "cleaning_summary": _cleaning_summary(ai_context),
        "key_insights": key_insights[:5],
        "recommended_actions": recommended_actions[:5],
        "risk_notes": risk_notes[:5],
    })
    return report.model_dump()


def _trend_note(trend: list[dict[str, Any]], metric: str) -> str:
    if len(trend) < 2:
        return f"Trend data for {metric} is limited; use the workbook charts for directional review."
    previous = float(trend[-2].get("value") or 0)
    latest = float(trend[-1].get("value") or 0)
    if previous == 0:
        return f"Latest {metric} trend value is {round(latest, 2)}."
    change = round((latest / previous - 1.0) * 100, 1)
    direction = "increased" if change >= 0 else "decreased"
    return f"Latest {metric} {direction} by {abs(change)}% versus the previous period."


def _cleaning_summary(ai_context: dict[str, Any]) -> str:
    actions = ai_context.get("cleaning_actions", [])
    if not actions:
        return "No material cleaning actions were required beyond deterministic profiling."
    names = []
    for action in actions[:5]:
        name = str(action.get("action", "cleaning")).replace("_", " ")
        if name not in names:
            names.append(name)
    return "Python applied deterministic cleaning steps including " + ", ".join(names) + "."
