from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from app.core.config import Settings, get_settings


def _ensure_project_root_on_path() -> None:
    project_root = Path(__file__).resolve().parents[4]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))


try:
    from dashboard_generator.aggregation import safe_avg, safe_sum
    from dashboard_generator.cleaning import CleaningResult, clean_csv
    from dashboard_generator.semantic_detection import classify_dataset
    from dashboard_generator.legacy_core import aggregate_data
    from dashboard_generator.theme import choose_template
except ModuleNotFoundError:
    _ensure_project_root_on_path()
    from dashboard_generator.aggregation import safe_avg, safe_sum
    from dashboard_generator.cleaning import CleaningResult, clean_csv
    from dashboard_generator.semantic_detection import classify_dataset
    from dashboard_generator.legacy_core import aggregate_data
    from dashboard_generator.theme import choose_template


@dataclass(frozen=True)
class AIContextLimits:
    max_ai_context_chars: int
    max_top_categories: int
    max_trend_points: int
    max_anomalies: int
    max_cleaning_actions: int = 20
    max_high_missing_columns: int = 10

    @classmethod
    def from_settings(cls, settings: Settings) -> "AIContextLimits":
        return cls(
            max_ai_context_chars=settings.max_ai_context_chars,
            max_top_categories=settings.max_top_categories,
            max_trend_points=settings.max_trend_points,
            max_anomalies=settings.max_anomalies,
            max_cleaning_actions=settings.max_cleaning_actions_for_ai,
            max_high_missing_columns=settings.max_high_missing_columns_for_ai,
        )


def prepare_python_analysis(csv_path: str | Path, *, template: str, output_format: str, settings: Settings | None = None) -> dict[str, Any]:
    settings = settings or get_settings()
    cleaning_result = clean_csv(csv_path)
    _validate_column_limit(cleaning_result, settings)

    quality = cleaning_result.quality_report
    spec = classify_dataset(cleaning_result.clean_df, quality)
    aggregations = aggregate_data(cleaning_result.clean_df, spec)
    selected_template = choose_template(spec).slug if template == "auto" else template
    ai_context = build_ai_context_json(
        cleaning_result,
        spec,
        aggregations,
        selected_template,
        output_format,
        AIContextLimits.from_settings(settings),
    )
    from .service import generate_ai_dashboard_report_result

    result = generate_ai_dashboard_report_result(ai_context, settings=settings)
    return {
        "cleaning_result": cleaning_result,
        "spec": spec,
        "aggregations": aggregations,
        "selected_template": selected_template,
        "ai_context": ai_context,
        "ai_report": result["ai_report"],
        "ai_metadata": result["metadata"],
    }


def validate_csv_shape(csv_path: str | Path, *, settings: Settings | None = None) -> None:
    settings = settings or get_settings()
    cleaning_result = clean_csv(csv_path)
    _validate_column_limit(cleaning_result, settings)


def build_ai_context_json(
    cleaning_result: CleaningResult,
    spec: dict[str, Any],
    aggregations: dict[str, Any],
    selected_template: str,
    output_format: str,
    limits: AIContextLimits | dict[str, Any],
) -> dict[str, Any]:
    if isinstance(limits, dict):
        limits = AIContextLimits(
            max_ai_context_chars=int(limits.get("max_ai_context_chars", 20_000)),
            max_top_categories=int(limits.get("max_top_categories", 10)),
            max_trend_points=int(limits.get("max_trend_points", 12)),
            max_anomalies=int(limits.get("max_anomalies", 5)),
            max_cleaning_actions=int(limits.get("max_cleaning_actions", 20)),
            max_high_missing_columns=int(limits.get("max_high_missing_columns", 10)),
        )

    quality = cleaning_result.quality_report
    df = cleaning_result.clean_df
    raw_df = cleaning_result.raw_df
    metric = spec.get("metric")
    recommended_template = choose_template(spec).slug
    total_before = max(int(raw_df.shape[0] * raw_df.shape[1]), 1)
    total_after = max(int(df.shape[0] * df.shape[1]), 1)
    missing_before = int(raw_df.isna().sum().sum())
    missing_after = int(df.isna().sum().sum())
    empty_columns = _empty_columns_from_log(cleaning_result.cleaning_log)
    high_missing_columns = _high_missing_columns(raw_df, limits.max_high_missing_columns)

    context = {
        "dataset_overview": {
            "detected_type": str(spec.get("dataset_type", "generic")),
            "rows_before": int(quality["original_rows"]),
            "rows_after": int(quality["clean_rows"]),
            "columns_before": int(quality["original_columns"]),
            "columns_after": int(quality["clean_columns"]),
            "main_metric": metric,
            "date_column": spec.get("date"),
            "primary_dimension": spec.get("primary_dimension"),
            "secondary_dimension": spec.get("secondary_dimension"),
            "quality_before": float(quality["quality_before"]),
            "quality_after": float(quality["quality_after"]),
        },
        "data_quality": {
            "missing_cells_before_pct": _pct(missing_before, total_before),
            "missing_cells_after_pct": _pct(missing_after, total_after),
            "duplicate_rows_removed": int(quality["duplicate_rows_removed"]),
            "empty_columns_removed": empty_columns[: limits.max_high_missing_columns],
            "columns_with_high_missing_rate": high_missing_columns,
            "detected_issues": _detected_issues(quality, high_missing_columns, empty_columns),
        },
        "cleaning_actions": _cleaning_actions(cleaning_result.cleaning_log, limits.max_cleaning_actions),
        "business_metrics": {
            "total_metric": _number_or_none(safe_sum(df, metric)) if metric else None,
            "average_metric": _number_or_none(safe_avg(df, metric)) if metric else None,
            "record_count": int(len(df)),
            "top_categories": _frame_distribution(aggregations.get("top_dim"), limits.max_top_categories),
            "trend": _trend(aggregations.get("trend"), limits.max_trend_points),
            "secondary_distribution": _frame_distribution(aggregations.get("second_dim"), limits.max_top_categories),
        },
        "anomalies": _anomalies(cleaning_result, spec, aggregations, limits.max_anomalies),
        "dashboard_context": {
            "selected_template": selected_template,
            "recommended_template": recommended_template,
            "recommendation_reason": _recommendation_reason(spec, recommended_template),
            "output_format": output_format,
        },
    }
    return _cap_context(context, limits.max_ai_context_chars)


def serialized_context_length(ai_context: dict[str, Any]) -> int:
    return len(json.dumps(ai_context, ensure_ascii=True, separators=(",", ":")))


def _validate_column_limit(cleaning_result: CleaningResult, settings: Settings) -> None:
    columns = int(cleaning_result.quality_report["original_columns"])
    if columns > settings.max_columns:
        raise ValueError(f"Uploaded CSV has too many columns ({columns} > {settings.max_columns})")


def _pct(value: int | float, total: int | float) -> float:
    return round(float(value) / max(float(total), 1.0) * 100, 2)


def _number_or_none(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number != number:
        return None
    return round(number, 4)


def _empty_columns_from_log(log: list[dict[str, Any]]) -> list[str]:
    for entry in log:
        if entry.get("action") == "drop_empty_columns":
            return [item.strip() for item in str(entry.get("details", "")).split(",") if item.strip()]
    return []


def _high_missing_columns(raw_df: Any, limit: int) -> list[dict[str, Any]]:
    if raw_df.empty:
        return []
    rows = max(len(raw_df), 1)
    items = []
    for column in raw_df.columns:
        pct = _pct(int(raw_df[column].isna().sum()), rows)
        if pct >= 20:
            items.append({"column": str(column)[:120], "missing_pct": pct})
    items.sort(key=lambda item: item["missing_pct"], reverse=True)
    return items[:limit]


def _detected_issues(quality: dict[str, Any], high_missing_columns: list[dict[str, Any]], empty_columns: list[str]) -> list[str]:
    issues = []
    if quality.get("duplicate_rows_removed", 0):
        issues.append(f"{quality['duplicate_rows_removed']} duplicate rows removed")
    if empty_columns:
        issues.append(f"{len(empty_columns)} empty columns removed")
    if high_missing_columns:
        issues.append(f"{len(high_missing_columns)} columns have at least 20% missing values")
    if quality.get("quality_after", 100) < 80:
        issues.append("Data quality remains below 80 after cleaning")
    return issues or ["No major deterministic data quality issues detected"]


def _cleaning_actions(log: list[dict[str, Any]], limit: int) -> list[dict[str, Any]]:
    out = []
    for entry in log[:limit]:
        out.append({
            "action": str(entry.get("action", "unknown"))[:80],
            "column": str(entry["column"])[:120] if entry.get("column") is not None else None,
            "details": str(entry.get("details", ""))[:300],
        })
    return out


def _frame_distribution(frame: Any, limit: int) -> list[dict[str, Any]]:
    if frame is None or getattr(frame, "empty", True):
        return []
    rows = frame.head(limit).copy()
    total = float(rows["Value"].sum()) if "Value" in rows else 0.0
    out = []
    for _, row in rows.iterrows():
        value = _number_or_none(row.get("Value")) or 0.0
        out.append({
            "name": str(row.get("Category", "N/A"))[:120],
            "value": value,
            "share_pct": _pct(value, total) if total else 0.0,
        })
    return out


def _trend(frame: Any, limit: int) -> list[dict[str, Any]]:
    if frame is None or getattr(frame, "empty", True):
        return []
    out = []
    for _, row in frame.tail(limit).iterrows():
        out.append({"period": str(row.get("Period", "N/A"))[:40], "value": _number_or_none(row.get("Value")) or 0.0})
    return out


def _anomalies(cleaning_result: CleaningResult, spec: dict[str, Any], aggregations: dict[str, Any], limit: int) -> list[dict[str, Any]]:
    quality = cleaning_result.quality_report
    anomalies = []
    if quality.get("duplicate_rows_removed", 0):
        anomalies.append({
            "type": "duplicates_removed",
            "column": "*",
            "details": f"{quality['duplicate_rows_removed']} duplicate rows were removed during cleaning.",
            "severity": "medium" if quality["duplicate_rows_removed"] > 5 else "low",
        })
    for column, missing in quality.get("missing_before", {}).items():
        pct = _pct(int(missing), max(quality.get("original_rows", 1), 1))
        if pct >= 30:
            anomalies.append({
                "type": "high_missing_rate",
                "column": str(column)[:120],
                "details": f"{pct}% missing values before cleaning.",
                "severity": "high" if pct >= 60 else "medium",
            })
    metric = spec.get("metric")
    if metric and metric in cleaning_result.clean_df.columns:
        anomalies.extend(_numeric_outlier_anomalies(cleaning_result.clean_df[metric], str(metric)))
    anomalies.extend(_category_dominance_anomalies(aggregations))
    anomalies.extend(_trend_anomalies(aggregations, str(metric or "metric")))
    return anomalies[:limit]


def _numeric_outlier_anomalies(series: Any, column: str) -> list[dict[str, Any]]:
    try:
        q1 = float(series.quantile(0.25))
        q3 = float(series.quantile(0.75))
        iqr = q3 - q1
        if iqr <= 0:
            return []
        outlier_count = int(((series < q1 - 1.5 * iqr) | (series > q3 + 1.5 * iqr)).sum())
        if not outlier_count:
            return []
        return [{
            "type": "metric_outliers",
            "column": column,
            "details": f"{outlier_count} potential outliers detected using IQR.",
            "severity": "medium",
        }]
    except Exception:
        return []


def _category_dominance_anomalies(aggregations: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for key in ("top_dim", "second_dim"):
        frame = aggregations.get(key)
        if frame is None or getattr(frame, "empty", True) or "Value" not in frame:
            continue
        total = float(frame["Value"].sum())
        if total <= 0:
            continue
        top = frame.iloc[0]
        share = float(top["Value"]) / total * 100
        if share > 80:
            out.append({
                "type": "category_dominance",
                "column": key,
                "details": f"Top category '{str(top.get('Category', 'N/A'))[:80]}' represents {round(share, 1)}% of the distribution.",
                "severity": "medium",
            })
    return out


def _trend_anomalies(aggregations: dict[str, Any], metric: str) -> list[dict[str, Any]]:
    trend = _trend(aggregations.get("trend"), 12)
    if len(trend) < 2 or not trend[-2]["value"]:
        return []
    change_pct = (trend[-1]["value"] / trend[-2]["value"] - 1.0) * 100
    if abs(change_pct) < 50:
        return []
    return [{
        "type": "large_trend_change",
        "column": metric,
        "details": f"Latest trend point changed by {round(change_pct, 1)}% versus previous period.",
        "severity": "high" if abs(change_pct) >= 100 else "medium",
    }]


def _recommendation_reason(spec: dict[str, Any], recommended_template: str) -> str:
    dataset_type = str(spec.get("dataset_type", "generic"))
    if recommended_template == "dark-saas":
        return f"Dataset was detected as {dataset_type}, which matches SaaS, sales, CRM or marketing reporting."
    if recommended_template == "fintech-executive":
        return f"Dataset contains financial or invoice-like signals detected as {dataset_type}."
    return "Dataset appears generic or consulting-oriented, so the readable light template is safest."


def _cap_context(context: dict[str, Any], max_chars: int) -> dict[str, Any]:
    serialized = json.dumps(context, ensure_ascii=True, separators=(",", ":"))
    if len(serialized) <= max_chars:
        return context
    pruning_steps = [
        lambda c: c["anomalies"].pop() if c.get("anomalies") else None,
        lambda c: c["cleaning_actions"].pop() if c.get("cleaning_actions") else None,
        lambda c: c["business_metrics"]["top_categories"].pop() if c.get("business_metrics", {}).get("top_categories") else None,
        lambda c: c["business_metrics"]["trend"].pop(0) if c.get("business_metrics", {}).get("trend") else None,
        lambda c: c["business_metrics"]["secondary_distribution"].pop() if c.get("business_metrics", {}).get("secondary_distribution") else None,
        lambda c: c["data_quality"]["columns_with_high_missing_rate"].pop() if c.get("data_quality", {}).get("columns_with_high_missing_rate") else None,
    ]
    while len(serialized) > max_chars:
        changed = False
        for step in pruning_steps:
            before = serialized
            step(context)
            serialized = json.dumps(context, ensure_ascii=True, separators=(",", ":"))
            changed = changed or serialized != before
            if len(serialized) <= max_chars:
                return context
        if not changed:
            context["cleaning_actions"] = []
            context["anomalies"] = []
            context["business_metrics"]["top_categories"] = []
            context["business_metrics"]["secondary_distribution"] = []
            context["business_metrics"]["trend"] = []
            context["data_quality"]["columns_with_high_missing_rate"] = []
            context["data_quality"]["detected_issues"] = []
            break
    serialized = json.dumps(context, ensure_ascii=True, separators=(",", ":"))
    if len(serialized) > max_chars:
        _truncate_strings(context, 40)
    return context


def _truncate_strings(value: Any, max_len: int) -> None:
    if isinstance(value, dict):
        for key, item in list(value.items()):
            if isinstance(item, str):
                value[key] = item[:max_len]
            else:
                _truncate_strings(item, max_len)
    elif isinstance(value, list):
        for item in value:
            _truncate_strings(item, max_len)
