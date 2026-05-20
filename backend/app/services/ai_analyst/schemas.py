from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class DatasetOverview(StrictModel):
    detected_type: str
    rows_before: int
    rows_after: int
    columns_before: int
    columns_after: int
    main_metric: str | None = None
    date_column: str | None = None
    primary_dimension: str | None = None
    secondary_dimension: str | None = None
    quality_before: float
    quality_after: float


class HighMissingColumn(StrictModel):
    column: str
    missing_pct: float


class DataQualitySummary(StrictModel):
    missing_cells_before_pct: float
    missing_cells_after_pct: float
    duplicate_rows_removed: int
    empty_columns_removed: list[str] = Field(default_factory=list)
    columns_with_high_missing_rate: list[HighMissingColumn] = Field(default_factory=list)
    detected_issues: list[str] = Field(default_factory=list)


class CleaningAction(StrictModel):
    action: str
    column: str | None = None
    details: str


class TopCategory(StrictModel):
    name: str
    value: float
    share_pct: float


class TrendPoint(StrictModel):
    period: str
    value: float


class BusinessMetrics(StrictModel):
    total_metric: float | None = None
    average_metric: float | None = None
    record_count: int
    top_categories: list[TopCategory] = Field(default_factory=list)
    trend: list[TrendPoint] = Field(default_factory=list)
    secondary_distribution: list[TopCategory] = Field(default_factory=list)


class Anomaly(StrictModel):
    type: str
    column: str
    details: str
    severity: Literal["low", "medium", "high"]


class DashboardContext(StrictModel):
    selected_template: str
    recommended_template: str
    recommendation_reason: str
    output_format: str


class AIContext(StrictModel):
    dataset_overview: DatasetOverview
    data_quality: DataQualitySummary
    cleaning_actions: list[CleaningAction] = Field(default_factory=list)
    business_metrics: BusinessMetrics
    anomalies: list[Anomaly] = Field(default_factory=list)
    dashboard_context: DashboardContext


class AIReport(StrictModel):
    dashboard_title: str
    executive_summary: str
    data_quality_summary: str
    cleaning_summary: str
    key_insights: list[str] = Field(min_length=1, max_length=5)
    recommended_actions: list[str] = Field(min_length=1, max_length=5)
    risk_notes: list[str] = Field(min_length=1, max_length=5)

    @field_validator(
        "dashboard_title",
        "executive_summary",
        "data_quality_summary",
        "cleaning_summary",
    )
    @classmethod
    def require_non_empty_string(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("AI report strings cannot be empty")
        return value

    @field_validator("key_insights", "recommended_actions", "risk_notes")
    @classmethod
    def require_non_empty_items(cls, value: list[str]) -> list[str]:
        cleaned = [str(item).strip() for item in value if str(item).strip()]
        if not cleaned:
            raise ValueError("AI report lists cannot be empty")
        return cleaned[:5]
