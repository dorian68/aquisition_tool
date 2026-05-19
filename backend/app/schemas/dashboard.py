from typing import Any, Literal

from pydantic import BaseModel, Field


class KpiSpec(BaseModel):
    id: str
    label: str
    calculation: str
    column: str | None = None
    format: str = "number"
    value: Any


class ChartSpec(BaseModel):
    id: str
    type: Literal["line", "bar", "column", "pie", "doughnut", "scatter"]
    title: str
    x: str
    y: str | None = None
    aggregation: str = "sum"
    limit: int | None = None


class DashboardSpecResponse(BaseModel):
    dashboard_id: str
    dashboard_type: str
    title: str
    subtitle: str
    theme: str
    kpis: list[dict[str, Any]] = Field(default_factory=list)
    charts: list[dict[str, Any]] = Field(default_factory=list)
    layout: dict[str, Any]
    insights: list[str] = Field(default_factory=list)


class PreviewResponse(BaseModel):
    dashboard_spec: dict[str, Any]
    preview_data: dict[str, Any]


class GenerateXlsxResponse(BaseModel):
    file_id: str
    download_url: str
    status: str = "ready"

