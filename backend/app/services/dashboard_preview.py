from typing import Any

import pandas as pd


class DashboardPreviewBuilder:
    def build(self, df: pd.DataFrame, dashboard_spec: dict[str, Any], dataset_profile: dict[str, Any]) -> dict[str, Any]:
        return {
            "kpis": dashboard_spec.get("kpis", []),
            "charts": [self._chart_payload(df, chart) for chart in dashboard_spec.get("charts", [])],
            "quality_report": {
                "quality_score": dataset_profile.get("quality_score"),
                "duplicate_rows": dataset_profile.get("duplicate_rows", 0),
                "columns_with_missing_values": [
                    column for column in dataset_profile.get("columns", []) if column.get("missing_rate", 0) > 0
                ],
                "outliers_simple": dataset_profile.get("outliers_simple", []),
            },
            "sample_rows": self._sample_rows(df),
        }

    def _chart_payload(self, df: pd.DataFrame, chart: dict[str, Any]) -> dict[str, Any]:
        labels, values = self._aggregate(df, chart)
        return {
            "id": chart["id"],
            "type": chart["type"],
            "title": chart["title"],
            "labels": labels,
            "series": [{"name": chart.get("y") or "Count", "data": values}],
        }

    def _aggregate(self, df: pd.DataFrame, chart: dict[str, Any]) -> tuple[list[str], list[float]]:
        x_col = chart.get("x")
        y_col = chart.get("y")
        if not x_col or x_col not in df:
            return [], []

        work = df[[x_col] + ([y_col] if y_col and y_col in df else [])].copy()
        work[x_col] = work[x_col].fillna("Unknown")

        parsed_dates = pd.to_datetime(work[x_col], errors="coerce", format="mixed")
        if parsed_dates.notna().mean() >= 0.75:
            work["_x"] = parsed_dates.dt.to_period("M").astype(str)
            sort_index = True
        else:
            work["_x"] = work[x_col].astype(str)
            sort_index = False

        if y_col and y_col in work:
            work["_y"] = pd.to_numeric(work[y_col], errors="coerce").fillna(0)
            aggregation = chart.get("aggregation", "sum")
            if aggregation == "average":
                grouped = work.groupby("_x")["_y"].mean()
            elif aggregation == "count":
                grouped = work.groupby("_x")["_y"].count()
            else:
                grouped = work.groupby("_x")["_y"].sum()
        else:
            grouped = work.groupby("_x").size()

        if sort_index:
            grouped = grouped.sort_index()
        else:
            grouped = grouped.sort_values(ascending=False)

        limit = chart.get("limit")
        if limit:
            grouped = grouped.head(int(limit))

        labels = [str(index) for index in grouped.index.tolist()]
        values = [float(value) for value in grouped.values.tolist()]
        return labels, values

    @staticmethod
    def _sample_rows(df: pd.DataFrame) -> list[dict[str, Any]]:
        clean = df.head(10).where(pd.notna(df.head(10)), None)
        rows: list[dict[str, Any]] = []
        for row in clean.to_dict(orient="records"):
            rows.append({key: DashboardPreviewBuilder._json_value(value) for key, value in row.items()})
        return rows

    @staticmethod
    def _json_value(value: Any) -> Any:
        if pd.isna(value):
            return None
        if hasattr(value, "isoformat"):
            return value.isoformat()
        if isinstance(value, (int, float, str, bool)):
            return value
        return str(value)
