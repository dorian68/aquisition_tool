import uuid
from datetime import datetime
from typing import Any

import pandas as pd


class DashboardPlanner:
    theme = "optiquant_dark_premium"

    def plan(self, dataset_profile: dict[str, Any], df: pd.DataFrame) -> dict[str, Any]:
        dashboard_type = self._detect_dashboard_type(dataset_profile)
        if dashboard_type == "invoice_overdue":
            return self._invoice_dashboard(dataset_profile, df)
        if dashboard_type == "sales_performance":
            return self._sales_dashboard(dataset_profile, df)
        return self._generic_dashboard(dataset_profile, df)

    def _sales_dashboard(self, profile: dict[str, Any], df: pd.DataFrame) -> dict[str, Any]:
        revenue_col = self._first_semantic(profile, "money") or self._first_metric(profile)
        date_col = self._first_semantic(profile, "date")
        product_col = self._first_semantic(profile, "product")
        customer_col = self._first_semantic(profile, "customer")
        region_col = self._first_semantic(profile, "region")
        status_col = self._first_semantic(profile, "status")

        total_revenue = self._sum(df, revenue_col)
        order_count = int(len(df))
        average_order_value = total_revenue / order_count if order_count else 0

        kpis = [
            self._kpi("total_revenue", "Total Revenue", "sum", revenue_col, "currency", total_revenue),
            self._kpi("number_of_orders", "Number of Orders", "count", None, "number", order_count),
            self._kpi("average_order_value", "Average Order Value", "average", revenue_col, "currency", average_order_value),
        ]
        if product_col:
            kpis.append(self._kpi("top_product", "Top Product", "top", product_col, "text", self._top_value(df, product_col, revenue_col)))
        if customer_col:
            kpis.append(self._kpi("top_customer", "Top Customer", "top", customer_col, "text", self._top_value(df, customer_col, revenue_col)))

        charts = []
        if date_col and revenue_col:
            charts.append(self._chart("revenue_over_time", "line", "Revenue Over Time", date_col, revenue_col, "sum"))
        if product_col and revenue_col:
            charts.append(self._chart("top_products", "bar", "Top Products by Revenue", product_col, revenue_col, "sum", 10))
        if region_col and revenue_col:
            charts.append(self._chart("revenue_by_region", "column", "Revenue by Region", region_col, revenue_col, "sum", 10))
        if status_col:
            charts.append(self._chart("order_status_mix", "doughnut", "Order Status Mix", status_col, None, "count", 8))
        charts = charts or self._fallback_charts(profile)

        insights = [
            f"Total revenue is {self._format_number(total_revenue)} across {order_count} rows.",
            self._missing_insight(profile),
        ]
        if product_col and revenue_col:
            insights.append(f"{self._top_value(df, product_col, revenue_col)} is the strongest product signal in this export.")

        return self._spec(
            dashboard_type="sales_performance",
            title="Sales Performance Dashboard",
            kpis=kpis,
            charts=charts,
            insights=insights,
        )

    def _invoice_dashboard(self, profile: dict[str, Any], df: pd.DataFrame) -> dict[str, Any]:
        amount_col = self._first_semantic(profile, "money") or self._first_metric(profile)
        due_date_col = self._first_semantic(profile, "due_date") or self._date_named(profile, "due")
        customer_col = self._first_semantic(profile, "customer")
        status_col = self._first_semantic(profile, "status")

        total_invoiced = self._sum(df, amount_col)
        paid_mask = self._status_contains(df, status_col, ("paid", "settled", "closed"))
        unpaid_mask = ~paid_mask if paid_mask is not None else pd.Series([False] * len(df), index=df.index)
        overdue_mask = self._overdue_mask(df, due_date_col) & unpaid_mask if due_date_col else pd.Series([False] * len(df), index=df.index)
        overdue_amount = float(pd.to_numeric(df.loc[overdue_mask, amount_col], errors="coerce").fillna(0).sum()) if amount_col else 0
        paid_amount = float(pd.to_numeric(df.loc[paid_mask, amount_col], errors="coerce").fillna(0).sum()) if amount_col and paid_mask is not None else 0

        kpis = [
            self._kpi("total_invoiced", "Total Invoiced", "sum", amount_col, "currency", total_invoiced),
            self._kpi("overdue_amount", "Overdue Amount", "sum", amount_col, "currency", overdue_amount),
            self._kpi("paid_amount", "Paid Amount", "sum", amount_col, "currency", paid_amount),
            self._kpi("unpaid_invoices", "Unpaid Invoices", "count", status_col, "number", int(unpaid_mask.sum())),
        ]

        charts = []
        if customer_col and amount_col:
            charts.append(self._chart("overdue_by_customer", "bar", "Overdue by Customer", customer_col, amount_col, "sum", 10))
        if status_col:
            charts.append(self._chart("payment_status", "doughnut", "Payment Status", status_col, None, "count", 8))
        if due_date_col and amount_col:
            charts.append(self._chart("monthly_invoicing_trend", "line", "Monthly Invoicing Trend", due_date_col, amount_col, "sum"))
        charts = charts or self._fallback_charts(profile)

        insights = [
            f"Open receivables represent {self._format_number(overdue_amount)} in overdue amount.",
            self._missing_insight(profile),
        ]
        return self._spec(
            dashboard_type="invoice_overdue",
            title="Invoice & Overdue Dashboard",
            kpis=kpis,
            charts=charts,
            insights=insights,
        )

    def _generic_dashboard(self, profile: dict[str, Any], df: pd.DataFrame) -> dict[str, Any]:
        metric_col = self._first_metric(profile)
        category_col = self._first_dimension(profile)
        total_metric = self._sum(df, metric_col)
        average_metric = self._mean(df, metric_col)

        kpis = [
            self._kpi("row_count", "Rows", "count", None, "number", profile["row_count"]),
            self._kpi("data_quality_score", "Data Quality Score", "score", None, "percent", profile["quality_score"]),
        ]
        if metric_col:
            kpis.extend(
                [
                    self._kpi("total_metric", f"Total {metric_col}", "sum", metric_col, "number", total_metric),
                    self._kpi("average_metric", f"Average {metric_col}", "average", metric_col, "number", average_metric),
                ]
            )
        if category_col:
            kpis.append(self._kpi("top_category", f"Top {category_col}", "top", category_col, "text", self._top_value(df, category_col, metric_col)))

        charts = self._fallback_charts(profile)
        insights = [
            f"The dataset contains {profile['row_count']} rows and {profile['column_count']} columns.",
            self._missing_insight(profile),
            "This executive view highlights the strongest available metric and category signals.",
        ]
        return self._spec("generic_executive", "Executive CSV Dashboard", kpis, charts, insights)

    def _detect_dashboard_type(self, profile: dict[str, Any]) -> str:
        semantics = {column["semantic_type"] for column in profile.get("columns", [])}
        names = " ".join(column["name"].lower() for column in profile.get("columns", []))
        if "invoice" in names or ("due_date" in semantics and "status" in semantics):
            return "invoice_overdue"
        if "money" in semantics and ({"date", "product", "customer", "region"} & semantics):
            return "sales_performance"
        return "generic_executive"

    def _fallback_charts(self, profile: dict[str, Any]) -> list[dict[str, Any]]:
        metric = self._first_metric(profile)
        date = self._first_semantic(profile, "date")
        dimension = self._first_dimension(profile)
        charts = []
        if date and metric:
            charts.append(self._chart("metric_over_time", "line", f"{metric} Over Time", date, metric, "sum"))
        if dimension and metric:
            charts.append(self._chart("top_dimension", "bar", f"Top {dimension} by {metric}", dimension, metric, "sum", 10))
        elif dimension:
            charts.append(self._chart("dimension_mix", "doughnut", f"{dimension} Mix", dimension, None, "count", 8))
        return charts

    @staticmethod
    def _spec(
        dashboard_type: str,
        title: str,
        kpis: list[dict[str, Any]],
        charts: list[dict[str, Any]],
        insights: list[str],
    ) -> dict[str, Any]:
        return {
            "dashboard_id": str(uuid.uuid4()),
            "dashboard_type": dashboard_type,
            "title": title,
            "subtitle": "Generated from your CSV export",
            "theme": DashboardPlanner.theme,
            "kpis": kpis,
            "charts": charts,
            "layout": {
                "type": "executive_dashboard",
                "sections": ["header", "kpi_row", "main_charts", "insights", "data_quality"],
            },
            "insights": [insight for insight in insights if insight],
            "generated_at": datetime.utcnow().isoformat(),
        }

    @staticmethod
    def _kpi(kpi_id: str, label: str, calculation: str, column: str | None, fmt: str, value: Any) -> dict[str, Any]:
        return {"id": kpi_id, "label": label, "calculation": calculation, "column": column, "format": fmt, "value": value}

    @staticmethod
    def _chart(
        chart_id: str,
        chart_type: str,
        title: str,
        x: str,
        y: str | None,
        aggregation: str,
        limit: int | None = None,
    ) -> dict[str, Any]:
        chart = {"id": chart_id, "type": chart_type, "title": title, "x": x, "y": y, "aggregation": aggregation}
        if limit:
            chart["limit"] = limit
        return chart

    @staticmethod
    def _first_metric(profile: dict[str, Any]) -> str | None:
        metrics = profile.get("detected_metrics") or []
        return metrics[0] if metrics else None

    @staticmethod
    def _first_dimension(profile: dict[str, Any]) -> str | None:
        dimensions = profile.get("detected_dimensions") or []
        return dimensions[0] if dimensions else None

    @staticmethod
    def _first_semantic(profile: dict[str, Any], semantic_type: str) -> str | None:
        for column in profile.get("columns", []):
            if column.get("semantic_type") == semantic_type:
                return column["name"]
        return None

    @staticmethod
    def _date_named(profile: dict[str, Any], needle: str) -> str | None:
        for column in profile.get("columns", []):
            if column.get("type") == "date" and needle in column["name"].lower():
                return column["name"]
        return None

    @staticmethod
    def _sum(df: pd.DataFrame, column: str | None) -> float:
        if not column or column not in df:
            return 0.0
        return float(pd.to_numeric(df[column], errors="coerce").fillna(0).sum())

    @staticmethod
    def _mean(df: pd.DataFrame, column: str | None) -> float:
        if not column or column not in df:
            return 0.0
        values = pd.to_numeric(df[column], errors="coerce").dropna()
        return float(values.mean()) if len(values) else 0.0

    @staticmethod
    def _top_value(df: pd.DataFrame, dimension: str | None, metric: str | None = None) -> str | None:
        if not dimension or dimension not in df:
            return None
        if metric and metric in df:
            grouped = pd.to_numeric(df[metric], errors="coerce").fillna(0).groupby(df[dimension].fillna("Unknown")).sum()
            if not grouped.empty:
                return str(grouped.sort_values(ascending=False).index[0])
        counts = df[dimension].fillna("Unknown").value_counts()
        return str(counts.index[0]) if not counts.empty else None

    @staticmethod
    def _status_contains(df: pd.DataFrame, status_col: str | None, values: tuple[str, ...]) -> pd.Series | None:
        if not status_col or status_col not in df:
            return None
        lowered = df[status_col].fillna("").astype(str).str.lower()
        return lowered.apply(lambda item: any(value in item for value in values))

    @staticmethod
    def _overdue_mask(df: pd.DataFrame, due_date_col: str | None) -> pd.Series:
        if not due_date_col or due_date_col not in df:
            return pd.Series([False] * len(df), index=df.index)
        due_dates = pd.to_datetime(df[due_date_col], errors="coerce", format="mixed")
        return due_dates.notna() & (due_dates < pd.Timestamp.utcnow().tz_localize(None))

    @staticmethod
    def _format_number(value: float) -> str:
        return f"{value:,.0f}"

    @staticmethod
    def _missing_insight(profile: dict[str, Any]) -> str:
        affected = [column["name"] for column in profile.get("columns", []) if column.get("missing_rate", 0) > 0]
        if not affected:
            return "No missing values were detected in the profiled sample."
        return f"Missing values were detected in {len(affected)} columns: {', '.join(affected[:3])}."
