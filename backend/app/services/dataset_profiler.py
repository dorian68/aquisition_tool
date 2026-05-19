from collections.abc import Iterable
from typing import Any

import pandas as pd


MONEY_KEYWORDS = ("revenue", "sales", "amount", "price", "total", "paid", "spend", "cost", "invoice")
QUANTITY_KEYWORDS = ("qty", "quantity", "units", "count", "orders")
DATE_KEYWORDS = ("date", "time", "created", "updated", "month", "year", "due")
REGION_KEYWORDS = ("country", "region", "state", "city", "market", "territory")
PRODUCT_KEYWORDS = ("product", "sku", "item", "category")
CUSTOMER_KEYWORDS = ("customer", "client", "account", "company")
STATUS_KEYWORDS = ("status", "stage", "state")
ID_KEYWORDS = ("id", "uuid", "reference", "ref", "number", "invoice")
CAMPAIGN_KEYWORDS = ("campaign", "channel", "source", "medium")


def _name_has(name: str, keywords: Iterable[str]) -> bool:
    lowered = name.lower().replace("-", "_").replace(" ", "_")
    return any(keyword in lowered for keyword in keywords)


def _json_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if hasattr(value, "isoformat"):
        return value.isoformat()
    if isinstance(value, (int, float, str, bool)):
        return value
    return str(value)


class DatasetProfiler:
    def profile(self, df: pd.DataFrame) -> dict[str, Any]:
        row_count = int(len(df))
        duplicate_rows = int(df.duplicated().sum()) if row_count else 0
        columns: list[dict[str, Any]] = []
        detected_dates: list[str] = []
        detected_metrics: list[str] = []
        detected_dimensions: list[str] = []
        outliers: list[dict[str, Any]] = []

        for column in df.columns:
            series = df[column]
            column_type = self._column_type(column, series)
            semantic_type = self._semantic_type(column, column_type, series)
            missing_rate = float(series.isna().mean()) if row_count else 0.0
            unique_count = int(series.nunique(dropna=True))
            sample_values = [_json_value(value) for value in series.dropna().head(3).tolist()]

            if column_type == "date":
                detected_dates.append(column)
            if column_type == "numeric" and semantic_type != "identifier":
                detected_metrics.append(column)
            if column_type in {"categorical", "text"} and semantic_type != "identifier":
                detected_dimensions.append(column)

            outlier_count = self._outlier_count(series) if column_type == "numeric" else 0
            if outlier_count:
                outliers.append({"column": column, "count": outlier_count})

            columns.append(
                {
                    "name": column,
                    "type": column_type,
                    "semantic_type": semantic_type,
                    "missing_rate": round(missing_rate, 4),
                    "unique_count": unique_count,
                    "sample_values": sample_values,
                }
            )

        avg_missing = sum(column["missing_rate"] for column in columns) / max(len(columns), 1)
        duplicate_rate = duplicate_rows / max(row_count, 1)
        quality_score = round(max(0.0, 100.0 - avg_missing * 45 - duplicate_rate * 30 - len(outliers) * 2), 1)

        return {
            "row_count": row_count,
            "column_count": int(len(df.columns)),
            "columns": columns,
            "detected_date_columns": detected_dates,
            "detected_metrics": detected_metrics,
            "detected_dimensions": detected_dimensions,
            "duplicate_rows": duplicate_rows,
            "outliers_simple": outliers,
            "quality_score": quality_score,
        }

    def _column_type(self, name: str, series: pd.Series) -> str:
        if pd.api.types.is_numeric_dtype(series):
            return "numeric"

        non_null = series.dropna()
        if non_null.empty:
            return "text"

        if _name_has(name, DATE_KEYWORDS):
            parsed = pd.to_datetime(non_null, errors="coerce", format="mixed")
            if parsed.notna().mean() >= 0.65:
                return "date"

        parsed = pd.to_datetime(non_null, errors="coerce", format="mixed")
        if parsed.notna().mean() >= 0.85:
            return "date"

        unique_ratio = non_null.nunique(dropna=True) / max(len(non_null), 1)
        if unique_ratio <= 0.5 or non_null.nunique(dropna=True) <= 50:
            return "categorical"
        return "text"

    def _semantic_type(self, name: str, column_type: str, series: pd.Series) -> str:
        if column_type == "date":
            if "due" in name.lower():
                return "due_date"
            return "date"
        if _name_has(name, MONEY_KEYWORDS) and column_type == "numeric":
            return "money"
        if _name_has(name, QUANTITY_KEYWORDS) and column_type == "numeric":
            return "quantity"
        if _name_has(name, REGION_KEYWORDS):
            return "region"
        if _name_has(name, PRODUCT_KEYWORDS):
            return "product"
        if _name_has(name, CUSTOMER_KEYWORDS):
            return "customer"
        if _name_has(name, STATUS_KEYWORDS):
            return "status"
        if _name_has(name, CAMPAIGN_KEYWORDS):
            return "campaign"
        if _name_has(name, ID_KEYWORDS):
            unique_ratio = series.nunique(dropna=True) / max(series.dropna().shape[0], 1)
            if unique_ratio >= 0.75 or "invoice" in name.lower():
                return "identifier"
        return "generic"

    @staticmethod
    def _outlier_count(series: pd.Series) -> int:
        numeric = pd.to_numeric(series, errors="coerce").dropna()
        if len(numeric) < 8:
            return 0
        q1 = numeric.quantile(0.25)
        q3 = numeric.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            return 0
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        return int(((numeric < lower) | (numeric > upper)).sum())
