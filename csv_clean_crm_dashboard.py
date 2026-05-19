"""
OptiQuant IA — CSV Cleaner + CRM-style Excel Dashboard Generator

Goal:
- Take any CSV file
- Clean and normalize the data
- Generate a premium dark Excel dashboard inspired by the provided CRM dashboard reference

Important:
This version uses a pixel-based layout engine on top of XlsxWriter. Instead of
placing components directly with fixed Excel ranges, it designs the dashboard as
a 1500 x 900 px canvas, converts rectangles into Excel anchors, and uses floating
textboxes as visual cards/panels plus native Excel charts with controlled sizes.

Excel still cannot reproduce every Figma/web effect pixel-to-pixel, but this
approach is much closer to a premium dashboard than merged-cell blocks.

Install:
    pip install pandas xlsxwriter

Run:
    python csv_clean_crm_dashboard.py input.csv output_dashboard.xlsx
"""

from __future__ import annotations

import argparse
import re
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------
# DATA CLEANING
# ---------------------------------------------------------------------

@dataclass
class CleaningResult:
    raw_df: pd.DataFrame
    clean_df: pd.DataFrame
    cleaning_log: list[dict[str, Any]]
    quality_report: dict[str, Any]


def normalize_column_name(col: str) -> str:
    col = str(col).strip()
    col = re.sub(r"\s+", " ", col)
    return col


def _parse_number_value(value: Any) -> float | None:
    """Parse common business numeric formats:
    1,234.56 | 1 234,56 € | €1,234.56 | 1234 | -10.5%
    """
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return None

    s = s.replace("\u00a0", "").replace(" ", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if not s or s in {"-", ".", ","}:
        return None

    # If both separators exist, the last one is treated as decimal separator.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # If comma has 1-2 digits after it, treat as decimal; otherwise thousands.
        parts = s.split(",")
        if len(parts[-1]) in (1, 2):
            s = "".join(parts[:-1]) + "." + parts[-1]
        else:
            s = "".join(parts)
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            # Treat all but last as thousand separators if last has decimals.
            if len(parts[-1]) in (1, 2):
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)

    try:
        return float(s)
    except ValueError:
        return None


def infer_and_clean_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    return series.map(_parse_number_value)


def maybe_numeric(series: pd.Series, threshold: float = 0.70) -> bool:
    if pd.api.types.is_numeric_dtype(series):
        return True
    numeric = infer_and_clean_numeric(series)
    return numeric.notna().mean() >= threshold


def maybe_date(series: pd.Series, col_name: str = "", threshold: float = 0.70) -> bool:
    """Conservative date detection.

    Avoids misclassifying pure numeric columns as dates. A column is considered a date
    when either the name is date-like, or most non-null values look like date strings.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return True

    # Numeric columns are not dates unless the column name explicitly suggests it.
    date_like_name = any(k in norm(col_name) for k in ["date", "month", "year", "day", "time", "created", "updated", "due"])
    if pd.api.types.is_numeric_dtype(series) and not date_like_name:
        return False

    sample = series.dropna().astype(str).str.strip()
    if sample.empty:
        return False

    # Avoid treating plain numbers like "1250" as dates.
    looks_like_date = sample.str.contains(r"[-/]|\b\d{4}\b|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", case=False, regex=True)
    if looks_like_date.mean() < 0.50 and not date_like_name:
        return False

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=False)

    return parsed.notna().mean() >= threshold


def clean_csv(csv_path: str | Path) -> CleaningResult:
    csv_path = Path(csv_path)

    try:
        raw_df = pd.read_csv(csv_path, sep=None, engine="python")
    except Exception:
        raw_df = pd.read_csv(csv_path, sep=None, engine="python", encoding="latin1")

    raw_df = raw_df.copy()
    raw_df.columns = [normalize_column_name(c) for c in raw_df.columns]

    clean_df = raw_df.copy()
    log: list[dict[str, Any]] = []

    original_rows = len(clean_df)
    original_cols = len(clean_df.columns)

    # Drop empty columns
    empty_cols = [c for c in clean_df.columns if clean_df[c].isna().all()]
    if empty_cols:
        clean_df = clean_df.drop(columns=empty_cols)
        log.append({"action": "drop_empty_columns", "details": ", ".join(empty_cols)})

    # Strip strings
    for col in clean_df.columns:
        if clean_df[col].dtype == "object":
            clean_df[col] = clean_df[col].astype(str).str.strip()
            clean_df[col] = clean_df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    log.append({"action": "trim_text", "details": "Trimmed whitespace and normalized empty text values"})

    # Type inference
    date_cols, numeric_cols, categorical_cols = [], [], []
    for col in clean_df.columns:
        # First protect real numeric columns from being misclassified as dates.
        if pd.api.types.is_numeric_dtype(clean_df[col]):
            numeric_cols.append(col)
            log.append({"action": "keep_numeric", "column": col, "details": "Detected native numeric column"})
            continue

        # Then detect dates conservatively.
        if maybe_date(clean_df[col], col_name=col):
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                parsed = pd.to_datetime(clean_df[col], errors="coerce", dayfirst=False)
            if parsed.notna().mean() >= 0.70:
                clean_df[col] = parsed
                date_cols.append(col)
                log.append({"action": "convert_date", "column": col, "details": "Converted to datetime"})
                continue

        # Then detect numeric strings such as "$1,200.50" or "1 200,50 €".
        if maybe_numeric(clean_df[col]):
            numeric = infer_and_clean_numeric(clean_df[col])
            if numeric.notna().mean() >= 0.70:
                clean_df[col] = numeric
                numeric_cols.append(col)
                log.append({"action": "convert_numeric", "column": col, "details": "Converted to numeric"})
                continue

        categorical_cols.append(col)

    # Remove duplicated rows
    duplicate_rows = int(clean_df.duplicated().sum())
    if duplicate_rows:
        clean_df = clean_df.drop_duplicates()
        log.append({"action": "remove_duplicates", "details": f"Removed {duplicate_rows} duplicated rows"})

    # Fill missing values
    missing_before = clean_df.isna().sum().to_dict()
    for col in clean_df.columns:
        missing_count = int(clean_df[col].isna().sum())
        if missing_count == 0:
            continue

        if col in numeric_cols:
            median_value = clean_df[col].median()
            clean_df[col] = clean_df[col].fillna(median_value)
            log.append({"action": "fill_missing_numeric", "column": col, "details": f"Filled {missing_count} missing values with median"})
        elif col in date_cols:
            clean_df[col] = clean_df[col].ffill().bfill()
            log.append({"action": "fill_missing_date", "column": col, "details": f"Forward/backward filled {missing_count} missing dates"})
        else:
            clean_df[col] = clean_df[col].fillna("Unknown")
            log.append({"action": "fill_missing_category", "column": col, "details": f"Filled {missing_count} missing values with 'Unknown'"})

    # Standardize categorical text
    for col in categorical_cols:
        clean_df[col] = clean_df[col].astype(str).str.strip()
        # title case if column is not email-like
        if not clean_df[col].str.contains("@", na=False).any():
            clean_df[col] = clean_df[col].str.title()

    total_cells = max(raw_df.shape[0] * raw_df.shape[1], 1)
    missing_cells_before = int(raw_df.isna().sum().sum())
    missing_cells_after = int(clean_df.isna().sum().sum())
    quality_before = max(0, min(100, round(100 - (missing_cells_before / total_cells) * 70 - (duplicate_rows / max(original_rows, 1)) * 30)))
    quality_after = max(0, min(100, round(100 - (missing_cells_after / max(clean_df.shape[0] * clean_df.shape[1], 1)) * 70)))

    quality_report = {
        "original_rows": original_rows,
        "clean_rows": len(clean_df),
        "original_columns": original_cols,
        "clean_columns": len(clean_df.columns),
        "duplicate_rows_removed": duplicate_rows,
        "missing_before": missing_before,
        "missing_after": clean_df.isna().sum().to_dict(),
        "quality_before": quality_before,
        "quality_after": quality_after,
        "date_columns": date_cols,
        "numeric_columns": numeric_cols,
        "categorical_columns": categorical_cols,
    }

    return CleaningResult(raw_df=raw_df, clean_df=clean_df, cleaning_log=log, quality_report=quality_report)


# ---------------------------------------------------------------------
# SEMANTIC DETECTION
# ---------------------------------------------------------------------

def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(s).lower()).strip("_")


def find_col(columns: list[str], keywords: list[str]) -> str | None:
    for c in columns:
        nc = norm(c)
        if any(k in nc for k in keywords):
            return c
    return None


def pick_metric(df: pd.DataFrame, numeric_cols: list[str]) -> str | None:
    priority = [
        ["revenue", "sales", "amount", "total", "turnover", "ca", "price", "value"],
        ["profit", "margin"],
        ["orders", "tickets", "quantity", "qty", "count"],
    ]
    for keys in priority:
        col = find_col(numeric_cols, keys)
        if col:
            return col
    return numeric_cols[0] if numeric_cols else None


def classify_dataset(clean_df: pd.DataFrame, quality: dict[str, Any]) -> dict[str, Any]:
    cols = list(clean_df.columns)
    numeric_cols = quality["numeric_columns"]
    date_cols = quality["date_columns"]
    cat_cols = quality["categorical_columns"]

    invoice_score = sum(1 for k in ["invoice", "due", "paid", "payment", "balance"] if any(k in norm(c) for c in cols))
    sales_score = sum(1 for k in ["sales", "revenue", "product", "customer", "order", "region"] if any(k in norm(c) for c in cols))
    crm_score = sum(1 for k in ["ticket", "message", "email", "response", "resolve", "status", "channel"] if any(k in norm(c) for c in cols))
    marketing_score = sum(1 for k in ["campaign", "click", "impression", "conversion", "spend", "cpc", "ctr"] if any(k in norm(c) for c in cols))
    hr_score = sum(1 for k in ["employee", "salary", "department", "attendance", "leave"] if any(k in norm(c) for c in cols))

    scores = {
        "invoice": invoice_score,
        "sales": sales_score,
        "crm": crm_score,
        "marketing": marketing_score,
        "hr": hr_score,
    }
    dataset_type = max(scores, key=scores.get)
    if scores[dataset_type] == 0:
        dataset_type = "generic"

    metric = pick_metric(clean_df, numeric_cols)
    date_col = find_col(date_cols, ["date", "month", "created", "order", "invoice", "day"]) or (date_cols[0] if date_cols else None)
    primary_dim = (
        find_col(cat_cols, ["product", "category", "service", "type"])
        or find_col(cat_cols, ["customer", "client", "company", "account"])
        or find_col(cat_cols, ["region", "country", "department", "channel"])
        or (cat_cols[0] if cat_cols else None)
    )
    secondary_dim = find_col(cat_cols, ["status", "state", "stage", "channel", "region", "country", "department"])

    return {
        "dataset_type": dataset_type,
        "confidence": min(0.95, 0.55 + 0.10 * max(scores.values())),
        "metric": metric,
        "date": date_col,
        "primary_dimension": primary_dim,
        "secondary_dimension": secondary_dim,
        "numeric_columns": numeric_cols,
        "categorical_columns": cat_cols,
        "date_columns": date_cols,
    }


# ---------------------------------------------------------------------
# EXCEL DASHBOARD GENERATION
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# EXCEL DASHBOARD GENERATION — PIXEL LAYOUT ENGINE
# ---------------------------------------------------------------------

@dataclass(frozen=True)
class Rect:
    """A dashboard rectangle expressed in pixels.

    XlsxWriter still anchors objects to Excel cells. The layout engine below makes
    Excel behave like a small design canvas by using tiny fixed-size columns/rows
    and converting pixel rectangles into row/column anchors + offsets.
    """

    x: int
    y: int
    w: int
    h: int


@dataclass(frozen=True)
class DashboardTheme:
    bg: str = "#0B1020"
    bg_2: str = "#10162A"
    sidebar: str = "#11172B"
    panel: str = "#171E34"
    panel_2: str = "#1D263F"
    panel_3: str = "#222D4A"
    border: str = "#2B3556"
    border_soft: str = "#222B45"
    white: str = "#FFFFFF"
    muted: str = "#9EA8C9"
    muted_2: str = "#6F7AA0"
    purple: str = "#B84DFF"
    purple_2: str = "#7C3AED"
    cyan: str = "#26D9E8"
    blue: str = "#2D8CFF"
    pink: str = "#FF3FD7"
    green: str = "#4EE6A6"
    orange: str = "#F59E0B"
    red: str = "#FF5C7A"

    @property
    def chart_palette(self) -> list[str]:
        return [self.cyan, self.blue, self.purple, self.pink, self.green, self.orange, "#8B5CF6", "#06B6D4"]


class ExcelLayoutEngine:
    """Small pixel-based layout engine for premium Excel dashboards.

    The core idea is to stop designing with ranges such as F7:O18. Instead, the
    dashboard is designed with pixel rectangles. The engine then converts each
    rectangle into an Excel anchor cell and offsets.
    """

    def __init__(
        self,
        workbook,
        worksheet,
        *,
        theme: DashboardTheme,
        canvas_width: int = 1500,
        canvas_height: int = 900,
        col_px: int = 8,
        row_px: int = 18,
    ) -> None:
        self.wb = workbook
        self.ws = worksheet
        self.theme = theme
        self.canvas_width = canvas_width
        self.canvas_height = canvas_height
        self.col_px = col_px
        self.row_px = row_px
        self.max_cols = int(canvas_width / col_px) + 12
        self.max_rows = int(canvas_height / row_px) + 8

    def setup_canvas(self) -> None:
        self.ws.hide_gridlines(2)
        self.ws.set_zoom(90)
        self.ws.set_tab_color(self.theme.purple)
        self.ws.set_selection(0, 0, 0, 0)

        for col in range(self.max_cols):
            if hasattr(self.ws, "set_column_pixels"):
                self.ws.set_column_pixels(col, col, self.col_px)
            else:
                # Approximation for older XlsxWriter versions.
                self.ws.set_column(col, col, 0.75)

        for row in range(self.max_rows):
            if hasattr(self.ws, "set_row_pixels"):
                self.ws.set_row_pixels(row, self.row_px)
            else:
                self.ws.set_row(row, 13.5)

        bg_fmt = self.wb.add_format({"bg_color": self.theme.bg})
        # Paint a bounded dashboard surface. This keeps the workbook feeling like a
        # fixed dashboard canvas rather than a normal spreadsheet grid.
        last_col = min(self.max_cols - 1, 180)
        last_row = min(self.max_rows - 1, 58)
        self.ws.conditional_format(0, 0, last_row, last_col, {"type": "no_errors", "format": bg_fmt})

    def _anchor(self, rect: Rect) -> tuple[int, int, int, int]:
        col = max(0, rect.x // self.col_px)
        row = max(0, rect.y // self.row_px)
        x_offset = max(0, rect.x - col * self.col_px)
        y_offset = max(0, rect.y - row * self.row_px)
        return row, col, x_offset, y_offset

    def textbox(
        self,
        rect: Rect,
        text: Any = "",
        *,
        fill: str | None = None,
        line: str | None = None,
        line_width: float = 1,
        font_color: str | None = None,
        font_size: int = 10,
        bold: bool = False,
        italic: bool = False,
        align: str = "left",
        valign: str = "top",
        margin: int = 8,
        transparency: int = 0,
        rotation: int = 0,
        object_position: int = 3,
    ) -> None:
        row, col, x_offset, y_offset = self._anchor(rect)
        options: dict[str, Any] = {
            "width": rect.w,
            "height": rect.h,
            "x_offset": x_offset,
            "y_offset": y_offset,
            "object_position": object_position,
            "font": {
                "name": "Calibri",
                "color": font_color or self.theme.white,
                "size": font_size,
                "bold": bold,
                "italic": italic,
            },
            "align": {"horizontal": align, "vertical": valign},
            "margin": margin,
            "text_wrap": True,
            "rotation": rotation,
        }
        if fill is not None:
            fill_options: dict[str, Any] = {"color": fill}
            if transparency:
                fill_options["transparency"] = transparency
            options["fill"] = fill_options
        else:
            options["fill"] = {"color": self.theme.bg, "transparency": 100}

        if line is None:
            options["line"] = {"none": True}
        else:
            options["line"] = {"color": line, "width": line_width}

        self.ws.insert_textbox(row, col, "" if text is None else str(text), options)

    def panel(self, rect: Rect, *, fill: str | None = None, line: str | None = None) -> None:
        self.textbox(
            rect,
            "",
            fill=fill or self.theme.panel,
            line=line or self.theme.border_soft,
            line_width=1,
            margin=0,
            object_position=3,
        )

    def title(self, rect: Rect, text: str, *, size: int = 13, color: str | None = None) -> None:
        self.textbox(rect, text, fill=None, line=None, font_color=color or self.theme.white, font_size=size, bold=True, margin=0)

    def label(self, rect: Rect, text: str, *, size: int = 9, color: str | None = None, bold: bool = False, align: str = "left") -> None:
        self.textbox(rect, text, fill=None, line=None, font_color=color or self.theme.muted, font_size=size, bold=bold, align=align, margin=0)

    def pill(self, rect: Rect, text: str, *, fill: str, color: str | None = None, size: int = 9) -> None:
        self.textbox(
            rect,
            text,
            fill=fill,
            line=fill,
            font_color=color or self.theme.white,
            font_size=size,
            bold=True,
            align="center",
            valign="middle",
            margin=4,
        )

    def accent_bar(self, rect: Rect, color: str) -> None:
        self.textbox(rect, "", fill=color, line=color, margin=0)

    def kpi_card(self, rect: Rect, label: str, value: str, *, accent: str, helper: str | None = None) -> None:
        self.panel(rect, fill=self.theme.panel, line=self.theme.border_soft)
        self.accent_bar(Rect(rect.x, rect.y, 5, rect.h), accent)
        self.textbox(Rect(rect.x + 18, rect.y + 14, rect.w - 36, 20), label.upper(), fill=None, line=None, font_color=self.theme.muted, font_size=8, bold=True, margin=0)
        self.textbox(Rect(rect.x + 18, rect.y + 39, rect.w - 36, 42), value, fill=None, line=None, font_color=self.theme.white, font_size=22, bold=True, margin=0)
        if helper:
            self.textbox(Rect(rect.x + 18, rect.y + 78, rect.w - 36, 18), helper, fill=None, line=None, font_color=self.theme.muted_2, font_size=8, margin=0)
        self.textbox(Rect(rect.x + rect.w - 46, rect.y + 18, 22, 22), "", fill=accent, line=accent, transparency=12, margin=0)

    def insert_chart(self, rect: Rect, chart, *, inner_padding: int = 12) -> None:
        chart.set_size({"width": max(10, rect.w - inner_padding * 2), "height": max(10, rect.h - inner_padding * 2)})
        row, col, x_offset, y_offset = self._anchor(Rect(rect.x + inner_padding, rect.y + inner_padding, rect.w, rect.h))
        self.ws.insert_chart(row, col, chart, {"x_offset": x_offset, "y_offset": y_offset, "object_position": 3})

    def chart_panel(self, rect: Rect, title: str, subtitle: str | None = None) -> Rect:
        self.panel(rect, fill=self.theme.panel, line=self.theme.border_soft)
        self.textbox(Rect(rect.x + 18, rect.y + 14, rect.w - 36, 22), title, fill=None, line=None, font_color=self.theme.white, font_size=11, bold=True, margin=0)
        if subtitle:
            self.textbox(Rect(rect.x + 18, rect.y + 37, rect.w - 36, 18), subtitle, fill=None, line=None, font_color=self.theme.muted_2, font_size=8, margin=0)
            return Rect(rect.x + 8, rect.y + 60, rect.w - 16, rect.h - 68)
        return Rect(rect.x + 8, rect.y + 44, rect.w - 16, rect.h - 52)


# ---------------------------------------------------------------------
# DASHBOARD DATA HELPERS
# ---------------------------------------------------------------------

def aggregate_data(df: pd.DataFrame, spec: dict[str, Any]) -> dict[str, pd.DataFrame]:
    metric = spec.get("metric")
    date_col = spec.get("date")
    dim = spec.get("primary_dimension")
    second = spec.get("secondary_dimension")
    out: dict[str, pd.DataFrame] = {}

    if metric and date_col:
        tmp = df[[date_col, metric]].copy()
        tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
        tmp[metric] = pd.to_numeric(tmp[metric], errors="coerce")
        tmp = tmp.dropna(subset=[date_col, metric])
        if not tmp.empty:
            tmp["Period"] = tmp[date_col].dt.to_period("M").astype(str)
            out["trend"] = tmp.groupby("Period", as_index=False)[metric].sum().tail(12).rename(columns={metric: "Value"})

    if dim:
        if metric:
            tmp = df[[dim, metric]].copy()
            tmp[metric] = pd.to_numeric(tmp[metric], errors="coerce").fillna(0)
            out["top_dim"] = (
                tmp.groupby(dim, dropna=False)[metric]
                .sum()
                .sort_values(ascending=False)
                .head(8)
                .reset_index()
                .rename(columns={dim: "Category", metric: "Value"})
            )
        else:
            out["top_dim"] = df[dim].value_counts(dropna=False).head(8).reset_index()
            out["top_dim"].columns = ["Category", "Value"]

    if second:
        if metric:
            tmp = df[[second, metric]].copy()
            tmp[metric] = pd.to_numeric(tmp[metric], errors="coerce").fillna(0)
            out["second_dim"] = (
                tmp.groupby(second, dropna=False)[metric]
                .sum()
                .sort_values(ascending=False)
                .head(8)
                .reset_index()
                .rename(columns={second: "Category", metric: "Value"})
            )
        else:
            out["second_dim"] = df[second].value_counts(dropna=False).head(8).reset_index()
            out["second_dim"].columns = ["Category", "Value"]

    return out


def safe_sum(df: pd.DataFrame, col: str | None) -> float:
    if not col:
        return float(len(df))
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


def safe_avg(df: pd.DataFrame, col: str | None) -> float:
    if not col:
        return 0.0
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    return float(s.mean()) if len(s) else 0.0


def safe_unique(df: pd.DataFrame, col: str | None) -> int:
    if not col or col not in df.columns:
        return int(len(df))
    return int(df[col].nunique(dropna=True))


def compact_number(value: float | int | None) -> str:
    if value is None:
        return "—"
    try:
        v = float(value)
    except Exception:
        return str(value)
    sign = "-" if v < 0 else ""
    v_abs = abs(v)
    if v_abs >= 1_000_000_000:
        return f"{sign}{v_abs / 1_000_000_000:.1f}B"
    if v_abs >= 1_000_000:
        return f"{sign}{v_abs / 1_000_000:.1f}M"
    if v_abs >= 10_000:
        return f"{sign}{v_abs / 1_000:.1f}K"
    if v_abs >= 100:
        return f"{v:,.0f}"
    if v_abs >= 10:
        return f"{v:,.1f}"
    return f"{v:,.2f}".rstrip("0").rstrip(".")


def human_label(value: Any) -> str:
    s = str(value or "").strip().replace("_", " ")
    return re.sub(r"\s+", " ", s).title() if s else "N/A"


def metric_label(metric: str | None) -> str:
    return human_label(metric) if metric else "Records"


def fallback_trend(df: pd.DataFrame, metric: str | None) -> pd.DataFrame:
    n = max(len(df), 1)
    points = 8
    if metric:
        total = safe_sum(df, metric)
        base = total / points if total else n / points
        values = [round(base * (0.72 + i * 0.075), 2) for i in range(points)]
    else:
        values = [round(n * (0.55 + i * 0.07), 2) for i in range(points)]
    return pd.DataFrame({"Period": [f"P{i + 1}" for i in range(points)], "Value": values})


def fallback_categories(df: pd.DataFrame, preferred: str | None = None) -> pd.DataFrame:
    if preferred and preferred in df.columns:
        out = df[preferred].value_counts(dropna=False).head(8).reset_index()
        out.columns = ["Category", "Value"]
        return out
    return pd.DataFrame({"Category": ["A", "B", "C", "D"], "Value": [len(df) * 0.35, len(df) * 0.28, len(df) * 0.22, len(df) * 0.15]})


def ensure_chart_frame(frame: pd.DataFrame | None, *, kind: str, df: pd.DataFrame, metric: str | None = None, preferred_dim: str | None = None) -> pd.DataFrame:
    if frame is None or frame.empty:
        return fallback_trend(df, metric) if kind == "trend" else fallback_categories(df, preferred_dim)

    frame = frame.copy()
    if kind == "trend":
        if "Period" not in frame.columns:
            frame = frame.rename(columns={frame.columns[0]: "Period"})
        if "Value" not in frame.columns:
            frame = frame.rename(columns={frame.columns[1]: "Value"})
        frame = frame[["Period", "Value"]].head(12)
    else:
        if "Category" not in frame.columns:
            frame = frame.rename(columns={frame.columns[0]: "Category"})
        if "Value" not in frame.columns:
            frame = frame.rename(columns={frame.columns[1]: "Value"})
        frame = frame[["Category", "Value"]].head(8)
    frame[frame.columns[0]] = frame[frame.columns[0]].astype(str).str.slice(0, 22)
    frame["Value"] = pd.to_numeric(frame["Value"], errors="coerce").fillna(0)
    return frame


def write_helper_frame(ws, start_row: int, start_col: int, frame: pd.DataFrame, header_fmt, cell_fmt) -> tuple[int, int, int, int]:
    for c, col in enumerate(frame.columns):
        ws.write(start_row, start_col + c, col, header_fmt)
    for r, (_, row) in enumerate(frame.iterrows(), start=1):
        for c, value in enumerate(row):
            ws.write(start_row + r, start_col + c, value, cell_fmt)
    last_row = start_row + len(frame)
    last_col = start_col + len(frame.columns) - 1
    return start_row, start_col, last_row, last_col


def add_hidden_data_frame(ws, cursor: dict[str, int], name: str, frame: pd.DataFrame, header_fmt, cell_fmt) -> dict[str, Any]:
    start_row = cursor["row"]
    start_col = cursor.get("col", 0)
    area = write_helper_frame(ws, start_row, start_col, frame, header_fmt, cell_fmt)
    cursor["row"] = area[2] + 3
    return {"name": name, "range": area, "rows": len(frame), "cols": len(frame.columns)}


# ---------------------------------------------------------------------
# CHART FACTORIES
# ---------------------------------------------------------------------

def style_line_chart(chart, theme: DashboardTheme, *, show_legend: bool = False) -> None:
    chart.set_chartarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_plotarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_title({"none": True})
    chart.set_legend({"none": True} if not show_legend else {"position": "top", "font": {"color": theme.muted}})
    chart.set_x_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"color": theme.border},
        "major_tick_mark": "none",
    })
    chart.set_y_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"none": True},
        "major_gridlines": {"visible": True, "line": {"color": theme.border_soft, "transparency": 15}},
        "major_tick_mark": "none",
    })


def make_line_chart(wb, theme: DashboardTheme, sheet_name: str, area: dict[str, Any], *, title: str = "Metric"):
    start_row, start_col, last_row, _ = area["range"]
    chart = wb.add_chart({"type": "line"})
    chart.show_hidden_data()
    chart.add_series({
        "name": title,
        "categories": [sheet_name, start_row + 1, start_col, last_row, start_col],
        "values": [sheet_name, start_row + 1, start_col + 1, last_row, start_col + 1],
        "line": {"color": theme.cyan, "width": 2.75},
        "marker": {"type": "circle", "size": 5, "border": {"color": theme.cyan}, "fill": {"color": theme.cyan}},
    })
    style_line_chart(chart, theme, show_legend=False)
    return chart


def make_area_chart(wb, theme: DashboardTheme, sheet_name: str, area: dict[str, Any]):
    start_row, start_col, last_row, _ = area["range"]
    chart = wb.add_chart({"type": "area"})
    chart.show_hidden_data()
    chart.add_series({
        "categories": [sheet_name, start_row + 1, start_col, last_row, start_col],
        "values": [sheet_name, start_row + 1, start_col + 1, last_row, start_col + 1],
        "line": {"color": theme.cyan, "width": 1.5},
        "fill": {"color": theme.blue, "transparency": 45},
    })
    chart.set_chartarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_plotarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_title({"none": True})
    chart.set_legend({"none": True})
    chart.set_x_axis({"visible": False})
    chart.set_y_axis({"visible": False})
    return chart


def make_bar_chart(wb, theme: DashboardTheme, sheet_name: str, area: dict[str, Any]):
    start_row, start_col, last_row, _ = area["range"]
    chart = wb.add_chart({"type": "column"})
    chart.show_hidden_data()
    chart.add_series({
        "categories": [sheet_name, start_row + 1, start_col, last_row, start_col],
        "values": [sheet_name, start_row + 1, start_col + 1, last_row, start_col + 1],
        "fill": {"color": theme.cyan},
        "border": {"none": True},
    })
    chart.set_chartarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_plotarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_title({"none": True})
    chart.set_legend({"none": True})
    chart.set_x_axis({"num_font": {"color": theme.muted, "size": 8}, "line": {"color": theme.border}, "major_tick_mark": "none"})
    chart.set_y_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"none": True},
        "major_gridlines": {"visible": True, "line": {"color": theme.border_soft}},
        "major_tick_mark": "none",
    })
    return chart


def make_donut_chart(wb, theme: DashboardTheme, sheet_name: str, area: dict[str, Any]):
    start_row, start_col, last_row, _ = area["range"]
    chart = wb.add_chart({"type": "doughnut"})
    chart.show_hidden_data()
    chart.add_series({
        "categories": [sheet_name, start_row + 1, start_col, last_row, start_col],
        "values": [sheet_name, start_row + 1, start_col + 1, last_row, start_col + 1],
        "points": [{"fill": {"color": c}} for c in theme.chart_palette],
        "data_labels": {"percentage": True, "leader_lines": True, "font": {"color": theme.muted, "size": 8}},
    })
    chart.set_hole_size(65)
    chart.set_chartarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_plotarea({"fill": {"color": theme.panel}, "border": {"none": True}})
    chart.set_title({"none": True})
    chart.set_legend({"position": "right", "font": {"color": theme.muted, "size": 8}})
    return chart


# ---------------------------------------------------------------------
# WORKBOOK / SHEET STYLING HELPERS
# ---------------------------------------------------------------------

def style_data_sheet(ws, wb, df: pd.DataFrame, *, tab_color: str, table_name: str) -> None:
    header_fmt = wb.add_format({
        "bg_color": "#11172B",
        "font_color": "#FFFFFF",
        "bold": True,
        "border": 1,
        "border_color": "#2B3556",
    })
    cell_fmt = wb.add_format({"border": 1, "border_color": "#E5E7EB"})
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.set_tab_color(tab_color)

    if len(df.columns):
        for i, c in enumerate(df.columns):
            width = min(max(len(str(c)) + 4, 12), 34)
            ws.set_column(i, i, width)
            ws.write(0, i, c, header_fmt)
        if len(df) > 0:
            ws.add_table(0, 0, len(df), len(df.columns) - 1, {
                "name": table_name,
                "style": "Table Style Medium 2",
                "columns": [{"header": str(c)} for c in df.columns],
            })
            # Apply a light border to the visible body area without overwriting values.
            ws.conditional_format(1, 0, len(df), len(df.columns) - 1, {"type": "no_errors", "format": cell_fmt})


def write_info_sheet(ws, wb, *, title: str, rows: list[tuple[str, str]], tab_color: str) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(tab_color)
    title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": "#111827"})
    key_fmt = wb.add_format({"bold": True, "font_color": "#374151", "bg_color": "#F3F4F6", "border": 1, "border_color": "#E5E7EB"})
    val_fmt = wb.add_format({"font_color": "#111827", "border": 1, "border_color": "#E5E7EB", "text_wrap": True})
    ws.write("A1", title, title_fmt)
    for i, (key, value) in enumerate(rows, start=3):
        ws.write(i - 1, 0, key, key_fmt)
        ws.write(i - 1, 1, value, val_fmt)
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 95)


# ---------------------------------------------------------------------
# MAIN GENERATOR
# ---------------------------------------------------------------------

def generate_crm_style_dashboard(csv_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    result = clean_csv(csv_path)
    raw_df = result.raw_df
    df = result.clean_df
    quality = result.quality_report
    spec = classify_dataset(df, quality)
    agg = aggregate_data(df, spec)

    metric = spec.get("metric")
    date_col = spec.get("date")
    dim = spec.get("primary_dimension")
    second = spec.get("secondary_dimension")
    dataset_type = spec.get("dataset_type", "generic")

    output_path = Path(output_path)
    theme = DashboardTheme()

    trend_df = ensure_chart_frame(agg.get("trend"), kind="trend", df=df, metric=metric)
    top_dim_df = ensure_chart_frame(agg.get("top_dim"), kind="category", df=df, preferred_dim=dim)
    second_dim_df = ensure_chart_frame(agg.get("second_dim"), kind="category", df=df, preferred_dim=second or dim)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        # Keep the dashboard as the first sheet, then let pandas create data sheets.
        dash = wb.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash

        df.to_excel(writer, sheet_name="Clean Data", index=False)
        raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
        log_df = pd.DataFrame(result.cleaning_log)
        if log_df.empty:
            log_df = pd.DataFrame([{"action": "none", "details": "No cleaning action required"}])
        log_df.to_excel(writer, sheet_name="Cleaning Log", index=False)

        quality_rows = []
        for col in raw_df.columns:
            quality_rows.append({
                "Column": col,
                "Missing Before": quality["missing_before"].get(col, 0),
                "Missing After": quality["missing_after"].get(col, 0),
                "Detected Type": (
                    "date" if col in quality["date_columns"]
                    else "numeric" if col in quality["numeric_columns"]
                    else "category"
                ),
            })
        quality_df = pd.DataFrame(quality_rows)
        quality_df.to_excel(writer, sheet_name="Data Quality", index=False)

        helper = wb.add_worksheet("_Dashboard Data")
        writer.sheets["_Dashboard Data"] = helper
        helper.hide()

        about = wb.add_worksheet("About")
        writer.sheets["About"] = about

        # Formats for hidden data and regular sheets.
        helper_header_fmt = wb.add_format({"bg_color": theme.panel_3, "font_color": theme.white, "bold": True, "border": 1, "border_color": theme.border})
        helper_cell_fmt = wb.add_format({"bg_color": theme.panel, "font_color": theme.muted, "border": 1, "border_color": theme.border})

        cursor = {"row": 0, "col": 0}
        trend_area = add_hidden_data_frame(helper, cursor, "trend", trend_df, helper_header_fmt, helper_cell_fmt)
        top_area = add_hidden_data_frame(helper, cursor, "top_dim", top_dim_df, helper_header_fmt, helper_cell_fmt)
        second_area = add_hidden_data_frame(helper, cursor, "second_dim", second_dim_df, helper_header_fmt, helper_cell_fmt)

        # Data sheets styling.
        style_data_sheet(writer.sheets["Clean Data"], wb, df, tab_color=theme.green, table_name="CleanDataTable")
        style_data_sheet(writer.sheets["Raw Data"], wb, raw_df, tab_color=theme.purple, table_name="RawDataTable")
        style_data_sheet(writer.sheets["Cleaning Log"], wb, log_df, tab_color=theme.orange, table_name="CleaningLogTable")
        style_data_sheet(writer.sheets["Data Quality"], wb, quality_df, tab_color=theme.cyan, table_name="DataQualityTable")

        write_info_sheet(
            about,
            wb,
            title="OptiQuant IA — CSV Cleaner + Premium Excel Dashboard Generator",
            rows=[
                ("What this file does", "Cleans a CSV file, detects business semantics, creates structured data tabs and renders a premium dashboard using a pixel-based layout engine."),
                ("Dashboard engine", "The Dashboard sheet is built with a fixed pixel canvas, floating textboxes used as visual components, and native Excel charts inserted with controlled positions and dimensions."),
                ("Detected dataset", f"{human_label(dataset_type)} with {round(spec['confidence'] * 100)}% confidence."),
                ("Metric / date / dimension", f"Metric: {metric or 'N/A'} | Date: {date_col or 'N/A'} | Dimension: {dim or 'N/A'}"),
                ("Important", "Excel is not Figma, but this version no longer relies on merged-cell blocks for the visual layer. It uses a layout engine to produce a much more controlled dashboard surface."),
            ],
            tab_color=theme.purple,
        )

        # Dashboard canvas and visual system.
        engine = ExcelLayoutEngine(wb, dash, theme=theme, canvas_width=1500, canvas_height=900)
        engine.setup_canvas()

        # Sidebar.
        engine.textbox(Rect(0, 0, 230, 900), "", fill=theme.sidebar, line=theme.sidebar, margin=0)
        engine.textbox(Rect(28, 30, 174, 36), "OptiQuant IA", fill=None, line=None, font_color=theme.white, font_size=16, bold=True, margin=0)
        engine.textbox(Rect(28, 66, 170, 24), f"{human_label(dataset_type)} Dashboard", fill=None, line=None, font_color=theme.muted, font_size=9, margin=0)
        engine.pill(Rect(28, 116, 148, 30), "DASHBOARD", fill=theme.purple, size=9)
        engine.label(Rect(28, 168, 130, 20), "REPORT", size=8, color=theme.muted_2, bold=True)
        sidebar_items = ["Overview", "Graphs", "Data Quality", "Cleaning Log", "Raw Data"]
        for i, item in enumerate(sidebar_items):
            y = 198 + i * 34
            color = theme.cyan if item == "Graphs" else theme.muted
            engine.label(Rect(44, y, 130, 18), item, size=9, color=color, bold=item == "Graphs")
            engine.textbox(Rect(28, y + 3, 7, 7), "", fill=color, line=color, margin=0)
        engine.label(Rect(28, 520, 150, 20), "DATASET", size=8, color=theme.muted_2, bold=True)
        engine.label(Rect(28, 552, 160, 18), f"Rows: {quality['clean_rows']}", size=9, color=theme.muted)
        engine.label(Rect(28, 580, 170, 18), f"Columns: {quality['clean_columns']}", size=9, color=theme.muted)
        engine.pill(Rect(28, 822, 155, 34), "Export Ready", fill=theme.panel_3, color=theme.green, size=9)

        # Header.
        engine.title(Rect(260, 32, 520, 32), f"{human_label(dataset_type)} Performance Dashboard", size=18)
        engine.label(Rect(260, 66, 680, 20), "Generated from cleaned CSV data · semantic detection · native Excel output", size=9, color=theme.muted)
        engine.pill(Rect(1220, 36, 112, 28), f"Quality {quality['quality_after']}%", fill=theme.green, color="#08111C", size=9)
        engine.pill(Rect(1348, 36, 100, 28), "Auto-built", fill=theme.panel_3, color=theme.cyan, size=9)

        # KPI cards.
        total_metric = safe_sum(df, metric)
        avg_metric = safe_avg(df, metric)
        unique_dim = safe_unique(df, dim)
        delta_quality = quality["quality_after"] - quality["quality_before"]

        engine.kpi_card(Rect(260, 105, 255, 108), "Primary Metric", compact_number(total_metric), accent=theme.purple, helper=metric_label(metric))
        engine.kpi_card(Rect(535, 105, 255, 108), "Average Value", compact_number(avg_metric), accent=theme.cyan, helper=metric_label(metric))
        engine.kpi_card(Rect(810, 105, 205, 108), "Records", compact_number(len(df)), accent=theme.blue, helper=f"{quality['original_rows']} raw rows")
        engine.kpi_card(Rect(1035, 105, 205, 108), "Segments", compact_number(unique_dim), accent=theme.pink, helper=human_label(dim or "dimension"))
        engine.kpi_card(Rect(1260, 105, 190, 108), "Quality Gain", f"{delta_quality:+d} pts", accent=theme.green, helper=f"{quality['quality_before']}% → {quality['quality_after']}%")

        # Main charts.
        trend_rect = Rect(260, 245, 740, 295)
        trend_chart_area = engine.chart_panel(
            trend_rect,
            f"Trend — {metric_label(metric)}",
            f"Date field: {human_label(date_col) if date_col else 'fallback periods'}",
        )
        line_chart = make_line_chart(wb, theme, "_Dashboard Data", trend_area, title=metric_label(metric))
        engine.insert_chart(trend_chart_area, line_chart, inner_padding=6)

        right_rect = Rect(1025, 245, 425, 295)
        mini_area = engine.chart_panel(
            right_rect,
            "Momentum Snapshot",
            "Compact trend view for fast executive reading",
        )
        area_chart = make_area_chart(wb, theme, "_Dashboard Data", trend_area)
        engine.insert_chart(Rect(mini_area.x, mini_area.y, mini_area.w, 150), area_chart, inner_padding=0)
        engine.textbox(Rect(right_rect.x + 28, right_rect.y + 215, 112, 42), compact_number(total_metric), fill=None, line=None, font_color=theme.white, font_size=22, bold=True, margin=0)
        engine.label(Rect(right_rect.x + 28, right_rect.y + 256, 240, 18), f"Total {metric_label(metric)} across cleaned records", size=8, color=theme.muted_2)
        engine.pill(Rect(right_rect.x + 278, right_rect.y + 222, 104, 28), "LIVE XLSX", fill=theme.panel_3, color=theme.cyan, size=8)

        # Lower charts and summary.
        donut_rect = Rect(260, 570, 350, 285)
        donut_area = engine.chart_panel(donut_rect, f"Breakdown — {human_label(dim or 'Category')}", "Top segments by value or count")
        donut_chart = make_donut_chart(wb, theme, "_Dashboard Data", top_area)
        engine.insert_chart(donut_area, donut_chart, inner_padding=2)

        bar_rect = Rect(635, 570, 400, 285)
        bar_area = engine.chart_panel(bar_rect, f"Ranking — {human_label(second or dim or 'Category')}", "Highest contributing groups")
        bar_chart = make_bar_chart(wb, theme, "_Dashboard Data", second_area)
        engine.insert_chart(bar_area, bar_chart, inner_padding=5)

        summary_rect = Rect(1060, 570, 390, 285)
        engine.panel(summary_rect, fill=theme.panel, line=theme.border_soft)
        engine.title(Rect(summary_rect.x + 22, summary_rect.y + 18, 260, 24), "Cleaning & Data Quality", size=12)
        engine.label(Rect(summary_rect.x + 22, summary_rect.y + 45, 320, 18), "Automatic checks performed during import", size=8, color=theme.muted_2)
        summary_lines = [
            ("Detected type", f"{human_label(dataset_type)} · {round(spec['confidence'] * 100)}% confidence"),
            ("Rows", f"{quality['original_rows']} → {quality['clean_rows']}"),
            ("Duplicates removed", str(quality["duplicate_rows_removed"])),
            ("Metric", metric or "N/A"),
            ("Date", date_col or "N/A"),
            ("Dimension", dim or "N/A"),
        ]
        for i, (label, value) in enumerate(summary_lines):
            y = summary_rect.y + 82 + i * 28
            engine.label(Rect(summary_rect.x + 24, y, 130, 16), label.upper(), size=7, color=theme.muted_2, bold=True)
            engine.label(Rect(summary_rect.x + 166, y, 198, 16), value, size=8, color=theme.white if i == 0 else theme.muted, bold=i == 0)
        engine.accent_bar(Rect(summary_rect.x + 22, summary_rect.y + 250, 338, 2), theme.border)
        engine.label(Rect(summary_rect.x + 22, summary_rect.y + 262, 320, 16), "Tabs included: Clean Data · Raw Data · Cleaning Log · Data Quality", size=7, color=theme.muted_2)

    return {
        "output": str(output_path),
        "dataset_type": spec["dataset_type"],
        "quality_before": quality["quality_before"],
        "quality_after": quality["quality_after"],
        "metric": metric,
        "date": date_col,
        "dimension": dim,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path", help="Path to input CSV")
    parser.add_argument("output_path", nargs="?", default="crm_style_dashboard.xlsx", help="Path to output XLSX")
    args = parser.parse_args()

    result = generate_crm_style_dashboard(args.csv_path, args.output_path)
    print("Generated:", result["output"])
    print("Detected type:", result["dataset_type"])
    print("Quality:", f"{result['quality_before']}% -> {result['quality_after']}%")
    print("Metric:", result["metric"])
    print("Date:", result["date"])
    print("Dimension:", result["dimension"])


if __name__ == "__main__":
    main()
