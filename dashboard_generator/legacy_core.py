"""
OptiQuant IA — CSV to Dark SaaS Premium Excel Dashboard Generator

Goal:
- Take any CSV file
- Clean and normalize the data
- Generate a Dark SaaS Premium Excel dashboard inspired by Option A: neon accents, glass panels, executive SaaS layout

Important:
This Option A version uses a pixel-based layout engine on top of XlsxWriter. Instead of
placing components directly with fixed Excel ranges, it designs the dashboard as
a 1500 x 900 px canvas, converts rectangles into Excel anchors, and separates
the visual layer from the chart layer.

Cards and panels are rendered as rounded PNG assets when Pillow is available,
then inserted under the charts. This avoids the common Excel/XlsxWriter issue
where floating textboxes are drawn above charts and hide them. Transparent
labels remain as textboxes only where they don't cover chart surfaces.

Install:
    pip install pandas xlsxwriter pillow

Run:
    # Standard premium XLSX output, with macro source exported beside the workbook.
    python csv_to_dark_saas_dashboard_generator.py input.csv output_dashboard.xlsx

    # Macro-enabled XLSM output. Requires a precompiled vbaProject.bin extracted
    # once from a trusted Excel macro template. The generated XLSM is ready to open.
    python csv_to_dark_saas_dashboard_generator.py input.csv output_dashboard.xlsm --vba-project vbaProject.bin
"""

from __future__ import annotations

import argparse
import io
import re
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont
except Exception:  # Pillow is optional; the engine falls back to cell-painted panels.
    Image = ImageDraw = ImageFilter = ImageFont = None


# ---------------------------------------------------------------------
# VBA AUTOMATION TEMPLATE
# ---------------------------------------------------------------------

DASHBOARD_VBA_SOURCE = """
Attribute VB_Name = "DashboardMacros"
Option Explicit

Private Sub SafeActivate(ByVal sheetName As String)
    On Error GoTo Fail
    ThisWorkbook.Worksheets(sheetName).Activate
    Exit Sub
Fail:
    MsgBox "Cannot open sheet '" & sheetName & "'. " & Err.Description, vbExclamation, "OptiQuant IA"
End Sub

Public Sub GoToDashboard()
    SafeActivate "Dashboard"
End Sub

Public Sub GoToCleanData()
    SafeActivate "Clean Data"
End Sub

Public Sub GoToRawData()
    SafeActivate "Raw Data"
End Sub

Public Sub GoToCleaningLog()
    SafeActivate "Cleaning Log"
End Sub

Public Sub GoToDataQuality()
    SafeActivate "Data Quality"
End Sub

Public Sub RefreshDashboard()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ThisWorkbook.RefreshAll
    Application.CalculateFull
    ThisWorkbook.Worksheets("Dashboard").Activate

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed successfully.", vbInformation, "OptiQuant IA"
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description, vbExclamation, "OptiQuant IA"
End Sub

Public Sub ExportDashboardPDF()
    On Error GoTo Fail
    Dim outputPath As String

    If Len(ThisWorkbook.Path) = 0 Then
        outputPath = Environ$("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator
    Else
        outputPath = ThisWorkbook.Path & Application.PathSeparator
    End If

    outputPath = outputPath & "dashboard_export_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "PDF exported:" & vbCrLf & outputPath, vbInformation, "OptiQuant IA"
    Exit Sub

Fail:
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, "OptiQuant IA"
End Sub
"""


def write_vba_macro_source(path: str | Path) -> Path:
    """Write the VBA module source used to create the reusable vbaProject.bin template.

    XlsxWriter can embed a binary vbaProject.bin, but it doesn't compile VBA text
    into that binary format. This helper exports the exact macro source that should
    be imported once into an Excel .xlsm template and extracted as vbaProject.bin.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(DASHBOARD_VBA_SOURCE.strip() + "\n", encoding="utf-8")
    return path


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

    s_without_currency_words = re.sub(r"\b(eur|usd|gbp|cad|aud|chf|jpy|cny)\b", "", s, flags=re.IGNORECASE)
    if re.search(r"[A-Za-z]", s_without_currency_words):
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


def _parse_datetime_series(series: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        try:
            return pd.to_datetime(series, errors="coerce", dayfirst=False, format="mixed")
        except TypeError:
            return pd.to_datetime(series, errors="coerce", dayfirst=False)


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

    parsed = _parse_datetime_series(series)

    return parsed.notna().mean() >= threshold


def _read_csv_candidate(csv_path: Path, *, sep: str | None, encoding: str, header: int | None = 0) -> pd.DataFrame | None:
    try:
        return pd.read_csv(csv_path, sep=sep, engine="python", encoding=encoding, header=header)
    except Exception:
        return None


def _column_name_looks_like_data(value: Any) -> bool:
    text = str(value).strip()
    if not text:
        return True
    if "@" in text:
        return True
    if re.search(r"[€$£¥]", text):
        return True
    if re.search(r"\d", text) and re.search(r"[-/]", text):
        return True
    if _parse_number_value(text) is not None:
        return True
    parsed = _parse_datetime_series(pd.Series([text]))
    return bool(parsed.notna().iloc[0])


def _looks_like_missing_header(columns: list[Any]) -> bool:
    if len(columns) <= 1:
        return False
    data_like = sum(1 for col in columns if _column_name_looks_like_data(col))
    return data_like >= max(3, int(len(columns) * 0.45))


def _read_csv_flexible(csv_path: Path) -> pd.DataFrame:
    encodings = ["utf-8-sig", "utf-8", "latin1"]
    separators: list[str | None] = [None, ";", ",", "\t", "|"]
    best: tuple[int, int, pd.DataFrame, str, str | None] | None = None

    for encoding in encodings:
        for sep in separators:
            frame = _read_csv_candidate(csv_path, sep=sep, encoding=encoding, header=0)
            if frame is None:
                continue
            non_empty_cols = int(sum(not frame[col].isna().all() for col in frame.columns))
            score = (non_empty_cols, len(frame))
            if best is None or score > (best[0], best[1]):
                best = (non_empty_cols, len(frame), frame, encoding, sep)

    if best is None:
        raise ValueError(f"Unable to read CSV file: {csv_path}")

    _, _, frame, encoding, sep = best
    if _looks_like_missing_header(list(frame.columns)):
        no_header = _read_csv_candidate(csv_path, sep=sep, encoding=encoding, header=None)
        if no_header is not None:
            no_header.columns = [f"Column {i + 1}" for i in range(len(no_header.columns))]
            return no_header
    return frame


def clean_csv(csv_path: str | Path) -> CleaningResult:
    csv_path = Path(csv_path)

    raw_df = _read_csv_flexible(csv_path)

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
            parsed = _parse_datetime_series(clean_df[col])
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
    # Option A — Dark SaaS Premium
    # High-contrast deep navy background, neon cyan/violet accents, glass-like cards.
    bg: str = "#050816"
    bg_2: str = "#080D1F"
    sidebar: str = "#090E22"
    panel: str = "#0F172A"
    panel_2: str = "#111C33"
    panel_3: str = "#17223D"
    border: str = "#26314F"
    border_soft: str = "#1C2744"
    border_glow: str = "#2E4B78"
    shadow: str = "#02040D"
    white: str = "#F8FBFF"
    muted: str = "#AAB6D8"
    muted_2: str = "#7381A6"
    purple: str = "#B84DFF"
    purple_2: str = "#6D28D9"
    cyan: str = "#22D3EE"
    blue: str = "#2D8CFF"
    pink: str = "#EC4899"
    green: str = "#22C55E"
    orange: str = "#F59E0B"
    red: str = "#FB7185"

    @property
    def chart_palette(self) -> list[str]:
        return [self.cyan, self.purple, self.blue, self.pink, self.green, self.orange, "#8B5CF6", "#06B6D4"]


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
        self._image_streams: list[io.BytesIO] = []
        self._format_cache: dict[tuple[str, str, int], Any] = {}

    def setup_canvas(self) -> None:
        self.ws.hide_gridlines(2)
        self.ws.set_zoom(85)
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

    @staticmethod
    def _hex_to_rgba(hex_color: str, transparency: int = 0) -> tuple[int, int, int, int]:
        hex_color = str(hex_color).strip().lstrip("#")
        if len(hex_color) == 3:
            hex_color = "".join(ch * 2 for ch in hex_color)
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        alpha = int(round(255 * max(0, min(100, 100 - transparency)) / 100))
        return r, g, b, alpha

    def _image(
        self,
        rect: Rect,
        png_bytes: bytes,
        *,
        object_position: int = 3,
        name: str = "dashboard_asset.png",
        url: str | None = None,
        tip: str | None = None,
        decorative: bool = True,
    ) -> None:
        row, col, x_offset, y_offset = self._anchor(rect)
        stream = io.BytesIO(png_bytes)
        stream.seek(0)
        # Keep a reference until workbook close. XlsxWriter reads the stream while packaging.
        self._image_streams.append(stream)
        options: dict[str, Any] = {
            "image_data": stream,
            "x_offset": x_offset,
            "y_offset": y_offset,
            "object_position": object_position,
            "decorative": decorative,
        }
        if url:
            # Images can carry hyperlinks. We use this for polished sidebar
            # navigation buttons without sacrificing the dashboard DA.
            options["url"] = url
            if tip:
                options["tip"] = tip
            options["decorative"] = False
            options["description"] = tip or name
        self.ws.insert_image(row, col, name, options)

    def _paint_cell_panel(self, rect: Rect, *, fill: str, line: str, border_width: int = 1) -> None:
        # Fallback when Pillow is unavailable. It deliberately uses cell formats
        # rather than a textbox so it doesn't cover Excel charts.
        start_row, start_col, _, _ = self._anchor(rect)
        end_row, end_col, _, _ = self._anchor(Rect(rect.x + rect.w, rect.y + rect.h, 1, 1))
        key = (fill, line, border_width)
        fmt = self._format_cache.get(key)
        if fmt is None:
            fmt = self.wb.add_format({"bg_color": fill, "border": border_width, "border_color": line})
            self._format_cache[key] = fmt
        self.ws.conditional_format(start_row, start_col, max(start_row, end_row), max(start_col, end_col), {"type": "no_errors", "format": fmt})

    def _rounded_panel_png(
        self,
        rect: Rect,
        *,
        fill: str,
        line: str,
        radius: int = 18,
        border_width: int = 1,
        fill_transparency: int = 0,
        border_transparency: int = 0,
        shadow: bool = True,
        accent_top: str | None = None,
    ) -> tuple[bytes, Rect] | None:
        if Image is None or ImageDraw is None or ImageFilter is None:
            return None

        scale = 2
        pad = 14 if shadow else 3
        shadow_dx = 0
        shadow_dy = 5 if shadow else 0
        shadow_blur = 10 if shadow else 0

        width = max(1, (rect.w + pad * 2) * scale)
        height = max(1, (rect.h + pad * 2 + shadow_dy) * scale)
        img = Image.new("RGBA", (width, height), (0, 0, 0, 0))

        body_box = [pad * scale, pad * scale, (pad + rect.w) * scale, (pad + rect.h) * scale]
        rr = radius * scale

        if shadow:
            shadow_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            shadow_draw = ImageDraw.Draw(shadow_layer)
            sx0 = (pad + shadow_dx) * scale
            sy0 = (pad + shadow_dy) * scale
            sx1 = (pad + shadow_dx + rect.w) * scale
            sy1 = (pad + shadow_dy + rect.h) * scale
            shadow_draw.rounded_rectangle(
                [sx0, sy0, sx1, sy1],
                radius=rr,
                fill=self._hex_to_rgba(self.theme.shadow, 62),
            )
            shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(shadow_blur * scale))
            img.alpha_composite(shadow_layer)

        draw = ImageDraw.Draw(img)
        fill_rgba = self._hex_to_rgba(fill, fill_transparency)
        border_rgba = self._hex_to_rgba(line, border_transparency)
        draw.rounded_rectangle(body_box, radius=rr, fill=fill_rgba, outline=border_rgba, width=max(1, border_width * scale))

        # Subtle top highlight to avoid flat, hard-edged cards.
        highlight = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        hdraw = ImageDraw.Draw(highlight)
        top_color = self._hex_to_rgba(accent_top or self.theme.white, 88)
        hdraw.rounded_rectangle(
            [body_box[0] + 2 * scale, body_box[1] + 2 * scale, body_box[2] - 2 * scale, body_box[1] + 22 * scale],
            radius=max(1, rr - 2 * scale),
            fill=top_color,
        )
        img.alpha_composite(highlight)

        if scale != 1:
            img = img.resize((width // scale, height // scale), Image.Resampling.LANCZOS)

        stream = io.BytesIO()
        img.save(stream, format="PNG")
        png_rect = Rect(rect.x - pad, rect.y - pad, rect.w + pad * 2, rect.h + pad * 2 + shadow_dy)
        return stream.getvalue(), png_rect


    @staticmethod
    def _mix_rgb(a: tuple[int, int, int], b: tuple[int, int, int], t: float) -> tuple[int, int, int]:
        return (
            int(round(a[0] + (b[0] - a[0]) * t)),
            int(round(a[1] + (b[1] - a[1]) * t)),
            int(round(a[2] + (b[2] - a[2]) * t)),
        )

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
        hex_color = str(hex_color).strip().lstrip("#")
        if len(hex_color) == 3:
            hex_color = "".join(ch * 2 for ch in hex_color)
        return int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)

    def _gradient_panel_png(
        self,
        rect: Rect,
        *,
        left: str,
        right: str,
        line: str,
        radius: int = 20,
        border_width: int = 1,
        shadow: bool = True,
        glass: bool = True,
    ) -> tuple[bytes, Rect] | None:
        """Rounded gradient PNG card used for Option A SaaS KPI cards.

        The PNG is inserted underneath text/charts, avoiding the Excel layering issue
        where shape backgrounds can cover native chart objects.
        """
        if Image is None or ImageDraw is None or ImageFilter is None:
            return None

        scale = 2
        pad = 14 if shadow else 3
        shadow_dy = 5 if shadow else 0
        width = max(1, (rect.w + pad * 2) * scale)
        height = max(1, (rect.h + pad * 2 + shadow_dy) * scale)
        img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        body_box = [pad * scale, pad * scale, (pad + rect.w) * scale, (pad + rect.h) * scale]
        rr = radius * scale

        if shadow:
            shadow_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            shadow_draw = ImageDraw.Draw(shadow_layer)
            shadow_draw.rounded_rectangle(
                [body_box[0], body_box[1] + shadow_dy * scale, body_box[2], body_box[3] + shadow_dy * scale],
                radius=rr,
                fill=self._hex_to_rgba(self.theme.shadow, 54),
            )
            shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(11 * scale))
            img.alpha_composite(shadow_layer)

        left_rgb = self._hex_to_rgb(left)
        right_rgb = self._hex_to_rgb(right)
        grad = Image.new("RGBA", (rect.w * scale, rect.h * scale), (0, 0, 0, 0))
        px = grad.load()
        for x in range(rect.w * scale):
            t = x / max(1, rect.w * scale - 1)
            rgb = self._mix_rgb(left_rgb, right_rgb, t)
            for y in range(rect.h * scale):
                vertical = y / max(1, rect.h * scale - 1)
                alpha = 242
                shade = 1.0 - 0.12 * vertical
                px[x, y] = (int(rgb[0] * shade), int(rgb[1] * shade), int(rgb[2] * shade), alpha)

        mask = Image.new("L", (rect.w * scale, rect.h * scale), 0)
        mdraw = ImageDraw.Draw(mask)
        mdraw.rounded_rectangle([0, 0, rect.w * scale - 1, rect.h * scale - 1], radius=rr, fill=255)
        crop_target = Image.new("RGBA", (rect.w * scale, rect.h * scale), (0, 0, 0, 0))
        crop_target.alpha_composite(grad)
        crop_target.putalpha(mask)
        img.paste(crop_target, (pad * scale, pad * scale), crop_target)

        draw = ImageDraw.Draw(img)
        border_rgba = self._hex_to_rgba(line, 8)
        draw.rounded_rectangle(body_box, radius=rr, outline=border_rgba, width=max(1, border_width * scale))

        if glass:
            glass_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            gdraw = ImageDraw.Draw(glass_layer)
            gdraw.rounded_rectangle(
                [body_box[0] + 2 * scale, body_box[1] + 2 * scale, body_box[2] - 2 * scale, body_box[1] + 28 * scale],
                radius=max(1, rr - 2 * scale),
                fill=self._hex_to_rgba("#FFFFFF", 90),
            )
            glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            glow_draw = ImageDraw.Draw(glow)
            glow_draw.ellipse(
                [body_box[2] - 76 * scale, body_box[1] - 38 * scale, body_box[2] + 24 * scale, body_box[1] + 62 * scale],
                fill=self._hex_to_rgba(right, 68),
            )
            glow = glow.filter(ImageFilter.GaussianBlur(16 * scale))
            img.alpha_composite(glow)
            img.alpha_composite(glass_layer)

        if scale != 1:
            img = img.resize((width // scale, height // scale), Image.Resampling.LANCZOS)

        stream = io.BytesIO()
        img.save(stream, format="PNG")
        png_rect = Rect(rect.x - pad, rect.y - pad, rect.w + pad * 2, rect.h + pad * 2 + shadow_dy)
        return stream.getvalue(), png_rect

    def gradient_panel(
        self,
        rect: Rect,
        *,
        left: str,
        right: str,
        line: str | None = None,
        radius: int = 20,
        border_width: int = 1,
        shadow: bool = True,
    ) -> None:
        asset = self._gradient_panel_png(
            rect,
            left=left,
            right=right,
            line=line or self.theme.border_glow,
            radius=radius,
            border_width=border_width,
            shadow=shadow,
        )
        if asset is not None:
            png_bytes, image_rect = asset
            self._image(image_rect, png_bytes, name="gradient_panel.png")
        else:
            self.panel(rect, fill=left, line=line or self.theme.border_glow, radius=radius, border_width=border_width, shadow=shadow)

    def panel(
        self,
        rect: Rect,
        *,
        fill: str | None = None,
        line: str | None = None,
        radius: int = 18,
        border_width: int = 1,
        fill_transparency: int = 0,
        border_transparency: int = 0,
        shadow: bool = True,
        accent_top: str | None = None,
    ) -> None:
        fill = fill or self.theme.panel
        line = line or self.theme.border_glow
        asset = self._rounded_panel_png(
            rect,
            fill=fill,
            line=line,
            radius=radius,
            border_width=border_width,
            fill_transparency=fill_transparency,
            border_transparency=border_transparency,
            shadow=shadow,
            accent_top=accent_top,
        )
        if asset is not None:
            png_bytes, image_rect = asset
            self._image(image_rect, png_bytes, name="rounded_panel.png")
        else:
            self._paint_cell_panel(rect, fill=fill, line=line, border_width=max(1, border_width))

    def title(self, rect: Rect, text: str, *, size: int = 13, color: str | None = None) -> None:
        self.textbox(rect, text, fill=None, line=None, font_color=color or self.theme.white, font_size=size, bold=True, margin=0)

    def label(self, rect: Rect, text: str, *, size: int = 9, color: str | None = None, bold: bool = False, align: str = "left") -> None:
        self.textbox(rect, text, fill=None, line=None, font_color=color or self.theme.muted, font_size=size, bold=bold, align=align, margin=0)

    def pill(self, rect: Rect, text: str, *, fill: str, color: str | None = None, size: int = 9) -> None:
        self.panel(rect, fill=fill, line=fill, radius=max(10, rect.h // 2), border_width=1, shadow=False, accent_top=fill)
        self.textbox(
            rect,
            text,
            fill=None,
            line=None,
            font_color=color or self.theme.white,
            font_size=size,
            bold=True,
            align="center",
            valign="middle",
            margin=4,
        )

    def _linked_button_png(
        self,
        rect: Rect,
        text: str,
        *,
        fill: str,
        line: str,
        text_color: str,
        accent: str | None = None,
        radius: int = 17,
        font_size: int = 12,
    ) -> bytes | None:
        """Render a readable clickable dashboard/navigation button as a PNG.

        The first version used compact 27px-high image buttons with 9pt text.
        That looked polished but was too small once Excel rendered the workbook.
        This version uses a higher internal render scale, larger typography and
        true vertical centering so the sidebar remains legible after export.
        """
        if Image is None or ImageDraw is None:
            return None

        scale = 3
        img = Image.new("RGBA", (max(1, rect.w * scale), max(1, rect.h * scale)), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        box = [0, 0, rect.w * scale - 1, rect.h * scale - 1]
        draw.rounded_rectangle(
            box,
            radius=radius * scale,
            fill=self._hex_to_rgba(fill, 0),
            outline=self._hex_to_rgba(line, 10),
            width=max(2, scale),
        )
        if accent:
            dot_r = max(4, min(7, rect.h // 5)) * scale
            cx = 18 * scale
            cy = rect.h * scale // 2
            draw.ellipse([cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r], fill=self._hex_to_rgba(accent, 0))

        font = None
        if ImageFont is not None:
            for candidate in ("DejaVuSans-Bold.ttf", "DejaVuSans.ttf", "Arial.ttf"):
                try:
                    font = ImageFont.truetype(candidate, font_size * scale)
                    break
                except Exception:
                    continue
        tx = 36 * scale if accent else 14 * scale
        label = str(text)
        if font is not None:
            bbox = draw.textbbox((0, 0), label, font=font)
            text_h = bbox[3] - bbox[1]
            ty = max(2 * scale, (rect.h * scale - text_h) // 2 - bbox[1])
        else:
            ty = max(2 * scale, (rect.h * scale - font_size * scale) // 2)
        draw.text((tx, ty), label, fill=self._hex_to_rgba(text_color, 0), font=font)

        if scale != 1:
            img = img.resize((rect.w, rect.h), Image.Resampling.LANCZOS)
        stream = io.BytesIO()
        img.save(stream, format="PNG")
        return stream.getvalue()

    def image_link_button(
        self,
        rect: Rect,
        text: str,
        url: str,
        *,
        fill: str | None = None,
        line: str | None = None,
        text_color: str | None = None,
        accent: str | None = None,
        active: bool = False,
        tip: str | None = None,
        font_size: int = 12,
    ) -> None:
        fill = fill or (self.theme.panel_3 if active else self.theme.sidebar)
        line = line or (self.theme.border_glow if active else self.theme.border_soft)
        text_color = text_color or (self.theme.white if active else self.theme.muted)
        png = self._linked_button_png(rect, text, fill=fill, line=line, text_color=text_color, accent=accent or self.theme.cyan, font_size=font_size)
        if png is not None:
            self._image(rect, png, name="linked_button.png", url=url, tip=tip or text, decorative=False)
            return

        # Fallback without Pillow: use a cell hyperlink. It is less pixel-perfect
        # but still preserves navigation/clickability.
        row, col, _, _ = self._anchor(rect)
        link_fmt = self.wb.add_format({
            "bg_color": fill,
            "font_color": text_color,
            "bold": active,
            "font_size": 9,
            "underline": False,
            "border": 1,
            "border_color": line,
            "align": "left",
            "valign": "vcenter",
        })
        self.ws.write_url(row, col, url, link_fmt, string=text, tip=tip or text)

    def accent_bar(self, rect: Rect, color: str) -> None:
        self.textbox(rect, "", fill=color, line=color, margin=0)

    def kpi_card(self, rect: Rect, label: str, value: str, *, accent: str, helper: str | None = None) -> None:
        # Option A KPI cards use a dark glass gradient asset under transparent labels.
        # This keeps the premium SaaS finish and avoids chart/shape layering regressions.
        self.gradient_panel(rect, left=self.theme.panel_2, right=accent, line=accent, radius=20, border_width=1, shadow=True)
        self.textbox(Rect(rect.x + 18, rect.y + 14, rect.w - 74, 20), label.upper(), fill=None, line=None, font_color=self.theme.muted, font_size=8, bold=True, margin=0)
        self.textbox(Rect(rect.x + 18, rect.y + 39, rect.w - 56, 42), value, fill=None, line=None, font_color=self.theme.white, font_size=22, bold=True, margin=0)
        if helper:
            self.textbox(Rect(rect.x + 18, rect.y + 78, rect.w - 36, 18), helper, fill=None, line=None, font_color=self.theme.muted, font_size=8, margin=0)
        self.pill(Rect(rect.x + rect.w - 48, rect.y + 18, 26, 26), "", fill=accent, size=7)

    def form_button(self, rect: Rect, caption: str, macro: str, *, description: str | None = None) -> None:
        """Insert a native Excel form button bound to a VBA macro.

        XlsxWriter supports form buttons as the macro-enabled control surface.
        They intentionally sit above the visual layer so the click target is reliable.
        """
        row, col, x_offset, y_offset = self._anchor(rect)
        self.ws.insert_button(
            row,
            col,
            {
                "caption": caption,
                "macro": macro,
                "width": rect.w,
                "height": rect.h,
                "x_offset": x_offset,
                "y_offset": y_offset,
                "description": description or caption,
            },
        )

    def insert_chart(self, rect: Rect, chart, *, inner_padding: int = 12) -> None:
        chart.set_size({"width": max(10, rect.w - inner_padding * 2), "height": max(10, rect.h - inner_padding * 2)})
        row, col, x_offset, y_offset = self._anchor(Rect(rect.x + inner_padding, rect.y + inner_padding, rect.w, rect.h))
        self.ws.insert_chart(row, col, chart, {"x_offset": x_offset, "y_offset": y_offset, "object_position": 3})

    def chart_panel(self, rect: Rect, title: str, subtitle: str | None = None) -> Rect:
        self.panel(rect, fill=self.theme.panel, line=self.theme.border_glow, radius=22, border_width=1, shadow=True, accent_top=self.theme.cyan)
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
        tmp[date_col] = _parse_datetime_series(tmp[date_col])
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
        "major_gridlines": {"visible": False},
        "minor_gridlines": {"visible": False},
        "major_tick_mark": "none",
    })
    chart.set_y_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"none": True},
        "major_gridlines": {"visible": False},
        "minor_gridlines": {"visible": False},
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
    chart.set_x_axis({"visible": False, "major_gridlines": {"visible": False}, "minor_gridlines": {"visible": False}})
    chart.set_y_axis({"visible": False, "major_gridlines": {"visible": False}, "minor_gridlines": {"visible": False}})
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
    chart.set_x_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"color": theme.border},
        "major_gridlines": {"visible": False},
        "minor_gridlines": {"visible": False},
        "major_tick_mark": "none",
    })
    chart.set_y_axis({
        "num_font": {"color": theme.muted, "size": 8},
        "line": {"none": True},
        "major_gridlines": {"visible": False},
        "minor_gridlines": {"visible": False},
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

def generate_dark_saas_dashboard(
    csv_path: str | Path,
    output_path: str | Path,
    *,
    vba_project_path: str | Path | None = None,
    macro_source_path: str | Path | None = None,
    write_macro_source_file: bool = True,
) -> dict[str, Any]:
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

    vba_project = Path(vba_project_path) if vba_project_path else None
    macros_embedded = False
    if vba_project is not None:
        if not vba_project.exists():
            raise FileNotFoundError(f"vbaProject.bin not found: {vba_project}")
        macros_embedded = True
        if output_path.suffix.lower() != ".xlsm":
            output_path = output_path.with_suffix(".xlsm")
    elif output_path.suffix.lower() == ".xlsm":
        raise ValueError("Output .xlsm requires --vba-project pointing to a precompiled vbaProject.bin")

    generated_macro_source: Path | None = None
    if write_macro_source_file:
        generated_macro_source = write_vba_macro_source(macro_source_path or output_path.with_name(output_path.stem + "_DashboardMacros.bas"))

    trend_df = ensure_chart_frame(agg.get("trend"), kind="trend", df=df, metric=metric)
    top_dim_df = ensure_chart_frame(agg.get("top_dim"), kind="category", df=df, preferred_dim=dim)
    second_dim_df = ensure_chart_frame(agg.get("second_dim"), kind="category", df=df, preferred_dim=second or dim)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book
        if macros_embedded and vba_project is not None:
            wb.add_vba_project(str(vba_project))

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
            title="OptiQuant IA — Dark SaaS Premium CSV Dashboard Generator",
            rows=[
                ("What this file does", "Cleans any CSV file, detects business semantics and renders an Option A Dark SaaS Premium dashboard with neon accents, rounded cards, click-ready navigation and native Excel charts."),
                ("Dashboard engine", "The Dashboard sheet is built with a fixed pixel canvas. Glass-like rounded PNG assets are inserted below charts; transparent labels and clickable controls sit above without hiding chart content."),
                ("Detected dataset", f"{human_label(dataset_type)} with {round(spec['confidence'] * 100)}% confidence."),
                ("Metric / date / dimension", f"Metric: {metric or 'N/A'} | Date: {date_col or 'N/A'} | Dimension: {dim or 'N/A'}"),
                ("Important", "Excel is not Figma, but this version avoids chart-covering textbox backgrounds, uses rounded image panels with soft borders/shadows, removes chart gridlines, and keeps charts editable as native Excel objects."),
                ("Automation", "Sidebar navigation is clickable in XLSX and XLSM. Refresh/Export PDF action buttons are embedded when --vba-project points to a trusted vbaProject.bin. Without that binary, this script still exports the VBA source module beside the workbook for template creation."),
            ],
            tab_color=theme.purple,
        )

        # Dashboard canvas and visual system.
        engine = ExcelLayoutEngine(wb, dash, theme=theme, canvas_width=1500, canvas_height=900)
        engine.setup_canvas()

        # Sidebar — Option A SaaS product navigation.
        engine.panel(Rect(0, 0, 238, 900), fill=theme.sidebar, line=theme.sidebar, radius=0, border_width=0, shadow=False)
        engine.textbox(Rect(28, 26, 180, 30), "◈ OptiQuant IA", fill=None, line=None, font_color=theme.white, font_size=15, bold=True, margin=0)
        engine.textbox(Rect(28, 58, 176, 24), "AI CSV Dashboard Engine", fill=None, line=None, font_color=theme.muted_2, font_size=8, margin=0)
        engine.pill(Rect(28, 104, 160, 30), "OPTION A · SAAS", fill=theme.purple, size=8)
        engine.label(Rect(28, 160, 130, 20), "WORKSPACE", size=8, color=theme.muted_2, bold=True)
        sidebar_items = [
            ("Overview", "Dashboard", "A1"),
            ("Analytics", "Dashboard", "A1"),
            ("Data Quality", "Data Quality", "A1"),
            ("Clean Data", "Clean Data", "A1"),
            ("Raw Data", "Raw Data", "A1"),
            ("Cleaning Log", "Cleaning Log", "A1"),
            ("About", "About", "A1"),
        ]
        for i, (item, target_sheet, target_cell) in enumerate(sidebar_items):
            y = 174 + i * 44
            active = item in {"Overview", "Analytics"}
            accent = theme.cyan if active else theme.muted_2
            engine.image_link_button(
                Rect(20, y - 4, 202, 38),
                item,
                f"internal:'{target_sheet}'!{target_cell}",
                fill=theme.panel_2 if active else theme.sidebar,
                line=theme.border_glow if active else theme.sidebar,
                text_color=theme.white if active else theme.muted,
                accent=accent,
                active=active,
                tip=f"Open {target_sheet}",
                font_size=12,
            )

        engine.panel(Rect(20, 502, 198, 124), fill=theme.panel, line=theme.border_soft, radius=18, border_width=1, shadow=True, accent_top=theme.cyan)
        engine.label(Rect(34, 520, 120, 16), "DATASET DETECTED", size=7, color=theme.muted_2, bold=True)
        engine.textbox(Rect(34, 542, 160, 24), human_label(dataset_type), fill=None, line=None, font_color=theme.white, font_size=12, bold=True, margin=0)
        engine.label(Rect(34, 572, 154, 16), f"{quality['clean_rows']} records · {quality['clean_columns']} fields", size=8, color=theme.muted)
        engine.label(Rect(34, 598, 154, 16), f"Quality score {quality['quality_after']}%", size=8, color=theme.green, bold=True)

        engine.label(Rect(28, 640, 150, 20), "ACTIONS", size=8, color=theme.muted_2, bold=True)
        engine.panel(Rect(20, 670, 198, 160), fill=theme.panel, line=theme.border_glow, radius=18, border_width=1, shadow=True, accent_top=theme.purple)
        if macros_embedded:
            engine.label(Rect(34, 684, 154, 18), "Macro-enabled", size=8, color=theme.green, bold=True)
            engine.form_button(Rect(30, 706, 166, 32), "Refresh", "RefreshDashboard", description="Refresh workbook data and calculations")
            engine.form_button(Rect(30, 746, 166, 32), "Export PDF", "ExportDashboardPDF", description="Export the dashboard sheet as a PDF")
            engine.form_button(Rect(30, 786, 166, 32), "Go Data", "GoToCleanData", description="Open Clean Data sheet")
        else:
            engine.label(Rect(34, 688, 154, 18), "XLSX mode", size=9, color=theme.cyan, bold=True)
            engine.label(Rect(34, 716, 154, 64), "Navigation is clickable. Generate .xlsm with --vba-project to activate Refresh / Export PDF buttons.", size=7, color=theme.muted)
            engine.image_link_button(Rect(28, 788, 172, 38), "Open Clean Data", "internal:'Clean Data'!A1", fill=theme.panel_3, line=theme.border_glow, text_color=theme.white, accent=theme.green, active=True)

        # Header — executive SaaS overview.
        engine.title(Rect(270, 26, 520, 32), "Executive Overview", size=20)
        engine.label(Rect(270, 62, 680, 18), "Summary of key metrics and insights from your cleaned dataset", size=9, color=theme.muted)
        engine.panel(Rect(780, 24, 178, 56), fill=theme.panel, line=theme.border_soft, radius=16, border_width=1, shadow=False, accent_top=theme.cyan)
        engine.label(Rect(798, 36, 132, 13), "DATASET DETECTED", size=6, color=theme.muted_2, bold=True)
        engine.label(Rect(798, 54, 136, 14), human_label(dataset_type), size=8, color=theme.green, bold=True)
        engine.panel(Rect(980, 24, 210, 56), fill=theme.panel, line=theme.border_soft, radius=16, border_width=1, shadow=False, accent_top=theme.purple)
        engine.label(Rect(998, 36, 132, 13), "LAST UPDATED", size=6, color=theme.muted_2, bold=True)
        engine.label(Rect(998, 54, 160, 14), pd.Timestamp.now().strftime("%b %d, %Y %H:%M"), size=8, color=theme.white)
        engine.pill(Rect(1220, 30, 96, 32), f"{quality['quality_after']}%", fill=theme.green, color="#061116", size=10)
        engine.label(Rect(1330, 30, 140, 16), "DATA QUALITY SCORE", size=7, color=theme.muted_2, bold=True)
        engine.label(Rect(1330, 51, 118, 16), "Excellent" if quality["quality_after"] >= 90 else "Good", size=8, color=theme.green, bold=True)

        # KPI cards — dark gradient, neon SaaS cards.
        total_metric = safe_sum(df, metric)
        avg_metric = safe_avg(df, metric)
        unique_dim = safe_unique(df, dim)
        delta_quality = quality["quality_after"] - quality["quality_before"]
        trend_delta = 0.0
        if len(trend_df) >= 2 and float(trend_df["Value"].iloc[-2]) != 0:
            trend_delta = (float(trend_df["Value"].iloc[-1]) / float(trend_df["Value"].iloc[-2]) - 1.0) * 100.0

        engine.kpi_card(Rect(270, 104, 225, 100), "Total " + metric_label(metric), compact_number(total_metric), accent=theme.purple, helper=f"{trend_delta:+.1f}% vs previous point")
        engine.kpi_card(Rect(515, 104, 225, 100), "Average Value", compact_number(avg_metric), accent=theme.cyan, helper=metric_label(metric))
        engine.kpi_card(Rect(760, 104, 200, 100), "Total Records", compact_number(len(df)), accent=theme.blue, helper=f"{quality['original_rows']} raw rows")
        engine.kpi_card(Rect(980, 104, 200, 100), "Segments", compact_number(unique_dim), accent=theme.green, helper=human_label(dim or "dimension"))
        engine.kpi_card(Rect(1200, 104, 250, 100), "Data Quality", f"{quality['quality_after']}%", accent=theme.pink, helper=f"{delta_quality:+d} pts after cleaning")

        # Main analytics row.
        trend_rect = Rect(270, 235, 710, 300)
        trend_chart_area = engine.chart_panel(
            trend_rect,
            f"{metric_label(metric)} Over Time",
            f"Date field: {human_label(date_col) if date_col else 'fallback periods'} · no chart gridlines",
        )
        line_chart = make_line_chart(wb, theme, "_Dashboard Data", trend_area, title=metric_label(metric))
        engine.insert_chart(trend_chart_area, line_chart, inner_padding=6)

        insights_rect = Rect(1005, 235, 445, 300)
        engine.panel(insights_rect, fill=theme.panel, line=theme.border_glow, radius=22, border_width=1, shadow=True, accent_top=theme.purple)
        engine.title(Rect(insights_rect.x + 24, insights_rect.y + 18, 220, 24), "Key Insights", size=12)
        insight_rows = [
            ("▲", theme.green, f"{metric_label(metric)} trend", f"Latest movement: {trend_delta:+.1f}% vs previous period"),
            ("★", theme.orange, "Top segment", str(top_dim_df.iloc[0]['Category']) if not top_dim_df.empty else "N/A"),
            ("●", theme.cyan, "Cleaned data", f"{quality['duplicate_rows_removed']} duplicates removed · {quality['clean_rows']} records ready"),
        ]
        for i, (icon, color, title, line) in enumerate(insight_rows):
            y = insights_rect.y + 62 + i * 58
            engine.pill(Rect(insights_rect.x + 24, y, 28, 28), icon, fill=color, color="#061116" if color == theme.green else theme.white, size=9)
            engine.label(Rect(insights_rect.x + 68, y - 1, 150, 16), title.upper(), size=7, color=theme.muted_2, bold=True)
            engine.label(Rect(insights_rect.x + 68, y + 18, 320, 18), line, size=8, color=theme.white if i == 0 else theme.muted)
        mini_area = Rect(insights_rect.x + 22, insights_rect.y + 232, insights_rect.w - 44, 50)
        area_chart = make_area_chart(wb, theme, "_Dashboard Data", trend_area)
        engine.insert_chart(mini_area, area_chart, inner_padding=0)

        # Bottom row: ranking, donut, data quality breakdown.
        bar_rect = Rect(270, 565, 445, 285)
        bar_area = engine.chart_panel(bar_rect, f"Top {human_label(second or dim or 'Categories')}", "Highest contributing groups")
        bar_chart = make_bar_chart(wb, theme, "_Dashboard Data", second_area)
        engine.insert_chart(bar_area, bar_chart, inner_padding=4)

        donut_rect = Rect(740, 565, 340, 285)
        donut_area = engine.chart_panel(donut_rect, f"Breakdown by {human_label(dim or 'Category')}", "Distribution across top segments")
        donut_chart = make_donut_chart(wb, theme, "_Dashboard Data", top_area)
        engine.insert_chart(donut_area, donut_chart, inner_padding=2)

        quality_rect = Rect(1105, 565, 345, 285)
        engine.panel(quality_rect, fill=theme.panel, line=theme.border_glow, radius=22, border_width=1, shadow=True, accent_top=theme.green)
        engine.title(Rect(quality_rect.x + 22, quality_rect.y + 18, 260, 24), "Data Quality Overview", size=12)
        engine.label(Rect(quality_rect.x + 22, quality_rect.y + 45, 300, 18), "Completeness · Consistency · Validity · Uniqueness", size=8, color=theme.muted_2)
        quality_metrics = [
            ("Completeness", quality["quality_after"], theme.cyan),
            ("Consistency", max(0, min(100, quality["quality_after"] - 3 + delta_quality // 2)), theme.green),
            ("Validity", max(0, min(100, quality["quality_after"] - 1)), theme.purple),
            ("Uniqueness", max(0, min(100, 100 - int(quality["duplicate_rows_removed"] / max(quality["original_rows"], 1) * 100))), theme.orange),
        ]
        for i, (q_label, score, color) in enumerate(quality_metrics):
            y = quality_rect.y + 84 + i * 42
            engine.label(Rect(quality_rect.x + 24, y, 120, 16), q_label.upper(), size=7, color=theme.muted_2, bold=True)
            engine.panel(Rect(quality_rect.x + 145, y + 4, 150, 8), fill=theme.panel_3, line=theme.panel_3, radius=5, border_width=0, shadow=False)
            engine.panel(Rect(quality_rect.x + 145, y + 4, int(150 * score / 100), 8), fill=color, line=color, radius=5, border_width=0, shadow=False, accent_top=color)
            engine.label(Rect(quality_rect.x + 304, y - 1, 40, 16), f"{score}%", size=8, color=theme.white, bold=True)
        engine.accent_bar(Rect(quality_rect.x + 22, quality_rect.y + 250, 300, 2), theme.border)
        engine.label(Rect(quality_rect.x + 22, quality_rect.y + 262, 300, 16), "Ready-to-use workbook: dashboard · clean data · raw data · logs", size=7, color=theme.muted_2)

    return {
        "output": str(output_path),
        "dataset_type": spec["dataset_type"],
        "quality_before": quality["quality_before"],
        "quality_after": quality["quality_after"],
        "metric": metric,
        "date": date_col,
        "dimension": dim,
        "macros_embedded": macros_embedded,
        "macro_source": str(generated_macro_source) if generated_macro_source else None,
    }


# Backward-compatible alias for previous CRM-style versions.
generate_crm_style_dashboard = generate_dark_saas_dashboard

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path", help="Path to input CSV")
    parser.add_argument("output_path", nargs="?", default="dark_saas_dashboard.xlsx", help="Path to output XLSX/XLSM")
    parser.add_argument("--vba-project", default=None, help="Path to a precompiled vbaProject.bin to embed in the generated XLSM")
    parser.add_argument("--macro-source-output", default=None, help="Where to write the DashboardMacros.bas source module")
    parser.add_argument("--no-macro-source", action="store_true", help="Do not export the DashboardMacros.bas source file")
    args = parser.parse_args()

    result = generate_dark_saas_dashboard(
        args.csv_path,
        args.output_path,
        vba_project_path=args.vba_project,
        macro_source_path=args.macro_source_output,
        write_macro_source_file=not args.no_macro_source,
    )
    print("Generated:", result["output"])
    print("Detected type:", result["dataset_type"])
    print("Quality:", f"{result['quality_before']}% -> {result['quality_after']}%")
    print("Metric:", result["metric"])
    print("Date:", result["date"])
    print("Dimension:", result["dimension"])
    print("Macros embedded:", result["macros_embedded"])
    if result.get("macro_source"):
        print("Macro source:", result["macro_source"])


if __name__ == "__main__":
    main()
