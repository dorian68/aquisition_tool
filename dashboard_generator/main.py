from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

from .aggregation import (
    aggregate_data,
    compact_number,
    ensure_chart_frame,
    human_label,
    metric_label,
    safe_avg,
    safe_sum,
    safe_unique,
)
from .cleaning import clean_csv
from .layout_engine import (
    ExcelLayoutEngine,
    Rect,
    add_hidden_data_frame,
    make_area_chart,
    make_bar_chart,
    make_donut_chart,
    make_line_chart,
    style_data_sheet,
    write_info_sheet,
)
from .semantic_detection import classify_dataset
from .theme import TEMPLATE_CHOICES, TemplateProfile, choose_template, get_template
from .vba import write_vba_macro_source


ASSET_VBA_PROJECT = Path(__file__).resolve().parent / "assets" / "vbaProject.bin"


def _resolve_profile(template: str, spec: dict[str, Any]) -> TemplateProfile:
    if template == "auto":
        return choose_template(spec)
    return get_template(template)


def _resolve_output_and_vba(output_path: str | Path, vba_project_path: str | Path | None) -> tuple[Path, Path | None, bool]:
    output = Path(output_path)
    if not output.suffix:
        output = output.with_suffix(".xlsx")

    vba_project = Path(vba_project_path) if vba_project_path else None
    if vba_project is None and output.suffix.lower() == ".xlsm" and ASSET_VBA_PROJECT.exists():
        vba_project = ASSET_VBA_PROJECT

    macros_embedded = False
    if vba_project is not None:
        if not vba_project.exists():
            raise FileNotFoundError(f"vbaProject.bin not found: {vba_project}")
        macros_embedded = True
        if output.suffix.lower() != ".xlsm":
            output = output.with_suffix(".xlsm")
    elif output.suffix.lower() == ".xlsm":
        raise ValueError("Output .xlsm requires --vba-project or dashboard_generator/assets/vbaProject.bin")

    return output, vba_project, macros_embedded


def _quality_frame(raw_df: pd.DataFrame, quality: dict[str, Any]) -> pd.DataFrame:
    rows = []
    for col in raw_df.columns:
        rows.append({
            "Column": col,
            "Missing Before": quality["missing_before"].get(col, 0),
            "Missing After": quality["missing_after"].get(col, 0),
            "Detected Type": (
                "date" if col in quality["date_columns"]
                else "numeric" if col in quality["numeric_columns"]
                else "category"
            ),
        })
    return pd.DataFrame(rows)


def _write_settings_sheet(
    ws,
    wb,
    *,
    profile: TemplateProfile,
    requested_template: str,
    spec: dict[str, Any],
    quality: dict[str, Any],
    generated_at: pd.Timestamp,
    client_name: str | None,
    macros_embedded: bool,
    output_path: Path,
) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(profile.theme.purple)

    title_fmt = wb.add_format({"bold": True, "font_size": 18, "font_color": "#111827"})
    key_fmt = wb.add_format({"bold": True, "font_color": "#374151", "bg_color": "#EEF2F7", "border": 1, "border_color": "#D1D5DB"})
    val_fmt = wb.add_format({"font_color": "#111827", "border": 1, "border_color": "#E5E7EB", "text_wrap": True})
    note_fmt = wb.add_format({"italic": True, "font_color": "#6B7280"})

    rows = [
        ("Template selected", profile.slug),
        ("Template requested", requested_template),
        ("Main metric", spec.get("metric") or "N/A"),
        ("Date column", spec.get("date") or "N/A"),
        ("Dimension column", spec.get("primary_dimension") or "N/A"),
        ("Secondary dimension", spec.get("secondary_dimension") or "N/A"),
        ("Generated at", generated_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Client name", client_name or "N/A"),
        ("Theme", profile.display_name),
        ("Dataset type", spec.get("dataset_type", "generic")),
        ("Detection confidence", f"{round(spec.get('confidence', 0) * 100)}%"),
        ("Data quality", f"{quality['quality_before']}% -> {quality['quality_after']}%"),
        ("Output format", output_path.suffix.lower().lstrip(".") or "xlsx"),
        ("Macros embedded", "yes" if macros_embedded else "no"),
    ]

    ws.write("A1", "Settings", title_fmt)
    ws.write("A2", "Control panel values used by dashboard formulas, macros and SaaS metadata.", note_fmt)
    for index, (key, value) in enumerate(rows, start=4):
        ws.write(index - 1, 0, key, key_fmt)
        ws.write(index - 1, 1, value, val_fmt)

    ws.set_column("A:A", 26)
    ws.set_column("B:B", 70)


def _write_ai_insights_sheet(ws, wb, *, ai_report: dict[str, Any], profile: TemplateProfile) -> None:
    theme = profile.theme
    ws.hide_gridlines(2)
    ws.set_tab_color(theme.purple)
    title_fmt = wb.add_format({"bold": True, "font_size": 20, "font_color": theme.white, "bg_color": theme.panel})
    section_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": theme.white, "bg_color": theme.purple, "border": 1, "border_color": theme.border})
    body_fmt = wb.add_format({"font_color": "#111827", "text_wrap": True, "valign": "top", "border": 1, "border_color": "#E5E7EB"})
    note_fmt = wb.add_format({"italic": True, "font_color": "#374151", "text_wrap": True, "bg_color": "#F3F4F6", "border": 1, "border_color": "#D1D5DB"})

    ws.merge_range("A1:D1", _safe_text(ai_report.get("dashboard_title"), "AI Dashboard Analysis"), title_fmt)
    ws.set_column("A:A", 24)
    ws.set_column("B:D", 46)

    row = 3
    sections = [
        ("Executive Summary", _safe_text(ai_report.get("executive_summary"))),
        ("Data Quality Summary", _safe_text(ai_report.get("data_quality_summary"))),
        ("Cleaning Summary", _safe_text(ai_report.get("cleaning_summary"))),
        ("Key Insights", _bullet_lines(ai_report.get("key_insights"))),
        ("Recommended Actions", _bullet_lines(ai_report.get("recommended_actions"))),
        ("Risk Notes", _bullet_lines(ai_report.get("risk_notes"))),
        ("AI Context Safety", "Raw CSV rows were not sent to the AI model. The AI report was generated from a compact Python-generated data profile."),
    ]
    for title, text in sections:
        ws.merge_range(row, 0, row, 3, title, section_fmt)
        row += 1
        fmt = note_fmt if title == "AI Context Safety" else body_fmt
        ws.merge_range(row, 0, row + 1, 3, text, fmt)
        ws.set_row(row, 42)
        ws.set_row(row + 1, 42)
        row += 3


def _safe_text(value: Any, fallback: str = "") -> str:
    text = str(value or fallback).strip()
    return text[:1200]


def _list_values(value: Any, limit: int) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item).strip()[:400] for item in value[:limit] if str(item).strip()]


def _bullet_lines(value: Any, limit: int = 8) -> str:
    items = _list_values(value, limit)
    return "\n".join(f"- {item}" for item in items) if items else "N/A"


def generate_dashboard(
    csv_path: str | Path,
    output_path: str | Path,
    *,
    template: str = "auto",
    vba_project_path: str | Path | None = None,
    macro_source_path: str | Path | None = None,
    write_macro_source_file: bool = True,
    client_name: str | None = None,
    hide_settings: bool = False,
    ai_report: dict[str, Any] | None = None,
) -> dict[str, Any]:
    result = clean_csv(csv_path)
    raw_df = result.raw_df
    df = result.clean_df
    quality = result.quality_report
    spec = classify_dataset(df, quality)
    profile = _resolve_profile(template, spec)
    agg = aggregate_data(df, spec)

    metric = spec.get("metric")
    date_col = spec.get("date")
    dim = spec.get("primary_dimension")
    second = spec.get("secondary_dimension")
    dataset_type = spec.get("dataset_type", "generic")
    generated_at = pd.Timestamp.now()

    output_path, vba_project, macros_embedded = _resolve_output_and_vba(output_path, vba_project_path)
    theme = profile.theme

    generated_macro_source: Path | None = None
    if write_macro_source_file:
        generated_macro_source = write_vba_macro_source(
            macro_source_path or output_path.with_name(output_path.stem + "_DashboardMacros.bas")
        )

    trend_df = ensure_chart_frame(agg.get("trend"), kind="trend", df=df, metric=metric)
    top_dim_df = ensure_chart_frame(agg.get("top_dim"), kind="category", df=df, preferred_dim=dim)
    second_dim_df = ensure_chart_frame(agg.get("second_dim"), kind="category", df=df, preferred_dim=second or dim)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book
        if macros_embedded and vba_project is not None:
            wb.add_vba_project(str(vba_project))

        dash = wb.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash

        df.to_excel(writer, sheet_name="Clean Data", index=False)
        raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
        log_df = pd.DataFrame(result.cleaning_log)
        if log_df.empty:
            log_df = pd.DataFrame([{"action": "none", "details": "No cleaning action required"}])
        log_df.to_excel(writer, sheet_name="Cleaning Log", index=False)

        quality_df = _quality_frame(raw_df, quality)
        quality_df.to_excel(writer, sheet_name="Data Quality", index=False)

        helper = wb.add_worksheet("_Dashboard Data")
        writer.sheets["_Dashboard Data"] = helper
        helper.hide()

        settings = wb.add_worksheet("Settings")
        writer.sheets["Settings"] = settings

        about = wb.add_worksheet("About")
        writer.sheets["About"] = about

        if ai_report:
            ai_sheet = wb.add_worksheet("AI Insights")
            writer.sheets["AI Insights"] = ai_sheet

        helper_header_fmt = wb.add_format({
            "bg_color": theme.panel_3,
            "font_color": theme.white,
            "bold": True,
            "border": 1,
            "border_color": theme.border,
        })
        helper_cell_fmt = wb.add_format({
            "bg_color": theme.panel,
            "font_color": theme.muted,
            "border": 1,
            "border_color": theme.border,
        })

        cursor = {"row": 0, "col": 0}
        trend_area = add_hidden_data_frame(helper, cursor, "trend", trend_df, helper_header_fmt, helper_cell_fmt)
        top_area = add_hidden_data_frame(helper, cursor, "top_dim", top_dim_df, helper_header_fmt, helper_cell_fmt)
        second_area = add_hidden_data_frame(helper, cursor, "second_dim", second_dim_df, helper_header_fmt, helper_cell_fmt)

        style_data_sheet(writer.sheets["Clean Data"], wb, df, tab_color=theme.green, table_name="CleanDataTable")
        style_data_sheet(writer.sheets["Raw Data"], wb, raw_df, tab_color=theme.purple, table_name="RawDataTable")
        style_data_sheet(writer.sheets["Cleaning Log"], wb, log_df, tab_color=theme.orange, table_name="CleaningLogTable")
        style_data_sheet(writer.sheets["Data Quality"], wb, quality_df, tab_color=theme.cyan, table_name="DataQualityTable")

        _write_settings_sheet(
            settings,
            wb,
            profile=profile,
            requested_template=template,
            spec=spec,
            quality=quality,
            generated_at=generated_at,
            client_name=client_name,
            macros_embedded=macros_embedded,
            output_path=output_path,
        )
        if hide_settings:
            settings.hide()

        if ai_report:
            _write_ai_insights_sheet(ai_sheet, wb, ai_report=ai_report, profile=profile)

        write_info_sheet(
            about,
            wb,
            title=profile.about_title,
            rows=[
                ("What this file does", profile.about_description),
                ("Dashboard engine", "The Dashboard sheet is built with a fixed pixel canvas. Rounded PNG assets are inserted below charts; transparent labels and clickable controls sit above without hiding chart content."),
                ("Detected dataset", f"{human_label(dataset_type)} with {round(spec['confidence'] * 100)}% confidence."),
                ("Metric / date / dimension", f"Metric: {metric or 'N/A'} | Date: {date_col or 'N/A'} | Dimension: {dim or 'N/A'}"),
                ("Template", f"{profile.option_label} - {profile.display_name} ({profile.slug})."),
                ("Automation", "Sidebar navigation is clickable in XLSX and XLSM. Refresh/Export PDF/Reset Filters macros are supported when --vba-project points to a trusted vbaProject.bin."),
            ],
            tab_color=theme.purple,
        )

        engine = ExcelLayoutEngine(wb, dash, theme=theme, canvas_width=1500, canvas_height=900)
        engine.setup_canvas()

        engine.panel(Rect(0, 0, 238, 900), fill=theme.sidebar, line=theme.sidebar, radius=0, border_width=0, shadow=False)
        engine.textbox(Rect(28, 26, 180, 30), "◈ OptiQuant IA", fill=None, line=None, font_color=theme.white, font_size=15, bold=True, margin=0)
        engine.textbox(Rect(28, 58, 176, 24), profile.sidebar_subtitle, fill=None, line=None, font_color=theme.muted_2, font_size=8, margin=0)
        engine.pill(Rect(28, 104, 160, 30), profile.sidebar_badge, fill=theme.purple, size=8)
        engine.label(Rect(28, 160, 130, 20), profile.sidebar_section, size=8, color=theme.muted_2, bold=True)
        sidebar_items = [
            ("Overview", "Dashboard", "A1"),
            ("Analytics", "Dashboard", "A1"),
            ("Data Quality", "Data Quality", "A1"),
            ("Clean Data", "Clean Data", "A1"),
            ("Raw Data", "Raw Data", "A1"),
            ("Settings", "Settings", "A1"),
            *([("AI Insights", "AI Insights", "A1")] if ai_report else []),
            ("Cleaning Log", "Cleaning Log", "A1"),
            ("About", "About", "A1"),
        ]
        for i, (item, target_sheet, target_cell) in enumerate(sidebar_items):
            y = 174 + i * 40
            active = item in {"Overview", "Analytics"}
            accent = theme.cyan if active else theme.muted_2
            engine.image_link_button(
                Rect(20, y - 4, 202, 36),
                item,
                f"internal:'{target_sheet}'!{target_cell}",
                fill=theme.panel_2 if active else theme.sidebar,
                line=theme.border_glow if active else theme.sidebar,
                text_color=theme.white if active else theme.muted,
                accent=accent,
                active=active,
                tip=f"Open {target_sheet}",
                font_size=11,
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
            engine.form_button(Rect(30, 786, 166, 32), "Reset Filters", "ResetFilters", description="Clear filters across workbook sheets")
        else:
            engine.label(Rect(34, 688, 154, 18), "XLSX mode", size=9, color=theme.cyan, bold=True)
            engine.label(Rect(34, 716, 154, 64), "Navigation is clickable. Generate .xlsm with --vba-project to activate Refresh / Export PDF buttons.", size=7, color=theme.muted)
            engine.image_link_button(Rect(28, 788, 172, 38), "Open Clean Data", "internal:'Clean Data'!A1", fill=theme.panel_3, line=theme.border_glow, text_color=theme.white, accent=theme.green, active=True)

        engine.title(Rect(270, 26, 520, 32), profile.header_title, size=20)
        engine.label(Rect(270, 62, 680, 18), profile.header_subtitle, size=9, color=theme.muted)
        engine.panel(Rect(780, 24, 178, 56), fill=theme.panel, line=theme.border_soft, radius=16, border_width=1, shadow=False, accent_top=theme.cyan)
        engine.label(Rect(798, 36, 132, 13), "DATASET DETECTED", size=6, color=theme.muted_2, bold=True)
        engine.label(Rect(798, 54, 136, 14), human_label(dataset_type), size=8, color=theme.green, bold=True)
        engine.panel(Rect(980, 24, 210, 56), fill=theme.panel, line=theme.border_soft, radius=16, border_width=1, shadow=False, accent_top=theme.purple)
        engine.label(Rect(998, 36, 132, 13), "LAST UPDATED", size=6, color=theme.muted_2, bold=True)
        engine.label(Rect(998, 54, 160, 14), generated_at.strftime("%b %d, %Y %H:%M"), size=8, color=theme.white)
        engine.pill(Rect(1220, 30, 96, 32), f"{quality['quality_after']}%", fill=theme.green, color="#061116", size=10)
        engine.label(Rect(1330, 30, 140, 16), "DATA QUALITY SCORE", size=7, color=theme.muted_2, bold=True)
        engine.label(Rect(1330, 51, 118, 16), "Excellent" if quality["quality_after"] >= 90 else "Good", size=8, color=theme.green, bold=True)

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
        engine.kpi_card(Rect(980, 104, 200, 100), profile.kpi_segment_label, compact_number(unique_dim), accent=theme.green, helper=human_label(dim or "dimension"))
        engine.kpi_card(Rect(1200, 104, 250, 100), "Data Quality", f"{quality['quality_after']}%", accent=theme.pink, helper=f"{delta_quality:+d} pts after cleaning")

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
        if ai_report:
            engine.title(Rect(insights_rect.x + 24, insights_rect.y + 18, 270, 24), "AI Executive Summary", size=12)
            engine.label(Rect(insights_rect.x + 24, insights_rect.y + 45, 360, 16), _safe_text(ai_report.get("dashboard_title"), "Dashboard Analysis"), size=8, color=theme.muted_2, bold=True)
            engine.label(Rect(insights_rect.x + 24, insights_rect.y + 68, 390, 42), _safe_text(ai_report.get("executive_summary"), "AI analysis unavailable."), size=8, color=theme.white)
            ai_lines = _list_values(ai_report.get("key_insights"), 3)
            action_lines = _list_values(ai_report.get("recommended_actions"), 2)
            for i, line in enumerate(ai_lines):
                y = insights_rect.y + 124 + i * 34
                engine.pill(Rect(insights_rect.x + 24, y, 24, 24), str(i + 1), fill=theme.cyan, color="#061116", size=8)
                engine.label(Rect(insights_rect.x + 60, y + 2, 340, 20), line, size=7, color=theme.muted)
            action_y = insights_rect.y + 226
            engine.label(Rect(insights_rect.x + 24, action_y, 160, 14), "RECOMMENDED ACTIONS", size=7, color=theme.muted_2, bold=True)
            for i, line in enumerate(action_lines):
                engine.label(Rect(insights_rect.x + 24, action_y + 20 + i * 18, 380, 16), f"- {line}", size=7, color=theme.green if i == 0 else theme.muted)
        else:
            engine.title(Rect(insights_rect.x + 24, insights_rect.y + 18, 220, 24), "Key Insights", size=12)
            insight_rows = [
                ("▲", theme.green, f"{metric_label(metric)} trend", f"Latest movement: {trend_delta:+.1f}% vs previous period"),
                ("★", theme.orange, "Top segment", str(top_dim_df.iloc[0]["Category"]) if not top_dim_df.empty else "N/A"),
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

        bar_rect = Rect(270, 565, 445, 285)
        bar_area = engine.chart_panel(bar_rect, f"Top {human_label(second or dim or 'Categories')}", profile.bar_subtitle)
        bar_chart = make_bar_chart(wb, theme, "_Dashboard Data", second_area)
        engine.insert_chart(bar_area, bar_chart, inner_padding=4)

        donut_rect = Rect(740, 565, 340, 285)
        donut_area = engine.chart_panel(donut_rect, f"Breakdown by {human_label(dim or 'Category')}", profile.donut_subtitle)
        donut_chart = make_donut_chart(wb, theme, "_Dashboard Data", top_area)
        engine.insert_chart(donut_area, donut_chart, inner_padding=2)

        quality_rect = Rect(1105, 565, 345, 285)
        engine.panel(quality_rect, fill=theme.panel, line=theme.border_glow, radius=22, border_width=1, shadow=True, accent_top=theme.green)
        engine.title(Rect(quality_rect.x + 22, quality_rect.y + 18, 260, 24), profile.quality_title, size=12)
        engine.label(Rect(quality_rect.x + 22, quality_rect.y + 45, 300, 18), profile.quality_subtitle, size=8, color=theme.muted_2)
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
        engine.label(Rect(quality_rect.x + 22, quality_rect.y + 262, 300, 16), profile.ready_line, size=7, color=theme.muted_2)

    return {
        "output": str(output_path),
        "requested_template": template,
        "selected_template": profile.slug,
        "template_name": profile.display_name,
        "dataset_type": spec["dataset_type"],
        "quality_before": quality["quality_before"],
        "quality_after": quality["quality_after"],
        "metric": metric,
        "date": date_col,
        "dimension": dim,
        "macros_embedded": macros_embedded,
        "macro_source": str(generated_macro_source) if generated_macro_source else None,
        "settings_sheet": "Settings",
        "ai_insights_sheet": "AI Insights" if ai_report else None,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a premium Excel dashboard from a CSV file.")
    parser.add_argument("csv_path", help="Path to input CSV")
    parser.add_argument("output_path", nargs="?", default="dashboard.xlsx", help="Path to output XLSX/XLSM")
    parser.add_argument("--template", choices=TEMPLATE_CHOICES, default="auto", help="Dashboard template to use")
    parser.add_argument("--vba-project", default=None, help="Path to a precompiled vbaProject.bin to embed in generated XLSM")
    parser.add_argument("--macro-source-output", default=None, help="Where to write the DashboardMacros.bas source module")
    parser.add_argument("--no-macro-source", action="store_true", help="Do not export the DashboardMacros.bas source file")
    parser.add_argument("--client-name", default=None, help="Optional client name written to the Settings sheet")
    parser.add_argument("--hide-settings", action="store_true", help="Hide the Settings control panel sheet")
    args = parser.parse_args()

    result = generate_dashboard(
        args.csv_path,
        args.output_path,
        template=args.template,
        vba_project_path=args.vba_project,
        macro_source_path=args.macro_source_output,
        write_macro_source_file=not args.no_macro_source,
        client_name=args.client_name,
        hide_settings=args.hide_settings,
    )
    print("Generated:", result["output"])
    print("Template:", result["selected_template"], f"({result['template_name']})")
    if result["requested_template"] == "auto":
        print("Template mode: auto")
    print("Detected type:", result["dataset_type"])
    print("Quality:", f"{result['quality_before']}% -> {result['quality_after']}%")
    print("Metric:", result["metric"])
    print("Date:", result["date"])
    print("Dimension:", result["dimension"])
    print("Macros embedded:", result["macros_embedded"])
    print("Settings sheet:", result["settings_sheet"])
    if result.get("macro_source"):
        print("Macro source:", result["macro_source"])


if __name__ == "__main__":
    main()
