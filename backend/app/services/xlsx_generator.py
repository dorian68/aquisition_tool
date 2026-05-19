from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd

from app.templates.themes.optiquant_dark import EXCEL_THEME


class XlsxDashboardGenerator:
    def generate(
        self,
        df: pd.DataFrame,
        dashboard_spec: dict[str, Any],
        preview_data: dict[str, Any],
        dataset_profile: dict[str, Any],
    ) -> bytes:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            formats = self._formats(workbook)

            dashboard_ws = workbook.add_worksheet("Dashboard")
            writer.sheets["Dashboard"] = dashboard_ws
            chart_data_ws = workbook.add_worksheet("_ChartData")
            writer.sheets["_ChartData"] = chart_data_ws

            self._write_dashboard(workbook, dashboard_ws, chart_data_ws, formats, dashboard_spec, preview_data, dataset_profile)
            self._write_data_sheet(writer, workbook, df, formats)
            self._write_insights_sheet(workbook, formats, dashboard_spec)
            self._write_quality_sheet(workbook, formats, dataset_profile)
            self._write_about_sheet(workbook, formats)
            chart_data_ws.hide()

        return output.getvalue()

    def _write_dashboard(
        self,
        workbook: Any,
        ws: Any,
        chart_data_ws: Any,
        formats: dict[str, Any],
        spec: dict[str, Any],
        preview: dict[str, Any],
        profile: dict[str, Any],
    ) -> None:
        ws.hide_gridlines(2)
        ws.set_tab_color(EXCEL_THEME["accent"])
        ws.set_default_row(24)
        ws.set_column("A:A", 3)
        ws.set_column("B:M", 14)
        ws.set_column("N:N", 3)

        ws.merge_range("B2:M3", "OptiQuant IA", formats["brand"])
        ws.merge_range("B4:M5", spec["title"], formats["title"])
        ws.merge_range("B6:M6", f"{spec['subtitle']} | Generated {datetime.utcnow():%Y-%m-%d %H:%M UTC}", formats["subtitle"])

        kpi_positions = [("B8:D10"), ("E8:G10"), ("H8:J10"), ("K8:M10"), ("B12:D14")]
        for kpi, cell_range in zip(spec.get("kpis", [])[:5], kpi_positions, strict=False):
            ws.merge_range(cell_range, f"{kpi['label']}\n{self._format_value(kpi.get('value'), kpi.get('format'))}", formats["kpi"])

        chart_positions = ["B16", "H16", "B33", "H33"]
        chart_row = 0
        for chart_payload, position in zip(preview.get("charts", [])[:4], chart_positions, strict=False):
            chart_row = self._write_chart_data(chart_data_ws, chart_row, chart_payload)
            chart = self._create_chart(workbook, chart_payload, chart_row - len(chart_payload.get("labels", [])), len(chart_payload.get("labels", [])))
            ws.insert_chart(position, chart, {"x_scale": 1.18, "y_scale": 1.15})
            chart_row += 2

        insight_start = 50
        ws.merge_range(insight_start, 1, insight_start, 12, "Executive Insights", formats["section"])
        for offset, insight in enumerate(spec.get("insights", [])[:5], start=1):
            ws.merge_range(insight_start + offset, 1, insight_start + offset, 12, f"- {insight}", formats["body"])

        ws.merge_range(
            insight_start + 8,
            1,
            insight_start + 8,
            12,
            f"Data quality score: {profile.get('quality_score', 0)} / 100",
            formats["footer"],
        )

    def _write_chart_data(self, ws: Any, start_row: int, payload: dict[str, Any]) -> int:
        ws.write(start_row, 0, payload.get("title"))
        ws.write(start_row + 1, 0, "Label")
        ws.write(start_row + 1, 1, payload.get("series", [{}])[0].get("name", "Value"))
        labels = payload.get("labels", [])
        values = payload.get("series", [{}])[0].get("data", [])
        for index, (label, value) in enumerate(zip(labels, values, strict=False), start=start_row + 2):
            ws.write(index, 0, label)
            ws.write(index, 1, value)
        return start_row + 2 + len(labels)

    def _create_chart(self, workbook: Any, payload: dict[str, Any], start_row: int, count: int) -> Any:
        chart_type = payload.get("type", "column")
        xlsx_type = {"bar": "bar", "line": "line", "pie": "pie", "doughnut": "doughnut", "column": "column"}.get(chart_type, "column")
        chart = workbook.add_chart({"type": xlsx_type})
        first = start_row + 2
        last = first + max(count - 1, 0)
        chart.add_series(
            {
                "name": payload.get("title"),
                "categories": ["_ChartData", first, 0, last, 0],
                "values": ["_ChartData", first, 1, last, 1],
                "fill": {"color": EXCEL_THEME["accent"]},
                "border": {"color": EXCEL_THEME["accent_dark"]},
            }
        )
        chart.set_title({"name": payload.get("title", "Chart"), "name_font": {"color": EXCEL_THEME["text"], "bold": True}})
        chart.set_legend({"position": "bottom", "font": {"color": EXCEL_THEME["muted"]}})
        chart.set_chartarea({"fill": {"color": EXCEL_THEME["surface"]}, "border": {"color": EXCEL_THEME["border"]}})
        chart.set_plotarea({"fill": {"color": EXCEL_THEME["surface_alt"]}, "border": {"color": EXCEL_THEME["border"]}})
        if xlsx_type not in {"pie", "doughnut"}:
            chart.set_x_axis({"name_font": {"color": EXCEL_THEME["muted"]}, "num_font": {"color": EXCEL_THEME["muted"]}})
            chart.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": EXCEL_THEME["border"]}}, "num_font": {"color": EXCEL_THEME["muted"]}})
        return chart

    def _write_data_sheet(self, writer: pd.ExcelWriter, workbook: Any, df: pd.DataFrame, formats: dict[str, Any]) -> None:
        df.to_excel(writer, sheet_name="Data", index=False, startrow=1)
        ws = writer.sheets["Data"]
        ws.set_tab_color(EXCEL_THEME["accent"])
        ws.write(0, 0, "Cleaned CSV Data", formats["sheet_title"])
        ws.freeze_panes(2, 0)
        if len(df.columns):
            columns = [{"header": str(column)} for column in df.columns]
            ws.add_table(1, 0, len(df) + 1, len(df.columns) - 1, {"columns": columns, "style": "Table Style Medium 4"})
            for index, column in enumerate(df.columns):
                width = min(max(len(str(column)) + 4, 12), 32)
                ws.set_column(index, index, width)

    def _write_insights_sheet(self, workbook: Any, formats: dict[str, Any], spec: dict[str, Any]) -> None:
        ws = workbook.add_worksheet("Insights")
        ws.set_tab_color(EXCEL_THEME["accent"])
        ws.set_column("A:A", 4)
        ws.set_column("B:H", 18)
        ws.merge_range("B2:H3", "Business Insights", formats["title"])
        row = 5
        for insight in spec.get("insights", []):
            ws.merge_range(row, 1, row, 7, insight, formats["body"])
            row += 2
        row += 1
        ws.merge_range(row, 1, row, 7, "Recommended next step: connect this spreadsheet source for automated refresh.", formats["section"])

    def _write_quality_sheet(self, workbook: Any, formats: dict[str, Any], profile: dict[str, Any]) -> None:
        ws = workbook.add_worksheet("Data Quality")
        ws.set_tab_color(EXCEL_THEME["warning"])
        ws.set_column("A:A", 4)
        ws.set_column("B:F", 24)
        ws.merge_range("B2:F3", "Data Quality Report", formats["title"])
        headers = ["Column", "Type", "Semantic Type", "Missing Rate", "Unique Count"]
        for col, header in enumerate(headers, start=1):
            ws.write(5, col, header, formats["table_header"])
        for row_index, column in enumerate(profile.get("columns", []), start=6):
            ws.write(row_index, 1, column["name"], formats["body"])
            ws.write(row_index, 2, column["type"], formats["body"])
            ws.write(row_index, 3, column.get("semantic_type"), formats["body"])
            ws.write(row_index, 4, column.get("missing_rate"), formats["body"])
            ws.write(row_index, 5, column.get("unique_count"), formats["body"])
        footer_row = 8 + len(profile.get("columns", []))
        ws.merge_range(footer_row, 1, footer_row, 5, f"Duplicate rows: {profile.get('duplicate_rows', 0)}", formats["footer"])
        ws.merge_range(footer_row + 1, 1, footer_row + 1, 5, "Suggested cleaning: review missing values, duplicates and numeric outliers before operational use.", formats["footer"])

    def _write_about_sheet(self, workbook: Any, formats: dict[str, Any]) -> None:
        ws = workbook.add_worksheet("About")
        ws.set_tab_color(EXCEL_THEME["accent"])
        ws.set_column("A:A", 4)
        ws.set_column("B:H", 18)
        ws.merge_range("B2:H3", "Generated by OptiQuant IA", formats["title"])
        ws.merge_range("B5:H6", "This workbook turns a raw CSV export into a decision-ready Excel dashboard.", formats["body"])
        ws.merge_range("B8:H9", "Connect your spreadsheet source to refresh this dashboard automatically.", formats["section"])
        ws.merge_range("B11:H11", "No macros were generated in this workbook.", formats["footer"])

    @staticmethod
    def _formats(workbook: Any) -> dict[str, Any]:
        return {
            "brand": workbook.add_format({"bold": True, "font_size": 14, "font_color": EXCEL_THEME["accent"], "bg_color": EXCEL_THEME["background"], "align": "left", "valign": "vcenter"}),
            "title": workbook.add_format({"bold": True, "font_size": 22, "font_color": EXCEL_THEME["text"], "bg_color": EXCEL_THEME["background"], "align": "left", "valign": "vcenter"}),
            "subtitle": workbook.add_format({"font_size": 10, "font_color": EXCEL_THEME["muted"], "bg_color": EXCEL_THEME["background"], "align": "left"}),
            "kpi": workbook.add_format({"bold": True, "font_size": 12, "font_color": EXCEL_THEME["text"], "bg_color": EXCEL_THEME["surface"], "border": 1, "border_color": EXCEL_THEME["border"], "align": "center", "valign": "vcenter", "text_wrap": True}),
            "section": workbook.add_format({"bold": True, "font_size": 12, "font_color": EXCEL_THEME["accent"], "bg_color": EXCEL_THEME["surface_alt"], "border": 1, "border_color": EXCEL_THEME["border"], "text_wrap": True}),
            "body": workbook.add_format({"font_size": 10, "font_color": EXCEL_THEME["text"], "bg_color": EXCEL_THEME["background"], "text_wrap": True, "valign": "top"}),
            "footer": workbook.add_format({"font_size": 9, "font_color": EXCEL_THEME["muted"], "bg_color": EXCEL_THEME["background"], "text_wrap": True}),
            "sheet_title": workbook.add_format({"bold": True, "font_size": 14, "font_color": EXCEL_THEME["background"], "bg_color": EXCEL_THEME["accent"]}),
            "table_header": workbook.add_format({"bold": True, "font_color": EXCEL_THEME["background"], "bg_color": EXCEL_THEME["accent"], "border": 1}),
        }

    @staticmethod
    def _format_value(value: Any, fmt: str | None) -> str:
        if fmt == "currency" and isinstance(value, (int, float)):
            return f"${value:,.0f}"
        if fmt == "percent" and isinstance(value, (int, float)):
            return f"{value:.1f}/100"
        if isinstance(value, float):
            return f"{value:,.1f}"
        if isinstance(value, int):
            return f"{value:,}"
        return str(value)

