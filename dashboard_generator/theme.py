from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class DashboardTheme:
    bg: str
    bg_2: str
    sidebar: str
    panel: str
    panel_2: str
    panel_3: str
    border: str
    border_soft: str
    border_glow: str
    shadow: str
    white: str
    muted: str
    muted_2: str
    purple: str
    purple_2: str
    cyan: str
    blue: str
    pink: str
    green: str
    orange: str
    red: str
    chart_palette_colors: tuple[str, ...]

    @property
    def chart_palette(self) -> list[str]:
        return list(self.chart_palette_colors)


@dataclass(frozen=True)
class TemplateProfile:
    slug: str
    display_name: str
    option_label: str
    theme_family: str
    theme: DashboardTheme
    about_title: str
    about_description: str
    sidebar_subtitle: str
    sidebar_badge: str
    sidebar_section: str
    header_title: str
    header_subtitle: str
    kpi_segment_label: str
    bar_subtitle: str
    donut_subtitle: str
    quality_title: str
    quality_subtitle: str
    ready_line: str


DARK_SAAS = TemplateProfile(
    slug="dark-saas",
    display_name="Dark SaaS Premium",
    option_label="Option A",
    theme_family="dark_saas",
    theme=DashboardTheme(
        bg="#050816",
        bg_2="#080D1F",
        sidebar="#090E22",
        panel="#0F172A",
        panel_2="#111C33",
        panel_3="#17223D",
        border="#26314F",
        border_soft="#1C2744",
        border_glow="#2E4B78",
        shadow="#02040D",
        white="#F8FBFF",
        muted="#AAB6D8",
        muted_2="#7381A6",
        purple="#B84DFF",
        purple_2="#6D28D9",
        cyan="#22D3EE",
        blue="#2D8CFF",
        pink="#EC4899",
        green="#22C55E",
        orange="#F59E0B",
        red="#FB7185",
        chart_palette_colors=("#22D3EE", "#B84DFF", "#2D8CFF", "#EC4899", "#22C55E", "#F59E0B", "#8B5CF6", "#06B6D4"),
    ),
    about_title="OptiQuant IA - Dark SaaS Premium CSV Dashboard Generator",
    about_description="Cleans any CSV file, detects business semantics and renders an Option A Dark SaaS Premium dashboard with neon accents, rounded cards, click-ready navigation and native Excel charts.",
    sidebar_subtitle="AI CSV Dashboard Engine",
    sidebar_badge="OPTION A · SAAS",
    sidebar_section="WORKSPACE",
    header_title="Executive Overview",
    header_subtitle="Summary of key metrics and insights from your cleaned dataset",
    kpi_segment_label="Segments",
    bar_subtitle="Highest contributing groups",
    donut_subtitle="Distribution across top segments",
    quality_title="Data Quality Overview",
    quality_subtitle="Completeness · Consistency · Validity · Uniqueness",
    ready_line="Ready-to-use workbook: dashboard · clean data · raw data · logs",
)

FINTECH_EXECUTIVE = TemplateProfile(
    slug="fintech-executive",
    display_name="Fintech Executive",
    option_label="Option B",
    theme_family="fintech_executive",
    theme=DashboardTheme(
        bg="#080A0C",
        bg_2="#0E1114",
        sidebar="#0B0D10",
        panel="#15191D",
        panel_2="#1C2227",
        panel_3="#242B31",
        border="#3A424A",
        border_soft="#252B31",
        border_glow="#BFA74A",
        shadow="#030405",
        white="#F8F7F0",
        muted="#B8B4A5",
        muted_2="#7F837B",
        purple="#D6B94C",
        purple_2="#8A6F24",
        cyan="#24A47E",
        blue="#2B746C",
        pink="#BFA74A",
        green="#2ECC8A",
        orange="#E5B84B",
        red="#D66A5C",
        chart_palette_colors=("#D6B94C", "#2ECC8A", "#24A47E", "#E5B84B", "#2B746C", "#6C7A63", "#A38A34", "#4A7C70"),
    ),
    about_title="OptiQuant IA - Fintech Executive CSV Dashboard Generator",
    about_description="Cleans any CSV file, detects business semantics and renders an Option B Fintech Executive dashboard with gold/emerald accents, rounded cards, click-ready navigation and native Excel charts.",
    sidebar_subtitle="Financial CSV Control Room",
    sidebar_badge="OPTION B · FINTECH",
    sidebar_section="CONTROL ROOM",
    header_title="Executive Financial Overview",
    header_subtitle="Key financial performance indicators and controls from your cleaned dataset",
    kpi_segment_label="Counterparties",
    bar_subtitle="Highest contributing revenue/control groups",
    donut_subtitle="Institutional split across key segments",
    quality_title="Controls & Data Quality",
    quality_subtitle="Completeness · Reconciliation · Validity · Uniqueness",
    ready_line="Ready-to-use workbook: financial dashboard · clean data · raw data · logs",
)

LIGHT_CONSULTING = TemplateProfile(
    slug="light-consulting",
    display_name="Light Consulting",
    option_label="Option C",
    theme_family="light_consulting",
    theme=DashboardTheme(
        bg="#F6F8FC",
        bg_2="#EEF3FA",
        sidebar="#FFFFFF",
        panel="#FFFFFF",
        panel_2="#F9FBFF",
        panel_3="#EEF5FF",
        border="#E2E8F0",
        border_soft="#E8EEF7",
        border_glow="#2F80ED",
        shadow="#AAB7C8",
        white="#101828",
        muted="#667085",
        muted_2="#98A2B3",
        purple="#2F80ED",
        purple_2="#B7D4FF",
        cyan="#2DD4BF",
        blue="#2563EB",
        pink="#F59E0B",
        green="#12B76A",
        orange="#F79009",
        red="#F04438",
        chart_palette_colors=("#2F80ED", "#2DD4BF", "#2563EB", "#12B76A", "#F79009", "#7C3AED", "#06B6D4", "#94A3B8"),
    ),
    about_title="OptiQuant IA - Light Consulting CSV Dashboard Generator",
    about_description="Cleans any CSV file, detects business semantics and renders an Option C Light Consulting dashboard with blue/teal/orange accents, rounded cards, click-ready navigation and native Excel charts.",
    sidebar_subtitle="Business CSV Workspace",
    sidebar_badge="OPTION C · CONSULTING",
    sidebar_section="CONTROL ROOM",
    header_title="Business Overview",
    header_subtitle="Clean executive summary, KPIs and analysis from your cleaned dataset",
    kpi_segment_label="Counterparties",
    bar_subtitle="Highest contributing business groups",
    donut_subtitle="Consulting split across key segments",
    quality_title="Controls & Data Quality",
    quality_subtitle="Completeness · Reconciliation · Validity · Uniqueness",
    ready_line="Ready-to-use workbook: consulting dashboard · clean data · raw data · logs",
)


TEMPLATES: dict[str, TemplateProfile] = {
    DARK_SAAS.slug: DARK_SAAS,
    FINTECH_EXECUTIVE.slug: FINTECH_EXECUTIVE,
    LIGHT_CONSULTING.slug: LIGHT_CONSULTING,
}
TEMPLATE_CHOICES = tuple(["auto", *TEMPLATES.keys()])


def get_template(slug: str) -> TemplateProfile:
    try:
        return TEMPLATES[slug]
    except KeyError as exc:
        choices = ", ".join(TEMPLATE_CHOICES)
        raise ValueError(f"Unknown template '{slug}'. Expected one of: {choices}") from exc


def _norm(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").lower()).strip("_")


def choose_template(spec: dict[str, Any]) -> TemplateProfile:
    dataset_type = _norm(spec.get("dataset_type"))
    columns = " ".join(_norm(c) for c in [
        *spec.get("numeric_columns", []),
        *spec.get("categorical_columns", []),
        *spec.get("date_columns", []),
        spec.get("metric"),
        spec.get("primary_dimension"),
        spec.get("secondary_dimension"),
    ])

    if dataset_type in {"invoice"} or any(k in columns for k in ("invoice", "payment", "balance", "paid", "due")):
        return FINTECH_EXECUTIVE
    if dataset_type in {"sales", "marketing", "crm"} or any(k in columns for k in ("campaign", "conversion", "subscriber", "mrr", "arr")):
        return DARK_SAAS
    if any(k in columns for k in ("revenue", "profit", "margin", "amount", "total")):
        return FINTECH_EXECUTIVE
    return LIGHT_CONSULTING
