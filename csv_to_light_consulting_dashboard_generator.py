from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from dashboard_generator.main import generate_dashboard


def generate_light_consulting_dashboard(
    csv_path: str | Path,
    output_path: str | Path,
    *,
    vba_project_path: str | Path | None = None,
    macro_source_path: str | Path | None = None,
    write_macro_source_file: bool = True,
    client_name: str | None = None,
    hide_settings: bool = False,
) -> dict[str, Any]:
    return generate_dashboard(
        csv_path,
        output_path,
        template="light-consulting",
        vba_project_path=vba_project_path,
        macro_source_path=macro_source_path,
        write_macro_source_file=write_macro_source_file,
        client_name=client_name,
        hide_settings=hide_settings,
    )


generate_crm_style_dashboard = generate_light_consulting_dashboard
generate_dark_saas_dashboard = generate_light_consulting_dashboard
generate_fintech_executive_dashboard = generate_light_consulting_dashboard


def main() -> None:
    parser = argparse.ArgumentParser(description="Compatibility wrapper for the Light Consulting dashboard template.")
    parser.add_argument("csv_path", help="Path to input CSV")
    parser.add_argument("output_path", nargs="?", default="light_consulting_dashboard.xlsx", help="Path to output XLSX/XLSM")
    parser.add_argument("--vba-project", default=None, help="Path to a precompiled vbaProject.bin to embed in generated XLSM")
    parser.add_argument("--macro-source-output", default=None, help="Where to write the DashboardMacros.bas source module")
    parser.add_argument("--no-macro-source", action="store_true", help="Do not export the DashboardMacros.bas source file")
    parser.add_argument("--client-name", default=None, help="Optional client name written to the Settings sheet")
    parser.add_argument("--hide-settings", action="store_true", help="Hide the Settings control panel sheet")
    args = parser.parse_args()

    result = generate_light_consulting_dashboard(
        args.csv_path,
        args.output_path,
        vba_project_path=args.vba_project,
        macro_source_path=args.macro_source_output,
        write_macro_source_file=not args.no_macro_source,
        client_name=args.client_name,
        hide_settings=args.hide_settings,
    )
    print("Generated:", result["output"])
    print("Template:", result["selected_template"])
    print("Detected type:", result["dataset_type"])
    print("Quality:", f"{result['quality_before']}% -> {result['quality_after']}%")
    print("Macros embedded:", result["macros_embedded"])
    if result.get("macro_source"):
        print("Macro source:", result["macro_source"])


if __name__ == "__main__":
    main()
