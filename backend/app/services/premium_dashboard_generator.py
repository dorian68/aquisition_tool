from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


def _ensure_project_root_on_path() -> None:
    project_root = Path(__file__).resolve().parents[3]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))


try:
    from dashboard_generator.main import generate_dashboard
    from dashboard_generator.theme import TEMPLATE_CHOICES
except ModuleNotFoundError:
    _ensure_project_root_on_path()
    from dashboard_generator.main import generate_dashboard
    from dashboard_generator.theme import TEMPLATE_CHOICES


class PremiumDashboardGenerationError(ValueError):
    pass


def generate_premium_dashboard(
    csv_path: str | Path,
    output_path: str | Path,
    *,
    template: str,
    vba_project_path: str | Path | None = None,
    client_name: str | None = None,
    hide_settings: bool = False,
) -> dict[str, Any]:
    try:
        return generate_dashboard(
            csv_path,
            output_path,
            template=template,
            vba_project_path=vba_project_path,
            write_macro_source_file=False,
            client_name=client_name,
            hide_settings=hide_settings,
        )
    except (FileNotFoundError, ValueError) as exc:
        raise PremiumDashboardGenerationError(str(exc)) from exc
