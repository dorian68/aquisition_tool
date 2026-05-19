from __future__ import annotations

from pathlib import Path
from typing import Any

from ..main import generate_dashboard
from ..theme import FINTECH_EXECUTIVE as PROFILE


def generate(csv_path: str | Path, output_path: str | Path, **kwargs: Any) -> dict[str, Any]:
    return generate_dashboard(csv_path, output_path, template=PROFILE.slug, **kwargs)
