from __future__ import annotations

import logging
import json
from typing import Any

from app.core.config import Settings, get_settings
from .fallback import build_deterministic_ai_report
from .graph import run_ai_analyst_graph
from .schemas import AIReport

logger = logging.getLogger(__name__)


def generate_ai_dashboard_report(ai_context: dict[str, Any], *, settings: Settings | None = None) -> dict[str, Any]:
    return generate_ai_dashboard_report_result(ai_context, settings=settings)["ai_report"]


def generate_ai_dashboard_report_result(ai_context: dict[str, Any], *, settings: Settings | None = None) -> dict[str, Any]:
    settings = settings or get_settings()
    if not settings.ai_analyst_enabled or not settings.openai_api_key:
        return _fallback_result(ai_context, settings=settings, errors=[])
    try:
        result = run_ai_analyst_graph(ai_context, settings=settings)
        report = AIReport.model_validate(result["ai_report"]).model_dump()
        metadata = result.get("metadata", {})
        metadata.setdefault("provider", "langgraph")
        metadata.setdefault("model", settings.ai_analyst_model)
        metadata["ai_context_chars"] = _serialized_context_length(ai_context)
        return {"ai_report": report, "metadata": metadata}
    except Exception as exc:  # noqa: BLE001 - AI must fail open.
        logger.warning("AI analyst service failed safely: %s", exc)
        if settings.ai_analyst_fail_open:
            return _fallback_result(ai_context, settings=settings, errors=[str(exc)])
        raise


def _fallback_result(ai_context: dict[str, Any], *, settings: Settings, errors: list[str]) -> dict[str, Any]:
    return {
        "ai_report": build_deterministic_ai_report(ai_context),
        "metadata": {
            "provider": "langgraph",
            "model": settings.ai_analyst_model if settings.openai_api_key else None,
            "used_fallback": True,
            "ai_context_chars": _serialized_context_length(ai_context),
            "errors": [error[:300] for error in errors[:3]],
        },
    }


def _serialized_context_length(ai_context: dict[str, Any]) -> int:
    return len(json.dumps(ai_context, ensure_ascii=True, separators=(",", ":")))
