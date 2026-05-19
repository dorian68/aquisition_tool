"""Unified CSV-to-Excel dashboard generator package."""

from .main import generate_dashboard
from .theme import TEMPLATE_CHOICES, get_template

__all__ = ["TEMPLATE_CHOICES", "generate_dashboard", "get_template"]
