"""Compatibility exports for the worship PPT generator."""

from src.backend.worship import (
    ChurchApp,
    ChurchAppController,
    ChurchAppUI,
    WorshipPPTGenerator,
    generate_worship_ppt,
)

__all__ = [
    "ChurchApp",
    "ChurchAppController",
    "ChurchAppUI",
    "WorshipPPTGenerator",
    "generate_worship_ppt",
]
