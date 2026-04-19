from __future__ import annotations

from typing import Any, Callable

import pandas as pd


def double_value(value: Any) -> Any:
    if pd.isna(value):
        return value
    return value * 2


def upper_text(value: Any) -> Any:
    if pd.isna(value):
        return value
    return str(value).upper()


def time_to_seconds(value: Any) -> Any:
    if pd.isna(value):
        return value
    parts = str(value).split(":")
    if len(parts) != 3:
        raise ValueError(f"Expected HH:MM:SS, got: {value}")
    hours, minutes, seconds = map(int, parts)
    return hours * 3600 + minutes * 60 + seconds


# Register all transform functions here. The config file uses these keys.
FUNCTION_REGISTRY: dict[str, Callable[[Any], Any]] = {
    "double_value": double_value,
    "upper_text": upper_text,
    "time_to_seconds": time_to_seconds,
}
