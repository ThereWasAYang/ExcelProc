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


def double_value_series(series: pd.Series) -> pd.Series:
    return series * 2


def upper_text_series(series: pd.Series) -> pd.Series:
    return series.astype("string").str.upper()


def time_to_seconds_series(series: pd.Series) -> pd.Series:
    timedeltas = pd.to_timedelta(series.astype("string"), errors="raise")
    seconds = timedeltas.dt.total_seconds()
    return seconds.astype("Int64")


# Register all transform functions here. The config file uses these keys.
FUNCTION_REGISTRY: dict[str, Callable[[Any], Any]] = {
    "double_value": double_value,
    "upper_text": upper_text,
    "time_to_seconds": time_to_seconds,
}


# Optional fast paths for large files. Keys match FUNCTION_REGISTRY.
VECTOR_FUNCTION_REGISTRY: dict[str, Callable[[pd.Series], pd.Series]] = {
    "double_value": double_value_series,
    "upper_text": upper_text_series,
    "time_to_seconds": time_to_seconds_series,
}
