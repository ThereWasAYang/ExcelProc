from __future__ import annotations

from datetime import time
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "inputs"


def build_test_frame(row_count: int = 100) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    categories = ["A", "B", "C", "D"]
    regions = ["East", "West", "North", "South"]
    channels = ["Online", "Store", "Partner"]
    priorities = ["High", "Medium", "Low"]
    segments = ["Retail", "Wholesale"]
    quarters = ["Q1", "Q2", "Q3", "Q4"]

    for idx in range(row_count):
        second_of_day = (idx * 137) % 86400
        hh = second_of_day // 3600
        mm = (second_of_day % 3600) // 60
        ss = second_of_day % 60
        rows.append(
            {
                "Time": time(hour=hh, minute=mm, second=ss).strftime("%H:%M:%S"),
                "Category": categories[idx % len(categories)],
                "Region": regions[(idx // len(categories)) % len(regions)],
                "Amount": 100 + idx * 3,
                "Score": round(60 + (idx % 17) * 1.5, 1),
                "Channel": channels[idx % len(channels)],
                "Priority": priorities[idx % len(priorities)],
                "Segment": segments[(idx // 2) % len(segments)],
                "Quarter": quarters[(idx // 8) % len(quarters)],
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    DEFAULT_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    frame = build_test_frame(100)
    csv_path = DEFAULT_INPUT_DIR / "test_input_100rows.csv"
    xlsx_path = DEFAULT_INPUT_DIR / "test_input_100rows.xlsx"

    frame.to_csv(csv_path, index=False)
    frame.to_excel(xlsx_path, index=False, sheet_name="Sheet1")

    print(f"Created {csv_path}")
    print(f"Created {xlsx_path}")


if __name__ == "__main__":
    main()
