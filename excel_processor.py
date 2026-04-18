from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "inputs"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs"


@dataclass
class TransformSpec:
    column: str
    func_name: str


@dataclass
class PivotSpec:
    filters: list[str]
    rows: list[str]
    columns: list[str]
    values: list[str]


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


FUNCTION_REGISTRY: dict[str, Callable[[Any], Any]] = {
    "double_value": double_value,
    "upper_text": upper_text,
    "time_to_seconds": time_to_seconds,
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Process CSV/XLSX, insert derived columns, and create a native Excel pivot table."
    )
    parser.add_argument("--input", help="Input csv/xlsx path.")
    parser.add_argument(
        "--config",
        help="Optional JSON config file. Supported keys: input, suffix, transforms, pivot_filters, pivot_rows, pivot_columns, pivot_values.",
    )
    parser.add_argument(
        "--output",
        help="Optional explicit output xlsx path. If omitted, save to outputs/<input_stem>_<suffix>.xlsx.",
    )
    parser.add_argument(
        "--suffix",
        help="Suffix appended to the source filename when --output is omitted, e.g. input_A.xlsx.",
    )
    parser.add_argument(
        "--sheet-name",
        default=None,
        help="Source sheet name when reading xlsx. Defaults to the first sheet.",
    )
    parser.add_argument(
        "--transforms",
        required=False,
        help=(
            "JSON array, e.g. "
            '[["A", "double_value"], ["C", "upper_text"]]'
        ),
    )
    parser.add_argument(
        "--pivot-filters",
        default="[]",
        help='JSON array of column letters for pivot filters, e.g. ["A"].',
    )
    parser.add_argument(
        "--pivot-rows",
        default="[]",
        help='JSON array of column letters for pivot rows, e.g. ["B", "C"].',
    )
    parser.add_argument(
        "--pivot-columns",
        default="[]",
        help='JSON array of column letters for pivot columns, e.g. ["D"].',
    )
    parser.add_argument(
        "--pivot-values",
        required=False,
        help='JSON array of column letters for pivot values, e.g. ["E", "F"].',
    )
    parser.add_argument(
        "--pivot-sheet-name",
        default="PivotTable",
        help="Name of the native Excel pivot table sheet.",
    )
    parser.add_argument(
        "--data-sheet-name",
        default="SourceData",
        help="Name of the normalized source data sheet written to the output workbook.",
    )
    return parser.parse_args()


def load_frame(input_path: Path, sheet_name: str | None) -> pd.DataFrame:
    suffix = input_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(input_path)
    if suffix in EXCEL_EXTENSIONS:
        return pd.read_excel(input_path, sheet_name=sheet_name if sheet_name is not None else 0)
    raise ValueError(f"Unsupported input type: {suffix}")


def save_workbook(base_frame: pd.DataFrame, output_path: Path, data_sheet_name: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        base_frame.to_excel(writer, index=False, sheet_name=data_sheet_name)


def normalize_json_array(raw: str, field_name: str) -> list[str]:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"{field_name} must be valid JSON.") from exc
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{field_name} must be a JSON string array.")
    return value


def normalize_transforms(raw: str) -> list[TransformSpec]:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError("transforms must be valid JSON.") from exc
    if not isinstance(value, list):
        raise ValueError("transforms must be a JSON array.")

    specs: list[TransformSpec] = []
    for item in value:
        if (
            isinstance(item, list)
            and len(item) == 2
            and all(isinstance(part, str) for part in item)
        ):
            specs.append(TransformSpec(column=item[0], func_name=item[1]))
            continue
        if isinstance(item, dict) and isinstance(item.get("column"), str) and isinstance(
            item.get("func"), str
        ):
            specs.append(TransformSpec(column=item["column"], func_name=item["func"]))
            continue
        raise ValueError(
            'Each transform must be ["A", "function_name"] or {"column": "A", "func": "function_name"}'
        )
    return specs


def load_config(config_path: Path | None) -> dict[str, Any]:
    if config_path is None:
        return {}
    with config_path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("Config file must contain a JSON object.")
    return data


def excel_column_to_index(column_ref: str) -> int:
    ref = column_ref.strip().upper()
    if not ref.isalpha():
        raise ValueError(f"Invalid column reference: {column_ref}")
    value = 0
    for char in ref:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def get_registered_function(func_name: str) -> Callable[[Any], Any]:
    func = FUNCTION_REGISTRY.get(func_name)
    if func is None:
        available = ", ".join(sorted(FUNCTION_REGISTRY))
        raise ValueError(f"Unknown transform function: {func_name}. Available: {available}")
    return func


def apply_transforms(frame: pd.DataFrame, specs: list[TransformSpec]) -> pd.DataFrame:
    result = frame.copy()
    for spec in specs:
        # Each transform uses the current worksheet structure after prior inserts.
        source_index = excel_column_to_index(spec.column)
        if source_index < 0 or source_index >= len(result.columns):
            raise IndexError(f"Column out of range: {spec.column}")
        source_name = result.columns[source_index]
        func = get_registered_function(spec.func_name)
        derived_name = f"{source_name}_{spec.func_name}"
        insert_index = source_index + 1
        result.insert(insert_index, derived_name, result[source_name].apply(func))
    return result


def column_letters_to_names(frame: pd.DataFrame, letters: list[str]) -> list[str]:
    names: list[str] = []
    for letter in letters:
        index = excel_column_to_index(letter)
        if index < 0 or index >= len(frame.columns):
            raise IndexError(f"Column out of range: {letter}")
        names.append(str(frame.columns[index]))
    return names


def create_excel_pivot_table(
    output_path: Path,
    data_sheet_name: str,
    pivot_sheet_name: str,
    spec: PivotSpec,
    frame: pd.DataFrame,
) -> None:
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "Creating a native Excel pivot table requires pywin32 in the py312 environment."
        ) from exc

    row_names = column_letters_to_names(frame, spec.rows)
    column_names = column_letters_to_names(frame, spec.columns)
    value_names = column_letters_to_names(frame, spec.values)
    filter_names = column_letters_to_names(frame, spec.filters)

    if not value_names:
        raise ValueError("pivot_values cannot be empty.")

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(output_path.resolve()))
        source_sheet = workbook.Worksheets(data_sheet_name)

        try:
            old_sheet = workbook.Worksheets(pivot_sheet_name)
            old_sheet.Delete()
        except Exception:
            pass

        last_row = source_sheet.Cells(source_sheet.Rows.Count, 1).End(-4162).Row
        last_col = source_sheet.Cells(1, source_sheet.Columns.Count).End(-4159).Column
        source_range = source_sheet.Range(source_sheet.Cells(1, 1), source_sheet.Cells(last_row, last_col))

        cache = workbook.PivotCaches().Create(SourceType=1, SourceData=source_range)
        pivot_sheet = workbook.Worksheets.Add()
        pivot_sheet.Name = pivot_sheet_name
        pivot_table = cache.CreatePivotTable(
            TableDestination=f"{pivot_sheet_name}!R3C1",
            TableName="GeneratedPivotTable",
        )

        for name in filter_names:
            field = pivot_table.PivotFields(name)
            field.Orientation = 3

        for name in row_names:
            field = pivot_table.PivotFields(name)
            field.Orientation = 1

        for name in column_names:
            field = pivot_table.PivotFields(name)
            field.Orientation = 2

        for name in value_names:
            field = pivot_table.PivotFields(name)
            pivot_table.AddDataField(field, f"Sum of {name}", -4157)

        workbook.Save()
    except Exception as exc:
        raise RuntimeError(
            "Failed to create a native Excel pivot table. Ensure Microsoft Excel is installed "
            "and the workbook can be opened locally."
        ) from exc
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def main() -> int:
    args = parse_args()
    config = load_config(Path(args.config).expanduser().resolve() if args.config else None)

    input_value = args.input if args.input is not None else config.get("input")
    if not input_value or not isinstance(input_value, str):
        raise ValueError("input is required, either via --input or --config.")

    input_path = Path(input_value).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    suffix = args.suffix if args.suffix is not None else config.get("suffix")
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        if not suffix or not isinstance(suffix, str):
            raise ValueError("suffix is required when --output is not provided.")
        output_path = (DEFAULT_OUTPUT_DIR / f"{input_path.stem}_{suffix}.xlsx").resolve()

    transforms_raw = args.transforms
    if transforms_raw is None and "transforms" in config:
        transforms_raw = json.dumps(config["transforms"], ensure_ascii=False)
    if transforms_raw is None:
        raise ValueError("transforms is required, either via --transforms or --config.")

    pivot_values_raw = args.pivot_values
    if pivot_values_raw is None and "pivot_values" in config:
        pivot_values_raw = json.dumps(config["pivot_values"], ensure_ascii=False)
    if pivot_values_raw is None:
        raise ValueError("pivot_values is required, either via --pivot-values or --config.")

    transforms = normalize_transforms(transforms_raw)
    pivot_spec = PivotSpec(
        filters=normalize_json_array(
            args.pivot_filters
            if args.pivot_filters != "[]"
            else json.dumps(config.get("pivot_filters", []), ensure_ascii=False),
            "pivot_filters",
        ),
        rows=normalize_json_array(
            args.pivot_rows
            if args.pivot_rows != "[]"
            else json.dumps(config.get("pivot_rows", []), ensure_ascii=False),
            "pivot_rows",
        ),
        columns=normalize_json_array(
            args.pivot_columns
            if args.pivot_columns != "[]"
            else json.dumps(config.get("pivot_columns", []), ensure_ascii=False),
            "pivot_columns",
        ),
        values=normalize_json_array(pivot_values_raw, "pivot_values"),
    )

    frame = load_frame(input_path, args.sheet_name)
    processed = apply_transforms(frame, transforms)
    save_workbook(processed, output_path, args.data_sheet_name)

    create_excel_pivot_table(
        output_path=output_path,
        data_sheet_name=args.data_sheet_name,
        pivot_sheet_name=args.pivot_sheet_name,
        spec=pivot_spec,
        frame=processed,
    )
    print("Created native Excel pivot table.")

    print(f"Saved output to: {output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
