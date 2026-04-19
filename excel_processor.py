from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd

from processors import FUNCTION_REGISTRY


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "inputs"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs"


@dataclass
class TransformSpec:
    column: str
    func_name: str
    column_mode: str = "header"


@dataclass
class ColumnRef:
    value: str
    mode: str = "header"


@dataclass
class PivotSpec:
    filters: list[ColumnRef]
    rows: list[ColumnRef]
    columns: list[ColumnRef]
    values: list["PivotValueSpec"]


@dataclass
class PivotValueSpec:
    column: str
    summary: str = "sum"
    custom_name: str | None = None
    number_format: str | None = None
    column_mode: str = "header"

PIVOT_SUMMARY_FUNCTIONS = {
    "sum": -4157,
    "count": -4112,
    "average": -4106,
    "avg": -4106,
    "max": -4136,
    "min": -4139,
    "product": -4149,
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Process CSV/XLSX, insert derived columns, and create a native Excel pivot table."
    )
    parser.add_argument("--input", help="Input csv/xlsx path.")
    parser.add_argument(
        "--config",
        help="Optional JSON config file. Supported keys: input, suffix, transforms, pivot_filters, pivot_rows, pivot_columns, pivot_values, pivot_value_settings.",
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
            '[["Amount", "double_value"], {"column_letter": "C", "func": "double_value"}]'
        ),
    )
    parser.add_argument(
        "--pivot-filters",
        default="[]",
        help='JSON array of pivot filter headers, e.g. ["Channel"].',
    )
    parser.add_argument(
        "--pivot-rows",
        default="[]",
        help='JSON array of pivot row headers, e.g. ["Region", "Category"].',
    )
    parser.add_argument(
        "--pivot-columns",
        default="[]",
        help='JSON array of pivot column headers, e.g. ["Segment"].',
    )
    parser.add_argument(
        "--pivot-values",
        required=False,
        help='JSON array of pivot value headers, e.g. ["Amount", "Score"].',
    )
    parser.add_argument(
        "--pivot-value-settings",
        required=False,
        help=(
            "Optional JSON array for native Excel PivotTable value field settings, "
            'e.g. [{"column": "E", "summary": "sum", "name": "Total E", "number_format": "#,##0.00"}].'
        ),
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
            specs.append(TransformSpec(column=item[0], column_mode="header", func_name=item[1]))
            continue
        if isinstance(item, dict) and isinstance(item.get("func"), str):
            if isinstance(item.get("header"), str):
                specs.append(
                    TransformSpec(
                        column=item["header"],
                        column_mode="header",
                        func_name=item["func"],
                    )
                )
                continue
            if isinstance(item.get("column_letter"), str):
                specs.append(
                    TransformSpec(
                        column=item["column_letter"],
                        column_mode="column_letter",
                        func_name=item["func"],
                    )
                )
                continue
            if isinstance(item.get("column"), str):
                specs.append(
                    TransformSpec(
                        column=item["column"],
                        column_mode="header",
                        func_name=item["func"],
                    )
                )
                continue
        raise ValueError(
            'Each transform must be ["ColumnTitle", "function_name"], '
            '{"header": "ColumnTitle", "func": "function_name"}, or '
            '{"column_letter": "A", "func": "function_name"}.'
        )
    return specs


def normalize_pivot_value_settings(raw: str) -> list[PivotValueSpec]:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError("pivot_value_settings must be valid JSON.") from exc
    if not isinstance(value, list):
        raise ValueError("pivot_value_settings must be a JSON array.")

    specs: list[PivotValueSpec] = []
    for item in value:
        if not isinstance(item, dict):
            raise ValueError(
                "Each pivot value setting must be an object like "
                '{"column": "E", "summary": "sum", "name": "Total E", "number_format": "#,##0.00"}.'
            )
        column = item.get("column")
        header = item.get("header")
        column_letter = item.get("column_letter")
        summary = item.get("summary", "sum")
        custom_name = item.get("name")
        number_format = item.get("number_format")
        if not isinstance(summary, str):
            raise ValueError("pivot value setting summary must be a string.")
        if custom_name is not None and not isinstance(custom_name, str):
            raise ValueError("pivot value setting name must be a string when provided.")
        if number_format is not None and not isinstance(number_format, str):
            raise ValueError("pivot value setting number_format must be a string when provided.")

        if isinstance(header, str):
            column_value = header
            column_mode = "header"
        elif isinstance(column_letter, str):
            column_value = column_letter
            column_mode = "column_letter"
        elif isinstance(column, str):
            column_value = column
            column_mode = "header"
        else:
            raise ValueError(
                "pivot value setting must contain 'header', 'column_letter', or legacy 'column'."
            )

        specs.append(
            PivotValueSpec(
                column=column_value,
                summary=summary.strip().lower(),
                custom_name=custom_name,
                number_format=number_format,
                column_mode=column_mode,
            )
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


def normalize_column_refs(raw: str, field_name: str) -> list[ColumnRef]:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"{field_name} must be valid JSON.") from exc
    if not isinstance(value, list):
        raise ValueError(f"{field_name} must be a JSON array.")

    refs: list[ColumnRef] = []
    for item in value:
        if isinstance(item, str):
            refs.append(ColumnRef(value=item, mode="header"))
            continue
        if isinstance(item, dict):
            if isinstance(item.get("header"), str):
                refs.append(ColumnRef(value=item["header"], mode="header"))
                continue
            if isinstance(item.get("column_letter"), str):
                refs.append(ColumnRef(value=item["column_letter"], mode="column_letter"))
                continue
            if isinstance(item.get("column"), str):
                refs.append(ColumnRef(value=item["column"], mode="header"))
                continue
        raise ValueError(
            f"Each item in {field_name} must be a header string, "
            '{"header": "ColumnTitle"}, or {"column_letter": "A"}.'
        )
    return refs


def get_registered_function(func_name: str) -> Callable[[Any], Any]:
    func = FUNCTION_REGISTRY.get(func_name)
    if func is None:
        available = ", ".join(sorted(FUNCTION_REGISTRY))
        raise ValueError(f"Unknown transform function: {func_name}. Available: {available}")
    return func


def resolve_transform_column_name(columns: pd.Index, spec: TransformSpec) -> str:
    if spec.column_mode == "header":
        if columns.duplicated().any():
            duplicate_names = sorted({str(name) for name in columns[columns.duplicated()]})
            raise ValueError(
                "Transform column lookup by title requires unique headers. "
                f"Duplicate headers found: {', '.join(duplicate_names)}"
            )
        if spec.column not in columns:
            available = ", ".join(map(str, columns))
            raise KeyError(
                f"Transform header not found: {spec.column}. Available headers: {available}"
            )
        return str(spec.column)

    if spec.column_mode == "column_letter":
        source_index = excel_column_to_index(spec.column)
        if source_index < 0 or source_index >= len(columns):
            raise IndexError(f"Transform column out of range: {spec.column}")
        return str(columns[source_index])

    raise ValueError(f"Unsupported transform column mode: {spec.column_mode}")


def resolve_column_ref_name(columns: pd.Index, ref: ColumnRef, label: str) -> str:
    if ref.mode == "header":
        if columns.duplicated().any():
            duplicate_names = sorted({str(name) for name in columns[columns.duplicated()]})
            raise ValueError(
                f"{label} lookup by title requires unique headers. "
                f"Duplicate headers found: {', '.join(duplicate_names)}"
            )
        if ref.value not in columns:
            available = ", ".join(map(str, columns))
            raise KeyError(
                f"{label} header not found: {ref.value}. Available headers: {available}"
            )
        return str(ref.value)

    if ref.mode == "column_letter":
        source_index = excel_column_to_index(ref.value)
        if source_index < 0 or source_index >= len(columns):
            raise IndexError(f"{label} column out of range: {ref.value}")
        return str(columns[source_index])

    raise ValueError(f"Unsupported {label} mode: {ref.mode}")


def apply_transforms(frame: pd.DataFrame, specs: list[TransformSpec]) -> pd.DataFrame:
    result = frame.copy()
    for spec in specs:
        source_name = resolve_transform_column_name(result.columns, spec)
        source_index = result.columns.get_loc(source_name)
        if not isinstance(source_index, int):
            raise ValueError(f"Transform column must resolve to a single column: {spec.column}")
        func = get_registered_function(spec.func_name)
        derived_name = f"{source_name}_{spec.func_name}"
        insert_index = source_index + 1
        result.insert(insert_index, derived_name, result[source_name].apply(func))
    return result


def column_refs_to_names(frame: pd.DataFrame, refs: list[ColumnRef], label: str) -> list[str]:
    names: list[str] = []
    for ref in refs:
        names.append(resolve_column_ref_name(frame.columns, ref, label))
    return names


def build_pivot_value_specs(
    pivot_values_raw: str | None,
    pivot_value_settings_raw: str | None,
) -> list[PivotValueSpec]:
    if pivot_value_settings_raw is not None:
        specs = normalize_pivot_value_settings(pivot_value_settings_raw)
        if not specs:
            raise ValueError("pivot_value_settings cannot be empty when provided.")
        return specs

    if pivot_values_raw is None:
        raise ValueError(
            "pivot_values or pivot_value_settings is required, either via CLI or --config."
        )

    refs = normalize_column_refs(pivot_values_raw, "pivot_values")
    if not refs:
        raise ValueError("pivot_values cannot be empty.")
    return [PivotValueSpec(column=ref.value, column_mode=ref.mode) for ref in refs]


def get_pivot_summary_function(summary: str) -> int:
    excel_function = PIVOT_SUMMARY_FUNCTIONS.get(summary.strip().lower())
    if excel_function is None:
        available = ", ".join(sorted(PIVOT_SUMMARY_FUNCTIONS))
        raise ValueError(
            f"Unsupported pivot value summary: {summary}. Available: {available}"
        )
    return excel_function


def default_data_field_name(source_name: str, summary: str) -> str:
    summary_labels = {
        "sum": "Sum of",
        "count": "Count of",
        "average": "Average of",
        "avg": "Average of",
        "max": "Max of",
        "min": "Min of",
        "product": "Product of",
    }
    prefix = summary_labels.get(summary.strip().lower(), summary.strip().title())
    return f"{prefix} {source_name}"


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
            "Creating a native Excel pivot table requires pywin32 in the current Python environment."
        ) from exc

    row_names = column_refs_to_names(frame, spec.rows, "pivot row")
    column_names = column_refs_to_names(frame, spec.columns, "pivot column")
    filter_names = column_refs_to_names(frame, spec.filters, "pivot filter")

    if not spec.values:
        raise ValueError("pivot_values cannot be empty.")

    value_field_specs: list[tuple[str, PivotValueSpec]] = []
    for value_spec in spec.values:
        value_name = resolve_column_ref_name(
            frame.columns,
            ColumnRef(value=value_spec.column, mode=value_spec.column_mode),
            "pivot value",
        )
        value_field_specs.append((value_name, value_spec))

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

        for source_name, value_spec in value_field_specs:
            field = pivot_table.PivotFields(source_name)
            data_field = pivot_table.AddDataField(
                field,
                value_spec.custom_name or default_data_field_name(source_name, value_spec.summary),
                get_pivot_summary_function(value_spec.summary),
            )
            if value_spec.number_format:
                data_field.NumberFormat = value_spec.number_format

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
    pivot_value_settings_raw = args.pivot_value_settings
    if pivot_value_settings_raw is None and "pivot_value_settings" in config:
        pivot_value_settings_raw = json.dumps(
            config["pivot_value_settings"], ensure_ascii=False
        )

    transforms = normalize_transforms(transforms_raw)
    pivot_spec = PivotSpec(
        filters=normalize_column_refs(
            args.pivot_filters
            if args.pivot_filters != "[]"
            else json.dumps(config.get("pivot_filters", []), ensure_ascii=False),
            "pivot_filters",
        ),
        rows=normalize_column_refs(
            args.pivot_rows
            if args.pivot_rows != "[]"
            else json.dumps(config.get("pivot_rows", []), ensure_ascii=False),
            "pivot_rows",
        ),
        columns=normalize_column_refs(
            args.pivot_columns
            if args.pivot_columns != "[]"
            else json.dumps(config.get("pivot_columns", []), ensure_ascii=False),
            "pivot_columns",
        ),
        values=build_pivot_value_specs(pivot_values_raw, pivot_value_settings_raw),
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
