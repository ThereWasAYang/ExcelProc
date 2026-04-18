# ExcelProc

`ExcelProc` processes a `csv` or `xlsx` file, inserts derived columns based on script-defined functions, and then creates a native Excel PivotTable in the output workbook.

## Features

1. Read `csv` or `xlsx`.
2. Apply multiple `(X, func)` transforms, where `X` is a column letter like `A` or `B`, and `func` is a function defined in `excel_processor.py`.
3. Save the result as `xlsx` only.
4. Output files are saved under `outputs/` by default, with the name `<input_stem>_<suffix>.xlsx`.
5. Build a native Excel PivotTable in a new sheet of the output workbook.

## Directory Layout

- `inputs/`: sample and test input files
- `outputs/`: generated output workbooks
- `excel_processor.py`: main script
- `generate_test_files.py`: test data generator

## Environment

The project uses the conda environment `py312`.

## Required Dependencies

```powershell
conda run -n py312 python -m pip install pandas openpyxl pywin32
```

The PivotTable is the real Excel PivotTable from the Excel UI. That means:

- `pywin32` must be installed in `py312`
- Microsoft Excel must be installed on the machine
- the script must be able to launch Excel locally

If any of these requirements is missing, the script will fail directly instead of generating a fallback summary sheet.

## Transform Functions

`func` is not a lambda expression. It must be the name of a function defined in [excel_processor.py](E:/Work/Pycharm/ExcelProc/excel_processor.py).

Current built-in examples:

- `double_value`
- `upper_text`
- `time_to_seconds`

You can add more complex business logic by defining more functions in `excel_processor.py` and then referencing the function names in config or CLI arguments.

## Usage

Example with CLI arguments:

```powershell
conda run -n py312 python .\excel_processor.py `
  --input .\inputs\test_input_100rows.csv `
  --suffix DEMO `
  --transforms "[[\"A\", \"time_to_seconds\"], [\"D\", \"double_value\"]]" `
  --pivot-filters "[\"B\"]" `
  --pivot-rows "[\"C\"]" `
  --pivot-columns "[\"B\"]" `
  --pivot-values "[\"D\"]"
```

Example with config only:

```powershell
conda run -n py312 python .\excel_processor.py `
  --config .\sample_config.json
```

## Config Example

```json
{
  "input": "inputs/sample.csv",
  "suffix": "TEST",
  "transforms": [
    ["C", "double_value"],
    ["E", "upper_text"]
  ],
  "pivot_filters": ["A"],
  "pivot_rows": ["B"],
  "pivot_columns": ["D"],
  "pivot_values": ["C"]
}
```

## Test Data

Generate test files with:

```powershell
conda run -n py312 python .\generate_test_files.py
```

Generated files:

- `inputs/test_input_100rows.csv`
- `inputs/test_input_100rows.xlsx`

Each file contains 100 rows, and the first column is a time string in `HH:MM:SS` format.
The generated test data also includes a `Channel` column that is used by the test config as a PivotTable report filter field.

## Arguments

- `--input`: input `csv/xlsx` path; can also be provided in `--config`
- `--config`: JSON config file; supported keys include `input`, `suffix`, `transforms`, `pivot_filters`, `pivot_rows`, `pivot_columns`, `pivot_values`
- `--output`: explicit output `.xlsx` path
- `--suffix`: suffix appended to the source filename when `--output` is omitted
- `--sheet-name`: source sheet name when the input is `xlsx`
- `--transforms`: JSON array of transforms
- `--pivot-filters`: JSON array of PivotTable report filter column letters
- `--pivot-rows`: JSON array of PivotTable row field column letters
- `--pivot-columns`: JSON array of PivotTable column field column letters
- `--pivot-values`: JSON array of PivotTable value field column letters
- `--pivot-sheet-name`: PivotTable sheet name, default `PivotTable`
- `--data-sheet-name`: normalized source data sheet name, default `SourceData`

## Notes

- Column references use Excel letters such as `A`, `B`, `AA`.
- Each transform is applied against the current worksheet structure after previous inserted columns.
- The generated pivot sheet is an actual Excel PivotTable, not a pandas summary table.
