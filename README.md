# ExcelProc

处理 `csv` 或 `xlsx` 文件，支持：

1. 读取源数据。
2. 根据多个 `(列索引, func)` 规则，在指定列右侧插入计算结果列。
3. 无论输入是 `csv` 还是 `xlsx`，结果都另存为 `xlsx`。
4. 默认输出文件名为 `<原文件名>_<suffix>.xlsx`，其中 `suffix` 由用户给定。
5. 在输出文件中创建透视页：优先创建 Excel 原生透视表；若当前环境不可用 Excel COM，则退化为写入同维度的汇总结果表。

## Environment

默认使用 conda 环境 `py312`。

## Install

```powershell
conda run -n py312 python -m pip install pandas openpyxl pywin32
```

如果本机没有安装 Excel，脚本仍可运行，但透视页会是静态汇总结果表，而不是可交互的 Excel 原生透视表。

## Usage

直接传参数：

```powershell
conda run -n py312 python .\excel_processor.py `
  --input .\test_input_100rows.csv `
  --suffix DEMO `
  --transforms "[[\"D\", \"lambda x: x * 2\"]]" `
  --pivot-filters "[\"B\"]" `
  --pivot-rows "[\"C\"]" `
  --pivot-columns "[\"B\"]" `
  --pivot-values "[\"D\"]"
```

这会生成 `test_input_100rows_DEMO.xlsx`。

在 PowerShell 下更推荐使用配置文件：

```powershell
conda run -n py312 python .\excel_processor.py `
  --input .\sample.csv `
  --config .\sample_config.json
```

## Test Data

生成测试文件：

```powershell
conda run -n py312 python .\generate_test_files.py
```

会生成：

- `test_input_100rows.csv`
- `test_input_100rows.xlsx`

两个文件均包含 100 行数据，第一列为 `HH:MM:SS` 格式时间。

## Arguments

- `--input`: 输入文件，支持 `csv/xlsx`
- `--output`: 显式指定输出 `xlsx` 路径；若不传，则按 `--suffix` 自动命名
- `--suffix`: 自动命名时追加到原文件名后的后缀
- `--sheet-name`: 输入为 `xlsx` 时可指定源 sheet
- `--transforms`: 变换规则 JSON 数组
- `--config`: JSON 配置文件，可包含 `suffix/transforms/pivot_filters/pivot_rows/pivot_columns/pivot_values`
- `--pivot-filters`: 透视表筛选器列索引数组
- `--pivot-rows`: 透视表行区域列索引数组
- `--pivot-columns`: 透视表列区域列索引数组
- `--pivot-values`: 透视表值区域列索引数组
- `--pivot-sheet-name`: 透视页名称，默认 `PivotTable`
- `--data-sheet-name`: 输出源数据 sheet 名称，默认 `SourceData`

## Notes

- 列索引使用 Excel 风格字母，如 `A`、`B`、`AA`
- `func` 以可调用表达式传入，例如 `lambda x: x * 2`
- 多个变换按给定顺序依次执行；后续列字母以当前表结构为准
