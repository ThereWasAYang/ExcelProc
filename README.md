# ExcelProc

`ExcelProc` 用于处理 `csv` 或 `xlsx` 文件，支持按指定函数插入新列，并在输出工作簿中创建 Excel 原生数据透视表。

## 功能说明

1. 读取 `csv` 或 `xlsx` 文件。
2. 支持多个 `(X, func)` 转换规则，其中 `X` 推荐使用列标题，`func` 是已注册的数据处理函数名。
3. 无论输入文件是 `csv` 还是 `xlsx`，输出都统一保存为 `xlsx`。
4. 默认将结果输出到 `outputs/` 目录，文件名格式为 `<原文件名>_<suffix>.xlsx`。
5. 在输出文件中新建一个 Excel 原生数据透视表工作表。

## 目录结构

- `inputs/`：输入样本和测试数据
- `configs/`：示例配置和测试配置
- `outputs/`：输出文件
- `scripts/`：辅助脚本
- `processors/`：数据处理函数
- `excel_processor.py`：主处理脚本
- `setup_env.bat`：Windows 下一键配置环境脚本

## 运行环境

项目依赖 Python 3.12。

如果你使用 conda、venv 或其他环境管理工具，请在你自己的本地 Python 环境中安装依赖并运行脚本。项目文档不依赖任何固定的环境名称。

## 依赖安装

```powershell
python -m pip install pandas openpyxl pywin32
```

## 一键配置环境

如果项目使用者不熟悉 Python 环境配置，建议直接双击运行：

[setup_env.bat](E:/Work/Pycharm/ExcelProc/setup_env.bat)

这个脚本会自动完成以下操作：

1. 检测本机可用的 Python
2. 在项目目录下创建虚拟环境 `.venv`
3. 自动升级 `pip`
4. 自动安装 `requirements.txt` 中的依赖
5. 自动创建 `inputs`、`configs`、`outputs` 目录（如果缺失）

运行完成后，可以直接使用下面的命令启动项目：

```powershell
.venv\Scripts\python.exe .\excel_processor.py --config .\configs\sample_config.json
```

如果双击脚本时提示没有 Python，需要先在电脑上安装 Python 3.12 或更高版本。

## Excel 依赖说明

数据透视表使用的是 Excel 自带的“插入 -> 数据透视表”能力，因此需要满足以下条件：

- 当前 Python 环境已安装 `pywin32`
- 本机已安装 Microsoft Excel
- 当前环境允许通过 Excel COM 启动 Excel

如果以上条件不满足，脚本会直接报错，不会退化成普通汇总表。

## 数据处理函数

数据处理函数现在统一放在 [processors/transform_functions.py](E:/Work/Pycharm/ExcelProc/processors/transform_functions.py) 中管理。

当前内置示例函数：

- `double_value`
- `upper_text`
- `time_to_seconds`

### 如何添加新的数据处理函数

1. 打开 [processors/transform_functions.py](E:/Work/Pycharm/ExcelProc/processors/transform_functions.py)
2. 新增一个函数，函数接收单个单元格值作为参数，并返回处理后的结果
3. 在同文件底部的 `FUNCTION_REGISTRY` 中注册这个函数
4. 在配置文件里的 `transforms` 中引用注册名

例如，新增一个把字符串前后空格去掉的函数：

```python
def strip_text(value: Any) -> Any:
    if pd.isna(value):
        return value
    return str(value).strip()
```

然后把它注册到 `FUNCTION_REGISTRY`：

```python
FUNCTION_REGISTRY = {
    "double_value": double_value,
    "upper_text": upper_text,
    "time_to_seconds": time_to_seconds,
    "strip_text": strip_text,
}
```

之后就可以在配置文件中这样使用：

```json
{
  "transforms": [
    ["Name", "strip_text"]
  ]
}
```

### 编写函数时的建议

- 优先处理空值，例如使用 `pd.isna(value)`
- 函数应只处理单个单元格值，不要直接依赖整列 DataFrame
- 返回值应是可写入 Excel 的普通值类型，如字符串、数字、布尔值或空值
- 如果函数只适用于特定格式，建议在函数内主动校验并抛出明确错误

## 使用方式

### 方式一：直接传参

```powershell
python .\excel_processor.py `
  --input .\inputs\test_input_100rows.csv `
  --suffix DEMO `
  --transforms "[[\"Time\", \"time_to_seconds\"], [\"Amount\", \"double_value\"]]" `
  --pivot-filters "[\"Channel\", \"Priority\"]" `
  --pivot-rows "[\"Region\", \"Category\"]" `
  --pivot-columns "[\"Segment\", \"Quarter\"]" `
  --pivot-value-settings "[{\"header\": \"Amount\", \"summary\": \"sum\", \"name\": \"Total Amount\", \"number_format\": \"#,##0\"}, {\"header\": \"Score\", \"summary\": \"average\", \"name\": \"Average Score\", \"number_format\": \"0.0\"}]"
```

### 方式二：使用配置文件

```powershell
python .\excel_processor.py `
  --config .\configs\sample_config.json
```

## 配置示例

```json
{
  "input": "inputs/sample.csv",
  "suffix": "TEST",
  "transforms": [
    ["Amount", "double_value"],
    ["Name", "upper_text"]
  ],
  "pivot_filters": ["Category"],
  "pivot_rows": ["Region"],
  "pivot_columns": ["Name"],
  "pivot_value_settings": [
    {
      "header": "Amount",
      "summary": "sum",
      "name": "Total Amount",
      "number_format": "#,##0.00"
    }
  ]
}
```

## 列索引写法

配置文件中的列索引目前支持两种表示方式。

### 方式一：按表头名称

这是推荐写法，含义最直接，也不会因为前面插入新列而影响后续规则。

```json
{
  "transforms": [
    ["Amount", "double_value"],
    {"header": "Name", "func": "upper_text"}
  ],
  "pivot_filters": ["Channel", "Priority"],
  "pivot_rows": ["Region", "Category"],
  "pivot_columns": ["Segment", "Quarter"],
  "pivot_value_settings": [
    {
      "header": "Amount",
      "summary": "sum",
      "name": "Total Amount"
    }
  ]
}
```

### 方式二：按 Excel 列字母

如果你确实希望按原始列位置指定，可以显式写成 `column_letter`。

```json
{
  "transforms": [
    {"column_letter": "A", "func": "time_to_seconds"},
    {"column_letter": "D", "func": "double_value"}
  ],
  "pivot_filters": [
    {"column_letter": "G"},
    {"column_letter": "H"}
  ],
  "pivot_rows": [
    {"column_letter": "C"},
    {"column_letter": "B"}
  ],
  "pivot_columns": [
    {"column_letter": "I"},
    {"column_letter": "J"}
  ],
  "pivot_value_settings": [
    {
      "column_letter": "D",
      "summary": "sum",
      "name": "Total Amount"
    }
  ]
}
```

## 测试数据

[scripts/generate_test_files.py](E:/Work/Pycharm/ExcelProc/scripts/generate_test_files.py) 仅用于生成测试样本。

执行命令：

```powershell
python .\scripts\generate_test_files.py
```

生成文件：

- `inputs/test_input_100rows.csv`
- `inputs/test_input_100rows.xlsx`

测试样本特点：

- 共 100 行
- 第一列为 `HH:MM:SS` 格式时间
- 包含 `Channel`、`Priority`、`Segment`、`Quarter` 等分类字段，可用于更充分地验证数据透视表的筛选、行、列区域配置

## 参数说明

- `--input`：输入文件路径，支持 `csv/xlsx`，也可以写在 `--config` 中
- `--config`：JSON 配置文件
  示例路径：`configs/sample_config.json`
- `--output`：显式指定输出 `.xlsx` 路径
- `--suffix`：未显式指定 `--output` 时，追加到输出文件名后的后缀
- `--sheet-name`：当输入为 `xlsx` 时，指定读取的源工作表
- `--transforms`：转换规则 JSON 数组
- `--pivot-filters`：数据透视表“筛选”区域字段数组，默认按表头名称解析
- `--pivot-rows`：数据透视表“行”区域字段数组，默认按表头名称解析
- `--pivot-columns`：数据透视表“列”区域字段数组，默认按表头名称解析
- `--pivot-values`：数据透视表“值”区域字段数组，默认按表头名称解析
- `--pivot-value-settings`：数据透视表“值字段设置”数组，支持 `header`、`column_letter`、兼容旧字段 `column`、以及 `summary`、`name`、`number_format`
- `--pivot-sheet-name`：数据透视表工作表名称，默认 `PivotTable`
- `--data-sheet-name`：输出源数据工作表名称，默认 `SourceData`

## 说明补充

- `transforms` 支持三种写法：
  `["Amount", "double_value"]`
  默认按表头 `Amount` 定位
  `{"header": "Amount", "func": "double_value"}`
  显式按表头定位
  `{"column_letter": "C", "func": "double_value"}`
  显式按 Excel 列字母定位
- 为兼容旧字典配置，`{"column": "Amount", "func": "double_value"}` 仍按“表头”解释
- 这样可以避免表头恰好叫 `A`、`B` 时与 Excel 列字母产生歧义
- `pivot_filters`、`pivot_rows`、`pivot_columns`、`pivot_values` 里的字符串现在都默认按表头解释
- 如果这些透视表字段确实需要按 Excel 列字母定位，请显式写成对象，例如：
  `{"column_letter": "A"}`
- `pivot_value_settings` 中推荐写 `{"header": "Amount", ...}`；如果必须按列字母，也可写 `{"column_letter": "C", ...}`
- 当同时提供 `pivot_values` 和 `pivot_value_settings` 时，优先使用 `pivot_value_settings`
- 当前支持的值汇总方式有：`sum`、`count`、`average`/`avg`、`max`、`min`、`product`
- 输出的数据透视表是 Excel 原生数据透视表，不是 pandas 汇总表
