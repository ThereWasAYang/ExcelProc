# ExcelProc

`ExcelProc` 用于处理 `csv` 或 `xlsx` 文件，支持按指定函数插入新列，并在输出工作簿中创建 Excel 原生数据透视表。

## 功能说明

1. 读取 `csv` 或 `xlsx` 文件。
2. 根据配置中的 `transforms` 对指定列逐行执行数据处理函数，并在源列右侧插入结果列。
3. 新插入列可以在配置中指定列标题和小数位数。
4. 无论输入文件是 `csv` 还是 `xlsx`，输出都统一保存为 `xlsx`。
5. 在输出文件中新建一个 Excel 原生数据透视表工作表。
6. 数据透视表值字段可以配置汇总方式、显示名称和小数位数。

## 目录结构

- `inputs/`：输入样本和测试数据
- `configs/`：示例配置和测试配置
- `outputs/`：输出文件
- `scripts/`：辅助脚本
- `processors/`：数据处理函数
- `excel_processor.py`：主处理脚本
- `setup_env.bat`：Windows 下一键配置环境脚本

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

## 依赖安装

如果不使用一键脚本，也可以自行安装依赖：

```powershell
python -m pip install pandas openpyxl pywin32
```

数据透视表使用的是 Excel 自带的“插入 -> 数据透视表”能力，因此需要满足以下条件：

- 当前 Python 环境已安装 `pywin32`
- 本机已安装 Microsoft Excel
- 当前环境允许通过 Excel COM 启动 Excel

如果以上条件不满足，脚本会直接报错，不会退化成普通汇总表。

## 配置文件

配置文件放在 `configs/` 目录。虽然文件扩展名仍是 `.json`，但脚本支持在配置文件中写 `//` 单行注释和 `/* ... */` 多行注释，便于说明每个参数的含义。

示例配置：

[configs/sample_config.json](E:/Work/Pycharm/ExcelProc/configs/sample_config.json)

测试配置：

[configs/test_config.json](E:/Work/Pycharm/ExcelProc/configs/test_config.json)

## transforms 配置

`transforms` 用于描述“从哪一列取值、调用哪个函数、把结果列插入到哪里”。

推荐写法：

```json
{
  "transforms": [
    {
      "header": "Amount",
      "func": "double_value",
      "title": "Amount Double",
      "decimals": 0
    }
  ]
}
```

字段说明：

- `header`：按表头名称定位源列
- `column_letter`：按 Excel 列字母定位源列，例如 `A`、`B`、`AA`
- `func`：数据处理函数名，必须已经注册到 `FUNCTION_REGISTRY`
- `title`：新插入列的列标题；不写时默认使用 `<源列标题>_<函数名>`
- `decimals`：新插入列的小数位数；为 `0` 时结果会自动转为整数

按 Excel 列字母定位的写法：

```json
{
  "transforms": [
    {
      "column_letter": "D",
      "func": "double_value",
      "title": "Amount Double",
      "decimals": 0
    }
  ]
}
```

## 数据处理函数

数据处理函数统一放在 [processors/transform_functions.py](E:/Work/Pycharm/ExcelProc/processors/transform_functions.py) 中管理。

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
    {
      "header": "Name",
      "func": "strip_text",
      "title": "Clean Name"
    }
  ]
}
```

编写函数时建议：

- 优先处理空值，例如使用 `pd.isna(value)`
- 函数应只处理单个单元格值，不要直接依赖整列 DataFrame
- 返回值应是可写入 Excel 的普通值类型，如字符串、数字、布尔值或空值
- 如果函数只适用于特定格式，建议在函数内主动校验并抛出明确错误

## 数据透视表配置

透视表字段默认按表头名称定位：

```json
{
  "pivot_filters": ["Channel", "Priority"],
  "pivot_rows": ["Region", "Category"],
  "pivot_columns": ["Segment", "Quarter"],
  "pivot_value_settings": [
    {
      "header": "Amount",
      "summary": "sum",
      "name": "Total Amount",
      "decimals": 0
    },
    {
      "header": "Score",
      "summary": "average",
      "name": "Average Score",
      "decimals": 1
    }
  ]
}
```

透视表字段说明：

- `pivot_filters`：数据透视表“筛选”区域字段
- `pivot_rows`：数据透视表“行”区域字段
- `pivot_columns`：数据透视表“列”区域字段
- `pivot_values`：数据透视表“值”区域字段，简单场景可用
- `pivot_value_settings`：数据透视表“值字段设置”，推荐使用

`pivot_value_settings` 字段说明：

- `header`：按表头名称定位值字段
- `column_letter`：按 Excel 列字母定位值字段
- `summary`：汇总方式，支持 `sum`、`count`、`average`/`avg`、`max`、`min`、`product`
- `name`：数据透视表中显示的值字段名称
- `number_format`：Excel 数字格式，例如 `#,##0.00`
- `decimals`：值字段显示的小数位数；为 `0` 时显示为整数

如果同时设置了 `decimals` 和 `number_format`，优先使用 `decimals` 自动生成格式。

按 Excel 列字母定位的示例：

```json
{
  "pivot_filters": [
    {"column_letter": "G"}
  ],
  "pivot_value_settings": [
    {
      "column_letter": "E",
      "summary": "sum",
      "name": "Total Amount",
      "decimals": 0
    }
  ]
}
```

## 小数位控制

项目支持两类小数位控制：

1. `transforms[].decimals`：控制新插入列的数据小数位
2. `pivot_value_settings[].decimals`：控制数据透视表值字段显示的小数位

当 `decimals` 设置为 `0` 时：

- 新插入列的结果会转为整数
- 数据透视表值字段会使用整数格式显示

## 使用方式

### 使用配置文件运行

```powershell
python .\excel_processor.py --config .\configs\sample_config.json
```

### 直接传参运行

```powershell
python .\excel_processor.py `
  --input .\inputs\test_input_100rows.csv `
  --suffix DEMO `
  --transforms "[{\"header\": \"Time\", \"func\": \"time_to_seconds\", \"title\": \"Time Seconds\", \"decimals\": 0}, {\"header\": \"Amount\", \"func\": \"double_value\", \"title\": \"Amount Double\", \"decimals\": 0}]" `
  --pivot-filters "[\"Channel\", \"Priority\"]" `
  --pivot-rows "[\"Region\", \"Category\"]" `
  --pivot-columns "[\"Segment\", \"Quarter\"]" `
  --pivot-value-settings "[{\"header\": \"Amount\", \"summary\": \"sum\", \"name\": \"Total Amount\", \"decimals\": 0}, {\"header\": \"Score\", \"summary\": \"average\", \"name\": \"Average Score\", \"decimals\": 1}]"
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
