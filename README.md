# 考勤数据处理工具

本项目包含两个 Python 脚本，用于管理和处理 Excel 格式的考勤数据。

## 功能介绍
-   `excel_inspector.py`: 用于检查 Excel 文件的结构和内容，帮助用户快速了解考勤数据的概览。
-   `fill_attendance.py`: 自动化将打卡月报中的考勤状态数据填充到月度考勤表中的指定位置，并进行状态符号转换。

## 安装
1.  **Python 环境**: 确保您的系统已安装 Python 3.x。
2.  **安装依赖**: 使用 pip 安装 `openpyxl` 库：
    ```bash
    pip install openpyxl
    ```

## 代码结构
```
.
├── 2025年5月考勤.xlsx             # 示例目标考勤文件
├── 上下班打卡_月报_20250501-20250526.xlsx # 示例源打卡月报文件
├── excel_inspector.py           # Excel 文件检查脚本
├── fill_attendance.py           # 考勤数据填充脚本
└── README.md                    # 项目说明文件
```

## 关键接口

### `excel_inspector.py`
-   `inspect_excel_file(file_path)`:
    *   **功能**: 检查指定 Excel 文件的所有工作表名称，打印每个工作表的前5行数据。
        *   如果工作表名为 '25年4月考勤（4.1-4.30）'，则打印 B 列（姓名列）的所有非空值。
        *   如果工作表名为 '上下班打卡_月报'，则打印 AV 到 BU 列的唯一值。
    *   **参数**:
        *   `file_path` (str): 要检查的 Excel 文件路径。

### `fill_attendance.py`
-   `fill_attendance_data(target_file, source_file, target_sheet_name, source_sheet_name)`:
    *   **功能**: 从源 Excel 文件中读取考勤数据，并根据姓名匹配，将其填充到目标 Excel 文件的指定工作表中，同时将考勤状态文本转换为预设的符号。
    *   **参数**:
        *   `target_file` (str): 目标 Excel 文件路径（月度考勤表）。
        *   `source_file` (str): 源 Excel 文件路径（打卡月报）。
        *   `target_sheet_name` (str): 目标工作表名称。
        *   `source_sheet_name` (str): 源工作表名称。

## 使用方法

1.  **准备 Excel 文件**: 确保您的目标考勤文件和源打卡月报文件存在于脚本可访问的路径。默认情况下，脚本会查找当前目录下的 `2025年5月考勤.xlsx` 和 `上下班打卡_月报_20250501-20250526.xlsx`。

2.  **运行 `excel_inspector.py`**:
    ```bash
    python excel_inspector.py
    ```
    此命令将检查默认的两个 Excel 文件并打印相关信息。

3.  **运行 `fill_attendance.py`**:
    ```bash
    python fill_attendance.py
    ```
    此命令将从源文件 (`上下班打卡_月报_20250501-20250526.xlsx` 的 '上下班打卡_月报' 工作表) 读取数据，并填充到目标文件 (`2025年5月考勤.xlsx` 的 '25年4月考勤（4.1-4.30）' 工作表) 中。

## 如何扩展

-   **新增考勤状态符号**: 如果有新的考勤状态需要转换，可以在 `fill_attendance.py` 的 `symbol_map` 字典中添加新的键值对。
-   **调整列或工作表**: 如果 Excel 文件的结构发生变化（例如，姓名列或数据列的位置改变，或工作表名称改变），需要相应修改 `excel_inspector.py` 和 `fill_attendance.py` 中硬编码的列索引和工作表名称。
-   **增加新功能**:
    *   可以添加新的函数到现有脚本中，或创建新的 Python 脚本来处理其他考勤相关的任务（例如，生成报告、数据校验等）。
    *   考虑将文件路径和工作表名称作为命令行参数，而不是硬编码在脚本中，以提高灵活性。
