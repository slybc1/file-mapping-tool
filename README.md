# 文件匹配与分类工具

文件匹配与分类工具提供 Windows 图形界面和 AIX Java 命令行版。

## AIX Java 版（v1.3.0）

AIX 版需要 Java 8 或更高版本。它会递归扫描源目录的所有子目录，按完整文件名精确匹配，并依据 Excel 中的年份、业务类型和编号归档。默认执行预检查，加入 `--apply` 后才移动文件。

- [下载 AIX Java 运行包](https://github.com/slybc1/file-mapping-tool/releases/download/v1.3.0/FileMappingTool-v1.3.0-AIX-Java.zip)
- [AIX 运行说明和命令示例](AIX-Java/README.md)

## Windows 文件包（v1.2.0）

- [Windows Python 运行包](https://github.com/slybc1/file-mapping-tool/releases/download/v1.2.0/FileMappingTool-v1.2.0-Python.zip)解压后先运行 `setup_source.bat`，之后运行 `run_source.bat`。安装、离线部署和故障排查见 [Python 运行说明](SOURCE_EDITION.md)。
- [Windows EXE 运行包](https://github.com/slybc1/file-mapping-tool/releases/download/v1.2.0/FileMappingTool-v1.2.0-Windows.zip)解压后可运行 `FileMappingTool.exe`。

## 程序用途

程序读取 Excel 第一行的列标题，由用户自由选择：

- 哪一列的值用于创建编号文件夹。工作表包含两个分组列和 `CTX_ID` 时，界面会自动选定 `CTX_ID`。
- 哪一列保存需要精确匹配的完整文件名。
- 在哪个文件夹中搜索文件。
- 文件移动到哪个分类结果文件夹。
- 是否同时搜索待匹配文件夹下的所有子文件夹。

程序先执行预检查，不会立即改动文件。用户确认后，才会把匹配文件移动到按 Excel 列值建立的文件夹中。

## 按年份、业务类型和编号分类

工作表同时包含 `FILE_CREATE_DATETIME` 和 `GC_CTX_TYPE_CD` 时，程序自动按每行数据生成 `结果目录/年份/业务类型/编号/文件名`。例如 `2025/PM/10001/a.pdf` 和 `2025/WO/10001/b.pdf`。

- 年份来自 `FILE_CREATE_DATETIME`。支持 Excel 日期单元格和年在前的日期文本，例如 `2025-12-31 23:59:59`。`2026-01-01 00:00:00` 归入 `2026`。
- 业务类型来自 `GC_CTX_TYPE_CD`，例如 `PM`、`WO` 或 `EQ`。编号优先读取 `CTX_ID`，工作表没有该列时使用界面选择的编号列，例如 `WO_NO`。
- 同一年、同一类型、同一编号的多行使用一个文件夹。同一编号出现在另一年或另一类型时，使用对应目录下的独立文件夹。
- 两个分组列都缺失时，按旧表结构直接在结果目录下建立编号文件夹。只有其中一列时，预检查会提示补齐另一列。相关列标题可以有不同大小写或首尾空格，重复标题会导致预检查失败。
- 日期缺失或无法识别、类型或编号缺失或不能用作文件夹名时，该行标记为 `Move failed`，可在结果 Excel 的 `Failure Reason` 栏查看原因。

## EXE 使用方法

1. 解压 Windows 压缩包，双击 `FileMappingTool.exe`。
2. 点击“Excel 原始数据”右侧的“选择”，选择 `.xlsx` 文件。
3. 选择工作表。
4. 核对“创建文件夹依据列”。界面自动选定 `CTX_ID` 时可直接使用，其他工作表可手动选择。
5. 在“文件名所在列”中选择保存完整文件名的列。
6. 选择“待匹配文件夹”。不限定盘符或文件夹名称。
7. 如果文件散落在下级目录，勾选“包含待匹配文件夹下的所有子文件夹”。默认不勾选，扫描更快。
8. 选择“分类结果文件夹”和“结果 Excel”保存位置。
9. 点击“1. 开始预检查”。核对匹配、未找到、冲突和无效行数量。
10. 确认无误后，点击“2. 确认并移动”。没有可移动文件时，点击“2. 生成结果”可查看各行状态。

## 处理规则

- 使用完整文件名精确匹配，包含扩展名。
- 文件名和扩展名严格区分大小写。`Report.pdf` 与 `report.PDF` 分别匹配。
- 源目录中存在多个同名文件时，会全部移动；目标重名自动追加 `_2`、`_3`，不会覆盖。
- 同一源文件对应多个目标目录时，每个目录保存独立副本，全部保存后删除源文件。失败或取消时保留源文件。该情况会增加存储占用。
- 同一目标目录的重复 Excel 行共享该目录内的处理结果。多个来源的同名文件全部保留并自动编号。目标中已有的同名文件也完整保留。
- 同一年内仅大小写不同的类型，或同一年同一类型内仅大小写不同的编号，可能指向同一个 Windows 文件夹。这些行会标记为 `Mapping conflict`。
- 未找到文件时标记为 `Not found`。
- 数据无效、权限不足、文件被占用或移动失败时标记为 `Move failed`，原因写入新增列。
- 原始 Excel 不会被覆盖。结果 Excel 追加 `Mapping Status` 和 `Failure Reason` 两列。
- 同时生成 UTF-8 CSV 移动日志。

## 大批量性能设计

- Excel 使用只读流式方式分批读取。
- 待匹配文件名和运行状态保存在临时 SQLite 数据库中，避免一次占用大量内存。
- 源文件夹只扫描一次，不读取文件内容。
- 匹配名称数量不超过一百万时使用内存集合加速查找。
- 扫描和移动在后台线程执行，界面保持响应，并支持暂停和取消。
- 源文件和目标文件夹位于同一盘符时，移动通常只更新目录信息，速度明显快于跨盘移动。

## 从源码运行

需要 64 位 Python 3.11 或更高版本。

如果电脑阻止未签名 EXE，可以使用源码运行版，详见 [SOURCE_EDITION.md](SOURCE_EDITION.md)。首次双击安装依赖：

```text
setup_source.bat
```

安装完成后，每次双击：

```text
run_source.bat
```

也可以在本文件夹打开命令提示符后运行：

```text
py -3 source_launcher.py setup
.venv\Scripts\python.exe source_launcher.py run
```

## 重新打包 EXE

双击：

```text
build_exe.bat
```

完成后，新程序位于：

```text
dist\FileMappingTool.exe
```

打包后的 EXE 不要求目标电脑安装 Python，但目前没有代码签名证书。如果公司策略阻止运行，请由 IT 审核，或使用公司允许的 Python 环境运行源码版。源码版仍需遵守设备的软件运行策略。

## 运行核心测试

安装依赖后，在本文件夹运行：

```text
.venv\Scripts\python.exe test_mapping_core.py
```

测试会在系统临时目录创建数据，验证年份与类型分类、自定义列、子文件夹扫描、文件移动和结果导出，然后自动清理临时数据。

## 源代码结构

- `file_mapping_tool.py`：PySide6 图形界面、进度显示和后台线程。
- `mapping_core.py`：Excel 流式读取、文件扫描、SQLite 匹配、移动、日志和结果 Excel。
- `requirements.txt`：Python 依赖版本范围。
- `FileMappingTool.spec`：PyInstaller 打包配置。
- `requirements-runtime.txt`：源码运行版的固定运行依赖，不含打包工具。
- `setup_source.bat`：首次创建独立 Python 环境并安装运行依赖。
- `run_source.bat`：使用已安装环境启动源码，失败时保留提示。
- `source_launcher.py`：环境安装、离线依赖支持、启动和诊断日志。
- `SOURCE_EDITION.md`：源码运行版使用和 IT 部署说明。
- `build_exe.bat`：安装依赖并生成独立 EXE。
- `test_mapping_core.py`：核心流程自动化测试。

## 注意事项

- 当前仅支持 `.xlsx`。
- 第一行必须是标题行。
- “创建文件夹依据列”和“文件名所在列”不能选择同一列。
- 分类结果文件夹不能与待匹配文件夹相同，也不能放在待匹配文件夹内部。
- 一个目标目录时执行移动。多个目标目录时分发独立副本后清理源文件。失败或取消后请核对日志及目标文件再重试，重试会编号保留已有文件。
