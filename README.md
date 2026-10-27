# 文件匹配与分类工具

本文件夹保存 `FileMappingTool.exe` 对应的 Python 源代码、运行依赖与打包脚本。

## 下载 v1.1.0

- [Python 运行版 ZIP](https://github.com/slybc1/file-mapping-tool/releases/download/v1.1.0/FileMappingTool-v1.1.0-Python.zip)：适合可以使用 Python 的电脑，解压后首次运行 `setup_source.bat`，之后运行 `run_source.bat`。
- [Windows EXE 压缩包](https://github.com/slybc1/file-mapping-tool/releases/download/v1.1.0/FileMappingTool-v1.1.0-Windows.zip)：解压后运行 `FileMappingTool.exe`。
- [完整发布说明](https://github.com/slybc1/file-mapping-tool/releases/tag/v1.1.0)。Python 运行版的安装、离线部署和故障排查见 [SOURCE_EDITION.md](SOURCE_EDITION.md)。

## 程序用途

程序读取 Excel 第一行的列标题，由用户自由选择：

- 哪一列的值用于创建分类文件夹。
- 哪一列保存需要精确匹配的完整文件名。
- 在哪个文件夹中搜索文件。
- 文件移动到哪个分类结果文件夹。
- 是否同时搜索待匹配文件夹下的所有子文件夹。

程序先执行预检查，不会立即改动文件。用户确认后，才会把匹配文件移动到按 Excel 列值建立的文件夹中。

## EXE 使用方法

1. 双击上一级目录中的 `FileMappingTool.exe`。
2. 点击“Excel 原始数据”右侧的“选择”，选择 `.xlsx` 文件。
3. 选择工作表。
4. 在“创建文件夹依据列”中选择用于建立文件夹的列。
5. 在“文件名所在列”中选择保存完整文件名的列。
6. 选择“待匹配文件夹”。不限定盘符或文件夹名称。
7. 如果文件散落在下级目录，勾选“包含待匹配文件夹下的所有子文件夹”。默认不勾选，扫描更快。
8. 选择“分类结果文件夹”和“结果 Excel”保存位置。
9. 点击“1. 开始预检查”。核对匹配、未找到、冲突和无效行数量。
10. 确认无误后，点击“2. 确认并移动”。

## 处理规则

- 使用完整文件名精确匹配，包含扩展名。
- v1.1.0：文件名严格区分大小写，Report.pdf 与 report.PDF 不会互相匹配。
- Excel 中同一个分类编号可以出现多行，多个文件会进入同一个文件夹。
- 源目录中存在多个同名文件时，会全部移动；目标重名自动追加 `_2`、`_3`，不会覆盖。
- 同一完整文件名对应多个 WO 时，每个 WO 都获得独立副本；全部成功后删除源文件。失败或取消时保留源文件。该情况会增加存储占用。
- 同一个 WO 的重复 Excel 行共享该 WO 内的处理结果；多个来源的同名文件全部保留并自动编号。已有目标文件（包括仅大小写不同的名称）不覆盖。
- 仅大小写不同的 WO 编号在 Windows 下可能指向同一目录，预检查标记为 `Mapping conflict`，避免归属混淆。
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

测试会在系统临时目录创建数据，验证自定义列、子文件夹扫描、移动文件、结果 Excel 和日志，然后自动清理临时数据。

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
- 单 WO 执行移动；多 WO 分发独立副本后清理源文件。正式运行前先查看预检查结果，重要文件建议另行备份。失败或取消后请核对日志及目标文件再重试，重试会编号保留已有文件。
