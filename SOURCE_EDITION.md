# Python 源码运行版（v1.1.0）

适用于可使用 Python、但未签名的 FileMappingTool.exe 被阻止的 Windows 电脑。
本版本复用现有界面和处理核心，文件在用户电脑本地处理。
Python 和依赖仍须符合单位的软件使用策略；若也被拦截，请由 IT 批准安装。

## 首次使用

1. 安装单位批准的 **64 位 Python**（推荐 3.12；最低 3.11）。已有 Python 可继续使用。
2. 将整个 ZIP **全部解压**到有写入权限的本地文件夹，不要在压缩包内启动。
3. 双击 `setup_source.bat`。第一次需要联网下载运行依赖，程序会在当前目录创建 `.venv`。
4. 看到 `Setup complete` 后，双击 `run_source.bat`，即可打开原有工具窗口。

之后每次只需双击 `run_source.bat`，无需再次安装依赖，处理文件无需联网。
请在固定位置安装；若移动整个文件夹或更换 Python，请重新解压并安装到新位置。
本包不包含 EXE，不安装 PyInstaller；也不会把依赖安装到全局 Python 中。

## 使用流程

选择 Excel → 选择工作表 → 选择 WO 列和文件名列 → 选择源目录和目标目录 → 预检查 → 确认移动。
完整说明见 README.md。沿用 v1.1.0 的严格大小写匹配和跨 WO 同名分发规则。
同一个源附件对应多个 WO 时，各 WO 保存独立副本，会占用额外磁盘空间；全部成功才删除源附件。
预检查不移动文件。请先用少量测试文件确认使用流程。

## 如果不允许双击 BAT，但允许 Python

在解压后的文件夹打开命令提示符，首次运行：

```text
py -3.12 source_launcher.py setup
```

日常启动：

```text
.venv\Scripts\python.exe source_launcher.py run
```

如果没有 `py` 命令，可以使用安装好的 Python 的完整路径，例如：

```text
"C:\Path\To\Python\python.exe" source_launcher.py setup
```

## 无外网电脑（由 IT 准备离线依赖）

在相同 Windows 架构及相同 Python 版本的联网电脑上，在解压目录运行：

```text
py -3.12 -m pip download --only-binary=:all: -r requirements-runtime.txt -d wheels
```

将 `wheels` 文件夹随源代码复制到用户电脑，运行：

```text
py -3.12 source_launcher.py setup --offline wheels
```

离线安装只读取本地 wheels，不访问包索引。

## 排查问题

- 未找到 Python：安装或选择单位批准的 Python，确认是 64 位。
- 安装依赖失败：检查网络、公司包镜像或使用上面的离线依赖流程。
- 启动失败：窗口会保留错误提示，详细记录在同目录 `source-launch.log`。
- 目录无法写入：将完整源码解压到用户有写权限的文件夹。
- 公司阻止 Python、脚本或依赖 DLL：请 IT 审核，本方案不会修改系统安全设置。

只检查运行环境和界面能否启动（不移动用户文件）：

```text
.venv\Scripts\python.exe source_launcher.py check
```

`source-launch.log` 可能含本机路径、依赖安装错误；向他人提供前请检查内容。
