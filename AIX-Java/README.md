# 文件匹配与分类工具 · AIX Java 版 v1.3.0

在 AIX 服务器的终端运行。部署包中的 `file-mapping-aix.jar` 已包含读取和写入 Excel 所需的 Apache POI 依赖。

## 部署和检查

将 ZIP 解压后，把 `file-mapping-aix.jar` 放到 AIX 服务器的应用目录。在服务器上运行以下命令，记录 AIX 版本并确认实际调用的 Java 版本。

```sh
oslevel -s
java -version
java -jar file-mapping-aix.jar --help
```

需要 Java 8 或更高版本。若服务器安装了多个版本，可将命令中的 `java` 替换为所选 Java 的绝对路径。随包的 `run-file-mapping.sh` 也可启动同目录的 JAR：

```sh
sh run-file-mapping.sh --help
```

脚本优先使用 `$JAVA_HOME/bin/java`，否则使用 `PATH` 中的 `java`。服务器应能读取输入工作簿和源目录。正式执行还需要在目标目录和报表目录写入，以及从源目录删除已完成归档的文件。

## 预检查和正式执行

先用一份小样表和可替换的测试文件跑通流程。省略 `--apply` 即可进入只读预检查。程序扫描源目录，在终端列出预检查结果。

```sh
java -jar file-mapping-aix.jar \
  --excel /data/mappings/map.xlsx \
  --source /data/attachments/in \
  --target /data/attachments/out
```

确认预检查结果后，加入 `--apply`。这次运行会重新预检查，再执行文件归档并生成结果表和日志。

```sh
java -jar file-mapping-aix.jar \
  --excel /data/mappings/map.xlsx \
  --source /data/attachments/in \
  --target /data/attachments/out \
  --result /data/reports/mapping-result.xlsx \
  --apply
```

目标目录应独立于源目录。每次运行使用新的结果文件名，把结果表和日志放在源目录之外。同一源目录同一时间只运行一个归档任务。

## Excel 和归档规则

- 输入为 `.xlsx`，第一行为表头。默认使用第一张工作表，可用 `--sheet "Data"` 指定。
- 文件名列支持常见表头，例如 `File Name`、`FILE_NAME`、`FILENAME`、`附件名称`。如需手动指定，使用 `--file-column "附件文件名"` 或 `--file-column C`。
- `FILE_CREATE_DATETIME` 与 `GC_CTX_TYPE_CD` 同时存在时，程序自动按 `年份/业务类型/编号` 归档。日期可使用 Excel 日期或明确的年月日顺序文本，例如 `2025-12-31 23:59:59`。类型取自表格，支持 `PM`、`WO`、`EQ` 等值。
- 分组表优先使用 `CTX_ID`。缺少该列时，可用 `--id-column WO_NO` 指定编号列，或让程序识别常见表头，例如 `WO_NO`、`Work Order Number` 或 `编号`。两个分类列都缺少时采用旧表的单层编号目录。只缺一个分类列时，预检查会提示缺少的表头。
- 同一年、同一类型、同一编号只建一个文件夹。该编号的不同附件归入同一目录。在其他年份或业务类型出现的相同编号会各自归档。
- 默认扫描源目录及所有子目录，跳过符号链接。按完整文件名和扩展名匹配，严格区分大小写。不同位置同名的源文件都会保留，目标重名按 `_2`、`_3` 编号。
- 同一附件名要归入多个目录时，每个目录获得独立副本。全部分发成功后，程序才移除源附件。

```text
/data/attachments/out/
└── 2025/
    ├── PM/
    │   └── 1001/Plan.pdf
    └── WO/
        └── 1001/Report.pdf
```

## 结果和退出码

`--apply` 会生成结果 Excel 与同目录的 `<结果文件名>_log.csv`。省略 `--result` 时，结果写到输入 Excel 旁。如果输入 Excel 位于源目录内，结果会写到目标目录。默认文件名是 `file-mapping-result-YYYYMMDD-HHMMSS-SSS.xlsx`。结果表保留工作表和单元格显示值，并在匹配工作表追加状态和原因。格式和公式保留在原始工作簿中。

| 退出码 | 含义 |
| --- | --- |
| `0` | 执行成功。未找到的附件标记为 `Not found`。 |
| `1` | 程序出错，或正式执行中存在 `Move failed`、`Mapping conflict` 的行。请检查终端输出和已生成的日志。 |
| `2` | 命令参数无效。可运行 `--help` 查看用法。 |

该版本以 Java 8 为兼容基线。正式归档前，请在实际 AIX 服务器上用测试目录验证预检查、归档、结果表和日志。

## 从源码构建

在装有 JDK 和 Maven 的开发机上，从仓库根目录运行 `mvn -f AIX-Java/pom.xml clean package`。构建产物为 `AIX-Java/target/file-mapping-aix.jar`。
