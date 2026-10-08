# 结算审核软件

工程结算审核工作台，包含 Windows 桌面版（C# / Windows Forms）及 Python 本地网页版本。源文件来自上传的 app.zip；两种入口均保留，尚未判定网页版本是否已弃用。

## Windows 桌面版

在 Windows PowerShell 中，从仓库根目录执行（需要 .NET Framework 4.x 的 64 位 C# 编译器）：

```powershell
powershell -ExecutionPolicy Bypass -File .\app\build.ps1
.\结算审核工作台.exe
```

构建脚本使用 `app/templates/结算审核模板.docx` 嵌入模板，并生成根目录 EXE 和配置文件。

运行现有回归测试：

```powershell
.\结算审核工作台.exe --test .\work\tests
```

首次测试请在独立目录中执行，并备份已有业务数据。当前云环境未运行 Windows 编译及桌面测试。

## Python 本地网页版本

建议使用 Python 3.10 或更新版本，在仓库根目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app\server.py
```

打开 http://127.0.0.1:8765 。默认使用随附占位模板，报告输出到仓库的 `outputs/生成报告`。如需配置，复制 `app/config.example.json` 为 `app/config.json`，可添加 `out_dir` 指定输出目录。

## 文件整理

保留全部源码、构建配置、模板及现有测试。未纳入 Python 缓存、用户工程 JSON、上传的 Excel、历史发布包、业务数据和历史测试输出；这些文件应在原电脑另行备份。

`app.zip` 中 `uploads/结算对比表模板_1789302844.xlsx` 与 `uploads/结算对比表模板_1789302916.xlsx` 的 SHA-256 相同。其他同名或同大小文件不据此认定重复。模板包含原有示例内容，未修改。

Python 报告生成回归测试：`python -m unittest discover -s app -p test_report.py`。导入时修复了缺失字段占位符导致报告生成循环的问题；网页/API 冒烟检查和报告导出检查已通过。
