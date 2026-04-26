# Markdown 转 Word（DOCX）便携工具（离线）

这个项目提供：

- GUI：选择 `.md` 文件 / 粘贴 Markdown / 拖拽文件 / 一键导出 `.docx`
- CLI：命令行批量转换
- 转换引擎：Pandoc（本地可执行文件，离线）

## 1. 准备 pandoc（Windows）

1. 下载 pandoc 的 Windows 安装包或 zip 包
2. 提取 `pandoc.exe`
3. 放到 `vendor\pandoc\pandoc.exe`

## 2. 直接运行（不打包）

安装依赖（仅需要一次）：

```powershell
python -m pip install -r requirements.txt
```

运行 GUI：

```powershell
python .\md2docx_gui.py
```

运行 CLI：

```powershell
python .\md2docx_cli.py .\demo.md
python .\md2docx_cli.py .\demo.md -o .\out.docx
```

## 3. 打包成便携版（文件夹版）

### 方式 A：在你自己的 Windows 电脑本地打包

确保 `vendor\pandoc\pandoc.exe` 已存在，然后执行：

```powershell
.\build_portable_windows.ps1
```

产物在：

- `dist\md2docx_gui\md2docx_gui.exe`
- `dist\md2docx_cli\md2docx_cli.exe`

把整个 `dist\md2docx_gui\` 文件夹拷贝到任意 Windows 电脑即可离线运行。

### 方式 B：不装 Python 也能“直接下载 exe”（推荐）

把本项目上传到 GitHub 后，用 GitHub Actions 在云端 Windows 机器自动打包，你只需要下载产物即可：

1. 在 GitHub 新建一个仓库，把本项目文件上传（包含 `.github/workflows/build-windows.yml`）
2. 打开仓库页面 → Actions → `Build Windows Portable` → Run workflow
3. 等待执行完成后，在该次运行页面下载 artifacts：
   - `md2docx_gui_portable`（里面就是 `md2docx_gui.exe` + 依赖）
   - `md2docx_cli_portable`

该工作流会自动下载 pandoc 并打进产物里；下载后的文件夹可离线运行。

## 4. 常见问题

- 提示找不到 pandoc：确认 `vendor\pandoc\pandoc.exe` 在同级目录结构下，或设置环境变量 `PANDOC_PATH`
- 拖拽不可用：确保已安装 `tkinterdnd2`，或先用“选择.md文件”功能
