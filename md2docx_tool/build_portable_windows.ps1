$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

if (!(Test-Path ".\vendor\pandoc\pandoc.exe")) {
  throw "缺少 vendor\pandoc\pandoc.exe。请先把 pandoc.exe 放到该目录。"
}

pyinstaller --noconfirm --clean --onedir --windowed `
  --name md2docx_gui `
  --add-data "vendor\pandoc\pandoc.exe;vendor\pandoc" `
  md2docx_gui.py

pyinstaller --noconfirm --clean --onedir `
  --name md2docx_cli `
  --add-data "vendor\pandoc\pandoc.exe;vendor\pandoc" `
  md2docx_cli.py

Write-Host ""
Write-Host "输出目录："
Write-Host "  dist\md2docx_gui\md2docx_gui.exe"
Write-Host "  dist\md2docx_cli\md2docx_cli.exe"
