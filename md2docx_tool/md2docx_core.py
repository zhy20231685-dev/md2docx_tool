from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional


class ConvertError(RuntimeError):
    pass


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def find_pandoc() -> Path:
    env = os.environ.get("PANDOC_PATH")
    if env:
        p = Path(env)
        if p.exists():
            return p

    candidates = [
        _app_base_dir() / "vendor" / "pandoc" / "pandoc.exe",
        _app_base_dir() / "vendor" / "pandoc" / "pandoc",
        Path.cwd() / "vendor" / "pandoc" / "pandoc.exe",
        Path.cwd() / "vendor" / "pandoc" / "pandoc",
    ]
    for c in candidates:
        if c.exists():
            return c

    raise FileNotFoundError(
        "未找到 pandoc 可执行文件。请把 pandoc.exe 放到 vendor/pandoc/ 下，或设置环境变量 PANDOC_PATH 指向 pandoc。"
    )


def convert_markdown_to_docx(
    *,
    markdown_path: Optional[Path] = None,
    markdown_text: Optional[str] = None,
    output_docx: Path,
    reference_docx: Optional[Path] = None,
    working_dir: Optional[Path] = None,
) -> None:
    if (markdown_path is None) == (markdown_text is None):
        raise ValueError("markdown_path 和 markdown_text 必须且只能提供一个")

    pandoc = find_pandoc()
    output_docx = Path(output_docx)
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    tmp_file: Optional[Path] = None
    try:
        if markdown_text is not None:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w", encoding="utf-8")
            tmp.write(markdown_text)
            tmp.close()
            tmp_file = Path(tmp.name)
            input_path = tmp_file
            cwd = working_dir or Path.cwd()
        else:
            input_path = Path(markdown_path).resolve()
            cwd = working_dir or input_path.parent

        args = [
            str(pandoc),
            str(input_path),
            "-f",
            "markdown",
            "-t",
            "docx",
            "-o",
            str(output_docx),
        ]

        if reference_docx:
            args += ["--reference-doc", str(Path(reference_docx).resolve())]

        proc = subprocess.run(
            args,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            shell=False,
        )
        if proc.returncode != 0:
            msg = (proc.stderr or proc.stdout or "").strip()
            raise ConvertError(msg or f"pandoc 执行失败，退出码 {proc.returncode}")
    finally:
        if tmp_file and tmp_file.exists():
            try:
                tmp_file.unlink()
            except OSError:
                pass

