from __future__ import annotations

import argparse
import sys
from pathlib import Path

from md2docx_core import ConvertError, convert_markdown_to_docx


def main() -> int:
    p = argparse.ArgumentParser(prog="md2docx")
    p.add_argument("input", help="Markdown 文件路径（.md）")
    p.add_argument("-o", "--output", help="输出 docx 路径；默认与输入同名", default=None)
    p.add_argument("--reference-docx", help="可选：Word模板 reference.docx", default=None)

    args = p.parse_args()

    in_path = Path(args.input).resolve()
    out_path = Path(args.output).resolve() if args.output else in_path.with_suffix(".docx")
    ref = Path(args.reference_docx).resolve() if args.reference_docx else None

    try:
        convert_markdown_to_docx(markdown_path=in_path, output_docx=out_path, reference_docx=ref)
        print(str(out_path))
        return 0
    except (FileNotFoundError, ConvertError) as e:
        print(str(e), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
