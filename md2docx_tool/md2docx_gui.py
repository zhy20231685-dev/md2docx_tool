from __future__ import annotations

import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from md2docx_core import ConvertError, convert_markdown_to_docx

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Markdown 转 Word（DOCX）")

        self.input_path: Path | None = None
        self.reference_docx: Path | None = None

        self.path_var = tk.StringVar(value="未选择文件（可粘贴Markdown或选择.md）")
        self.ref_var = tk.StringVar(value="未选择模板（可选）")
        self.status_var = tk.StringVar(value="就绪")

        top = tk.Frame(root)
        top.pack(fill="x", padx=10, pady=10)

        tk.Label(top, textvariable=self.path_var, anchor="w").pack(fill="x")

        btns = tk.Frame(top)
        btns.pack(fill="x", pady=(8, 0))

        tk.Button(btns, text="选择.md文件", command=self.pick_md).pack(side="left")
        tk.Button(btns, text="选择模板reference.docx", command=self.pick_ref).pack(side="left", padx=(8, 0))
        tk.Button(btns, text="导出DOCX", command=self.export).pack(side="right")

        tk.Label(root, textvariable=self.ref_var, anchor="w").pack(fill="x", padx=10)

        mid = tk.Frame(root)
        mid.pack(fill="both", expand=True, padx=10, pady=10)

        self.text = tk.Text(mid, wrap="word")
        self.text.pack(fill="both", expand=True)

        bottom = tk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        tk.Label(bottom, textvariable=self.status_var, anchor="w").pack(fill="x")

        if TkinterDnD is not None and DND_FILES is not None and hasattr(root, "drop_target_register"):
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<Drop>>", self.on_drop)

    def pick_md(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Markdown", "*.md"), ("All files", "*.*")])
        if not p:
            return
        self.input_path = Path(p)
        self.path_var.set(str(self.input_path))
        self.status_var.set("已选择文件")

    def pick_ref(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Word docx", "*.docx"), ("All files", "*.*")])
        if not p:
            return
        self.reference_docx = Path(p)
        self.ref_var.set(f"模板：{self.reference_docx}")

    def on_drop(self, event) -> None:
        raw = (event.data or "").strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        p = raw.split()[0]
        if p.lower().endswith(".md"):
            self.input_path = Path(p)
            self.path_var.set(str(self.input_path))
            self.status_var.set("已拖入文件")
        else:
            messagebox.showwarning("提示", "请拖入 .md 文件")

    def export(self) -> None:
        try:
            out = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word docx", "*.docx")])
            if not out:
                return
            out_path = Path(out)

            self.status_var.set("转换中...")
            self.root.update_idletasks()

            if self.input_path is not None:
                convert_markdown_to_docx(
                    markdown_path=self.input_path,
                    output_docx=out_path,
                    reference_docx=self.reference_docx,
                )
            else:
                md = self.text.get("1.0", "end").strip()
                if not md:
                    messagebox.showwarning("提示", "没有内容：请选择.md文件或在文本框粘贴Markdown")
                    self.status_var.set("就绪")
                    return
                convert_markdown_to_docx(
                    markdown_text=md,
                    output_docx=out_path,
                    reference_docx=self.reference_docx,
                )

            self.status_var.set("完成")
            messagebox.showinfo("完成", f"已导出：{out_path}")
        except (ConvertError, FileNotFoundError) as e:
            self.status_var.set("失败")
            messagebox.showerror("转换失败", str(e))
        except Exception:
            self.status_var.set("失败")
            messagebox.showerror("异常", traceback.format_exc())


def main() -> None:
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.geometry("900x600")
    root.mainloop()


if __name__ == "__main__":
    main()

