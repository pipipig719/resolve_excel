from __future__ import annotations

import os
import shlex
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "\u996e\u7247\u5bfc\u5165\u5de5\u5177 GUI"

CN_FINAL_NAME = "\u6000\u5b81\u996e\u7247\u8d27\u4f4d\u5bfc\u5165\u6700\u7ec8\u6587\u4ef6.xlsx"
CN_BACKUP_NAME = "\u6000\u5b81\u996e\u7247\u8d27\u4f4d\u5bfc\u5165\u5907\u4efd\u6587\u4ef6.xlsx"

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False


def parse_dnd_files(raw: str, root: tk.Misc) -> list[Path]:
    # Tk DnD payload can be "{C:/a b.xlsx} {C:/c.xlsx}" or plain.
    try:
        parts = root.tk.splitlist(raw)
    except tk.TclError:
        parts = shlex.split(raw)
    files: list[Path] = []
    for p in parts:
        text = p.strip().strip("{}").strip()
        if text:
            files.append(Path(text).expanduser())
    return files


class GuiApp:
    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("860x560")

        self.project_dir = Path(__file__).resolve().parent

        self.inventory_var = tk.StringVar()
        self.source_var = tk.StringVar()
        self.template_var = tk.StringVar(value=self._default_template())
        self.output_dir_var = tk.StringVar(value=str(self.project_dir))
        self.status_var = tk.StringVar(value="Ready")

        self.run_btn: ttk.Button | None = None
        self.widgets_to_disable: list[tk.Widget] = []

        self._build_layout()

    def _default_template(self) -> str:
        for f in self.project_dir.glob("*.xlsx"):
            if "\u6a21\u677f" in f.name:
                return str(f)
        return ""

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text="\u62d6\u62fd\u6216\u9009\u62e9\u6587\u4ef6\u540e\uff0c\u4e00\u952e\u751f\u6210\u6700\u7ec8\u6587\u4ef6\u4e0e\u5907\u4efd\u6587\u4ef6",
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        row = 1
        self._add_file_row(
            parent=container,
            row=row,
            label="\u5e93\u5b58\u6587\u4ef6",
            var=self.inventory_var,
            browse_callback=self._browse_inventory,
            file_types=[("Excel", "*.xlsx"), ("All", "*.*")],
        )
        row += 1

        self._add_file_row(
            parent=container,
            row=row,
            label="\u914d\u65b9\u95f4\u6570\u636e",
            var=self.source_var,
            browse_callback=self._browse_source,
            file_types=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        row += 1

        self._add_file_row(
            parent=container,
            row=row,
            label="\u5bfc\u5165\u6a21\u677f",
            var=self.template_var,
            browse_callback=self._browse_template,
            file_types=[("Excel", "*.xlsx"), ("All", "*.*")],
        )
        row += 1

        self._add_dir_row(
            parent=container,
            row=row,
            label="\u8f93\u51fa\u76ee\u5f55",
            var=self.output_dir_var,
            browse_callback=self._browse_output_dir,
        )
        row += 1

        hint_text = (
            "\u62d6\u62fd\u652f\u6301\u5df2\u542f\u7528 (tkinterdnd2)"
            if HAS_DND
            else "\u5f53\u524d\u672a\u542f\u7528\u62d6\u62fd\uff0c\u53ef\u5148\u7528 Browse\u3002\u5982\u9700\u62d6\u62fd\uff0c\u6267\u884c uv sync \u5b89\u88c5 tkinterdnd2\u3002"
        )
        hint = ttk.Label(container, text=hint_text)
        hint.grid(row=row, column=0, columnspan=3, sticky="w", pady=(4, 8))
        row += 1

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=row, column=0, columnspan=3, sticky="ew")
        self.run_btn = ttk.Button(btn_frame, text="\u5f00\u59cb\u751f\u6210", command=self.on_run_clicked)
        self.run_btn.pack(side=tk.LEFT)
        self.widgets_to_disable.append(self.run_btn)

        open_btn = ttk.Button(
            btn_frame,
            text="\u6253\u5f00\u8f93\u51fa\u76ee\u5f55",
            command=self.open_output_dir,
        )
        open_btn.pack(side=tk.LEFT, padx=(8, 0))
        self.widgets_to_disable.append(open_btn)
        row += 1

        self.log_text = tk.Text(container, height=18, wrap="word")
        self.log_text.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        self.log_text.configure(state=tk.DISABLED)
        container.rowconfigure(row, weight=1)
        row += 1

        status = ttk.Label(container, textvariable=self.status_var)
        status.grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        var: tk.StringVar,
        browse_callback,
        file_types: list[tuple[str, str]],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        self.widgets_to_disable.append(entry)

        btn = ttk.Button(parent, text="Browse", command=lambda: browse_callback(file_types))
        btn.grid(row=row, column=2, sticky="e", pady=4)
        self.widgets_to_disable.append(btn)

        self._register_dnd(entry, var)

    def _add_dir_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        var: tk.StringVar,
        browse_callback,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        self.widgets_to_disable.append(entry)

        btn = ttk.Button(parent, text="Browse", command=browse_callback)
        btn.grid(row=row, column=2, sticky="e", pady=4)
        self.widgets_to_disable.append(btn)

    def _register_dnd(self, widget: tk.Widget, var: tk.StringVar) -> None:
        if not HAS_DND:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: self._on_drop_file(e, var))
        except Exception:
            pass

    def _on_drop_file(self, event, var: tk.StringVar) -> None:
        files = parse_dnd_files(event.data, self.root)
        if files:
            var.set(str(files[0]))

    def _browse_inventory(self, file_types: list[tuple[str, str]]) -> None:
        path = filedialog.askopenfilename(
            title="\u9009\u62e9\u5e93\u5b58\u6587\u4ef6",
            filetypes=file_types,
        )
        if path:
            self.inventory_var.set(path)

    def _browse_source(self, file_types: list[tuple[str, str]]) -> None:
        path = filedialog.askopenfilename(
            title="\u9009\u62e9\u914d\u65b9\u95f4\u6570\u636e\u6587\u4ef6",
            filetypes=file_types,
        )
        if path:
            self.source_var.set(path)

    def _browse_template(self, file_types: list[tuple[str, str]]) -> None:
        path = filedialog.askopenfilename(
            title="\u9009\u62e9\u6a21\u677f\u6587\u4ef6",
            filetypes=file_types,
        )
        if path:
            self.template_var.set(path)

    def _browse_output_dir(self) -> None:
        path = filedialog.askdirectory(title="\u9009\u62e9\u8f93\u51fa\u76ee\u5f55")
        if path:
            self.output_dir_var.set(path)

    def append_log(self, text: str) -> None:
        self.root.after(0, self._append_log_ui, text)

    def _append_log_ui(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def set_status(self, text: str) -> None:
        self.root.after(0, self.status_var.set, text)

    def set_running_state(self, running: bool) -> None:
        state = tk.DISABLED if running else tk.NORMAL
        for w in self.widgets_to_disable:
            try:
                w.configure(state=state)
            except Exception:
                pass

    def open_output_dir(self) -> None:
        output_dir = Path(self.output_dir_var.get()).expanduser()
        if not output_dir.exists():
            messagebox.showwarning("Warning", "Output directory does not exist.")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(str(output_dir))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(output_dir)], check=False)
            else:
                subprocess.run(["xdg-open", str(output_dir)], check=False)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to open directory: {exc}")

    def on_run_clicked(self) -> None:
        inventory = Path(self.inventory_var.get().strip()).expanduser()
        source = Path(self.source_var.get().strip()).expanduser()
        template = Path(self.template_var.get().strip()).expanduser()
        output_dir = Path(self.output_dir_var.get().strip()).expanduser()

        if not inventory.exists():
            messagebox.showerror("Error", "Inventory file is missing.")
            return
        if not source.exists():
            messagebox.showerror("Error", "Source pharmacy file is missing.")
            return
        if not template.exists():
            messagebox.showerror("Error", "Template file is missing.")
            return

        if inventory.suffix.lower() != ".xlsx":
            messagebox.showerror("Error", "Inventory file must be .xlsx")
            return
        if source.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showerror("Error", "Source file must be .xlsx or .xls")
            return
        if template.suffix.lower() != ".xlsx":
            messagebox.showerror("Error", "Template file must be .xlsx")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        backup_dir = output_dir / "source"
        backup_dir.mkdir(parents=True, exist_ok=True)

        final_output = output_dir / CN_FINAL_NAME
        backup_output = backup_dir / CN_BACKUP_NAME

        self.set_running_state(True)
        self.set_status("Running...")
        self.append_log("=" * 60)
        self.append_log(f"Inventory: {inventory}")
        self.append_log(f"Source:    {source}")
        self.append_log(f"Template:  {template}")
        self.append_log(f"Output:    {final_output}")
        self.append_log(f"Backup:    {backup_output}")

        worker = threading.Thread(
            target=self._run_pipeline_thread,
            args=(inventory, source, template, final_output, backup_output),
            daemon=True,
        )
        worker.start()

    def _run_command(self, cmd: list[str]) -> int:
        self.append_log(f"$ {' '.join(shlex.quote(x) for x in cmd)}")
        proc = subprocess.Popen(
            cmd,
            cwd=self.project_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            self.append_log(line.rstrip())
        return proc.wait()

    def _run_pipeline_thread(
        self,
        inventory: Path,
        source: Path,
        template: Path,
        final_output: Path,
        backup_output: Path,
    ) -> None:
        report_dir = (final_output.parent / ".tmp_reports").resolve()
        try:
            cmd1 = [
                sys.executable,
                str(self.project_dir / "convert_inventory.py"),
                "--inventory",
                str(inventory),
                "--template",
                str(template),
                "--source-for-match",
                str(source),
                "--output",
                str(final_output),
                "--report-dir",
                str(report_dir),
            ]
            code1 = self._run_command(cmd1)
            if code1 != 0:
                raise RuntimeError(f"convert_inventory failed with exit code {code1}")

            cmd2 = [
                sys.executable,
                str(self.project_dir / "convert_source_backup.py"),
                "--source",
                str(source),
                "--template",
                str(template),
                "--output",
                str(backup_output),
            ]
            code2 = self._run_command(cmd2)
            if code2 != 0:
                raise RuntimeError(f"convert_source_backup failed with exit code {code2}")

            if report_dir.exists():
                shutil.rmtree(report_dir, ignore_errors=True)

            self.append_log("[OK] Done.")
            self.set_status("Done")
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success",
                    f"Final: {final_output}\nBackup: {backup_output}",
                ),
            )
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"[ERROR] {exc}")
            self.set_status("Failed")
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.root.after(0, lambda: self.set_running_state(False))


def main() -> None:
    if HAS_DND and TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = GuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
