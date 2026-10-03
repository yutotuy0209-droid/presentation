#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Small desktop front-end for auto_deck.py.

The UI is intentionally thin: it collects source/output/preset options and delegates
all generation logic to the deterministic CLI pipeline.
"""
from __future__ import annotations

from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class AutoDeckUi(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Presentation Auto Deck")
        self.geometry("760x520")
        self.minsize(680, 460)

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.preset_var = tk.StringVar(value="research-progress")
        self.aspect_var = tk.StringVar(value="4:3")
        self.footer_var = tk.StringVar()
        self.render_var = tk.BooleanVar(value=False)
        self.strict_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="入力ファイルを選択してください")
        self.messages: queue.Queue[tuple[str, str]] = queue.Queue()

        self._build()
        self.after(80, self._poll)

    def _build(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        root = ttk.Frame(self, padding=18)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(7, weight=1)

        ttk.Label(root, text="PowerPoint Auto Deck", font=("Yu Gothic UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            root,
            text="Markdown / JSON / TXT → レイアウト → lint → PPTX → 任意でPowerPoint実レンダリング",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(root, text="入力").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(root, textvariable=self.source_var).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(root, text="選択", command=self._pick_source).grid(row=2, column=2)

        ttk.Label(root, text="出力").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(root, textvariable=self.output_var).grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(root, text="選択", command=self._pick_output).grid(row=3, column=2)

        options = ttk.Frame(root)
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        ttk.Label(options, text="Preset").pack(side=tk.LEFT)
        ttk.Combobox(
            options,
            textvariable=self.preset_var,
            values=("research-progress", "scientific", "story-driven"),
            width=18,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(6, 18))
        ttk.Label(options, text="比率").pack(side=tk.LEFT)
        ttk.Combobox(
            options, textvariable=self.aspect_var, values=("4:3", "16:9"),
            width=7, state="readonly"
        ).pack(side=tk.LEFT, padx=(6, 18))
        ttk.Checkbutton(options, text="PowerPointでPNG確認", variable=self.render_var).pack(side=tk.LEFT)
        ttk.Checkbutton(options, text="Strict", variable=self.strict_var).pack(side=tk.LEFT, padx=12)

        footer = ttk.Frame(root)
        footer.grid(row=5, column=0, columnspan=3, sticky="ew", pady=5)
        footer.columnconfigure(1, weight=1)
        ttk.Label(footer, text="フッター").grid(row=0, column=0, sticky="w")
        ttk.Entry(footer, textvariable=self.footer_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        action = ttk.Frame(root)
        action.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Label(action, textvariable=self.status_var).pack(side=tk.LEFT)
        self.generate_button = ttk.Button(action, text="生成", command=self._generate)
        self.generate_button.pack(side=tk.RIGHT)

        self.log = tk.Text(root, height=12, wrap="word", font=("Consolas", 9))
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self.log.configure(state=tk.DISABLED)

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[
                ("Presentation input", "*.md *.markdown *.json *.txt"),
                ("All files", "*.*"),
            ]
        )
        if not path:
            return
        self.source_var.set(path)
        if not self.output_var.get():
            source = Path(path)
            self.output_var.set(str(source.with_suffix(".pptx")))

    def _pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")]
        )
        if path:
            self.output_var.set(path)

    def _generate(self) -> None:
        source_text = self.source_var.get().strip()
        output_text = self.output_var.get().strip()
        if not source_text:
            messagebox.showerror("入力", "入力ファイルを指定してください。")
            return
        if not output_text:
            messagebox.showerror("出力", "出力先を指定してください。")
            return
        source = Path(source_text)
        output = Path(output_text)
        if not source.exists():
            messagebox.showerror("入力", "入力ファイルが見つかりません。")
            return
        self.generate_button.configure(state=tk.DISABLED)
        self.status_var.set("生成中")
        self._append("---- generate ----\n")

        def work() -> None:
            script = Path(__file__).resolve().parent / "auto_deck.py"
            cmd = [
                sys.executable, str(script), str(source),
                "--output", str(output),
                "--preset", self.preset_var.get(),
                "--aspect", self.aspect_var.get(),
            ]
            if self.footer_var.get().strip():
                cmd.extend(["--footer", self.footer_var.get().strip()])
            if self.render_var.get():
                cmd.append("--render")
            if not self.strict_var.get():
                cmd.append("--no-strict")

            proc = subprocess.run(cmd, text=True, capture_output=True, check=False)
            if proc.stdout:
                self.messages.put(("log", proc.stdout))
            if proc.stderr:
                self.messages.put(("log", proc.stderr))
            self.messages.put(("done", str(proc.returncode)))

        threading.Thread(target=work, daemon=True).start()

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self._append(payload)
                elif kind == "done":
                    ok = payload == "0"
                    self.status_var.set("完了" if ok else "生成エラー")
                    self.generate_button.configure(state=tk.NORMAL)
                    if ok:
                        messagebox.showinfo("PowerPoint Auto Deck", "生成が完了しました。")
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _append(self, text: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)


if __name__ == "__main__":
    AutoDeckUi().mainloop()
