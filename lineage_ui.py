"""
lineage_ui.py
=============
Graphical front-end for data_lineage.py — lets non-technical users build
client-formatted lineage workbooks from Atlan impact reports.

Launch with:  python lineage_ui.py        (or: python data_lineage.py --gui)

No dependencies beyond the engine's (pandas + openpyxl); the UI itself is
pure standard-library Tkinter. Settings persist to lineage_settings.json
next to this file, and the same file works as --config for the CLI.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import data_lineage as engine
except ImportError as exc:
    # Launched (e.g. by pythonw) on an interpreter missing pandas/openpyxl —
    # there's no console to show the traceback, so surface it as a dialog.
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Missing Python library",
        f"The Data Lineage Builder could not start because a required "
        f"library is missing:\n\n    {exc}\n\n"
        f"Install the dependencies with:\n\n    pip install pandas openpyxl")
    raise SystemExit(1)

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'lineage_settings.json')

ACCENT = '#7030A0'        # client purple
ACCENT_LIGHT = '#F2CEEF'  # client pink


class QueueLogHandler(logging.Handler):
    """Routes engine log records into the UI thread via a queue."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class LineageApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Data Lineage Builder")
        root.geometry("980x720")
        root.minsize(840, 600)
        self._set_window_icon()

        self.log_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.written_files: list[str] = []

        self._build_styles()
        self._build_layout()
        self._load_settings()
        self._poll_log_queue()

    def _set_window_icon(self):
        """Use app.ico for the title-bar icon (matches any packaged .exe icon).
        Resolves from a PyInstaller bundle (sys._MEIPASS) or a plain checkout."""
        try:
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            ico = os.path.join(base, 'app.ico')
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
        except Exception:
            pass

    # ------------------------------------------------------------------ UI --
    def _build_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use('vista')
        except tk.TclError:
            pass
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'),
                        foreground=ACCENT)
        style.configure('Sub.TLabel', font=('Segoe UI', 9), foreground='#666666')
        style.configure('Run.TButton', font=('Segoe UI', 11, 'bold'))
        style.configure('TNotebook.Tab', font=('Segoe UI', 10), padding=(14, 6))

    def _build_layout(self):
        header = ttk.Frame(self.root, padding=(16, 12, 16, 4))
        header.pack(fill='x')
        ttk.Label(header, text="Data Lineage Builder", style='Title.TLabel').pack(anchor='w')
        ttk.Label(header,
                  text="Turn Atlan impact reports into client-formatted lineage workbooks",
                  style='Sub.TLabel').pack(anchor='w')

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=12, pady=8)
        self._build_run_tab()
        self._build_rules_tab()
        self._build_format_tab()

        # status bar
        bar = ttk.Frame(self.root, padding=(16, 4, 16, 8))
        bar.pack(fill='x')
        self.progress = ttk.Progressbar(bar, mode='indeterminate', length=180)
        self.progress.pack(side='left')
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(bar, textvariable=self.status_var).pack(side='left', padx=10)
        self.open_folder_btn = ttk.Button(bar, text="Open output folder",
                                          command=self._open_folder, state='disabled')
        self.open_folder_btn.pack(side='right')
        self.open_file_btn = ttk.Button(bar, text="Open workbook",
                                        command=self._open_file, state='disabled')
        self.open_file_btn.pack(side='right', padx=6)

    # ---- tab 1: run ----------------------------------------------------------
    def _build_run_tab(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text='  Run  ')

        files_frame = ttk.LabelFrame(tab, text="Atlan impact reports (.csv / .xlsx)",
                                     padding=8)
        files_frame.pack(fill='x')
        list_row = ttk.Frame(files_frame)
        list_row.pack(fill='x')
        self.files_list = tk.Listbox(list_row, height=5, selectmode='extended',
                                     activestyle='dotbox')
        self.files_list.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(list_row, command=self.files_list.yview)
        scroll.pack(side='left', fill='y')
        self.files_list.config(yscrollcommand=scroll.set)
        btns = ttk.Frame(list_row)
        btns.pack(side='left', fill='y', padx=(8, 0))
        ttk.Button(btns, text="Add files…", command=self._add_files).pack(fill='x')
        ttk.Button(btns, text="Remove", command=self._remove_files).pack(fill='x', pady=4)
        ttk.Button(btns, text="Clear", command=lambda: self.files_list.delete(0, 'end')
                   ).pack(fill='x')

        out_frame = ttk.Frame(tab)
        out_frame.pack(fill='x', pady=(10, 0))
        ttk.Label(out_frame, text="Output workbook:").pack(side='left')
        self.output_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.output_var).pack(
            side='left', fill='x', expand=True, padx=8)
        ttk.Button(out_frame, text="Browse…", command=self._pick_output).pack(side='left')
        ttk.Label(tab, text="Leave blank to save next to the first input file.",
                  style='Sub.TLabel').pack(anchor='w')

        opts = ttk.LabelFrame(tab, text="Options", padding=8)
        opts.pack(fill='x', pady=(10, 0))
        self.opt_detailed = tk.BooleanVar(value=True)
        self.opt_separate = tk.BooleanVar(value=False)
        self.opt_rules_detailed = tk.BooleanVar(value=False)
        self.opt_unformatted = tk.BooleanVar(value=False)
        self.opt_drop_no_lineage = tk.BooleanVar(value=True)
        for col, (text, var) in enumerate([
                ("Include Detailed sheets", self.opt_detailed),
                ("Detailed sheets in a separate file", self.opt_separate),
                ("Apply pruning rules to Detailed", self.opt_rules_detailed),
                ("Unformatted raw export", self.opt_unformatted),
                ("Ignore queries with no upstream lineage", self.opt_drop_no_lineage)]):
            ttk.Checkbutton(opts, text=text, variable=var).grid(
                row=col // 2, column=col % 2, sticky='w', padx=6, pady=2)

        self.run_btn = ttk.Button(tab, text="▶  Build Lineage Workbook",
                                  style='Run.TButton', command=self._run)
        self.run_btn.pack(pady=12)

        log_frame = ttk.LabelFrame(tab, text="Log", padding=4)
        log_frame.pack(fill='both', expand=True)
        self.log_text = tk.Text(log_frame, height=10, state='disabled',
                                font=('Consolas', 9), background='#FAF7FC')
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        log_scroll.pack(side='left', fill='y')
        self.log_text.config(yscrollcommand=log_scroll.set)

    # ---- tab 2: pruning rules -------------------------------------------------
    def _build_rules_tab(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text='  Pruning Rules  ')
        ttk.Label(tab, text="These rules stop 'trivial' objects from being expanded "
                            "in the Summary sheets (one entry per line).",
                  style='Sub.TLabel').pack(anchor='w', pady=(0, 8))

        ns_frame = ttk.LabelFrame(tab, padding=8)
        ns_frame.pack(fill='both', expand=True)
        self.rule_ns_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(ns_frame, variable=self.rule_ns_enabled,
                        text="Block namespaces (database, database.schema, "
                             "or database.schema.object)").pack(anchor='w')
        self.ns_text = tk.Text(ns_frame, height=5, font=('Consolas', 9))
        self.ns_text.pack(fill='both', expand=True, pady=(4, 0))

        grid = ttk.Frame(tab)
        grid.pack(fill='both', expand=True, pady=(10, 0))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        tbl_frame = ttk.LabelFrame(grid, padding=8)
        tbl_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        self.rule_tbl_enabled = tk.BooleanVar(value=True)
        row = ttk.Frame(tbl_frame)
        row.pack(anchor='w', fill='x')
        ttk.Checkbutton(row, variable=self.rule_tbl_enabled,
                        text="Stop expanding these tables at level ≥").pack(side='left')
        self.rule_tbl_limit = tk.IntVar(value=8)
        ttk.Spinbox(row, from_=1, to=99, width=4,
                    textvariable=self.rule_tbl_limit).pack(side='left', padx=4)
        ttk.Label(tbl_frame, text='("ALL" = every table)', style='Sub.TLabel').pack(anchor='w')
        self.tbl_text = tk.Text(tbl_frame, height=6, font=('Consolas', 9))
        self.tbl_text.pack(fill='both', expand=True, pady=(4, 0))

        vw_frame = ttk.LabelFrame(grid, padding=8)
        vw_frame.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        self.rule_vw_enabled = tk.BooleanVar(value=True)
        row = ttk.Frame(vw_frame)
        row.pack(anchor='w', fill='x')
        ttk.Checkbutton(row, variable=self.rule_vw_enabled,
                        text="Stop expanding these views at level ≥").pack(side='left')
        self.rule_vw_limit = tk.IntVar(value=8)
        ttk.Spinbox(row, from_=1, to=99, width=4,
                    textvariable=self.rule_vw_limit).pack(side='left', padx=4)
        ttk.Label(vw_frame, text='("ALL" = every view)', style='Sub.TLabel').pack(anchor='w')
        self.vw_text = tk.Text(vw_frame, height=6, font=('Consolas', 9))
        self.vw_text.pack(fill='both', expand=True, pady=(4, 0))

    # ---- tab 3: formatting & sources -----------------------------------------
    def _build_format_tab(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text='  Formatting & Sources  ')

        fmt = ttk.LabelFrame(tab, text="Workbook appearance", padding=8)
        fmt.pack(fill='x')
        self.fmt_links = tk.BooleanVar(value=True)
        self.fmt_merge = tk.BooleanVar(value=True)
        self.fmt_pad = tk.BooleanVar(value=True)
        self.fmt_freeze = tk.BooleanVar(value=False)
        for col, (text, var) in enumerate([
                ("Hyperlink duplicates to their expansion", self.fmt_links),
                ("Merge Query cells over their block", self.fmt_merge),
                ("Pad empty cells like the client file", self.fmt_pad),
                ("Freeze header rows & Query column", self.fmt_freeze)]):
            ttk.Checkbutton(fmt, text=text, variable=var).grid(
                row=col // 2, column=col % 2, sticky='w', padx=6, pady=2)
        colors = ttk.Frame(fmt)
        colors.grid(row=2, column=0, columnspan=2, sticky='w', pady=(6, 0))
        ttk.Label(colors, text="Header fill:").pack(side='left')
        self.fmt_header_fill = tk.StringVar(value=engine.EXCEL_FORMAT['header_fill'])
        ttk.Entry(colors, textvariable=self.fmt_header_fill, width=8).pack(side='left', padx=4)
        ttk.Label(colors, text="Sub-header fill:").pack(side='left', padx=(10, 0))
        self.fmt_sub_fill = tk.StringVar(value=engine.EXCEL_FORMAT['subheader_fill'])
        ttk.Entry(colors, textvariable=self.fmt_sub_fill, width=8).pack(side='left', padx=4)
        ttk.Label(colors, text="Font:").pack(side='left', padx=(10, 0))
        self.fmt_font = tk.StringVar(value=engine.EXCEL_FORMAT['font_name'])
        ttk.Entry(colors, textvariable=self.fmt_font, width=16).pack(side='left', padx=4)

        src = ttk.LabelFrame(
            tab, text='"List of sources" labels — prefix = label, one per line '
                      '(e.g. PROD_DATALAKE.CRM_MSCRM = CRM)', padding=8)
        src.pack(fill='both', expand=True, pady=(10, 0))
        self.src_text = tk.Text(src, height=8, font=('Consolas', 9))
        self.src_text.pack(fill='both', expand=True)
        fb_row = ttk.Frame(src)
        fb_row.pack(anchor='w', pady=(6, 0))
        ttk.Label(fb_row, text="When no label matches, use:").pack(side='left')
        self.src_fallback = tk.StringVar(value=engine.SOURCE_LABEL_FALLBACK)
        ttk.Combobox(fb_row, textvariable=self.src_fallback, width=10,
                     values=('schema', 'db', 'blank'), state='readonly'
                     ).pack(side='left', padx=6)

        ttk.Button(tab, text="Save settings", command=self._save_settings).pack(
            anchor='e', pady=(8, 0))

    # ------------------------------------------------------------- actions --
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Atlan impact reports",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")])
        existing = set(self.files_list.get(0, 'end'))
        for p in paths:
            if p not in existing:
                self.files_list.insert('end', p)

    def _remove_files(self):
        for idx in reversed(self.files_list.curselection()):
            self.files_list.delete(idx)

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Output workbook", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.output_var.set(path)

    def _collect_rules(self) -> dict:
        def lines(widget):
            return [ln.strip() for ln in widget.get('1.0', 'end').splitlines() if ln.strip()]
        return {
            'enable_namespace_blocking': self.rule_ns_enabled.get(),
            'blocked_namespaces': lines(self.ns_text),
            'enable_high_level_tables': self.rule_tbl_enabled.get(),
            'high_level_tables_limit': self.rule_tbl_limit.get(),
            'high_level_tables': lines(self.tbl_text),
            'enable_high_level_views': self.rule_vw_enabled.get(),
            'high_level_views_limit': self.rule_vw_limit.get(),
            'high_level_views': lines(self.vw_text),
            'apply_rules_to_detailed': self.opt_rules_detailed.get(),
        }

    def _collect_format(self) -> dict:
        fmt = dict(engine.EXCEL_FORMAT)
        fmt.update({
            'hyperlink_duplicates': self.fmt_links.get(),
            'merge_query_blocks': self.fmt_merge.get(),
            'pad_empty_cells': self.fmt_pad.get(),
            'freeze_panes': self.fmt_freeze.get(),
            'header_fill': self.fmt_header_fill.get().strip() or '7030A0',
            'subheader_fill': self.fmt_sub_fill.get().strip() or 'F2CEEF',
            'font_name': self.fmt_font.get().strip() or 'Messina Sans',
        })
        return fmt

    def _collect_source_labels(self) -> dict:
        labels = {}
        for line in self.src_text.get('1.0', 'end').splitlines():
            if '=' in line:
                prefix, _, label = line.partition('=')
                if prefix.strip():
                    labels[prefix.strip()] = label.strip()
        return labels

    def _run(self):
        files = list(self.files_list.get(0, 'end'))
        if not files:
            messagebox.showwarning("No input", "Add at least one Atlan impact report first.")
            return
        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            messagebox.showerror("File not found", "\n".join(missing))
            return
        self._save_settings(silent=True)
        self.run_btn.config(state='disabled')
        self.open_file_btn.config(state='disabled')
        self.open_folder_btn.config(state='disabled')
        self.progress.start(12)
        self.status_var.set("Building…")
        self._log_clear()

        kwargs = dict(
            input_files=files,
            output_path=self.output_var.get().strip() or None,
            rules=self._collect_rules(),
            fmt=self._collect_format(),
            source_labels=self._collect_source_labels(),
            source_fallback=self.src_fallback.get(),
            include_detailed=self.opt_detailed.get(),
            separate_detailed=self.opt_separate.get(),
            unformatted=self.opt_unformatted.get(),
            drop_no_lineage=self.opt_drop_no_lineage.get(),
        )
        self.worker = threading.Thread(target=self._worker_main, args=(kwargs,), daemon=True)
        self.worker.start()

    def _worker_main(self, kwargs):
        handler = QueueLogHandler(self.log_queue)
        handler.setFormatter(logging.Formatter('%(message)s'))
        engine.log.addHandler(handler)
        engine.log.setLevel(logging.INFO)
        try:
            written = engine.build_workbooks(**kwargs)
            self.written_files = written
            self.log_queue.put(('DONE', written))
        except Exception as exc:
            self.log_queue.put(('ERROR', str(exc)))
        finally:
            engine.log.removeHandler(handler)

    # ------------------------------------------------------------- logging --
    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple):
                    kind, payload = item
                    self.progress.stop()
                    self.run_btn.config(state='normal')
                    if kind == 'DONE':
                        self.status_var.set("Done — " + os.path.basename(payload[0]))
                        self._log_line("\nFinished. Output:")
                        for path in payload:
                            self._log_line(f"  {path}")
                        self.open_file_btn.config(state='normal')
                        self.open_folder_btn.config(state='normal')
                    else:
                        self.status_var.set("Failed")
                        self._log_line(f"\nERROR: {payload}")
                        messagebox.showerror("Build failed", payload)
                else:
                    self._log_line(item)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_log_queue)

    def _log_line(self, text: str):
        self.log_text.config(state='normal')
        self.log_text.insert('end', text + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')

    def _log_clear(self):
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.config(state='disabled')

    def _open_file(self):
        if self.written_files:
            os.startfile(self.written_files[0])

    def _open_folder(self):
        if self.written_files:
            subprocess.Popen(['explorer', '/select,', self.written_files[0]])

    # ------------------------------------------------------------ settings --
    def _save_settings(self, silent: bool = False):
        cfg = {
            'trivial_rules': self._collect_rules(),
            'excel_format': self._collect_format(),
            'source_labels': self._collect_source_labels(),
            'source_label_fallback': self.src_fallback.get(),
            'ui': {
                'files': list(self.files_list.get(0, 'end')),
                'output': self.output_var.get(),
                'include_detailed': self.opt_detailed.get(),
                'separate_detailed': self.opt_separate.get(),
                'unformatted': self.opt_unformatted.get(),
                'drop_no_lineage': self.opt_drop_no_lineage.get(),
            },
        }
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as fh:
                json.dump(cfg, fh, indent=2)
            if not silent:
                self.status_var.set(f"Settings saved to {os.path.basename(SETTINGS_FILE)}")
        except OSError as exc:
            if not silent:
                messagebox.showerror("Could not save settings", str(exc))

    def _load_settings(self):
        rules = dict(engine.TRIVIAL_RULES)
        labels = dict(engine.SOURCE_LABELS)
        ui = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, encoding='utf-8') as fh:
                    cfg = json.load(fh)
                rules.update(cfg.get('trivial_rules', {}))
                labels = cfg.get('source_labels', labels)
                fmt = cfg.get('excel_format', {})
                self.fmt_links.set(fmt.get('hyperlink_duplicates', True))
                self.fmt_merge.set(fmt.get('merge_query_blocks', True))
                self.fmt_pad.set(fmt.get('pad_empty_cells', True))
                self.fmt_freeze.set(fmt.get('freeze_panes', False))
                self.fmt_header_fill.set(fmt.get('header_fill', '7030A0'))
                self.fmt_sub_fill.set(fmt.get('subheader_fill', 'F2CEEF'))
                self.fmt_font.set(fmt.get('font_name', 'Messina Sans'))
                self.src_fallback.set(cfg.get('source_label_fallback', 'schema'))
                ui = cfg.get('ui', {})
            except (OSError, json.JSONDecodeError):
                pass

        self.rule_ns_enabled.set(rules.get('enable_namespace_blocking', True))
        self.rule_tbl_enabled.set(rules.get('enable_high_level_tables', True))
        self.rule_vw_enabled.set(rules.get('enable_high_level_views', True))
        self.rule_tbl_limit.set(rules.get('high_level_tables_limit', 8))
        self.rule_vw_limit.set(rules.get('high_level_views_limit', 8))
        self.ns_text.insert('1.0', '\n'.join(rules.get('blocked_namespaces', [])))
        self.tbl_text.insert('1.0', '\n'.join(rules.get('high_level_tables', [])))
        self.vw_text.insert('1.0', '\n'.join(rules.get('high_level_views', [])))
        self.src_text.insert('1.0', '\n'.join(f'{k} = {v}' for k, v in labels.items()))

        for path in ui.get('files', []):
            if os.path.exists(path):
                self.files_list.insert('end', path)
        self.output_var.set(ui.get('output', ''))
        self.opt_detailed.set(ui.get('include_detailed', True))
        self.opt_separate.set(ui.get('separate_detailed', False))
        self.opt_rules_detailed.set(rules.get('apply_rules_to_detailed', False))
        self.opt_unformatted.set(ui.get('unformatted', False))
        self.opt_drop_no_lineage.set(ui.get('drop_no_lineage', True))


def run():
    root = tk.Tk()
    LineageApp(root)
    root.mainloop()


if __name__ == '__main__':
    run()
