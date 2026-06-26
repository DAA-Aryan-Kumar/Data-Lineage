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

# When frozen by PyInstaller, __file__ lives in a temp extraction dir, so keep
# the settings next to the executable instead (a stable, writable location).
_APP_DIR = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
            else os.path.dirname(os.path.abspath(__file__)))
SETTINGS_FILE = os.path.join(_APP_DIR, 'lineage_settings.json')

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
        self._files: list[str] = []   # input report paths, in processing order

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

        self._add_watermark()

    def _watermark_text(self) -> str:
        """Optional corner watermark, OFF by default. Enabled (e.g. "AK") via,
        in order of precedence: the LINEAGE_WATERMARK env var, a 'watermark.txt'
        bundled into a packaged build, or a 'watermark' key in the local
        settings file. Lets one codebase produce a plain or watermarked build."""
        txt = os.environ.get('LINEAGE_WATERMARK', '').strip()
        if txt:
            return txt
        if getattr(sys, 'frozen', False):
            bundled = os.path.join(getattr(sys, '_MEIPASS', ''), 'watermark.txt')
            try:
                if os.path.exists(bundled):
                    with open(bundled, encoding='utf-8') as fh:
                        return fh.read().strip()
            except OSError:
                pass
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, encoding='utf-8') as fh:
                    return str(json.load(fh).get('watermark', '')).strip()
        except (OSError, json.JSONDecodeError):
            pass
        return ''

    def _add_watermark(self):
        """Draw a small, muted maker's mark in the top-right corner if one is
        configured (see _watermark_text). No-op otherwise."""
        text = self._watermark_text()
        if not text:
            return
        mark = tk.Label(self.root, text=text, font=('Segoe UI', 8),
                        foreground='#C2C2C2', background=self.root.cget('background'))
        mark.place(relx=1.0, rely=0.0, x=-10, y=6, anchor='ne')
        mark.lift()

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
        ttk.Button(btns, text="Clear", command=self._clear_files).pack(fill='x')

        out_frame = ttk.Frame(tab)
        out_frame.pack(fill='x', pady=(10, 0))
        ttk.Label(out_frame, text="Output workbook:").pack(side='left')
        self.output_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.output_var).pack(
            side='left', fill='x', expand=True, padx=8)
        ttk.Button(out_frame, text="Browse…", command=self._pick_output).pack(side='left')
        ttk.Label(tab, text="Leave blank to save as \"Data Lineage Workbook.xlsx\" "
                            "next to the first input file.",
                  style='Sub.TLabel').pack(anchor='w')

        opts = ttk.LabelFrame(tab, text="Options", padding=8)
        opts.pack(fill='x', pady=(10, 0))
        self.opt_detailed = tk.BooleanVar(value=True)
        self.opt_separate = tk.BooleanVar(value=False)
        self.opt_drop_no_lineage = tk.BooleanVar(value=True)
        for col, (text, var) in enumerate([
                ("Include Detailed sheets", self.opt_detailed),
                ("Detailed sheets in a separate file", self.opt_separate),
                ("Ignore queries with no upstream lineage", self.opt_drop_no_lineage)]):
            ttk.Checkbutton(opts, text=text, variable=var).grid(
                row=col // 2, column=col % 2, sticky='w', padx=6, pady=2)

        actions = ttk.Frame(tab)
        actions.pack(fill='x', pady=12)
        self.run_btn = ttk.Button(actions, text="▶  Build Lineage Workbook",
                                  style='Run.TButton', command=self._run)
        self.run_btn.pack(side='left')
        ttk.Button(actions, text="Save settings",
                   command=self._save_settings).pack(side='right')

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
        ttk.Label(
            tab,
            text="Trim noise from the lineage. Each box takes one glob pattern "
                 "per line, matched against DATABASE.SCHEMA.OBJECT "
                 "(e.g.  *.DATAWAREHOUSE.*,  PROD_*.*.BRANCH*,  *.CRM*.* ). "
                 "Entries with fewer than three segments auto-fill with * "
                 "(TEMP_DB = TEMP_DB.*.*,  DB.SCHEMA = DB.SCHEMA.*).",
            style='Sub.TLabel', wraplength=900, justify='left').pack(anchor='w', pady=(0, 8))

        grid = ttk.Frame(tab)
        grid.pack(fill='both', expand=True)
        for i in (0, 1):
            grid.columnconfigure(i, weight=1, uniform='rules')
            grid.rowconfigure(i, weight=1)

        self.ex_case = tk.BooleanVar(value=False)
        self.blk_case = tk.BooleanVar(value=False)
        self.tbl_case = tk.BooleanVar(value=False)
        self.vw_case = tk.BooleanVar(value=False)

        def make_section(r, c, title, info, case_var, level_var=None):
            f = ttk.LabelFrame(grid, text=title, padding=8)
            f.grid(row=r, column=c, sticky='nsew',
                   padx=(0, 5) if c == 0 else (5, 0),
                   pady=(0, 5) if r == 0 else (5, 0))
            if level_var is None:
                ttk.Label(f, text=info, style='Sub.TLabel',
                          wraplength=380, justify='left').pack(anchor='w')
            else:
                row = ttk.Frame(f)
                row.pack(anchor='w', fill='x')
                ttk.Label(row, text=info, style='Sub.TLabel').pack(side='left')
                ttk.Spinbox(row, from_=0, to=99, width=4,
                            textvariable=level_var).pack(side='left', padx=4)
            txt = tk.Text(f, height=6, font=('Consolas', 9))
            txt.pack(fill='both', expand=True, pady=(4, 2))
            ttk.Checkbutton(f, text="Match case (exact — for quoted identifiers)",
                            variable=case_var).pack(anchor='w')
            return txt

        self.rule_tbl_limit = tk.IntVar(value=8)
        self.rule_vw_limit = tk.IntVar(value=8)
        self.ex_text = make_section(
            0, 0, "Exclude entirely",
            "Hidden from the output and not expanded.", self.ex_case)
        self.blk_text = make_section(
            0, 1, "Block expansion",
            "Shown in the output, but upstream is not expanded.", self.blk_case)
        self.tbl_text = make_section(
            1, 0, "Stop expanding tables",
            "Matching TABLES shown but not expanded beyond level",
            self.tbl_case, self.rule_tbl_limit)
        self.vw_text = make_section(
            1, 1, "Stop expanding views",
            "Matching VIEWS shown but not expanded beyond level",
            self.vw_case, self.rule_vw_limit)

        apply_row = ttk.Frame(tab)
        apply_row.pack(fill='x', pady=(10, 0))
        self.apply_summary = tk.BooleanVar(value=True)
        self.apply_detailed = tk.BooleanVar(value=False)
        ttk.Checkbutton(apply_row, text="Apply pruning rules to Summary sheets",
                        variable=self.apply_summary).pack(side='left', padx=(0, 18))
        ttk.Checkbutton(apply_row, text="Apply pruning rules to Detailed sheets",
                        variable=self.apply_detailed).pack(side='left')

    # ---- tab 3: formatting & sources -----------------------------------------
    def _build_format_tab(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text='  Formatting & Sources  ')
        self._fmt_lockable = []   # widgets disabled while "Unformatted" is on

        self.opt_unformatted = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            tab, text="Unformatted raw export  (a plain wide dump — disables every "
                      "styling option below except the source labels)",
            variable=self.opt_unformatted).pack(anchor='w')

        fmt = ttk.LabelFrame(tab, text="Workbook appearance", padding=8)
        fmt.pack(fill='x', pady=(8, 0))
        self.fmt_links = tk.BooleanVar(value=True)
        self.fmt_merge = tk.BooleanVar(value=True)
        self.fmt_pad = tk.BooleanVar(value=True)
        self.fmt_freeze = tk.BooleanVar(value=False)
        for col, (text, var) in enumerate([
                ("Hyperlink duplicates to their expansion", self.fmt_links),
                ("Merge Query cells over their block", self.fmt_merge),
                ("Pad empty cells like the client file", self.fmt_pad),
                ("Freeze header rows & Query column", self.fmt_freeze)]):
            cb = ttk.Checkbutton(fmt, text=text, variable=var)
            cb.grid(row=col // 2, column=col % 2, sticky='w', padx=6, pady=2)
            self._fmt_lockable.append(cb)
        colors = ttk.Frame(fmt)
        colors.grid(row=2, column=0, columnspan=2, sticky='w', pady=(6, 0))
        ttk.Label(colors, text="Header fill:").pack(side='left')
        self.fmt_header_fill = tk.StringVar(value=engine.EXCEL_FORMAT['header_fill'])
        self.fmt_sub_fill = tk.StringVar(value=engine.EXCEL_FORMAT['subheader_fill'])
        self.fmt_font = tk.StringVar(value=engine.EXCEL_FORMAT['font_name'])
        e1 = ttk.Entry(colors, textvariable=self.fmt_header_fill, width=8)
        e1.pack(side='left', padx=4)
        ttk.Label(colors, text="Sub-header fill:").pack(side='left', padx=(10, 0))
        e2 = ttk.Entry(colors, textvariable=self.fmt_sub_fill, width=8)
        e2.pack(side='left', padx=4)
        ttk.Label(colors, text="Font:").pack(side='left', padx=(10, 0))
        e3 = ttk.Entry(colors, textvariable=self.fmt_font, width=16)
        e3.pack(side='left', padx=4)
        self._fmt_lockable += [e1, e2, e3]

        src = ttk.LabelFrame(
            tab, text='"List of sources" labels — glob = label, one per line, '
                      'first match wins (e.g.  *.CRM_MSCRM.* = CRM)', padding=8)
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
        self.src_case = tk.BooleanVar(value=False)
        ttk.Checkbutton(fb_row, text="Match case", variable=self.src_case).pack(
            side='left', padx=(14, 0))

        self.opt_unformatted.trace_add('write', self._toggle_unformatted)
        self._toggle_unformatted()

    def _toggle_unformatted(self, *_):
        """Grey out the styling controls when Unformatted export is selected
        (source labels stay editable — they apply to the raw dump too)."""
        state = 'disabled' if self.opt_unformatted.get() else 'normal'
        for w in self._fmt_lockable:
            try:
                w.configure(state=state)
            except tk.TclError:
                pass

    # ------------------------------------------------------------- actions --
    def _refresh_files(self):
        """Redraw the listbox with 1-based numbers showing processing order."""
        self.files_list.delete(0, 'end')
        for i, path in enumerate(self._files, 1):
            self.files_list.insert('end', f'{i:>2}.  {path}')

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Atlan impact reports",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")])
        for p in paths:
            if p not in self._files:
                self._files.append(p)
        self._refresh_files()

    def _remove_files(self):
        for idx in sorted(self.files_list.curselection(), reverse=True):
            del self._files[idx]
        self._refresh_files()

    def _clear_files(self):
        self._files.clear()
        self._refresh_files()

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
            'exclude_patterns': lines(self.ex_text),
            'exclude_match_case': self.ex_case.get(),
            'block_patterns': lines(self.blk_text),
            'block_match_case': self.blk_case.get(),
            'table_patterns': lines(self.tbl_text),
            'table_level': self.rule_tbl_limit.get(),
            'table_match_case': self.tbl_case.get(),
            'view_patterns': lines(self.vw_text),
            'view_level': self.rule_vw_limit.get(),
            'view_match_case': self.vw_case.get(),
            'apply_to_summary': self.apply_summary.get(),
            'apply_to_detailed': self.apply_detailed.get(),
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
        files = list(self._files)
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
            source_match_case=self.src_case.get(),
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
            'source_label_match_case': self.src_case.get(),
            'ui': {
                'files': list(self._files),
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
                self.src_case.set(cfg.get('source_label_match_case', False))
                ui = cfg.get('ui', {})
            except (OSError, json.JSONDecodeError):
                pass

        self.rule_tbl_limit.set(rules.get('table_level', 8))
        self.rule_vw_limit.set(rules.get('view_level', 8))
        self.ex_case.set(rules.get('exclude_match_case', False))
        self.blk_case.set(rules.get('block_match_case', False))
        self.tbl_case.set(rules.get('table_match_case', False))
        self.vw_case.set(rules.get('view_match_case', False))
        self.ex_text.insert('1.0', '\n'.join(rules.get('exclude_patterns', [])))
        self.blk_text.insert('1.0', '\n'.join(rules.get('block_patterns', [])))
        self.tbl_text.insert('1.0', '\n'.join(rules.get('table_patterns', [])))
        self.vw_text.insert('1.0', '\n'.join(rules.get('view_patterns', [])))
        self.src_text.insert('1.0', '\n'.join(f'{k} = {v}' for k, v in labels.items()))
        self.apply_summary.set(rules.get('apply_to_summary', True))
        self.apply_detailed.set(rules.get('apply_to_detailed', False))

        self._files = [p for p in ui.get('files', []) if os.path.exists(p)]
        self._refresh_files()
        self.output_var.set(ui.get('output', ''))
        self.opt_detailed.set(ui.get('include_detailed', True))
        self.opt_separate.set(ui.get('separate_detailed', False))
        self.opt_unformatted.set(ui.get('unformatted', False))
        self.opt_drop_no_lineage.set(ui.get('drop_no_lineage', True))


def run():
    root = tk.Tk()
    LineageApp(root)
    root.mainloop()


if __name__ == '__main__':
    run()
